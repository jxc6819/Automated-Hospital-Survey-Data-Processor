import os
import re
import json
import threading
import tkinter as tk
from tkinter import messagebox
from difflib import SequenceMatcher

import gspread
from oauth2client.service_account import ServiceAccountCredentials


# Google APIs scope for Sheets + Drive access
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

# Put your service account JSON in a local file and point to it with an env var:
#   export GOOGLE_SERVICE_ACCOUNT_JSON="service_account.json"
#
# Do NOT commit the JSON file. Add it to .gitignore.
DEFAULT_SERVICE_ACCOUNT_PATH = "service_account.json"


def load_service_account_dict() -> dict:
    path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", DEFAULT_SERVICE_ACCOUNT_PATH)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Service account JSON not found: {path}\n"
            "Set GOOGLE_SERVICE_ACCOUNT_JSON or place service_account.json next to this script."
        )
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Basic sanity check so failures are obvious
    required = ["type", "client_email", "private_key"]
    missing = [k for k in required if k not in data]
    if missing:
        raise ValueError(f"Service account JSON is missing fields: {missing}")

    return data


def authorize_client() -> gspread.Client:
    key_data = load_service_account_dict()
    creds = ServiceAccountCredentials.from_json_keyfile_dict(key_data, SCOPE)
    return gspread.authorize(creds)


def find_best_column(keyword: str, headers: list[str], fuzzy_cutoff: float = 0.6) -> int:
    # Returns 1-based column index, or -1 if no match
    norm_q = re.sub(r"\s+", " ", keyword.lower().replace("?", "").strip())

    # First pass: simple substring match
    for i, h in enumerate(headers):
        norm_h = re.sub(r"\s+", " ", h.lower().replace("?", "").strip())
        if norm_q in norm_h:
            return i + 1

    # Second pass: fuzzy match
    best_score, best_idx = 0.0, -1
    for i, h in enumerate(headers):
        norm_h = re.sub(r"\s+", " ", h.lower().replace("?", "").strip())
        score = SequenceMatcher(None, norm_q, norm_h).ratio()
        if score > best_score:
            best_score, best_idx = score, i

    return (best_idx + 1) if best_score >= fuzzy_cutoff else -1


# These get set inside main_function()
attendance_sheet = None
master_sheet = None
pre_survey = None
post_survey = None

attendance_values = []
master_values = []
pre_values = []
post_values = []

pre_headers = []
post_headers = []
questions = []


def get_names() -> list[str]:
    names = master_sheet.col_values(1)
    return [n for n in names if n and n != "Identifier"]


def find_pre_column(keyword: str) -> int:
    for i, h in enumerate(pre_headers):
        if keyword in h:
            return i + 1
    return -1


def find_post_column(keyword: str) -> int:
    for i, h in enumerate(post_headers):
        if keyword in h:
            return i + 1
    return -1


def find_pre_row_by_id(identifier: str) -> int:
    col = find_pre_column("Personal Identifier:")
    if col == -1:
        return -1
    column_values = [row[col - 1] if len(row) >= col else "" for row in pre_values]
    for i, value in enumerate(column_values):
        if value.lower() == identifier.lower():
            return i + 1
    return -1


def find_post_row_by_id(identifier: str) -> int:
    col = find_post_column("Personal Identifier:")
    if col == -1:
        return -1
    column_values = [row[col - 1] if len(row) >= col else "" for row in post_values]
    for i, value in enumerate(column_values):
        if value.lower() == identifier.lower():
            return i + 1
    return -1


def get_questions_from_master() -> list[str]:
    header_row = master_sheet.row_values(4)
    out = []
    for header in header_row:
        question = header.split("\n")[0]
        if question != "Identifier" and "# of sessions attended" not in question:
            out.append(question)
    return out


def compile_post_response(row: int, question: str) -> str:
    col = find_post_column(question)
    if col == -1:
        col = find_best_column(question, post_headers, fuzzy_cutoff=0.35)
    if col == -1 or row == -1:
        return "ERROR"
    return post_values[row - 1][col - 1]


def compile_pre_response(row: int, question: str) -> str:
    col = find_pre_column(question)
    if col == -1:
        col = find_best_column(question, pre_headers, fuzzy_cutoff=0.35)
    if col == -1 or row == -1:
        return "ERROR"
    return pre_values[row - 1][col - 1]


def filter_responses(responses: list[str]) -> list[str]:
    # Clears out anything that looks like a timestamp (Form export weirdness)
    cleaned = []
    for r in responses:
        if ":" in r and "/" in r:
            cleaned.append(" ")
        else:
            cleaned.append(r)
    return cleaned


def compile_responses(identifier: str) -> list[str]:
    pre_row = find_pre_row_by_id(identifier)
    post_row = find_post_row_by_id(identifier)

    responses = []
    post_triggered = False

    for q in questions:
        if not post_triggered and questions.index(q) >= 36:
            post_triggered = True

        if post_triggered:
            responses.append(compile_post_response(post_row, q))
        else:
            responses.append(compile_pre_response(pre_row, q))

    return filter_responses(responses)


def response_to_number_helper(response: str, responses: list[str]):
    if response is None:
        return ""

    r = response.lower().strip()

    match r:
        case "pgy1": return 0
        case "pgy2": return 1
        case "pgy3": return 2

        case "20-25": return 0
        case "26-32": return 1
        case "33-40": return 2
        case "41+": return 3

        case "single/non partnered": return 0
        case "married/partnered": return 1

        case "man": return 0
        case "woman": return 1
        case "transgender": return 2
        case "non-binary, gender non-conforming, or genderqueer": return 3
        case "preferred response not listed": return 4

        case "allergy & immunology": return 0
        case "cardiology": return 1
        case "endocrinology": return 2
        case "geriatrics": return 3
        case "gi": return 4
        case "heme/onc": return 5
        case "hospital medicine": return 6
        case "infectious disease": return 7
        case "nephrology": return 8
        case "palliative care": return 9
        case "pulm/crit": return 10
        case "primary care": return 11
        case "rheumatology": return 12
        case "i don't plan to practice": return 13
        case "undecided": return 14

        case "ccu": return 0
        case "ed": return 1
        case "elective": return 2
        case "elmhurst": return 3
        case "micu": return 4
        case "nights": return 5
        case "sinai floors": return 6
        case "senior role": return 7
        case "va floors": return 8
        case "va icu": return 9

        case "no": return 0
        case "yes": return 1

        case "strongly disagree": return 0
        case "disagree": return 1
        case "neutral": return 2
        case "agree": return 3
        case "strongly agree": return 4

        case "not at all": return 0
        case "somewhat true": return 1
        case "moderately true": return 2
        case "very true": return 3
        case "completely true": return 4

        case "very little": return 1
        case "moderately": return 2
        case "a lot": return 3
        case "extremely": return 4

        case "i feel completely burned out": return 0
        case "my symptoms of burnout won't go away. i think about work frustrations a lot.": return 1
        case "i am definitely burning out and have more than one symptom of burnout, e.g. emotional exhaustion and depersonalization.": return 2
        case "i am very stressed and may be suffering some burnout symptoms, such as emotional exhaustion or depersonalization.": return 3
        case "i am under stress, and don't always have as much energy as i did, but i don't feel burned out.": return 4
        case "i enjoy my work. i have no symptoms of burnout.": return 5

        case "not of interest to me": return 0
        case "too busy with clinical duties": return 1
        case "too busy with admin": return 2
        case "too busy with other stuff": return 3
        case "i like to keep my lunch hour free": return 4
        case "n/a--i have attended all of them": return 5

        case "narrative medicine faculty from columbia university (current facilitators)": return 0
        case "mount sinai faculty with experience/interest in narrative medicine": return 1
        case "mount sinai residents with experience/interest in narrative medicine": return 2

        case "only pgy1's": return 0
        case "pgy1's, pgy2's, and pgy3's": return 1
        case "only pgy1s": return 0
        case "pgy1s, pgy2s, and pgy3s": return 1

        case "close reading and discussion": return 0
        case "writing exercise and discussion": return 1
        case "n/a--have not attended": return 2
        case "n/a - did not attend": return 5

    # Free-text "one word" response: return the text as-is (lowercased)
    if r == "" or r == " ":
        return ""
    if r in [x.lower() for x in responses]:
        idx = [x.lower() for x in responses].index(r)
        if "write one word" in questions[idx].lower():
            return r

    return "ERROR"


def response_to_number(responses: list[str]) -> list:
    return [response_to_number_helper(r, responses) for r in responses]


def count_attendance(identifier: str) -> int:
    ident = identifier.lower()
    return sum(1 for v in attendance_values if v.lower() == ident)


def get_user_numbers(identifier: str) -> list:
    responses = compile_responses(identifier)
    nums = response_to_number(responses)
    return [count_attendance(identifier)] + nums


def identifier_getter() -> list[str]:
    identities = []
    for v in attendance_values:
        lv = v.lower()
        if lv not in identities and "first two letters" not in lv:
            identities.append(lv)
    return identities


def write_identities():
    identities = identifier_getter()
    from gspread.utils import rowcol_to_a1

    start_a1 = rowcol_to_a1(5, 1)
    end_a1 = rowcol_to_a1(len(identities) + 4, 1)
    master_sheet.update(f"{start_a1}:{end_a1}", [[i] for i in identities])


def write_user_numbers(identifier: str, numbers: list):
    first_col_values = [row[0] for row in master_values]
    target_row = first_col_values.index(identifier) + 1

    from gspread.utils import rowcol_to_a1
    start_a1 = rowcol_to_a1(target_row, 2)
    end_a1 = rowcol_to_a1(target_row, 2 + len(numbers) - 1)

    master_sheet.update(f"{start_a1}:{end_a1}", [numbers])


def filter_names(names: list[str]) -> list[str]:
    # Only process rows where column B is still blank
    col_b = [row[1] if len(row) > 1 else "" for row in master_values]
    out = []
    for i, name in enumerate(names):
        # master headers are on row 4, data starts at row 5 => offset by 4
        if i + 4 < len(col_b) and not col_b[i + 4]:
            out.append(name)
    return out


def main_function(att_url: str, master_url: str, pre_url: str, post_url: str):
    client = authorize_client()

    global attendance_sheet, master_sheet, pre_survey, post_survey
    attendance_sheet = client.open_by_url(att_url).sheet1
    master_sheet = client.open_by_url(master_url).sheet1
    pre_survey = client.open_by_url(pre_url).sheet1
    post_survey = client.open_by_url(post_url).sheet1

    global attendance_values, master_values, pre_values, post_values
    attendance_values = attendance_sheet.col_values(2)
    master_values = master_sheet.get_all_values()
    pre_values = pre_survey.get_all_values()
    post_values = post_survey.get_all_values()

    global pre_headers, post_headers, questions
    pre_headers = pre_survey.row_values(1)
    post_headers = post_survey.row_values(1)
    questions = get_questions_from_master()

    # If the master sheet doesn't have IDs filled in yet, seed from attendance
    if len(master_sheet.cell(5, 1).value) < 2:
        write_identities()
        # refresh after writing
        master_values[:] = master_sheet.get_all_values()

    names = filter_names(get_names())
    for name in names:
        nums = get_user_numbers(name)
        write_user_numbers(name, nums)


def launch_gui():
    root = tk.Tk()
    root.title("Survey Data Processor")

    labels = [
        "Attendance Sheet URL",
        "Master Sheet URL",
        "Pre-Survey URL",
        "Post-Survey URL",
    ]

    vars_ = []
    for i, text in enumerate(labels):
        tk.Label(root, text=text).grid(row=i, column=0, sticky="e", padx=6, pady=6)
        var = tk.StringVar()
        vars_.append(var)
        tk.Entry(root, textvariable=var, width=80).grid(row=i, column=1, padx=6, pady=6)

    attendance_var, master_var, pre_var, post_var = vars_

    run_btn = tk.Button(root, text="Run")
    run_btn.grid(row=len(labels), columnspan=2, pady=10)

    def run_clicked():
        run_btn.config(state=tk.DISABLED)

        def task():
            try:
                main_function(
                    attendance_var.get().strip(),
                    master_var.get().strip(),
                    pre_var.get().strip(),
                    post_var.get().strip(),
                )
                messagebox.showinfo("Success", "Done.")
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                run_btn.config(state=tk.NORMAL)

        threading.Thread(target=task, daemon=True).start()

    run_btn.config(command=run_clicked)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()
