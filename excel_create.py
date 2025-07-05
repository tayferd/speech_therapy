import os
import re
import pandas as pd
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog, messagebox

import os
import re
import pandas as pd
from PyPDF2 import PdfReader

def extract_communication_goals(text, student_name=None):
    import re

    comm_blocks = re.findall(
        r"Domain\(s\)/TSAA\(s\):\s*Communication(.*?)(?=\nDomain\(s\)/TSAA\(s\):|\nAssessments|\nAccommodations|Page \d|\Z)",
        text, re.DOTALL | re.IGNORECASE
    )

    results = []

    for section in comm_blocks:
        goal_blocks = re.split(r"\bGoal:\s*", section, flags=re.IGNORECASE)

        for block in goal_blocks[1:]:
            goal_text = re.split(
                r"(?:\nShort-term Objectives|Assessment Procedures|Progress Reported|\n[A-Z][a-z]+:|\Z)",
                block, maxsplit=1, flags=re.IGNORECASE
            )[0].strip().replace('\n', ' ')

            # FIRST: try to extract formally marked benchmarks
            benchmark_blocks = re.findall(
                r"Short-term Objectives or Benchmarks:(.*?)(?=\n(?:Short-term Objectives or Benchmarks:|Goal:|Domain\(s\)|Assessments|Accommodations|Page \d|\Z))",
                block, re.DOTALL | re.IGNORECASE
            )

            subgoals = []
            for bblock in benchmark_blocks:
                lines = bblock.strip().splitlines()
                for line in lines:
                    line = line.strip()
                    if (
                        len(line.split()) > 5 and
                        " will " in line.lower() and
                        not re.search(r"(student be|participate|assessment|instruction|comment)", line, re.IGNORECASE)
                    ):
                        subgoals.append(line)

            # IF NO FORMAL BENCHMARKS, fallback to find will-statements as subgoals
            if not subgoals:
                lines = block.strip().splitlines()
                for line in lines:
                    line = line.strip()
                    if (
                        len(line.split()) > 5 and
                        " will " in line.lower() and
                        not re.search(r"(student be|participate|assessment|instruction|comment)", line, re.IGNORECASE)
                    ):
                        subgoals.append(line)

            if goal_text:
                results.append({
                    "goal": goal_text,
                    "subgoals": subgoals
                })

    return results

def analyze_pdfs_and_generate_excel(pdf_folder, output_excel_path):
    all_data = []
    max_goals = 0
    max_subgoals_per_goal = {}

    for filename in os.listdir(pdf_folder):
        if filename.lower().endswith(".pdf"):
            path = os.path.join(pdf_folder, filename)
            reader = PdfReader(path)
            text = "\n".join(page.extract_text() or "" for page in reader.pages)

            student_id = re.search(r"\b(\d{10})\b", text)
            student_id = student_id.group(1) if student_id else "NoID"

            name_match = re.search(r"Student:\s+([A-Z ,'-]+)", text)
            name = name_match.group(1).title().replace(",", "").replace("  ", " ") if name_match else "Unknown Student"
            name_parts = name.split()
            first_name = name_parts[0] if name_parts else "Unknown"
            last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else "Student"

            goals = extract_communication_goals(text, name)
            max_goals = max(max_goals, len(goals))

            for i, g in enumerate(goals):
                max_subgoals_per_goal[i] = max(max_subgoals_per_goal.get(i, 0), len(g['subgoals']))

            all_data.append({
                'first_name': first_name,
                'last_name': last_name,
                'id': student_id,
                'goals': goals
            })

    columns = ['First Name', 'Last Name', 'Student ID']
    for g_idx in range(max_goals):
        columns.append(f'Goal {g_idx+1}')
        for s_idx in range(max_subgoals_per_goal.get(g_idx, 0)):
            columns.append(f'Benchmark {g_idx+1}.{s_idx+1}')

    rows = []
    for student in all_data:
        row = {
            'First Name': student['first_name'],
            'Last Name': student['last_name'],
            'Student ID': student['id']
        }
        for g_idx in range(max_goals):
            if g_idx < len(student['goals']):
                goal = student['goals'][g_idx]
                row[f'Goal {g_idx+1}'] = goal['goal']
                for s_idx in range(max_subgoals_per_goal[g_idx]):
                    key = f'Benchmark {g_idx+1}.{s_idx+1}'
                    if s_idx < len(goal['subgoals']):
                        row[key] = goal['subgoals'][s_idx]
                    else:
                        row[key] = ''
            else:
                row[f'Goal {g_idx+1}'] = ''
                for s_idx in range(max_subgoals_per_goal[g_idx]):
                    row[f'Benchmark {g_idx+1}.{s_idx+1}'] = ''
        rows.append(row)

    df = pd.DataFrame(rows, columns=columns)
    df.to_excel(output_excel_path, index=False)
    return output_excel_path

def process_pdfs_to_excel(folder_path):
    rows = []
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            reader = PdfReader(pdf_path)
            text = "\n".join(page.extract_text() or "" for page in reader.pages)

            # Extract student ID
            id_match = re.search(r"\b(\d{10})\b", text)
            student_id = id_match.group(1) if id_match else "NoID"

            # Extract name using improved logic
            name = extract_name(text)
            name_parts = name.split()
            first_name = name_parts[0] if name_parts else "Unknown"
            last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else "Student"

            goals_data = extract_communication_goals(text, name)

            row = {
                "First Name": first_name,
                "Last Name": last_name,
                "Student ID": student_id
            }

            for i in range(2):
                if i < len(goals_data):
                    row[f"Goal {i+1}"] = goals_data[i]['goal']
                    for j in range(3):
                        if j < len(goals_data[i]['subgoals']):
                            row[f"Benchmark {i+1}.{j+1}"] = goals_data[i]['subgoals'][j]
                        else:
                            row[f"Benchmark {i+1}.{j+1}"] = ""
                else:
                    row[f"Goal {i+1}"] = ""
                    for j in range(3):
                        row[f"Benchmark {i+1}.{j+1}"] = ""

            rows.append(row)

    df = pd.DataFrame(rows)
    excel_path = os.path.join(folder_path, "iep_goals_summary.xlsx")
    df.to_excel(excel_path, index=False)
    return excel_path

# GUI Setup
def select_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_var.set(folder)

def run_extraction():
    folder = folder_var.get()
    if folder:
        try:
            output_path = os.path.join(folder, "iep_goals_summary.xlsx")
            analyze_pdfs_and_generate_excel(folder, output_path)
            messagebox.showinfo("Success", f"Excel file created:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
    else:
        messagebox.showwarning("No Folder Selected", "Please select a folder to continue.")

def extract_name(text):
    # Match "Student: LASTNAME, FIRSTNAME [MIDDLENAME]"
    match = re.search(r"Student:\s+([A-Z\s'-]+),\s*([A-Z][a-zA-Z'-]+)", text)
    if match:
        last = match.group(1).strip().title()
        first = match.group(2).strip().title()
        return f"{first} {last}"

    # Fallback to "Firstname will..."
    fallback = re.search(r"\b([A-Z][a-z]+)\s+will\b", text)
    if fallback:
        return f"{fallback.group(1)}"

    return "Unknown Student"


# Create GUI window
root = tk.Tk()
root.title("IEP Goal Extractor")

folder_var = tk.StringVar()

tk.Label(root, text="Selected Folder:").pack()
tk.Entry(root, textvariable=folder_var, width=60).pack(padx=10)
tk.Button(root, text="Select Folder", command=select_folder).pack(pady=5)
tk.Button(root, text="Run Extraction", command=run_extraction).pack(pady=5)

root.mainloop()
