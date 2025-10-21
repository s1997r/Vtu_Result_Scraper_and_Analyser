'''
from bs4 import BeautifulSoup
import pandas as pd

def parse_student_result(html_path):
    """Parse student result HTML and return DataFrame"""
    try:
        with open(html_path, "r", encoding="utf-8") as f:
            student_html = f.read()

        soup = BeautifulSoup(student_html, "html.parser")

        # Step 1: Extract Student Info
        student_info = {}
        rows = soup.select("table.table-bordered tr")
        for row in rows:
            cells = row.find_all("td")
            if len(cells) == 2:
                key = cells[0].get_text(strip=True).replace(":", "")
                value = cells[1].get_text(strip=True).replace(":", "")
                student_info[key] = value

        # Step 2: Extract Subject Results
        subject_rows = soup.select(".divTableRow")[1:]  # Skip header
        subject_data = {}
        total_internal = 0
        total_full = 0

        for row in subject_rows:
            cells = row.select(".divTableCell")
            if len(cells) >= 7:
                code = cells[0].get_text(strip=True)
                subject_data.update({
                    f"{code}_SubjectName": cells[1].get_text(strip=True),
                    f"{code}_InternalMarks": cells[2].get_text(strip=True),
                    f"{code}_ExternalMarks": cells[3].get_text(strip=True),
                    f"{code}_Total": cells[4].get_text(strip=True),
                    f"{code}_Result": cells[5].get_text(strip=True),
                    f"{code}_UpdatedOn": cells[6].get_text(strip=True),
                })

                try:
                    full = int(cells[4].get_text(strip=True))
                    total_full += full
                except ValueError:
                    pass

        # Step 3: Merge into one flat row
        flat_row = {**student_info, **subject_data}
        flat_row["Total_Full_Marks"] = total_full

        # Step 4: Create DataFrame
        return pd.DataFrame([flat_row])
    except Exception as e:
        print(f"Error parsing student data: {e}")
        return pd.DataFrame()
'''

from bs4 import BeautifulSoup
import pandas as pd


def parse_student_result(html_path):
    """Parse student result HTML and return DataFrame"""
    try:
        with open(html_path, "r", encoding="utf-8") as f:
            student_html = f.read()

        soup = BeautifulSoup(student_html, "html.parser")

        # Step 1: Extract Student Info
        student_info = {}
        rows = soup.select("table.table-bordered tr")
        for row in rows:
            cells = row.find_all("td")
            if len(cells) == 2:
                key = cells[0].get_text(strip=True).replace(":", "")
                value = cells[1].get_text(strip=True).replace(":", "")
                student_info[key] = value

        # Step 2: Extract Subject Results
        subject_rows = soup.select(".divTableRow")[1:]  # Skip header
        subject_data = {}
        total_full = 0

        for row in subject_rows:
            cells = row.select(".divTableCell")
            if len(cells) >= 7:
                code = cells[0].get_text(strip=True)

                # Skip if the code is just a header like "Subject Code"
                if code.lower() == "subject code":
                    continue

                subject_data.update({
                    f"{code}_SubjectName": cells[1].get_text(strip=True),
                    f"{code}_InternalMarks": cells[2].get_text(strip=True),
                    f"{code}_ExternalMarks": cells[3].get_text(strip=True),
                    f"{code}_Total": cells[4].get_text(strip=True),
                    f"{code}_Result": cells[5].get_text(strip=True),
                    f"{code}_UpdatedOn": cells[6].get_text(strip=True),
                })

                try:
                    full = int(cells[4].get_text(strip=True))
                    total_full += full
                except ValueError:
                    pass

        # Step 3: Merge into one flat row
        flat_row = {**student_info, **subject_data}
        flat_row["Total_Full_Marks"] = total_full
        df = pd.DataFrame([flat_row]).fillna("")
        # Step 4: Create DataFrame
        return df
    except Exception as e:
        print(f"Error parsing student data: {e}")
        return pd.DataFrame()


