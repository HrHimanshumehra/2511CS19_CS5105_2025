# Student Group Maker (Simplified Version)

This is a **Streamlit web app** that helps create student groups automatically from an uploaded Excel file.  
It generates two types of groups (**Branchwise Mix** and **Uniform Mix**) and provides group statistics with easy download options.

---

## Features
- Upload an Excel file with student details (`Roll`, `Name`, `Email`).  
- Extracts **department code** from roll numbers automatically.  
- Creates **department-wise CSV files**.  
- Generates:
  - **Branchwise Mix Groups** → Students from the same department spread across groups.  
  - **Uniform Mix Groups** → Balanced groups with students from different departments.  
- Exports:
  - Group statistics (`output.xlsx`)  
  - Department-wise files, branchwise groups, uniform groups  
  - A single ZIP file containing everything  

---

## Input Format
The Excel file must have **at least 3 columns** in this order:

| Roll   | Name         | Email             |
|--------|-------------|-------------------|
| CS101  | Alice Kumar | alice@email.com  |
| ME202  | Ravi Sharma | ravi@email.com   |
| EE305  | Priya Singh | priya@email.com  |

- The **Roll** column should include department code letters (`CS`, `ME`, `EE`, etc).  
- Extra columns (if any) will be ignored.  

---

## How to Run
1. Install dependencies:
   ```bash
   pip install streamlit pandas openpyxl
   ```
2. Save the code as `app.py`.  
3. Run the app:
   ```bash
   streamlit run app.py
   ```
4. Open in your browser → http://localhost:8501  

---

## Outputs
The app generates:
- `departments/` → Department-wise student lists  
- `branchwise_groups/` → Branchwise mixed groups  
- `uniform_groups/` → Uniformly mixed groups  
- `output.xlsx` → Group statistics in Excel format  
- `all_groups.zip` → ZIP file with all the above  

---

## Usage Steps
1. Upload Excel file in the app.  
2. Enter the number of groups for Branchwise & Uniform mix.  
3. Click **Run Grouping**.  
4. Preview groups and statistics inside the app.  
5. Download Excel or ZIP outputs.  

---

## Requirements
- Python 3.8+  
- Libraries:
  - `streamlit`  
  - `pandas`  
  - `openpyxl`  

---

This tool makes it easy for teachers, admins, or coordinators to divide students into groups fairly and efficiently.
