import streamlit as st
import pandas as pd
import os
import shutil
import math
import zipfile
import io



def clear_and_make(path):
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path)



def get_department(roll):
    roll = str(roll)
    dept = ""
    for ch in roll:
        if ch.isalpha():
            dept += ch
    return  dept



def stats_from_folder(folder):
    files = [f for f in os.listdir(folder) if f.endswith(".csv")]
    files.sort(key=lambda x: int(''.join([c for c in x if c.isdigit()]) or 9999))
    data = {}
    for f in files:
        try:
            df = pd.read_csv(os.path.join(folder, f))
        except:
            continue
        if  "Roll" not in df.columns:
            continue
        df["Department"] = df["Roll"].apply(get_department)
        counts = df["Department"].value_counts().to_dict()
        counts["Total"] = len(df)
        data[f.replace(".csv", "").upper()] = counts
    if data:
        df_stats = pd.DataFrame(data).T.fillna(0).astype(int)
        if "Total" in df_stats.columns:
            cols = [c for c in df_stats.columns if c != "Total"] + ["Total"]
            df_stats = df_stats[cols]
        return df_stats
    else:
        return pd.DataFrame()



def make_branchwise_groups(df, k):
    groups = [[] for _ in range(k)]
    dept_groups = {d: list(rows.to_dict("records")) for d, rows in df.groupby("Department")}
    total = len(df)
    size = math.ceil(total / k)
    g = 0
    while any(len(v) > 0 for v in dept_groups.values()):
        for d in list(dept_groups.keys()):
            if dept_groups[d]:
                student = dept_groups[d].pop(0)
                groups[g].append(student)
                if len(groups[g]) >= size:
                    g =(g +1) % k
    return groups



def  make_uniform_groups(dept_dfs, k ):
    total =sum(len(v) for v in dept_dfs.values())
    size = math.ceil(total / k)
    groups = [[] for _ in range(k)]
    group_sizes =[0] * k
    g =0
    remain = {d: df.copy() for d, df in dept_dfs.items()}

    while any(len(x) > 0 for x in remain.values()):
        
        dept = max(remain, key=lambda d: len(remain[d]))
        if len(remain[dept]) == 0:
            continue
        take =min(size - group_sizes[g], len(remain[dept]))
        part = remain[dept].iloc[:take].to_dict("records")
        groups[g].extend(part)
        group_sizes[g] += take
        remain[dept] = remain[dept].iloc[take:].reset_index(drop=True)
        if group_sizes[g] >= size:
            g =(g + 1) % k
    return groups



def zip_all(folders, extra_file=None):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for key, folder in folders.items():
            for root, _, files in os.walk(folder):
                for f in files:
                    full = os.path.join(root, f)
                    arc = os.path.join(key, os.path.relpath(full, folder))
                    z.write(full, arc)
        if extra_file and os.path.exists(extra_file):
            z.write(extra_file, os.path.basename(extra_file))
    buffer.seek(0)
    return buffer

  
st.title(" Student Group Maker (Simplified Version)")

file = st.file_uploader("Upload Excel file", type=["xlsx"])
k1 = st.number_input("Groups (Branchwise Mix)", min_value=1, value=3)
k2 = st.number_input("Groups (Uniform Mix)", min_value=1, value=3)

if st.button("Run Grouping"):
    if file is None:
        st.error("Please upload file")
    else:
        df = pd.read_excel(file)
        df = df.iloc[:, :3] 
        df.columns = ["Roll", "Name", "Email"]
        df["Department"] = df["Roll"].apply(get_department)

        
        folder_full = "departments"
        clear_and_make(folder_full)
        for d, part in df.groupby("Department"):
            part[["Roll", "Name", "Email"]].to_csv(os.path.join(folder_full, f"{d}.csv"), index=False)

        
        dept_dict = {}
        for f in os.listdir(folder_full):
            if f.endswith(".csv"):
                d = f.replace(".csv", "")
                dept_dict[d] = pd.read_csv(os.path.join(folder_full, f))

        
        branch_groups = make_branchwise_groups(df[["Roll","Name","Email","Department"]], int(k1))
        folder_branch = "branchwise_groups"
        clear_and_make(folder_branch)
        for i, g in enumerate(branch_groups, 1):
            pd.DataFrame(g)[["Roll","Name","Email"]].to_csv(os.path.join(folder_branch, f"g{i}.csv"), index=False)

        
        uniform_groups = make_uniform_groups(dept_dict, int(k2))
        folder_uniform = "uniform_groups"
        clear_and_make(folder_uniform)
        for i, g in enumerate(uniform_groups, 1):
            pd.DataFrame(g)[["Roll","Name","Email"]].to_csv(os.path.join(folder_uniform, f"g{i}.csv"), index=False)

        
        stats_branch = stats_from_folder(folder_branch)
        stats_uniform = stats_from_folder(folder_uniform)
        excel_out = "output.xlsx"
        with pd.ExcelWriter(excel_out, engine="openpyxl") as w:
            r = 0
            if not stats_branch.empty:
                pd.DataFrame([["Branchwise Mix"]]).to_excel(w, sheet_name="Stats", startrow=r, header=False, index=False)
                r += 2
                stats_branch.to_excel(w, sheet_name="Stats", startrow=r)
                r += len(stats_branch) + 3
            if not stats_uniform.empty:
                pd.DataFrame([["Uniform Mix"]]).to_excel(w, sheet_name="Stats", startrow=r, header=False, index=False)
                r += 2
                stats_uniform.to_excel(w, sheet_name="Stats", startrow=r)

        st.success(" Groups created!")

        
        st.write("###  Branchwise Stats")
        st.dataframe(stats_branch)
        st.write("### Uniform Stats")
        st.dataframe(stats_uniform)

        
        st.write("### Branchwise Groups")
        for f in sorted(os.listdir(folder_branch)):
            with st.expander(f):
                st.dataframe(pd.read_csv(os.path.join(folder_branch, f)).head(20))
        st.write("### Uniform Groups")
        for f in sorted(os.listdir(folder_uniform)):
            with st.expander(f):
                st.dataframe(pd.read_csv(os.path.join(folder_uniform, f)).head(20))
        st.write("### Departments")
        for f in sorted(os.listdir(folder_full)):
            with st.expander(f):
                st.dataframe(pd.read_csv(os.path.join(folder_full, f)).head(20))

        
        with open(excel_out, "rb") as f:
            st.download_button(" Download Excel", f, file_name="output.xlsx")
        allzip = zip_all({"departments": folder_full, "branchwise_groups": folder_branch, "uniform_groups": folder_uniform}, extra_file=excel_out)
        st.download_button(" Download All", allzip, file_name="all_groups.zip", mime="application/zip")

