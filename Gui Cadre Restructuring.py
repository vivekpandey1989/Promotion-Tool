import pandas as pd
from datetime import datetime
import math
import tkinter as tk
from tkinter import scrolledtext, messagebox

# -------- FILE PATH --------
file_path = r"D:\Vivek Drive\ED Documents\Personal\Personal_Self\Promotion\Promotion Analysis_Latest.xlsx"

# =========================================================
# QUOTA FUNCTION
# =========================================================
def select_by_quota(pool, vacancies):
    pool = pool.sort_values('Seniority')

    sc = math.floor(vacancies * 0.15)
    st = math.floor(vacancies * 0.075)
    ur = vacancies - sc - st

    sc_sel = pool[pool['Category'] == 'SC'].head(sc)
    st_sel = pool[pool['Category'] == 'ST'].head(st)

    remaining_pool = pool.drop(sc_sel.index).drop(st_sel.index)
    ur_sel = remaining_pool.head(ur)

    selected = pd.concat([sc_sel, st_sel, ur_sel])
    return selected.sort_values('Seniority')

# =========================================================
# MAIN LOGIC
# =========================================================
def run_calculation():
    try:
        # -------- DATE --------
        date_str = date_entry.get().strip()
        try:
            target_date = datetime.strptime(date_str, "%d-%m-%Y")
        except:
            messagebox.showerror("Error", "Enter date in DD-MM-YYYY format")
            return

        # -------- INPUT --------
        total_posts = {
            "ADE": int(ade_entry.get()),
            "JD": int(jd_entry.get()),
            "DD": int(dd_entry.get()),
            "AD": int(ad_entry.get()),
            "EO": int(eo_entry.get())
        }

        rr = {
            "ADE": float(ade_rr.get()) / 100,
            "JD": float(jd_rr.get()) / 100,
            "DD": float(dd_rr.get()) / 100,
            "AD": float(ad_rr.get()) / 100,
            "EO": float(eo_rr.get()) / 100
        }

        dept = {k: math.floor(total_posts[k] * rr[k]) for k in total_posts}

        # -------- LOAD DATA --------
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()

        df.rename(columns={
            'Name (S/Shri)': 'Name',
            'Date of Retirement': 'Retirement Date'
        }, inplace=True)

        df['Post'] = df['Post'].str.strip().replace({'AD-s': 'AD'})
        df['Category'] = df['Category'].fillna('UR').str.strip().str.upper()
        df['Category'] = df['Category'].replace({'GEN': 'UR', 'GENERAL': 'UR'})

        # Remove retirees
        df['Retirement Date'] = pd.to_datetime(df['Retirement Date'], errors='coerce')
        df = df[df['Retirement Date'] > target_date].copy()

        df = df.reset_index(drop=True)
        df['Seniority'] = df.index + 1

        # Vigilance filter
        df['Vigilance'] = df['Vigilance'].fillna('').str.lower()

        eligible = df[df['Vigilance'] != 'yes'].copy().sort_values('Seniority')
        blocked = df[df['Vigilance'] == 'yes'].copy()

        remaining = eligible.copy()
        final = []

        # =========================================================
        # ADE / JD / DD (PURE SENIORITY)
        # =========================================================
        ade = remaining.head(dept["ADE"]).copy()
        ade['New Post'] = 'ADE'
        remaining = remaining.drop(ade.index)

        jd = remaining.head(dept["JD"]).copy()
        jd['New Post'] = 'JD'
        remaining = remaining.drop(jd.index)

        dd = remaining.head(dept["DD"]).copy()
        dd['New Post'] = 'DD'
        remaining = remaining.drop(dd.index)

        final.extend([ade, jd, dd])

        # =========================================================
        # AD (FULL LIST + 8% + RESERVATION)
        # =========================================================
        ad_target = dept["AD"]

        reserve_8 = math.floor(ad_target * 0.08)
        ad_usable = ad_target - reserve_8

        ad_sel = select_by_quota(remaining, ad_usable).copy()
        ad_sel['New Post'] = 'AD'

        remaining = remaining.drop(ad_sel.index)
        final.append(ad_sel)

        # =========================================================
        # EO (FULL LIST + RESERVATION)
        # =========================================================
        eo_target = dept["EO"]

        eo_sel = select_by_quota(remaining, eo_target).copy()
        eo_sel['New Post'] = 'EO'

        remaining = remaining.drop(eo_sel.index)
        final.append(eo_sel)

        # =========================================================
        # REMAINING + BLOCKED
        # =========================================================
        remaining['New Post'] = remaining['Post']
        final.append(remaining)

        blocked['New Post'] = blocked['Post']
        final.append(blocked)

        # =========================================================
        # FINAL ASSEMBLY
        # =========================================================
        final_df = pd.concat(final)

        order = ["ADE", "JD", "DD", "AD", "EO", "AEO"]
        order_map = {p: i for i, p in enumerate(order)}

        final_df['Order'] = final_df['New Post'].map(order_map).fillna(999)
        final_df = final_df.sort_values(['Order', 'Seniority'])

        final_df['Final Seniority'] = range(1, len(final_df) + 1)

        # =========================================================
        # SUMMARY COUNTS
        # =========================================================
        post_counts = final_df['New Post'].value_counts()

        ade_count = post_counts.get('ADE', 0)
        jd_count = post_counts.get('JD', 0)
        dd_count = post_counts.get('DD', 0)
        ad_count = post_counts.get('AD', 0)
        eo_count = post_counts.get('EO', 0)

        # =========================================================
        # OUTPUT DATA
        # =========================================================
        output_df = final_df[
            ['Final Seniority', 'Rank in Post', 'Name', 'Date of Birth',
             'Category', 'Batch and Rank', 'New Post']
        ]

        # =========================================================
        # DISPLAY
        # =========================================================
        output_box.delete(1.0, tk.END)

        summary = f"""
Total Posts After Allocation:

ADE : {ade_count}
JD  : {jd_count}
DD  : {dd_count}
AD  : {ad_count}
EO  : {eo_count}

------------------------------------------------------------
"""
        output_box.insert(tk.END, summary)

        header = "{:<15} {:<10} {:<25} {:<12} {:<10} {:<22} {:<10}\n".format(
            "Final Sen", "Rank", "Name", "DOB", "Cat", "Batch&Rank", "Post"
        )
        output_box.insert(tk.END, header)
        output_box.insert(tk.END, "-" * 120 + "\n")

        for _, row in output_df.iterrows():
            line = "{:<15} {:<10} {:<25} {:<12} {:<10} {:<22} {:<10}\n".format(
                str(row['Final Seniority']),
                str(row['Rank in Post']),
                str(row['Name'])[:25],
                str(row['Date of Birth'])[:10],
                str(row['Category']),
                str(row['Batch and Rank'])[:22],
                str(row['New Post'])
            )
            output_box.insert(tk.END, line)

    except Exception as e:
        messagebox.showerror("Error", str(e))


# =========================================================
# GUI
# =========================================================
root = tk.Tk()
root.title("Promotion Tool")
root.geometry("1100x700")

def add_field(label, row):
    tk.Label(root, text=label).grid(row=row, column=0, sticky='w')
    entry = tk.Entry(root)
    entry.grid(row=row, column=1)
    return entry

date_entry = add_field("Date (DD-MM-YYYY):", 0)

ade_entry = add_field("ADE posts:", 1)
jd_entry = add_field("JD posts:", 2)
dd_entry = add_field("DD posts:", 3)
ad_entry = add_field("AD posts:", 4)
eo_entry = add_field("EO posts:", 5)

ade_rr = add_field("ADE %:", 6)
jd_rr = add_field("JD %:", 7)
dd_rr = add_field("DD %:", 8)
ad_rr = add_field("AD %:", 9)
eo_rr = add_field("EO %:", 10)

tk.Button(root, text="Calculate", command=run_calculation).grid(row=11, column=0, columnspan=2)

output_box = scrolledtext.ScrolledText(root, width=140, height=30)
output_box.grid(row=12, column=0, columnspan=2)

root.mainloop()