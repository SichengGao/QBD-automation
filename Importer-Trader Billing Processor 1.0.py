import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from openpyxl import load_workbook
import os
import re
from collections import Counter, defaultdict

CONFIG_FILE = "config.txt"
DEBUG = False  # set True to print diagnostics to console

# ==============================
# NEW UI OPTIONS（以后扩展改这里）
# ==============================
IMPORTER_OPTIONS = [
    "Coastmax",
    "AL Ingots"
]

TRADER_OPTIONS = [
    "GC Aluminum, Inc",
    "Phoenix International Trading Inc."
]

# ==============================
# SAFE ACCOUNTING REFERENCE MAP
# ==============================
raw_reference_map = {
    "material, materials": "50000",
    "international freight, freight costs (ocean), freight cost ocean": "51300",
    "delivery": "55100",
    "fuel surcharge": "56000",
    "overweight": "55800",
    "destination fee, destination terminal handling charges, dest terminal handling charges": "55900",
    "freight insurance": "51000",
    "courier costs (air), courier cost air, courier air": "51200",
    "customs clearance & admin, customs clearance and admin": "51400",
    "isf fee, isf fees": "51500",
    "duties, duty, custom duty 7501, customs 7501, customs": "59240",
    "aes fee": "55700",
    "drayage": "51600",
    "destination drayage, drayage (destination)": "59120",
    "transload, transload and final delivery": "55600",
    "pre pull, pre-pull": "59230",
    "exam, customs exam fee": "59130",
    "detention": "59140",
    "dry run": "59160",
    "storage": "59170",
    "demurrage, destination demurrage": "59180",
    "destination line demurrage": "55300",
    "per diem": "59190",
    "chassis, destination chassis fee": "59150",
    "terminal fee": "59200",
    "pier pass, destination pierpass, destination pier pass, destination pier diem": "59110",
    "handling fees, handling fee": "59210",
    "service fees": "53000",
    "others, others_round up": "59000",
    "others_round up": "59100",
    "exwork, ex-work": "59250",
    "warehouse in/out, warehouse in out": "59260",
    "bond renewal": "51800",
    "commissions paid": "52000",
    "ams": "59220",
    "die, dies, tooling": "51900"
}

# Flatten map
reference_map = {}
for key_string, code in raw_reference_map.items():
    for k in key_string.split(","):
        reference_map[k.strip().lower()] = code

sorted_keywords = sorted(reference_map.keys(), key=len, reverse=True)

# ------------------------------
# Helpers for column detection
# ------------------------------
def header_index_by_names(ws, names):
    headers = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[1]]
    for name in names:
        nl = name.lower()
        for i, h in enumerate(headers):
            if nl in h:
                return i
    return None

def detect_memo_column(ws, container_re, sample_rows=50):
    num_cols = len(ws[1])
    scores = [0] * num_cols
    max_check = min(sample_rows, ws.max_row - 1)
    for r in range(2, 2 + max_check):
        row = list(ws[r])
        for c_idx in range(num_cols):
            val = row[c_idx].value if row[c_idx] is not None else ""
            s = str(val or "").strip().lower()
            if "material" in s or "service fees" in s:
                scores[c_idx] += 2
            elif container_re.search(s):
                scores[c_idx] += 3
    best = max(range(num_cols), key=lambda i: scores[i])
    if scores[best] == 0:
        return None
    return best

def detect_amount_column(ws, sample_rows=50):
    num_cols = len(ws[1])
    scores = [0] * num_cols
    max_check = min(sample_rows, ws.max_row - 1)
    for r in range(2, 2 + max_check):
        row = list(ws[r])
        for c_idx in range(num_cols):
            val = row[c_idx].value if row[c_idx] is not None else None
            if isinstance(val, (int, float)):
                scores[c_idx] += 3
            else:
                s = str(val or "").strip().replace(",", "")
                try:
                    float(s)
                    scores[c_idx] += 2
                except:
                    pass
    best = max(range(num_cols), key=lambda i: scores[i])
    if scores[best] == 0:
        return None
    return best

# ==============================
# CORE UPDATE LOGIC
# ==============================
def update_excel(file_path, selected_importer, selected_trader):

    wb = load_workbook(file_path)
    ws = wb.active

    col_vendor = header_index_by_names(ws, ["vendor"]) or 0
    col_code   = header_index_by_names(ws, ["expense account", "account code"])
    col_memo   = header_index_by_names(ws, ["expense memo", "memo", "description"])
    col_class  = header_index_by_names(ws, ["expense class", "class"])
    col_amt    = header_index_by_names(ws, ["expense amount", "amount", "value"])
    col_desc   = header_index_by_names(ws, ["product/service description", "description"]) or None
    col_addr   = header_index_by_names(ws, ["address line 1", "address"])
    col_terms  = header_index_by_names(ws, ["terms"])
    col_bill   = header_index_by_names(ws, ["bill no"]) or 1
    col_customer = header_index_by_names(ws, ["expense customer", "customer"])

    original_rows = list(ws.iter_rows(min_row=2, values_only=True))
    max_len = len(ws[1])
    new_rows = []

    container_pattern = re.compile(r"[A-Z]{4}[-_ ]?\d{7}", re.IGNORECASE)

    if col_memo is None:
        detected = detect_memo_column(ws, container_pattern)
        col_memo = detected if detected is not None else 6

    if col_amt is None:
        detected_amt = detect_amount_column(ws)
        col_amt = detected_amt if detected_amt is not None else 3

    if col_code is None: col_code = 4
    if col_class is None: col_class = 9
    if col_desc is None: col_desc = 10
    if col_addr is None: col_addr = 11
    if col_terms is None: col_terms = 12

    # ==============================
    # STEP 0 — REMOVE NON 12500
    # ==============================
    filtered_rows = []
    for raw_row in original_rows:
        row = list(raw_row) + [None] * (max_len - len(raw_row))
        acc = str(row[col_code] or "").strip()
        if acc and "12500" not in acc:
            continue
        filtered_rows.append(row)

    # ==============================
    # NEW — TRADER FILTER
    # ==============================
    trader_bills = set()
    for row in filtered_rows:
        customer = str(row[col_customer] or "").strip()
        bill = str(row[col_bill]).strip()
        if customer == selected_trader:
            trader_bills.add(bill)

    trader_filtered = []
    for row in filtered_rows:
        customer = str(row[col_customer] or "").strip()
        bill = str(row[col_bill]).strip()
        if customer == selected_trader:
            trader_filtered.append(row)
        elif customer == "" and bill in trader_bills:
            trader_filtered.append(row)

    filtered_rows = trader_filtered

    # ==============================
    # BUILD BILL LOOKUP
    # ==============================
    bill_lookup = defaultdict(list)
    for row in filtered_rows:
        bill_lookup[str(row[col_bill]).strip()].append(row)

    # ==============================
    # STEP 1 — MATERIAL DUPLICATION
    # ==============================
    for raw_row in filtered_rows:
        row = list(raw_row) + [None] * (max_len - len(raw_row))
        memo = str(row[col_memo] or "").strip()
        memo_lower = memo.lower()
        
        # 【新增】：提前获取 Class 列的值，防止集装箱号被单独写在 Class 列
        class_val = str(row[col_class] or "").strip()

        new_rows.append(row.copy())

        if "material" in memo_lower and "service fees" not in memo_lower:
            # ==============================
            # 新逻辑：合并 Memo 和 Class 判断，并加上严格的前后边界断言
            # 防止 8位数字(如VLCG25110674) 或 包含非标前缀的字符串 被部分匹配
            # ==============================
            combined_text = memo + " " + class_val
            shipment_match = re.search(r"(?<![a-zA-Z])[a-zA-Z]{4}[-_ ]?\d{7}(?!\d)", combined_text)
            service_amount = 750.00 if shipment_match else 100.00

            service_row = row.copy()
            service_row[col_memo] = re.sub(r"material", "Service Fees", memo, flags=re.IGNORECASE)
            service_row[col_code] = "53000"
            service_row[col_amt] = service_amount
            service_row[col_desc] = "Service Charges"
            service_row[col_addr] = selected_importer
            if not service_row[col_terms]:
                service_row[col_terms] = "Due on receipt"

            new_rows.append(service_row)

    ws.delete_rows(2, ws.max_row)

    rows_processed = 0
    matches_found = 0

    # ==============================
    # STEP 2 — CODING / NORMALIZATION
    # ==============================
    for row in new_rows:
        rows_processed += 1

        row[col_vendor] = selected_importer

        memo = str(row[col_memo] or "").strip()
        memo_lower = memo.lower()
        matched = False

        try:
            amt_val = float(str(row[col_amt] or "0").replace(",", ""))
        except:
            amt_val = 0

        if amt_val == 0:
            bill = str(row[col_bill]).strip()
            siblings = bill_lookup.get(bill, [])
            for sib in siblings:
                sib_memo = str(sib[col_memo] or "").lower()
                for keyword in sorted_keywords:
                    if keyword in sib_memo:
                        row[col_code] = reference_map[keyword]
                        matched = True
                        break
                if matched:
                    break

        if not matched:
            if "service fees" in memo_lower:
                row[col_code] = "53000"
                matched = True
            else:
                for keyword in sorted_keywords:
                    if keyword in memo_lower:
                        row[col_code] = reference_map[keyword]
                        matched = True
                        matches_found += 1
                        break

        if not matched:
            row[col_code] = "99000"

        if row[col_class] and row[col_memo]:
            cls = str(row[col_class]).strip()
            orig = str(row[col_memo]).strip()
            if not orig.startswith(cls + "_"):
                row[col_memo] = f"{cls}_{orig}"

        ws.append(row)

    # ==============================
    # SAVE
    # ==============================
    folder, original = os.path.split(file_path)
    name, ext = os.path.splitext(original)
    safe_trader = selected_trader.replace(" ", "_").replace(",", "")
    new_path = os.path.join(folder, f"{name}_ready_for_trader_filtered_by_{safe_trader}{ext}")
    wb.save(new_path)

    messagebox.showinfo(
        "Success",
        f"Processed: {rows_processed}\n"
        f"Matched Codes: {matches_found}\n\n"
        f"Saved to:\n{new_path}"
    )

# ==============================
# GUI
# ==============================
def load_default_path():
    if os.path.isfile(CONFIG_FILE):
        with open(CONFIG_FILE) as f:
            return f.read().strip()
    return ""

def save_default_path():
    p = entry_file_path.get()
    if not os.path.isfile(p):
        messagebox.showwarning("Warning", "Invalid file.")
        return
    with open(CONFIG_FILE, "w") as f:
        f.write(p)
    messagebox.showinfo("Saved", "✅ Default path saved.")

def browse_file():
    p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if p:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, p)

def run_update():
    p = entry_file_path.get()
    if not os.path.isfile(p):
        messagebox.showwarning("Warning", "Select valid file.")
        return
    update_excel(p, importer_var.get(), trader_var.get())

root = tk.Tk()
root.title("Combined Bill Processor")
root.geometry("620x330")
root.resizable(False, False)

tk.Label(root, text="Excel File Path:").pack(pady=(10, 0))
entry_file_path = tk.Entry(root, width=80)
entry_file_path.pack(pady=5)
entry_file_path.insert(0, load_default_path())

tk.Button(root, text="Browse...", command=browse_file).pack()

tk.Button(
    root,
    text="Save as default path",
    command=save_default_path,
    bg="#2196F3",
    fg="white"
).pack(pady=(10, 5))

tk.Label(root, text="Importer:").pack(pady=(10, 0))
importer_var = tk.StringVar(value=IMPORTER_OPTIONS[0])
ttk.Combobox(root, textvariable=importer_var, values=IMPORTER_OPTIONS, state="readonly").pack()

tk.Label(root, text="Trader:").pack(pady=(10, 0))
trader_var = tk.StringVar(value=TRADER_OPTIONS[0])
ttk.Combobox(root, textvariable=trader_var, values=TRADER_OPTIONS, state="readonly").pack(pady=(0, 15))

tk.Button(
    root,
    text="Run Update",
    command=run_update,
    bg="#4CAF50",
    fg="white",
    height=2,
    width=20
).pack(pady=(20, 5))

root.mainloop()
