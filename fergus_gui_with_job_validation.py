import xml.etree.ElementTree as ET
import re
import requests
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import defaultdict

# --- CONFIG ---
FERGUS_API_KEY = "fergPAT_998c3012-8f08-42bbdf4f57ac-57a4-4c6d8eae554c-2673-407a-888d-6e82066071ad2410a92a"
FERGUS_API_BASE = "https://api.fergus.com"
DEFAULT_SALES_ACCOUNT_ID = 128381
TOLERANCE = 0.01
# ---------------

ns = {
    'ss': 'urn:schemas-microsoft-com:office:spreadsheet'
}

def parse_currency(value):
    if not value or not re.search(r'\d', value):
        return 0.0
    return float(re.sub(r'[^\d.]', '', value))

def extract_job_number(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    rows = root.findall('.//ss:Worksheet/ss:Table/ss:Row', ns)
    for row in rows:
        cells = row.findall('./ss:Cell', ns)
        cursor = 0
        for cell in cells:
            idx = cell.attrib.get('{%s}Index' % ns['ss'])
            cursor = int(idx) - 1 if idx else cursor
            if cursor == 14:
                value = cell.find('./ss:Data', ns)
                if value is not None and value.text and value.text.strip():
                    return value.text.strip()
            cursor += 1
    return ""

def get_job_details(job_number):
    headers = {
        "Authorization": f"Bearer {FERGUS_API_KEY}"
    }
    response = requests.get(f"{FERGUS_API_BASE}/jobs", headers=headers)
    try:
        data = response.json().get("data", [])
    except Exception as e:
        messagebox.showerror("Error", f"Failed to fetch jobs: {e}")
        return None

    for job in data:
        if str(job.get("jobNo")) == job_number:
            return {
                "id": job.get("id"),
                "jobNo": job.get("jobNo"),
                "description": job.get("description", ""),
                "longDescription": job.get("longDescription", ""),
                "customer": job.get("customer", {}).get("customerFullName", ""),
                "quoteAccepted": job.get("activeQuote", {}).get("isAccepted", False)
            }

    return None

def extract_quote_items(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    rows = root.findall('.//ss:Worksheet/ss:Table/ss:Row', ns)

    sections = defaultdict(list)

    for i, row in enumerate(rows):
        cells = row.findall('./ss:Cell', ns)
        data = {}
        cursor = 0
        for cell in cells:
            idx = cell.attrib.get('{%s}Index' % ns['ss'])
            cursor = int(idx) - 1 if idx else cursor
            value = cell.find('./ss:Data', ns)
            value_text = value.text.strip() if value is not None and value.text else ''
            data[cursor] = value_text
            cursor += 1

        name = data.get(0, '')
        description = data.get(1, '')
        quantity = float(data.get(2, '0'))
        units = data.get(3, '')
        hours = float(data.get(5, '0'))
        cost_each = parse_currency(data.get(9, ''))
        price_each = parse_currency(data.get(10, ''))
        total = parse_currency(data.get(11, ''))
        item_type = data.get(12, '')
        section_name = data.get(15, '') or "General"

        final_description = name if name else description
        is_labour = item_type.lower() == "labor"

        if units.lower() == "hours":
            final_quantity = quantity
        else:
            final_quantity = hours if hours > 0 else quantity

        if not is_labour and price_each > 0:
            corrected_quantity = total / price_each
            if abs(corrected_quantity - final_quantity) > TOLERANCE:
                final_quantity = round(corrected_quantity, 2)

        if final_description and price_each > 0:
            item = {
                "itemName": final_description,
                "itemQuantity": final_quantity,
                "itemPrice": price_each,
                "itemCost": cost_each,
                "discountRate": 0,
                "sortOrder": i
            }

            if is_labour:
                item["isLabour"] = True
            else:
                item["salesAccountId"] = DEFAULT_SALES_ACCOUNT_ID

            sections[section_name].append(item)

    return sections

def build_sections_payload(sections_dict):
    sections = []
    for i, (section_name, items) in enumerate(sections_dict.items()):
        non_labour_names = [
            item["itemName"] for item in items if item["itemName"] and not item.get("isLabour")
        ]
        section = {
            "name": section_name,
            "sectionLineItemMultiplier": 1,
            "parentSectionId": 0,
            "description": "\n".join(non_labour_names),
            "sortOrder": i,
            "lineItems": items,
            "sections": []
        }
        sections.append(section)
    return sections

def push_quote(xml_path, job_number):
    sectioned_items = extract_quote_items(xml_path)
    job_info = get_job_details(job_number)
    if not job_info:
        messagebox.showerror("Error", f"Job number '{job_number}' not found.")
        return

    headers = {
        "Authorization": f"Bearer {FERGUS_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "title": job_info["description"],
        "description": "",
        "dueDays": 180,
        "versionNumber": None,
        "sections": build_sections_payload(sectioned_items)
    }

    url = f"{FERGUS_API_BASE}/jobs/{job_info['id']}/quotes"
    response = requests.post(url, headers=headers, json=payload)

    if response.status_code in (200, 201):
        messagebox.showinfo("Success", "✅ Quote pushed to Fergus successfully!")
    else:
        messagebox.showerror("Error", f"❌ Failed to push quote Status: {response.status_code} {response.text}")

def select_file_and_start():
    file_path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
    if not file_path:
        return
    job_number = extract_job_number(file_path)
    job_entry.delete(0, tk.END)
    job_entry.insert(0, job_number)
    root.selected_file = file_path
    show_job_info(job_number)

def show_job_info(job_number):
    job_info = get_job_details(job_number)
    if not job_info:
        job_info_label.config(text="❌ Job not found.")
    else:
        info_text = (
            "Job No: {job_info['jobNo']}
"
            "Description: {job_info['description']}
"
            "Customer: {job_info['customer']}
"
            "Quote Accepted: {'✅ Yes' if job_info['quoteAccepted'] else '❌ No'}"
        )
        job_info_label.config(text=info_text)

def on_submit():
    file_path = getattr(root, 'selected_file', None)
    if not file_path:
        messagebox.showerror("Error", "No XML file selected.")
        return
    job_number = job_entry.get().strip()
    if not job_number:
        messagebox.showerror("Error", "Job number is empty.")
        return
    push_quote(file_path, job_number)

# GUI
root = tk.Tk()
root.title("Fergus Quote Uploader")
root.geometry("480x320")

select_button = tk.Button(root, text="Select XML File", command=select_file_and_start)
select_button.pack(pady=10)

label = tk.Label(root, text="Job Number:")
label.pack()

job_entry = tk.Entry(root, width=30)
job_entry.pack(pady=5)

job_info_label = tk.Label(root, text="", justify="left")
job_info_label.pack(pady=10)

submit_button = tk.Button(root, text="Push Quote to Fergus", command=on_submit)
submit_button.pack(pady=20)

root.mainloop()
