# FergusQuoteUploader_improved.py
# Changes added:
# - Centralised currency/qty reconciliation (compute_line_values)
# - Alternating row colours + consistent numeric alignment
# - Click-to-sort for both preview tables
# - Button spacing/ordering polish
# - Per-page width & height memory to prevent jumpiness
# - Safer job number extraction from Job folder name ("6811 - Project" -> "6811")
# - Preflight validation before push (missing prices/qty, negative numbers)
# - Save outgoing payload and preview CSV to disk for debugging

import os
import re
import json
import csv
import requests
import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkfont
import win32com.client
from datetime import datetime
import webbrowser
import tempfile

FERGUS_API_KEY = "fergPAT_998c3012-8f08-42bbdf4f57ac-57a4-4c6d8eae554c-2673-407a-888d-6e82066071ad2410a92a"
FERGUS_API_BASE = "https://api.fergus.com"
# Web app base (change if your region uses a different domain)
FERGUS_WEB_BASE = "https://go.fergus.com"
DEFAULT_SALES_ACCOUNT_ID = 128381
TOLERANCE = 0.01

print("Starting Python PlanSwift plugin…")
ps = win32com.client.Dispatch("PlanSwift9.PlanSwift")
print("✅ Connected to PlanSwift")
if not ps.IsJobOpen():
    print("❌ No job open."); raise SystemExit

def is_number(val):
    try: return float(str(val).replace(",","").strip()) != 0
    except: return False

def safe_get_property(item, name):
    # Fewer COM calls: try GetPropertyResultAsString then PropertyItem
    try:
        val = item.GetPropertyResultAsString(name)
        if val: return val
    except: pass
    try:
        prop = item.PropertyItem(name)
        return str(prop.Value)
    except: pass
    return ""

def get_units(item, property_name):
    try:
        prop = item.GetProperty(property_name)
        return str(prop.Units or prop.InputUnits or "")
    except: return ""

def find_child_by_name(parent, target_name):
    try:
        count = parent.ChildCount()
        for i in range(count):
            child = parent.ChildItem(i)
            if child and target_name.lower() in (child.Name or "").lower():
                return child
    except Exception as e:
        print(f"⚠️ Error searching for '{target_name}': {e}")
    return None

def extract_digits(s):
    # Extract leading digit run to be safe: "6811 - Project" -> "6811"
    if not s: return s
    m = re.search(r"(\d+)", str(s))
    return m.group(1) if m else s

print("🔍 Attempting to access root…")
root = ps.Root(); print("✅ Got root folder")
job = find_child_by_name(root,"Job")
if not job: print("❌ 'Job' folder not found."); raise SystemExit
takeoff = find_child_by_name(job,"Takeoff")
if not takeoff: print("❌ 'Takeoff' folder not found under 'Job'."); raise SystemExit
print(f"✅ Scanning takeoff folder: {takeoff.Name}")
GLOBAL_JOB_NUMBER = extract_digits(safe_get_property(job,"Name"))
print(f"✅ Extracted job number from Job folder name: {GLOBAL_JOB_NUMBER}")

def collect_items_with_estimate_data(item, results, depth=0):
    try:
        item_type = safe_get_property(item,"Type")
        # Exclude generic "section" containers
        if "section" in (item_type or "").lower() and not re.search(r"(area|count)\\b", (item_type or ""), flags=re.I):
            return
        row = {
            "Name": safe_get_property(item,"Name"),
            "Description": safe_get_property(item,"Description"),
            "Group": safe_get_property(item,"Group"),
            "Qty": (safe_get_property(item,"Qty") or "").strip(),
            "Units": get_units(item,"Qty"),
            "Hours": safe_get_property(item,"Hours"),
            "Price Each": safe_get_property(item,"Price Each"),
            "Cost Each": safe_get_property(item,"Cost Each"),
            "Price Total": (safe_get_property(item,"Price Total") or safe_get_property(item,"Total") or safe_get_property(item,"Result")),
            "Takeoff": (safe_get_property(item,"Takeoff") or safe_get_property(item,"Takeoff Name")),
            "Job Number": safe_get_property(item,"Job Number"),
            "Type": item_type
        }
        has_relevant_data = (
            (row["Group"] or "").strip()
            or is_number(row["Price Each"])
            or is_number(row["Cost Each"])
            or is_number(row["Price Total"])
            or is_number(row["Hours"])
        )
        if has_relevant_data:
            results.append(row)
        if item.HasChildren():
            for i in range(item.ChildCount()):
                child = item.ChildItem(i)
                if child: collect_items_with_estimate_data(child, results, depth+1)
    except Exception as e:
        print(f"⚠️ Error reading item or children: {e}")

ALL_ITEMS = []; collect_items_with_estimate_data(takeoff, ALL_ITEMS)
if not ALL_ITEMS: print("❌ No relevant takeoff items found."); raise SystemExit
print(f"✅ Found {len(ALL_ITEMS)} items with estimating data")

def parse_currency(val):
    try: return float(re.sub(r"[^0-9.]", "", str(val))) if re.search(r"[0-9]", str(val)) else 0.0
    except: return 0.0

def compute_line_values(row):
    """
    Returns tuple: (name, qty, cost, price, total, is_labour_bool)
    Applies standard reconciliation:
    - Prefer Name else Description for display
    - qty = Qty else Hours (numeric)
    - Adjust qty if |price*qty - total| > TOLERANCE and price > 0  -> qty = total/price
    """
    name = row.get("Name") or row.get("Description") or ""
    price = parse_currency(row.get("Price Each",""))
    cost  = parse_currency(row.get("Cost Each",""))
    total = parse_currency(row.get("Price Total",""))
    qty_raw = row.get("Qty")
    if qty_raw is None or str(qty_raw).strip() == "":
        qty_raw = row.get("Hours") or "0"
    try:
        qty = float(str(qty_raw).replace(",","").strip() or "0")
    except:
        qty = 0.0
    if price > 0 and abs(price*qty - total) > TOLERANCE and total > 0:
        qty = total / price
    is_labour = (row.get("Type") or "").strip().lower() in ("labor", "labour")
    return name, qty, cost, price, price*qty, is_labour

def group_items(items):
    grouped = {}
    for it in items:
        g = it.get("Group") or "General"
        grouped.setdefault(g, []).append(it)
    return grouped

def validate_items(items):
    problems = []
    for idx, row in enumerate(items, 1):
        name, qty, cost, price, line_total, is_labour = compute_line_values(row)
        if not name:
            problems.append(f"Row {idx}: missing name/description.")
        if price < 0 or cost < 0 or qty < 0:
            problems.append(f"Row {idx}: negative values are not allowed (qty/cost/price).")
        if price == 0 and line_total == 0:
            problems.append(f"Row {idx}: price each and total are zero.")
    return problems

def get_job_details(job_number):
    headers={"Authorization": f"Bearer {FERGUS_API_KEY}"}
    r=requests.get(f"{FERGUS_API_BASE}/jobs?filterSearchText={job_number}", headers=headers)
    if r.status_code!=200: return None
    data=r.json().get("data",[])
    for j in data:
        if str(j.get("jobNo"))==str(job_number):
            return {"id": j.get("id"), "jobNo": j.get("jobNo"),
                    "description": j.get("description",""),
                    "customer": j.get("customer",{}).get("customerFullName",""),
                    "quoteAccepted": j.get("activeQuote",{}).get("isAccepted",False)}
    return None

def get_existing_quotes(job_id):
    headers={"Authorization": f"Bearer {FERGUS_API_KEY}"}
    try:
        r=requests.get(f"{FERGUS_API_BASE}/jobs/{job_id}/quotes", headers=headers); r.raise_for_status()
        return r.json().get("data",[])
    except requests.RequestException: return []

def build_sections_payload(items):
    grouped_items = group_items(items)
    sections = []
    sidx = 0
    for g, rows in grouped_items.items():
        lineitems = []
        desc_lines = []
        for i, row in enumerate(rows):
            name, qty, cost, price, line_total, is_labour = compute_line_values(row)
            if not name or price <= 0:
                continue
            li = {
                "itemName": name,
                "itemQuantity": round(qty, 2),
                "itemPrice": price,
                "itemCost": cost,
                "discountRate": 0,
                "sortOrder": i
            }
            if is_labour:
                li["isLabour"] = True
            else:
                li["salesAccountId"] = DEFAULT_SALES_ACCOUNT_ID
                desc_lines.append(f"- {name}")
            lineitems.append(li)
        sections.append({
            "name": g,
            "description": "\r\n".join(desc_lines),
            "sortOrder": sidx,
            "sectionLineItemMultiplier": 1,
            "parentSectionId": 0,
            "lineItems": lineitems,
            "sections": []
        })
        sidx += 1
    return sections

def export_preview_csv(items, path):
    headers = ["Group","Name","Qty","Units","Hours","Cost Each","Price Each","Line Total"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for row in items:
            name, qty, cost, price, line_total, _ = compute_line_values(row)
            writer.writerow([row.get("Group") or "General", name, f"{qty:.2f}", row.get("Units",""), row.get("Hours",""),
                             f"{cost:.2f}", f"{price:.2f}", f"{line_total:.2f}"])


def push_quote(job_id, title, items, quote_id=None, job_no_for_web=None, parent=None):
    """
    Submit the quote to Fergus.
    - No persistent debug files are written.
    - After a successful create/update, opens the Fergus web app to the Job's Quotes page.
    """
    payload = {"title": title, "description": "", "dueDays": 180, "sections": build_sections_payload(items)}
    headers = {"Authorization": f"Bearer {FERGUS_API_KEY}", "Content-Type": "application/json"}

    try:
        if quote_id:
            url = f"{FERGUS_API_BASE}/jobs/{job_id}/quotes/{quote_id}"
            r = requests.put(url, headers=headers, json=payload)
        else:
            url = f"{FERGUS_API_BASE}/jobs/{job_id}/quotes"
            r = requests.post(url, headers=headers, json=payload)

        r.raise_for_status()
        # Try to extract the quote id from the response if provided
        quote_web_id = None
        try:
            resp = r.json()
            quote_web_id = (resp.get("data") or {}).get("id") or resp.get("id")
        except Exception:
            pass

        popup = tk.Toplevel(parent) if parent else tk.Toplevel()
        popup.title("Success")
        ttk.Label(popup, text="✅ Quote submitted successfully.\nOpening Fergus…").pack(padx=20, pady=20)
        # Center and raise popup (non-blocking)
        popup.update_idletasks()
        w, h = popup.winfo_width(), popup.winfo_height()
        if w <= 1:
            w = popup.winfo_reqwidth()
        if h <= 1:
            h = popup.winfo_reqheight()
        if parent:
            parent.update_idletasks()
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            if pw <= 1 or ph <= 1:
                pw, ph = parent.winfo_reqwidth(), parent.winfo_reqheight()
            x = px + (pw - w)//2
            y = py + (ph - h)//2
            try:
                popup.transient(parent)
            except Exception as transient_error:
                print(f"⚠️ Failed to mark popup transient: {transient_error}")
        else:
            sw, sh = popup.winfo_screenwidth(), popup.winfo_screenheight()
            x, y = (sw - w)//2, (sh - h)//3

        popup.geometry(f"{int(w)}x{int(h)}+{int(x)}+{int(y)}")

        try:
            popup.attributes("-topmost", True)
        except Exception as attributes_error:
            print(f"⚠️ Failed to set popup attributes: {attributes_error}")

        # Best-effort: open the job's Quotes page (works for most tenants)
        try:
            target_job_no = str(job_no_for_web or GLOBAL_JOB_NUMBER)
            target_url = f"https://app.fergus.com/jobs/view/{target_job_no}/quote"
            webbrowser.open(f"https://app.fergus.com/jobs/view/{str(job_no_for_web or GLOBAL_JOB_NUMBER)}/quote", new=2)
        except Exception:
            pass
        import sys
        popup.after(1000, lambda: (popup.destroy(), sys.exit(0)))

    except requests.RequestException as e:
        messagebox.showerror("Error", f"❌ Failed to submit quote: {e}")

class ResizeManager:
    """Centralized, flicker-free resizing with debouncing and optional width/height lock."""
    def __init__(self, root: tk.Tk):
        self.root = root
        self.pending = None
        self.in_resize = False
        self.locked_width = None
        self.locked_height = None
        self.centered_once = False
        self.last_geom = None

    def lock_width_to_current(self):
        self.root.update_idletasks()
        self.locked_width = self.root.winfo_width()

    def lock_height_to_current(self):
        self.root.update_idletasks()
        self.locked_height = self.root.winfo_height()

    def unlock(self):
        self.locked_width = None
        self.locked_height = None

    def cancel(self):
        if self.pending:
            try: self.root.after_cancel(self.pending)
            except Exception: pass
            self.pending=None

    def schedule(self, frame: tk.Widget, delay=60, min_w=820, min_h=480, max_w=1800, max_h=1000, center_threshold=64):
        self.cancel()
        self.pending = self.root.after(delay, lambda: self._do_resize(frame, min_w, min_h, max_w, max_h, center_threshold))

    def _do_resize(self, frame, min_w, min_h, max_w, max_h, center_threshold):
        if self.in_resize: return
        self.in_resize = True
        try:
            self.root.update_idletasks()
            req_w = frame.winfo_reqwidth()
            req_h = frame.winfo_reqheight()
            pad_w, pad_h = 40, 32  # borders/scrollbars
            w = max(min_w, min(req_w + pad_w, max_w))
            h = max(min_h, min(req_h + pad_h, max_h))
            if self.locked_width is not None:
                w = max(min_w, min(self.locked_width, max_w))
            if self.locked_height is not None:
                h = max(min_h, min(self.locked_height, max_h))
            new_geom = f"{int(w)}x{int(h)}"
            if self.last_geom != new_geom:
                prev_w = self.root.winfo_width()
                prev_h = self.root.winfo_height()
                self.root.geometry(new_geom)
                self.last_geom = new_geom
                if not self.centered_once or abs(prev_w - w) > center_threshold or abs(prev_h - h) > center_threshold:
                    self.center()
                    self.centered_once = True
        finally:
            self.in_resize = False

    def center(self):
        self.root.update_idletasks()
        w = self.root.winfo_width(); h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

def attach_treeview_sorting(tree: ttk.Treeview, numeric_cols=None):
    """Enable click-to-sort on a ttk.Treeview."""
    numeric_cols = set(numeric_cols or [])
    sort_state = {"col": None, "reverse": False}
    def sort_by(col):
        nonlocal sort_state
        data = [(tree.set(k, col), k) for k in tree.get_children("")]
        if col in numeric_cols:
            def to_num(s):
                try:
                    return float(re.sub(r"[^0-9.\-]", "", s) or 0)
                except:
                    return 0.0
            data.sort(key=lambda t: to_num(t[0]), reverse=sort_state["col"] == col and not sort_state["reverse"])
        else:
            data.sort(key=lambda t: t[0], reverse=sort_state["col"] == col and not sort_state["reverse"])
        for idx, (_, k) in enumerate(data):
            tree.move(k, "", idx)
            # re-apply striping tags
            tree.item(k, tags=("odd" if idx % 2 else "even",))
        sort_state = {"col": col, "reverse": sort_state["col"] == col and not sort_state["reverse"]}
    for c in tree["columns"]:
        tree.heading(c, text=c, command=lambda col=c: sort_by(col))


def attach_grouped_sorting(tree: ttk.Treeview, group_tag="section", subtotal_tag="subtotal", numeric_cols=None):
    """
    Enable click-to-sort that keeps groups (section -> rows -> subtotal) together
    and only sorts the item rows inside each group.
    """
    numeric_cols = set(numeric_cols or [])
    sort_state = {"col": None, "reverse": False}

    def to_num(s):
        try:
            return float(re.sub(r"[^0-9.\-]", "", s) or 0)
        except:
            return 0.0

    def sort_by(col):
        nonlocal sort_state
        # Gather top-level children order
        rows = list(tree.get_children(""))
        i = 0
        while i < len(rows):
            rid = rows[i]
            tags = set(tree.item(rid, "tags") or [])
            if group_tag in tags:
                # Start of a group: collect until next group or end
                group_start_index = i
                j = i + 1
                group_rows = []
                subtotal_row = None
                while j < len(rows):
                    rj = rows[j]
                    tj = set(tree.item(rj, "tags") or [])
                    if group_tag in tj:
                        break
                    if subtotal_tag in tj:
                        subtotal_row = rj
                    else:
                        group_rows.append(rj)
                    j += 1

                # Sort group_rows by column
                if col in numeric_cols:
                    group_rows.sort(
                        key=lambda rid_: to_num(tree.set(rid_, col)),
                        reverse=(sort_state["col"] == col and not sort_state["reverse"])
                    )
                else:
                    group_rows.sort(
                        key=lambda rid_: tree.set(rid_, col),
                        reverse=(sort_state["col"] == col and not sort_state["reverse"])
                    )

                # Reinsert in the same segment order: section header, sorted rows, subtotal (if any)
                idx = group_start_index + 1
                for k, gr in enumerate(group_rows):
                    tree.move(gr, "", idx + k)
                if subtotal_row is not None:
                    tree.move(subtotal_row, "", group_start_index + 1 + len(group_rows))

                # Advance i to next group header
                # Refresh rows snapshot after moves
                rows = list(tree.get_children(""))
                i = j
            else:
                i += 1

        # re-apply striping (skip section/subtotal)
        rows = list(tree.get_children(""))
        visible_idx = 0
        for rid in rows:
            tags = set(tree.item(rid, "tags") or [])
            if group_tag in tags or subtotal_tag in tags:
                continue
            new_tag = ("odd" if (visible_idx % 2) else "even")
            tree.item(rid, tags=tuple(tags.union({new_tag})))
            visible_idx += 1

        sort_state = {"col": col, "reverse": sort_state["col"] == col and not sort_state["reverse"]}

    for c in tree["columns"]:
        tree.heading(c, text=c, command=lambda col=c: sort_by(col))
class WizardApp(tk.Tk):
    def __init__(self, all_items):
        super().__init__()
        self.geometry("900x560"); self.minsize(360, 320)
        self.title("PlanSwift → Fergus Wizard")
        self._configure_theme()
        self.takeoff_width = None
        self.fergus_width = None
        self.takeoff_height = None
        self.fergus_height = None
        self.resizer = ResizeManager(self)

        self.all_items = all_items
        self.selected_groups = []
        self.show_takeoff_preview = False
        self.filtered_items = list(all_items)
        self.job_info = None
        self.existing_quotes = []

        container = ttk.Frame(self); container.pack(fill=tk.BOTH, expand=True)
        self.pages = {}
        for Page in (GroupSelectPage, TakeoffPreviewPage, FergusPreviewPage):
            frame = Page(parent=container, controller=self)
            self.pages[Page.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        container.grid_rowconfigure(0, weight=1); container.grid_columnconfigure(0, weight=1)
        for frame in self.pages.values(): frame.grid_remove()
        self.show_page("GroupSelectPage")

    def _configure_theme(self):
        try:
            style=ttk.Style(); style.theme_use("clam")
            base=tkfont.nametofont("TkDefaultFont"); base.configure(size=10)
            style.configure("TLabel", padding=(2,2), foreground="#111")
            style.configure("TButton", padding=(12,8), foreground="#111")
            style.configure("Primary.TButton", padding=(12,8), foreground="white", background="#2563eb")
            style.map("Primary.TButton",
                background=[("!disabled","#2563eb"),("pressed","#1d4ed8"),("active","#1e40af")],
                foreground=[("!disabled","white"),("pressed","white"),("active","white")])
            style.configure("Treeview", rowheight=22)
        except Exception as e:
            print("Theme config failed:", e)

    @staticmethod
    def bind_mousewheel(widget, yview_cmd):
        def _on_mousewheel(event):
            delta = 0
            if getattr(event, "num", None) == 5 or event.delta < 0: delta = 1
            else: delta = -1
            yview_cmd("scroll", delta, "units"); return "break"
        widget.bind_all("<MouseWheel>", _on_mousewheel, add=True)
        widget.bind_all("<Button-4>", _on_mousewheel, add=True)
        widget.bind_all("<Button-5>", _on_mousewheel, add=True)

    def show_page(self, name):
        for f in self.pages.values():
            f.grid_remove()
        frame = self.pages[name]
        frame.grid(row=0, column=0, sticky="nsew")
        frame.tkraise()

        # Lock remembered dimensions
        if name == "TakeoffPreviewPage":
            self.resizer.locked_width  = int(self.takeoff_width) if self.takeoff_width else None
            self.resizer.locked_height = int(self.takeoff_height) if self.takeoff_height else None
        elif name == "FergusPreviewPage":
            self.resizer.locked_width  = int(self.fergus_width) if self.fergus_width else None
            self.resizer.locked_height = int(self.fergus_height) if self.fergus_height else None
        else:
            self.resizer.unlock()

        if hasattr(frame, "on_show"):
            frame.on_show()

        self.resizer.schedule(
            frame,
            delay=80,
            min_w=860 if name != "GroupSelectPage" else 520,
            min_h=440
        )

    def update_filter(self, selected_groups):
        self.selected_groups = selected_groups
        self.filtered_items = [i for i in self.all_items if ((i.get("Group") or "Unassigned") in selected_groups)] if selected_groups else list(self.all_items)

class GroupSelectPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent); self.controller=controller
        hdr=ttk.Label(self, text="Select Groups", font=("Segoe UI",14,"bold"))
        hdr.pack(anchor="w", padx=20, pady=(20,10))
        body=ttk.Frame(self); body.pack(fill=tk.BOTH, expand=True, padx=20)
        left=ttk.Frame(body); left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(left, text="Choose groups to include:").pack(anchor="w", pady=(0,6))
        self.groups_container=ttk.Frame(left); self.groups_container.pack(anchor="w")
        right=ttk.Frame(body); right.pack(side=tk.LEFT, fill=tk.Y, padx=(24,0))
        self.preview_var=tk.BooleanVar(value=False)
        ttk.Checkbutton(right, text="Show takeoff items preview before quote", variable=self.preview_var).pack(anchor="w")
        footer=ttk.Frame(self); footer.pack(fill=tk.X, padx=20, pady=(10,20))
        ttk.Button(footer, text="Back", command=lambda: self.controller.destroy()).pack(side=tk.LEFT)
        ttk.Button(footer, text="Next", style="Primary.TButton", command=self._on_next).pack(side=tk.RIGHT)

    def on_show(self):
        for w in self.groups_container.winfo_children(): w.destroy()
        unique_groups=sorted(set((item.get("Group") or "").strip() or "Unassigned" for item in self.controller.all_items))
        self.group_vars={}; cols=2
        for idx,g in enumerate(unique_groups):
            var=tk.BooleanVar(value=(g in self.controller.selected_groups)); self.group_vars[g]=var
            r,c=divmod(idx,cols)
            ttk.Checkbutton(self.groups_container, text=g, variable=var).grid(row=r, column=c, sticky="w", padx=(0,24), pady=2)
        self.preview_var.set(self.controller.show_takeoff_preview)
        self.controller.resizer.schedule(self, delay=60, min_w=520, min_h=440, max_w=1200, max_h=900)

    def _on_next(self):
        selected=[g for g,v in self.group_vars.items() if v.get()]
        self.controller.update_filter(selected)
        self.controller.show_takeoff_preview=bool(self.preview_var.get())
        self.controller.show_page("TakeoffPreviewPage" if self.controller.show_takeoff_preview else "FergusPreviewPage")

class TakeoffPreviewPage(ttk.Frame):
    COLS=["Name","Description","Takeoff","Qty","Units","Hours","Cost Each","Price Each","Price Total","Group"]

    def __init__(self,parent,controller):
        super().__init__(parent); self.controller=controller
        hdr=ttk.Label(self, text="Takeoff Items", font=("Segoe UI",14,"bold"))
        hdr.pack(anchor="w", padx=20, pady=(20,10))

        outer=ttk.Frame(self); outer.pack(fill=tk.BOTH, expand=True, padx=20)
        tree_frame=ttk.Frame(outer); tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree=ttk.Treeview(tree_frame, columns=self.COLS, show="headings", selectmode="browse")
        vbar=ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)

        for c in self.COLS:
            self.tree.heading(c, text=c)
            anchor = "e" if c in ("Qty","Hours","Cost Each","Price Each","Price Total") else "w"
            default_w = 110
            if c in ("Name", "Description"):
                default_w = 260
            elif c in ("Takeoff", "Group"):
                default_w = 140
            self.tree.column(c, width=default_w, anchor=anchor, stretch=(c in ("Name","Description")))

        # Row striping
        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("odd", background="#f6f7fb")

        footer=ttk.Frame(self); footer.pack(fill=tk.X, padx=20, pady=(10,20))
        ttk.Button(footer, text="Back", command=lambda: self.controller.show_page("GroupSelectPage")).pack(side=tk.LEFT)
        btn_next = ttk.Button(footer, text="Next", style="Primary.TButton", command=lambda: self.controller.show_page("FergusPreviewPage"))
        btn_next.pack(side=tk.RIGHT, padx=(8,0))

        WizardApp.bind_mousewheel(self.tree, self.tree.yview)

        self.tree.bind("<Configure>", lambda e: self._stretch_wide_cols())

        # Sorting
        attach_treeview_sorting(self.tree, numeric_cols={"Qty","Hours","Cost Each","Price Each","Price Total"})

    def _required_width(self):
        fixed = 0
        for c in self.COLS:
            if c in ("Name", "Description"): continue
            try:
                fixed += int(self.tree.column(c, option='width'))
            except Exception: pass
        name_min = 320; desc_min = 420; scrollbar_w = 18; padding = 80
        return fixed + name_min + desc_min + scrollbar_w + padding

    def _stretch_wide_cols(self):
        try:
            self.update_idletasks()
            total_w = self.tree.winfo_width()
            fixed_total = 0
            for c in self.COLS:
                if c in ("Name", "Description"): continue
                fixed_total += int(self.tree.column(c, option='width'))
            padding = 28
            remaining = max(200, total_w - fixed_total - padding)
            name_share = int(remaining * 0.55)
            desc_share = max(160, remaining - name_share)
            self.tree.column("Name", width=name_share)
            self.tree.column("Description", width=desc_share)
        except Exception:
            pass

    def on_show(self):
        self.tree.delete(*self.tree.get_children())
        for idx, row in enumerate(self.controller.filtered_items):
            tags = ("odd" if idx % 2 else "even",)
            vals = [str(row.get(c,"")) for c in self.COLS]
            self.tree.insert("", tk.END, values=vals, tags=tags)

        self.after(30, self._stretch_wide_cols)
        min_w = max(980, self._required_width())
        self.controller.resizer.schedule(self, delay=60, min_w=min_w, min_h=560, max_w=1920, max_h=1100)

        # Capture stable size for later page returns
        self.after(180, self._capture_size)

    def _capture_size(self):
        self.controller.resizer.lock_width_to_current()
        self.controller.resizer.lock_height_to_current()
        self.controller.takeoff_width  = int(self.controller.resizer.locked_width or self.controller.winfo_width())
        self.controller.takeoff_height = int(self.controller.resizer.locked_height or self.controller.winfo_height())

class FergusPreviewPage(ttk.Frame):
    HEADERS=["Name","Qty","Cost Each","Price Each","Line Total"]
    def __init__(self,parent,controller):
        super().__init__(parent); self.controller=controller; self.quote_id_lookup={}
        self.hdr=ttk.Label(self, text="Quote Preview", font=("Segoe UI",14,"bold"))
        self.hdr.pack(anchor="w", padx=20, pady=(20,10))
        self.job_frame=ttk.Frame(self); self.job_frame.pack(fill=tk.X, padx=20)
        self.lbl_jobno=ttk.Label(self.job_frame, text="Job No:"); self.lbl_jobno.grid(row=0,column=0,sticky="w",pady=(0,2))
        self.lbl_desc=ttk.Label(self.job_frame, text=""); self.lbl_desc.grid(row=1,column=0,sticky="w")
        self.lbl_cust=ttk.Label(self.job_frame, text=""); self.lbl_cust.grid(row=2,column=0,sticky="w")
        self.lbl_acc=ttk.Label(self.job_frame, text=""); self.lbl_acc.grid(row=3,column=0,sticky="w",pady=(0,8))
        # Override job number UI
        ttk.Label(self.job_frame, text="Override Job No:").grid(row=0, column=2, sticky="e", padx=(20,4))
        self.override_var = tk.StringVar(value="")
        self.entry_override = ttk.Entry(self.job_frame, textvariable=self.override_var, width=14)
        self.entry_override.grid(row=0, column=3, sticky="w")
        self.entry_override.bind("<Return>", lambda e: self._load_override_job())
        self.btn_override = ttk.Button(self.job_frame, text="Load", command=self._load_override_job)
        self.btn_override.grid(row=0, column=4, sticky="w", padx=(6,0))

        qf=ttk.Frame(self); qf.pack(fill=tk.X, padx=20, pady=(0,10))
        ttk.Label(qf, text="Existing Quotes:").grid(row=0,column=0,sticky="w")
        self.dropdown_var=tk.StringVar(); self.dropdown=ttk.Combobox(qf, textvariable=self.dropdown_var, state="readonly", width=60)
        self.dropdown.grid(row=0,column=1,sticky="w",padx=(8,0))

        table_wrap=ttk.Frame(self); table_wrap.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0,10))
        self.table=ttk.Treeview(table_wrap, columns=self.HEADERS, show="headings", selectmode="browse")
        vbar=ttk.Scrollbar(table_wrap, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=vbar.set)
        self.table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); vbar.pack(side=tk.RIGHT, fill=tk.Y)

        for i,c in enumerate(self.HEADERS):
            self.table.heading(c, text=c)
            anchor = "e" if i>0 else "w"
            w = 300 if i==0 else 110
            self.table.column(c, width=w, anchor=anchor, stretch=(i==0))

        self.table.tag_configure("section", font=("Segoe UI",10,"bold"))
        self.table.tag_configure("subtotal", font=("Segoe UI",10,"bold"))
        self.table.tag_configure("even", background="#ffffff")
        self.table.tag_configure("odd", background="#f6f7fb")

        self.total_var=tk.StringVar(value="Total Quote Price: $0.00")
        total_frame=ttk.Frame(self); total_frame.pack(fill=tk.X, padx=20)
        ttk.Label(total_frame, textvariable=self.total_var, font=("Segoe UI",11,"bold")).pack(anchor="e", pady=(0,6))

        footer=ttk.Frame(self); footer.pack(fill=tk.X, padx=20, pady=(8,16))
        ttk.Button(footer, text="Back", command=self._on_back).pack(side=tk.LEFT)
        ttk.Button(footer, text="Cancel", command=self._on_cancel).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(footer, text="Update Selected Quote", style="Primary.TButton", command=self._on_update).pack(side=tk.RIGHT, padx=(8,0))
        ttk.Button(footer, text="Create New Quote", command=self._on_create).pack(side=tk.RIGHT)

        WizardApp.bind_mousewheel(self.table, self.table.yview)
        attach_grouped_sorting(self.table, group_tag="section", subtotal_tag="subtotal", numeric_cols={"Qty","Cost Each","Price Each","Line Total"})

    def on_show(self):
        if not self.controller.job_info:
            ji=get_job_details(GLOBAL_JOB_NUMBER)
            if not ji:
                messagebox.showerror("Error", f"❌ Job {GLOBAL_JOB_NUMBER} not found in Fergus."); self.controller.show_page("GroupSelectPage"); return
            self.controller.job_info=ji; self.controller.existing_quotes=get_existing_quotes(ji["id"]) or []
        self._refresh_job_header_and_quotes()

        # Pre-fill override field with the current job number for convenience
        try:
            self.override_var.set(str(self.controller.job_info.get("jobNo", "")))
        except Exception:
            pass

        # Build table
        self.table.delete(*self.table.get_children())
        grouped=group_items(self.controller.filtered_items)
        grand_total=0.0
        row_idx = 0
        for group_name, items in grouped.items():
            self.table.insert("", "end", values=(group_name,"","","",""), tags=("section",))
            row_idx += 1
            section_total=0.0
            for row in items:
                name, qty, cost, price, line_total, _ = compute_line_values(row)
                if not name: continue
                section_total+=line_total; grand_total+=line_total
                tags=("odd" if (row_idx % 2) else "even",)
                self.table.insert("", "end",
                    values=(name, f"{qty:.2f}", f"${cost:,.2f}", f"${price:,.2f}", f"${line_total:,.2f}"),
                    tags=tags)
                row_idx += 1
            self.table.insert("", "end", values=("", "", "", "Subtotal", f"${section_total:,.2f}"), tags=("subtotal",))
            row_idx += 1
        self.total_var.set(f"Total Quote Price: ${grand_total:,.2f}")
        self.table.configure(height=22)

        self.after(60, self._stretch_name_col)
        self.controller.resizer.schedule(self, delay=70, min_w=1000, min_h=700, max_w=1920, max_h=1200)
        self.after(180, self._capture_size)

    def _stretch_name_col(self):
        try:
            self.update_idletasks()
            total_w = self.table.winfo_width()
            fixed_w = 0
            for c in self.HEADERS[1:]:
                fixed_w += int(self.table.column(c, option='width'))
            padding = 24
            name_w = max(200, total_w - fixed_w - padding)
            self.table.column(self.HEADERS[0], width=name_w)
        except Exception:
            pass

    def _capture_size(self):
        self.controller.resizer.lock_width_to_current()
        self.controller.resizer.lock_height_to_current()
        self.controller.fergus_width  = int(self.controller.resizer.locked_width or self.controller.winfo_width())
        self.controller.fergus_height = int(self.controller.resizer.locked_height or self.controller.winfo_height())

    def _on_back(self):
        if self.controller.show_takeoff_preview: self.controller.show_page("TakeoffPreviewPage")
        else: self.controller.show_page("GroupSelectPage")
    def _on_cancel(self): self.controller.destroy()

    def _preflight(self):
        problems = validate_items(self.controller.filtered_items)
        if problems:
            msg = "Please review the following issues before continuing:\\n\\n" + "\\n".join("• " + p for p in problems[:20])
            if len(problems) > 20:
                msg += f"\\n…and {len(problems)-20} more."
            messagebox.showwarning("Validation issues", msg)
            return False
        return True

    def _on_create(self):
        if not self._preflight(): return
        job=self.controller.job_info
        if not job: return
        push_quote(
            job["id"],
            job.get("description") or "",
            self.controller.filtered_items,
            job_no_for_web=job.get("jobNo"),
            parent=self.controller,
        )

    def _on_update(self):
        if not self._preflight(): return
        job=self.controller.job_info
        if not job: return
        selected=self.dropdown.get()
        if not selected: messagebox.showerror("Error","❌ No quote selected to update."); return
        if "Accepted" in selected:
            messagebox.showwarning("Blocked","✅ This quote is already accepted and cannot be updated."); return
        quote_id=self.quote_id_lookup.get(selected)
        if not quote_id: messagebox.showerror("Error","❌ Could not resolve selected quote."); return
        if not messagebox.askyesno("Confirm", f"Update quote {selected}?"): return
        push_quote(
            job["id"],
            job.get("description") or "",
            self.controller.filtered_items,
            quote_id=quote_id,
            job_no_for_web=job.get("jobNo"),
            parent=self.controller,
        )


    def _load_override_job(self):
        """Load a manually-entered job number and switch the target job + quotes list."""
        raw = (self.override_var.get() or "").strip()
        if not raw:
            messagebox.showwarning("Missing", "Please enter a job number to override.")
            return
        override_no = extract_digits(raw)
        ji = get_job_details(override_no)
        if not ji:
            messagebox.showerror("Not found", f"❌ Job {override_no} not found in Fergus.")
            return
        self.controller.job_info = ji
        self.controller.existing_quotes = get_existing_quotes(ji["id"]) or []
        self._refresh_job_header_and_quotes()

    def _refresh_job_header_and_quotes(self):
        """Refresh labels and quotes dropdown from controller.job_info / existing_quotes."""
        job = self.controller.job_info or {}
        self.lbl_jobno.configure(text=f"Job No: {job.get('jobNo','')}")
        self.lbl_desc.configure(text=job.get('description') or "")
        self.lbl_cust.configure(text=f"Customer: {job.get('customer') or ''}")
        self.lbl_acc.configure(text=f"Quote Accepted: {'✅ Yes' if job.get('quoteAccepted') else '❌ No'}")

        self.quote_id_lookup.clear()
        formatted = []
        for q in sorted(self.controller.existing_quotes, key=lambda x: x.get("versionNumber",0)):
            version = q.get("versionNumber")
            created = (q.get("lastModified") or "").split("T")[0]
            status = []
            if q.get("isAccepted"): status.append("Accepted")
            if q.get("isSent"): status.append("Sent")
            if not status: status.append("Draft")
            label = f"v{version} - {'/'.join(status)} ({created})"
            formatted.append(label)
            self.quote_id_lookup[label] = q.get("id")
        self.dropdown["values"] = formatted
        self.dropdown.set(formatted[-1] if formatted else "")
if __name__=="__main__":
    try:
        app=WizardApp(ALL_ITEMS); app.mainloop()
    except Exception as e:
        import traceback, sys
        tb=traceback.format_exc(); print(">>> Caught exception:"+tb)
        try:
            if not hasattr(sys,'ps1'):
                root=tk.Tk(); root.withdraw(); messagebox.showerror("Script Crashed", tb); root.destroy()
        except Exception: pass