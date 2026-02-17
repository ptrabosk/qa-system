#!/usr/bin/env python3
import csv
import json
import re
import subprocess
import unicodedata
from copy import deepcopy
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox


def normalize_text(value):
    if value is None:
        return ""
    return unicodedata.normalize("NFKC", str(value))


def parse_json_text(text):
    if text is None:
        return None
    raw = str(text).strip()
    if not raw:
        return None
    try:
        return json.loads(raw)
    except Exception:
        return None


def read_json_object(path):
    if not path.exists():
        return {}
    raw = path.read_text(encoding="utf-8").strip()
    if not raw:
        return {}
    return json.loads(raw)


def write_json_object(path, value):
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False), encoding="utf-8")


def is_message_like(item):
    if not isinstance(item, dict):
        return False
    keys = set(item.keys())
    return bool({"message_text", "message_type", "content", "role"} & keys)


def get_scenario_count(payload):
    if payload is None:
        return 0
    if isinstance(payload, list):
        if not payload:
            return 0
        if all(is_message_like(x) for x in payload):
            return 1
        return len(payload)
    if isinstance(payload, dict) and "scenarios" in payload:
        scenarios = payload.get("scenarios")
        if isinstance(scenarios, list):
            if not scenarios:
                return 0
            if all(is_message_like(x) for x in scenarios):
                return 1
            return len(scenarios)
        if isinstance(scenarios, dict):
            return len(scenarios)
    return 0


def get_template_count(payload):
    if payload is None:
        return 0
    if isinstance(payload, list):
        return len(payload)
    if isinstance(payload, dict) and isinstance(payload.get("templates"), list):
        return len(payload["templates"])
    return 0


def resolve_default_working_folder(script_path):
    candidate = script_path.parent
    for _ in range(6):
        if (candidate / "scenarios.json").exists() or (candidate / "templates.json").exists():
            return candidate
        if candidate.parent == candidate:
            break
        candidate = candidate.parent
    return script_path.parent


def convert_to_string_array(value):
    out = []
    if value is None:
        return out
    if isinstance(value, list):
        items = value
    elif isinstance(value, dict):
        items = list(value.values())
    elif hasattr(value, "__dict__") and not isinstance(value, str):
        items = list(vars(value).values())
    else:
        items = [value]
    for item in items:
        text = normalize_text(item).strip()
        if text and text not in {"{}", "[]"}:
            out.append(text)
    return out


def unique_trimmed_string_array(value):
    seen = set()
    result = []
    for item in convert_to_string_array(value):
        text = normalize_text(item).strip()
        if not text:
            continue
        key = text.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(text)
    return result


def parse_list_like_text(text):
    raw = normalize_text(text).strip()
    if not raw or raw == "[]":
        return []

    parsed = parse_json_text(raw)
    if parsed is not None:
        return convert_to_string_array(parsed)

    matches = re.findall(r"'([^']*)'|\"([^\"]*)\"", raw)
    if matches:
        out = []
        for left, right in matches:
            val = (left or right).strip()
            if val:
                out.append(val)
        return out

    fallback = raw.strip("[]")
    if not fallback:
        return []
    parts = re.split(r"[,\n\r]+", fallback)
    return [p.strip(" \"'") for p in parts if p.strip(" \"'")]


def normalize_guideline_category_key(heading):
    if not heading or not str(heading).strip():
        return "important"
    key = normalize_text(heading).strip().lower()
    key = re.sub(r"^[^a-z0-9]+", "", key)
    key = key.replace("&", "and")
    key = re.sub(r"[^a-z0-9]+", "_", key).strip("_")
    if re.search(r"send.*cs", key):
        return "send_to_cs"
    if re.search(r"^escalate$|^escalation$|escalat", key):
        return "escalate"
    if re.search(r"^tone$", key):
        return "tone"
    if re.search(r"template", key):
        return "templates"
    if re.search(r"do.*and.*don|dos_and_donts|don_ts|donts", key):
        return "dos_and_donts"
    if re.search(r"drive.*purchase", key):
        return "drive_to_purchase"
    if re.search(r"promo", key):
        return "promo_and_exclusions"
    return key or "important"


def has_styled_math_chars(text):
    if not text:
        return False
    return any(0x1D400 <= ord(ch) <= 0x1D7FF for ch in text)


def parse_company_notes_to_categories(notes_text):
    raw = normalize_text(notes_text).strip()
    if not raw:
        return {}
    notes = {"important": []}
    current_key = "important"
    for line in re.split(r"\r?\n", raw):
        item = line.strip()
        if not item:
            continue
        if item.startswith("#"):
            current_key = normalize_guideline_category_key(item.lstrip("#").strip())
            notes.setdefault(current_key, [])
            continue
        if item.startswith("•"):
            item = item[1:].strip()
        if item.startswith("-"):
            item = item[1:].strip()
        if not item:
            continue
        if has_styled_math_chars(item):
            item = f"**{normalize_text(item).strip()}**"
        notes.setdefault(current_key, []).append(normalize_text(item).strip())
    return {k: v for k, v in notes.items() if v}


def normalize_scenario_notes(notes_value):
    if notes_value is None:
        return {}
    if not isinstance(notes_value, dict):
        return {}

    notes_out = {}
    key_order = []
    for raw_key, raw_val in notes_value.items():
        key = normalize_guideline_category_key(normalize_text(raw_key).strip())
        if key not in notes_out:
            notes_out[key] = []
            key_order.append(key)
        for item in convert_to_string_array(raw_val):
            txt = normalize_text(item).strip()
            if not txt:
                continue
            heading_match = re.match(r"^\*{0,2}\s*#\s*(.+)$", txt)
            if heading_match:
                moved = normalize_guideline_category_key(heading_match.group(1))
                if moved not in notes_out:
                    notes_out[moved] = []
                    key_order.append(moved)
                continue
            notes_out[key].append(txt)

    if "important" in notes_out:
        keep = []
        for item in notes_out["important"]:
            txt = normalize_text(item).strip()
            if re.search(r"send\s*to\s*cs|cssupport@|post-purchase|shipping inquiries on a current order", txt, re.I):
                notes_out.setdefault("send_to_cs", [])
                if "send_to_cs" not in key_order:
                    key_order.append("send_to_cs")
                notes_out["send_to_cs"].append(txt)
                continue
            if txt == "**":
                continue
            keep.append(txt)
        notes_out["important"] = keep

    clean = {}
    for key in key_order:
        if key not in notes_out:
            continue
        arr = unique_trimmed_string_array(notes_out[key])
        if arr:
            clean[key] = arr
    return clean


def normalize_scenario_record_for_storage(scenario):
    if not isinstance(scenario, dict):
        return {}
    out = deepcopy(scenario)

    right_panel = {}
    if isinstance(out.get("rightPanel"), dict):
        right_panel.update(out["rightPanel"])

    if "source" in out and "source" not in right_panel:
        right_panel["source"] = out.pop("source")
    if "browsingHistory" in out and "browsingHistory" not in right_panel:
        right_panel["browsingHistory"] = out.pop("browsingHistory")
    if "browsing_history" in out and "browsingHistory" not in right_panel:
        right_panel["browsingHistory"] = out.pop("browsing_history")
    if "orders" in out and "orders" not in right_panel:
        right_panel["orders"] = out.pop("orders")
    if "templatesUsed" in out and "templates" not in right_panel:
        right_panel["templates"] = out.pop("templatesUsed")
    if right_panel:
        out["rightPanel"] = right_panel

    blocklisted_source = out.get("blocklisted_words", out.get("blocklistedWords", []))
    out["blocklisted_words"] = unique_trimmed_string_array(blocklisted_source)
    out.pop("blocklistedWords", None)

    escalation_source = out.get("escalation_preferences", out.get("escalationPreferences", []))
    out["escalation_preferences"] = unique_trimmed_string_array(escalation_source)
    out.pop("escalationPreferences", None)

    notes_value = out.get("notes", out.get("guidelines"))
    out["notes"] = normalize_scenario_notes(notes_value)
    out.pop("guidelines", None)

    return out


def convert_scenario_container_to_list(container):
    if container is None:
        return []
    if isinstance(container, list):
        return list(container)
    if isinstance(container, dict) and "scenarios" in container:
        sc = container["scenarios"]
        if isinstance(sc, list):
            return list(sc)
        if isinstance(sc, dict):
            return list(sc.values())
    return []


def merge_scenarios_by_id(existing, incoming):
    result = [normalize_scenario_record_for_storage(item) for item in (existing or [])]
    id_to_index = {}
    for i, item in enumerate(result):
        sid = normalize_text(item.get("id", "")).strip()
        if sid and sid not in id_to_index:
            id_to_index[sid] = i

    updated = 0
    added = 0

    for item in (incoming or []):
        item_norm = normalize_scenario_record_for_storage(item)
        incoming_id = normalize_text(item_norm.get("id", "")).strip()
        if incoming_id and incoming_id in id_to_index:
            idx = id_to_index[incoming_id]
            base = normalize_scenario_record_for_storage(result[idx])
            merged = {**base, **item_norm}
            if base.get("rightPanel") or item_norm.get("rightPanel"):
                merged["rightPanel"] = {**(base.get("rightPanel") or {}), **(item_norm.get("rightPanel") or {})}
            result[idx] = normalize_scenario_record_for_storage(merged)
            updated += 1
            continue
        result.append(item_norm)
        added += 1
        if incoming_id:
            id_to_index[incoming_id] = len(result) - 1

    return {"scenarios": result, "updated": updated, "added": added}


def get_obj_prop_value(obj, names):
    if not isinstance(obj, dict):
        return None
    for name in names:
        if name in obj:
            return obj[name]
    return None


def normalize_message_media(media):
    result = []
    if media is None:
        return result
    items = media if isinstance(media, list) else [media]
    for item in items:
        text = normalize_text(item).strip()
        if not text:
            continue
        if text.startswith("[") and text.endswith("]"):
            nested = parse_json_text(text)
            if isinstance(nested, list):
                for nested_item in nested:
                    nested_text = normalize_text(nested_item).strip()
                    if nested_text:
                        result.append(nested_text)
                continue
        result.append(text)
    return result


def convert_csv_row_to_scenario(row):
    conversation = []
    conversation_raw = normalize_text(row.get("CONVERSATION_JSON", ""))
    conversation_parsed = parse_json_text(conversation_raw)
    if isinstance(conversation_parsed, list):
        for msg in conversation_parsed:
            if not isinstance(msg, dict):
                continue
            media = get_obj_prop_value(msg, ["message_media", "media"])
            text_raw = get_obj_prop_value(msg, ["message_text", "content"])
            type_raw = get_obj_prop_value(msg, ["message_type", "role"])
            entry = {
                "message_media": normalize_message_media(media),
                "message_text": normalize_text(text_raw),
                "message_type": normalize_text(type_raw).lower(),
            }
            date_time = normalize_text(get_obj_prop_value(msg, ["date_time", "dateTime", "timestamp"])).strip()
            if date_time:
                entry["date_time"] = date_time
            msg_id = normalize_text(get_obj_prop_value(msg, ["message_id", "id"])).strip()
            if msg_id:
                entry["message_id"] = msg_id
            conversation.append(entry)

    browsing_history = []
    products_parsed = parse_json_text(normalize_text(row.get("LAST_5_PRODUCTS", "")))
    if isinstance(products_parsed, list):
        for p in products_parsed:
            if not isinstance(p, dict):
                continue
            name = normalize_text(p.get("product_name")).strip()
            link = normalize_text(p.get("product_link")).strip()
            view_date = normalize_text(p.get("view_date")).strip()
            if not name and not link:
                continue
            item = {"item": name if name else link}
            if link:
                item["link"] = link
            if view_date:
                item["timeAgo"] = view_date
            browsing_history.append(item)

    orders_out = []
    orders_parsed = parse_json_text(normalize_text(row.get("ORDERS", "")))
    if isinstance(orders_parsed, list):
        for order in orders_parsed:
            if not isinstance(order, dict):
                continue
            items_out = []
            if isinstance(order.get("products"), list):
                for prod in order["products"]:
                    if not isinstance(prod, dict):
                        continue
                    item_out = {"name": normalize_text(prod.get("product_name")).strip()}
                    price_value = prod.get("product_price", prod.get("price"))
                    if normalize_text(price_value).strip():
                        item_out["price"] = price_value
                    prod_link = normalize_text(prod.get("product_link")).strip()
                    if prod_link:
                        item_out["productLink"] = prod_link
                    items_out.append(item_out)
            order_out = {
                "orderNumber": normalize_text(order.get("order_number")).strip(),
                "orderDate": normalize_text(order.get("order_date")).strip(),
                "items": items_out,
            }
            order_link = normalize_text(order.get("order_status_url")).strip()
            if order_link:
                order_out["link"] = order_link
            if normalize_text(order.get("total")).strip():
                order_out["total"] = order.get("total")
            orders_out.append(order_out)

    right_panel = {
        "source": {
            "label": "Website",
            "value": normalize_text(row.get("COMPANY_WEBSITE")).strip(),
            "date": "",
        }
    }
    if browsing_history:
        right_panel["browsingHistory"] = browsing_history
    if orders_out:
        right_panel["orders"] = orders_out

    notes = parse_company_notes_to_categories(normalize_text(row.get("COMPANY_NOTES", "")).strip())

    return {
        "id": normalize_text(row.get("SEND_ID")).strip(),
        "companyName": normalize_text(row.get("COMPANY_NAME")).strip(),
        "companyWebsite": normalize_text(row.get("COMPANY_WEBSITE")).strip(),
        "agentName": normalize_text(row.get("PERSONA")).strip(),
        "messageTone": normalize_text(row.get("MESSAGE_TONE")).strip(),
        "conversation": conversation,
        "notes": notes,
        "rightPanel": right_panel,
        "escalation_preferences": convert_to_string_array(parse_list_like_text(normalize_text(row.get("ESCALATION_TOPICS")))),
        "blocklisted_words": convert_to_string_array(parse_list_like_text(normalize_text(row.get("BLOCKLISTED_WORDS")))),
    }


class ContentManagerMacApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Scenario & Template Manager (macOS)")
        self.root.geometry("780x430")
        self.root.minsize(740, 400)

        script_path = Path(__file__).resolve()
        self.current_folder = resolve_default_working_folder(script_path)

        self.build_ui()
        self.refresh_meta()

    def scenarios_path(self):
        return self.current_folder / "scenarios.json"

    def templates_path(self):
        return self.current_folder / "templates.json"

    def set_status(self, message, is_error=False):
        self.status_text.config(state="normal")
        self.status_text.delete("1.0", tk.END)
        self.status_text.insert(tk.END, message)
        self.status_text.config(fg="#b02318" if is_error else "#33476b", state="disabled")

    def build_ui(self):
        top = tk.Frame(self.root, padx=16, pady=14)
        top.pack(fill="x")

        title = tk.Label(top, text="Scenario & Template Manager", font=("Helvetica", 15, "bold"))
        title.grid(row=0, column=0, sticky="w")

        choose_btn = tk.Button(top, text="Choose Folder", width=14, command=self.choose_folder)
        choose_btn.grid(row=0, column=1, sticky="e", padx=(8, 0))
        top.grid_columnconfigure(0, weight=1)

        self.folder_label = tk.Label(top, text=f"Folder: {self.current_folder}", anchor="w")
        self.folder_label.grid(row=1, column=0, columnspan=2, sticky="we", pady=(6, 0))

        middle = tk.Frame(self.root, padx=16, pady=8)
        middle.pack(fill="x")

        scenarios_box = tk.LabelFrame(middle, text="Scenarios", padx=10, pady=10)
        scenarios_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.scenarios_meta = tk.Label(scenarios_box, text="Items: 0")
        self.scenarios_meta.pack(anchor="w", pady=(0, 8))
        tk.Button(scenarios_box, text="Upload JSON / CSV", width=18, command=self.import_scenarios).pack(side="left")
        tk.Button(scenarios_box, text="Clear Scenarios", width=18, command=self.clear_scenarios).pack(side="left", padx=8)

        templates_box = tk.LabelFrame(middle, text="Templates", padx=10, pady=10)
        templates_box.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self.templates_meta = tk.Label(templates_box, text="Items: 0")
        self.templates_meta.pack(anchor="w", pady=(0, 8))
        tk.Button(templates_box, text="Upload JSON / CSV", width=18, command=self.import_templates).pack(side="left")
        tk.Button(templates_box, text="Clear Templates", width=18, command=self.clear_templates).pack(side="left", padx=8)

        middle.grid_columnconfigure(0, weight=1)
        middle.grid_columnconfigure(1, weight=1)

        actions = tk.Frame(self.root, padx=16, pady=6)
        actions.pack(fill="x")
        tk.Button(actions, text="Open Current Folder", width=20, command=self.open_current_folder).pack(anchor="w")

        status_group = tk.LabelFrame(self.root, text="Status", padx=10, pady=10)
        status_group.pack(fill="both", expand=True, padx=16, pady=(8, 14))
        self.status_text = tk.Text(status_group, height=4, wrap="word", state="disabled")
        self.status_text.pack(fill="both", expand=True)
        self.set_status("Ready.")

    def refresh_meta(self):
        try:
            scenarios_json = read_json_object(self.scenarios_path())
            templates_json = read_json_object(self.templates_path())
            self.scenarios_meta.config(text=f"Items: {get_scenario_count(scenarios_json)}")
            self.templates_meta.config(text=f"Items: {get_template_count(templates_json)}")
        except Exception as exc:
            self.set_status(f"Failed to read JSON files: {exc}", is_error=True)

    def choose_folder(self):
        selected = filedialog.askdirectory(initialdir=str(self.current_folder))
        if not selected:
            return
        self.current_folder = Path(selected)
        self.folder_label.config(text=f"Folder: {self.current_folder}")
        self.refresh_meta()
        self.set_status(f"Connected folder: {self.current_folder}")

    def import_templates(self):
        file_path = filedialog.askopenfilename(
            initialdir=str(self.current_folder),
            filetypes=[
                ("Template sources", "*.json *.csv"),
                ("JSON files", "*.json"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not file_path:
            return
        try:
            src = Path(file_path)
            if src.suffix.lower() == ".csv":
                templates = []
                with src.open("r", encoding="utf-8-sig", newline="") as fh:
                    rows = csv.DictReader(fh)
                    for row in rows:
                        name = next((normalize_text(row.get(k)).strip() for k in ["TEMPLATE_TITLE", "TEMPLATE_NAME", "NAME", "TEMPLATE", "TITLE"] if normalize_text(row.get(k)).strip()), "")
                        content = next((normalize_text(row.get(k)).strip() for k in ["TEMPLATE_TEXT", "CONTENT", "TEMPLATE_CONTENT", "BODY", "TEXT", "MESSAGE"] if normalize_text(row.get(k)).strip()), "")
                        shortcut = next((normalize_text(row.get(k)).strip() for k in ["SHORTCUT", "CODE", "KEYWORD"] if normalize_text(row.get(k)).strip()), "")
                        company = next((normalize_text(row.get(k)).strip() for k in ["COMPANY_NAME", "COMPANY", "BRAND"] if normalize_text(row.get(k)).strip()), "")
                        template_id = next((normalize_text(row.get(k)).strip() for k in ["TEMPLATE_ID", "ID"] if normalize_text(row.get(k)).strip()), "")

                        if not name or not content:
                            continue
                        tpl = {"name": name, "content": content}
                        if template_id:
                            tpl["id"] = template_id
                        if shortcut:
                            tpl["shortcut"] = shortcut
                        if company:
                            tpl["companyName"] = company
                        templates.append(tpl)
                write_json_object(self.templates_path(), {"templates": templates})
                self.refresh_meta()
                self.set_status(f"templates.json updated from CSV ({len(templates)} template(s)).")
                return

            parsed = json.loads(src.read_text(encoding="utf-8"))
            write_json_object(self.templates_path(), parsed)
            self.refresh_meta()
            self.set_status(f"templates.json updated from {src.name}.")
        except Exception as exc:
            self.set_status(f"Invalid JSON for templates.json: {exc}", is_error=True)

    def import_scenarios(self):
        file_path = filedialog.askopenfilename(
            initialdir=str(self.current_folder),
            filetypes=[
                ("Scenario sources", "*.json *.csv"),
                ("JSON files", "*.json"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not file_path:
            return
        try:
            existing_obj = read_json_object(self.scenarios_path())
            existing_list = convert_scenario_container_to_list(existing_obj)

            src = Path(file_path)
            if src.suffix.lower() == ".csv":
                incoming = []
                with src.open("r", encoding="utf-8-sig", newline="") as fh:
                    rows = csv.DictReader(fh)
                    for row in rows:
                        incoming.append(convert_csv_row_to_scenario(row))
                merged = merge_scenarios_by_id(existing_list, incoming)
                write_json_object(self.scenarios_path(), {"scenarios": merged["scenarios"]})
                self.refresh_meta()
                self.set_status(f"scenarios.json updated from CSV. Added: {merged['added']}, Updated: {merged['updated']}.")
                return

            parsed = json.loads(src.read_text(encoding="utf-8"))
            incoming_list = convert_scenario_container_to_list(parsed)
            if not incoming_list:
                raise ValueError("No scenarios found in selected file.")
            merged = merge_scenarios_by_id(existing_list, incoming_list)
            write_json_object(self.scenarios_path(), {"scenarios": merged["scenarios"]})
            self.refresh_meta()
            self.set_status(f"scenarios.json updated from {src.name}. Added: {merged['added']}, Updated: {merged['updated']}.")
        except Exception as exc:
            self.set_status(f"Failed to import scenarios source: {exc}", is_error=True)

    def clear_scenarios(self):
        if not messagebox.askyesno("Confirm Clear", 'Clear scenarios.json and reset it to { "scenarios": [] }?'):
            return
        try:
            write_json_object(self.scenarios_path(), {"scenarios": []})
            self.refresh_meta()
            self.set_status("scenarios.json cleared.")
        except Exception as exc:
            self.set_status(f"Failed to clear scenarios.json: {exc}", is_error=True)

    def clear_templates(self):
        if not messagebox.askyesno("Confirm Clear", 'Clear templates.json and reset it to { "templates": [] }?'):
            return
        try:
            write_json_object(self.templates_path(), {"templates": []})
            self.refresh_meta()
            self.set_status("templates.json cleared.")
        except Exception as exc:
            self.set_status(f"Failed to clear templates.json: {exc}", is_error=True)

    def open_current_folder(self):
        try:
            subprocess.run(["open", str(self.current_folder)], check=False)
        except Exception as exc:
            self.set_status(f"Could not open folder: {exc}", is_error=True)


def main():
    root = tk.Tk()
    app = ContentManagerMacApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
