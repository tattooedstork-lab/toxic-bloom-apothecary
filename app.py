import os
import json
from datetime import datetime
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from google import genai
from google.genai import types
from pydantic import BaseModel, Field

st.set_page_config(page_title="Toxic Bloom Apothecary", page_icon="🧪", layout="wide")

SHEET_URL = "https://docs.google.com/spreadsheets/d/1VqqSIPf28ScSR1EEI7vcduvJot2cuSzjgBPGkbEk2jY/edit"
KEY_FILE = "gemini_key.txt"

GRAMS_PER_OZ = 28.3495
GRAMS_PER_LB = 453.592
GRAMS_PER_KG = 1000.0
CATEGORIES = ["Soap Bars", "Dog Soap", "Bath Bombs", "Chapsticks"]

# --- OPTIONAL PIN LOCK FOR CLOUD HOSTING ---
if "APP_PIN" in st.secrets:
    if not st.session_state.get("pin_unlocked", False):
        st.title("🧪 Toxic Bloom Apothecary — Studio Login")
        entered_pin = st.text_input("Enter Studio PIN", type="password")
        if st.button("Unlock Studio"):
            if entered_pin == str(st.secrets["APP_PIN"]):
                st.session_state["pin_unlocked"] = True
                st.rerun()
            else:
                st.error("Incorrect PIN.")
        st.stop()

# --- RESOLVE GEMINI API KEY (CLOUD SECRETS OR LOCAL FILE) ---
st.sidebar.header("🔑 App Settings")
saved_key = ""
if "GEMINI_API_KEY" in st.secrets:
    saved_key = str(st.secrets["GEMINI_API_KEY"]).strip()
    api_key_input = saved_key
    st.sidebar.success("Gemini API Key connected via Cloud Secrets!")
else:
    if os.path.exists(KEY_FILE):
        with open(KEY_FILE, "r") as f:
            saved_key = f.read().strip()

    api_key_input = st.sidebar.text_input(
        "Paste Gemini API Key",
        value=saved_key,
        type="password"
    )

    if api_key_input and api_key_input != saved_key:
        with open(KEY_FILE, "w") as f:
            f.write(api_key_input.strip())
        st.sidebar.success("API Key saved!")

st.sidebar.divider()
if st.sidebar.button("🔄 Force Sync from Google Sheets"):
    st.cache_data.clear()
    st.rerun()

# --- GOOGLE SHEETS CONNECTION & CACHING ---
@st.cache_resource
def get_gsheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if os.path.exists("key.json"):
        creds = Credentials.from_service_account_file("key.json", scopes=scopes)
    else:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)

@st.cache_resource
def init_db():
    client = get_gsheet_client()
    url = st.secrets.get("SHEET_URL", SHEET_URL) if hasattr(st, "secrets") else SHEET_URL
    sh = client.open_by_url(url)
    existing_tabs = [ws.title for ws in sh.worksheets()]

    schemas = {
        "Inventory": [
            "Item Name", "Item Type", "Vendor", "Purchase Date",
            "Purchase Qty", "Purchase Unit", "Total Cost ($)",
            "Cost per Gram ($)", "Cost per Oz ($)", "Cost per Piece ($)"
        ],
        "Recipes": [
            "Recipe Name", "Category", "Batch Yield", "Labor & Overhead ($)",
            "Batch Material Cost ($)", "Total Batch Cost ($)", "Cost Per Item ($)",
            "Suggested Wholesale ($)", "Suggested Retail ($)", "Ingredients JSON", "Created Date"
        ],
        "Batches": [
            "Lot Number", "Date Made", "Recipe Name", "Category",
            "Units Produced", "Total Batch Cost ($)", "Cost Per Unit ($)", "Notes"
        ],
    }

    for tab_name, headers in schemas.items():
        if tab_name not in existing_tabs:
            ws = sh.add_worksheet(title=tab_name, rows="500", cols=str(len(headers)))
            ws.append_row(headers)
        else:
            ws = sh.worksheet(tab_name)
            if not ws.row_values(1):
                ws.append_row(headers)
    return sh

@st.cache_data(ttl=300)
def load_tab(tab_name: str) -> pd.DataFrame:
    sh = init_db()
    ws = sh.worksheet(tab_name)
    return pd.DataFrame(ws.get_all_records())

def calc_unit_costs(qty: float, unit: str, total_cost: float):
    if qty <= 0:
        return 0.0, 0.0, 0.0
    if unit == "pcs":
        return 0.0, 0.0, round(total_cost / qty, 4)
    if unit == "g":
        total_grams = qty
    elif unit == "oz":
        total_grams = qty * GRAMS_PER_OZ
    elif unit == "lb":
        total_grams = qty * GRAMS_PER_LB
    elif unit == "kg":
        total_grams = qty * GRAMS_PER_KG
    else:
        total_grams = qty

    cost_per_g = total_cost / total_grams
    cost_per_oz = cost_per_g * GRAMS_PER_OZ
    return round(cost_per_g, 5), round(cost_per_oz, 4), 0.0

def batch_upsert_inventory(sh, rows_list: list[dict]):
    ws = sh.worksheet("Inventory")
    records = ws.get_all_records()
    name_to_idx = {
        str(rec["Item Name"]).strip().lower(): idx
        for idx, rec in enumerate(records, start=2)
    }

    new_rows = []
    for row_dict in rows_list:
        row_values = [
            row_dict["Item Name"], row_dict["Item Type"], row_dict["Vendor"],
            row_dict["Purchase Date"], row_dict["Purchase Qty"], row_dict["Purchase Unit"],
            row_dict["Total Cost ($)"], row_dict["Cost per Gram ($)"],
            row_dict["Cost per Oz ($)"], row_dict["Cost per Piece ($)"]
        ]
        key = row_dict["Item Name"].strip().lower()
        if key in name_to_idx:
            r_idx = name_to_idx[key]
            ws.update(values=[row_values], range_name=f"A{r_idx}:J{r_idx}")
        else:
            new_rows.append(row_values)

    if new_rows:
        ws.append_rows(new_rows)
    st.cache_data.clear()

class ReceiptLineItem(BaseModel):
    item_name: str = Field(description="Clean ingredient or packaging name")
    item_type: str = Field(description="Either 'Ingredient' or 'Packaging'")
    quantity: float = Field(description="Numeric total quantity purchased")
    unit: str = Field(description="Must be one of: 'g', 'oz', 'lb', 'kg', 'pcs'")
    line_total_cost: float = Field(description="Total price paid for this line item in USD including shipping/tax share")

class ExtractedReceipt(BaseModel):
    vendor: str = Field(description="Supplier or store name")
    purchase_date: str = Field(description="Purchase date in YYYY-MM-DD format")
    items: list[ReceiptLineItem]

def parse_receipt_with_gemini(uploaded_file, api_key):
    client = genai.Client(api_key=api_key)
    file_bytes = uploaded_file.getvalue()
    mime_type = uploaded_file.type
    prompt = (
        "Extract all soapmaking, bath bomb, skincare ingredients, and packaging items from this receipt/invoice. "
        "Convert fluid ounces to 'oz', pounds to 'lb', grams to 'g', kilograms to 'kg', and countable items to 'pcs'. "
        "Distribute any shipping and tax proportionally across line_total_cost."
    )
    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=[types.Part.from_bytes(data=file_bytes, mime_type=mime_type), prompt],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ExtractedReceipt,
        ),
    )
    return json.loads(response.text)

st.title("🧪 Toxic Bloom Apothecary — Studio Manager")

if not os.path.exists("key.json") and "gcp_service_account" not in st.secrets:
    st.warning("Credentials not found. Add `gcp_service_account` to Streamlit Cloud Secrets or place `key.json` locally.")
    st.stop()

try:
    sh = init_db()
    inv_df = load_tab("Inventory")
    recipes_df = load_tab("Recipes")
    batches_df = load_tab("Batches")
except Exception as e:
    st.error(f"Could not connect to Google Sheet. Details: {e}")
    st.stop()

tab_inv, tab_scan, tab_recipe, tab_batch = st.tabs([
    "📦 Ingredient & Packaging Library",
    "🧾 AI Receipt Scanner",
    "🧼 Recipe & Pricing Builder",
    "📋 Batch & Lot Logger"
])

with tab_inv:
    st.subheader("Current Inventory & Unit Costs")
    if not inv_df.empty:
        st.dataframe(inv_df, use_container_width=True, hide_index=True)
    else:
        st.info("Your inventory is empty. Add an item below or scan a receipt!")

    st.divider()
    st.subheader("➕ Manually Add or Update Item")
    with st.form("manual_inv_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        item_name = c1.text_input("Item Name (e.g., Beef Tallow, Citric Acid, Lip Tube)")
        item_type = c2.selectbox("Item Type", ["Ingredient", "Packaging"])
        vendor = c3.text_input("Vendor")

        c4, c5, c6, c7 = st.columns(4)
        p_qty = c4.number_input("Quantity Purchased", min_value=0.01, value=16.0, step=1.0)
        p_unit = c5.selectbox("Unit", ["oz", "lb", "g", "kg", "pcs"])
        p_cost = c6.number_input("Total Price Paid ($)", min_value=0.01, value=15.00, step=0.50)
        p_date = c7.date_input("Purchase Date", value=datetime.today())

        if st.form_submit_button("Save to Inventory"):
            if item_name.strip():
                c_g, c_oz, c_pcs = calc_unit_costs(p_qty, p_unit, p_cost)
                row_data = {
                    "Item Name": item_name.strip(),
                    "Item Type": item_type,
                    "Vendor": vendor.strip(),
                    "Purchase Date": str(p_date),
                    "Purchase Qty": p_qty,
                    "Purchase Unit": p_unit,
                    "Total Cost ($)": round(p_cost, 2),
                    "Cost per Gram ($)": c_g,
                    "Cost per Oz ($)": c_oz,
                    "Cost per Piece ($)": c_pcs,
                }
                batch_upsert_inventory(sh, [row_data])
                st.success(f"Saved {item_name}!")
                st.rerun()

with tab_scan:
    st.subheader("🧾 Upload Supplier Receipt or Invoice")
    if not api_key_input:
        st.warning("👈 Add your Gemini API Key in Secrets or the left sidebar to enable receipt scanning.")
    uploaded = st.file_uploader("Upload Receipt", type=["png", "jpg", "jpeg", "pdf"])

    if uploaded and api_key_input and st.button("✨ Extract Receipt with Gemini"):
        with st.spinner("Reading invoice and calculating unit weights..."):
            try:
                st.session_state["parsed_receipt"] = parse_receipt_with_gemini(uploaded, api_key_input.strip())
            except Exception as e:
                st.error(f"Receipt extraction failed: {e}")

    if "parsed_receipt" in st.session_state:
        data = st.session_state["parsed_receipt"]
        st.markdown(f"**Vendor:** `{data.get('vendor', 'Unknown')}` | **Date:** `{data.get('purchase_date', '')}`")
        edited_items = st.data_editor(
            pd.DataFrame(data["items"]),
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "unit": st.column_config.SelectboxColumn("Unit", options=["g", "oz", "lb", "kg", "pcs"], required=True),
                "item_type": st.column_config.SelectboxColumn("Type", options=["Ingredient", "Packaging"], required=True),
            }
        )
        if st.button("✅ Approve & Sync All Items to Google Sheets"):
            rows_to_sync = []
            for _, r in edited_items.iterrows():
                c_g, c_oz, c_pcs = calc_unit_costs(float(r["quantity"]), r["unit"], float(r["line_total_cost"]))
                rows_to_sync.append({
                    "Item Name": str(r["item_name"]).strip(),
                    "Item Type": r["item_type"],
                    "Vendor": data.get("vendor", ""),
                    "Purchase Date": data.get("purchase_date", str(datetime.today().date())),
                    "Purchase Qty": float(r["quantity"]),
                    "Purchase Unit": r["unit"],
                    "Total Cost ($)": round(float(r["line_total_cost"]), 2),
                    "Cost per Gram ($)": c_g,
                    "Cost per Oz ($)": c_oz,
                    "Cost per Piece ($)": c_pcs,
                })
            batch_upsert_inventory(sh, rows_to_sync)
            del st.session_state["parsed_receipt"]
            st.success("All receipt items synced to Inventory!")
            st.rerun()

with tab_recipe:
    st.subheader("🧼 Formulation & Batch Cost Calculator")

    if inv_df.empty:
        st.warning("Add ingredients in the Library or scan a receipt first!")
    else:
        if "draft_lines" not in st.session_state:
            st.session_state["draft_lines"] = []

        col_meta1, col_meta2, col_meta3, col_meta4 = st.columns([2, 1.5, 1, 1])
        r_name = col_meta1.text_input("Recipe Name")
        r_cat = col_meta2.selectbox("Product Category", CATEGORIES)
        r_yield = col_meta3.number_input("Batch Yield (Items Made)", min_value=1, value=10, step=1)
        r_labor = col_meta4.number_input("Batch Labor / Overhead ($)", min_value=0.0, value=10.0, step=1.0)

        st.markdown("#### Add Ingredients & Packaging to Batch")
        ac1, ac2, ac3, ac4 = st.columns([3, 1.2, 1.2, 1])
        selected_item = ac1.selectbox("Select Inventory Item", inv_df["Item Name"].tolist())
        item_row = inv_df[inv_df["Item Name"] == selected_item].iloc[0]

        is_piece = str(item_row["Purchase Unit"]) == "pcs"
        allowed_units = ["pcs"] if is_piece else ["g", "oz"]
        used_amt = ac2.number_input("Amount Used", min_value=0.01, value=1.0 if is_piece else 10.0, step=1.0)
        used_unit = ac3.selectbox("Measure Unit", allowed_units)

        ac4.write("")
        ac4.write("")
        if ac4.button("➕ Add Line", use_container_width=True):
            if used_unit == "g":
                line_cost = used_amt * float(item_row["Cost per Gram ($)"])
            elif used_unit == "oz":
                line_cost = used_amt * float(item_row["Cost per Oz ($)"])
            else:
                line_cost = used_amt * float(item_row["Cost per Piece ($)"])

            st.session_state["draft_lines"].append({
                "Item Name": selected_item,
                "Amount": used_amt,
                "Unit": used_unit,
                "Line Cost ($)": round(line_cost, 4)
            })

        if st.session_state["draft_lines"]:
            draft_df = pd.DataFrame(st.session_state["draft_lines"])
            st.dataframe(draft_df, use_container_width=True, hide_index=True)

            b_undo, b_clear, _ = st.columns([1.5, 1.5, 5])
            if b_undo.button("↩️ Remove Last Line"):
                st.session_state["draft_lines"].pop()
                st.rerun()
            if b_clear.button("🗑️ Clear Recipe Lines"):
                st.session_state["draft_lines"] = []
                st.rerun()

            material_cost = draft_df["Line Cost ($)"].sum()
            total_batch_cost = material_cost + r_labor
            cost_per_item = total_batch_cost / r_yield

            st.divider()
            m1, m2 = st.columns(2)
            ws_mult = m1.slider("Wholesale Multiplier (x Cost)", 1.5, 4.0, 2.0, 0.1)
            rt_mult = m2.slider("Retail Multiplier (x Cost)", 2.0, 6.0, 4.0, 0.1)

            sug_wholesale = cost_per_item * ws_mult
            sug_retail = cost_per_item * rt_mult

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Materials + Packaging", f"${material_cost:.2f}")
            k2.metric("Total Batch Cost", f"${total_batch_cost:.2f}")
            k3.metric("Cost Per Item", f"${cost_per_item:.2f}")
            k4.metric("Retail / Wholesale", f"${sug_retail:.2f} / ${sug_wholesale:.2f}")

            if st.button("💾 Save Master Recipe to Google Sheets"):
                if r_name.strip():
                    sh.worksheet("Recipes").append_row([
                        r_name.strip(), r_cat, r_yield, round(r_labor, 2),
                        round(material_cost, 2), round(total_batch_cost, 2),
                        round(cost_per_item, 2), round(sug_wholesale, 2),
                        round(sug_retail, 2), json.dumps(st.session_state["draft_lines"]),
                        str(datetime.today().date())
                    ])
                    st.cache_data.clear()
                    st.session_state["draft_lines"] = []
                    st.success(f"Saved recipe {r_name}!")
                    st.rerun()
                else:
                    st.warning("Enter a Recipe Name before saving.")

    if not recipes_df.empty:
        st.divider()
        st.subheader("📖 Saved Master Recipes")
        display_cols = [c for c in recipes_df.columns if c != "Ingredients JSON"]
        st.dataframe(recipes_df[display_cols], use_container_width=True, hide_index=True)

with tab_batch:
    st.subheader("📋 Production Batch Log")

    if recipes_df.empty:
        st.info("Save a recipe first to log production batches.")
    else:
        with st.form("log_batch_form", clear_on_submit=True):
            b1, b2, b3 = st.columns(3)
            chosen_rec = b1.selectbox("Select Master Recipe", recipes_df["Recipe Name"].tolist())
            rec_info = recipes_df[recipes_df["Recipe Name"] == chosen_rec].iloc[0]

            units_made = b2.number_input("Actual Units Produced", min_value=1, value=int(rec_info["Batch Yield"]))
            date_made = b3.date_input("Production Date", value=datetime.today())
            notes = st.text_input("Batch Notes")

            if st.form_submit_button("Log Production Batch"):
                cat_prefix = "".join([w[0] for w in str(rec_info["Category"]).split()]).upper()
                lot_num = f"TBA-{cat_prefix}-{date_made.strftime('%Y%m%d')}-{len(batches_df)+1:03d}"
                batch_cost = float(rec_info["Total Batch Cost ($)"])
                unit_cost = round(batch_cost / units_made, 2)

                sh.worksheet("Batches").append_row([
                    lot_num, str(date_made), chosen_rec, str(rec_info["Category"]),
                    units_made, batch_cost, unit_cost, notes
                ])
                st.cache_data.clear()
                st.success(f"Logged Batch {lot_num}!")
                st.rerun()

    if not batches_df.empty:
        st.dataframe(batches_df, use_container_width=True, hide_index=True)
