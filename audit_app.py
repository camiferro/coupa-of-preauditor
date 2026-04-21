import streamlit as st
import pdfplumber
import pandas as pd
import re
import logging

# --- Configure Logging ---
# Suppress pdfminer graphical warnings
logging.getLogger("pdfminer").setLevel(logging.ERROR)

# Helper function to normalize strings for robust comparison
def normalize_name(name):
    return re.sub(r'[^a-z0-9]', '', str(name).lower())

# --- Data Extraction Engine ---
def extract_master_data(pdf_file):
    """Universal extractor for all OF types."""
    data = {"text": "", "fees": {}, "products": [], "yearly_schedule": {}, "msa_comment": "", "dates": {}, "billing_info": {}}
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            data["text"] += page.extract_text() + "\n"
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # Capture products and quantities from standard Coupa tables
                    # Account for "Included in the Above Total" and grouped rows due to PDF formatting
                    if row and len(row) >= 3:
                        row_str = str(row)
                        if "USD" in row_str or "Included" in row_str:
                            # Split by newline to handle multiple products merged into a single PDF cell
                            names = [n.strip() for n in str(row[0]).split('\n') if n.strip() and n.strip() != "Product Name"]
                            qtys = [q.strip() for q in str(row[2]).split('\n') if q.strip() and q.strip() != "Qty."]
                            
                            for i, name in enumerate(names):
                                qty_str = qtys[i] if i < len(qtys) else (qtys[-1] if qtys else "1")
                                # Extract numeric quantity
                                qty_num = re.sub(r'[^\d.]', '', qty_str)
                                qty_val = float(qty_num) if qty_num else 1.0
                                
                                data["products"].append({"name": name, "qty": qty_val, "qty_str": qty_str})

    # Clean text: remove newlines and normalize multiple spaces into a single space
    clean_text = re.sub(r'\s+', ' ', data["text"].replace('\n', ' '))

    # Date Extraction - Handle Coupa's multi-line layout formatting
    start_match = re.search(r"Subscription Start Date:\s*(.*?)\n", data["text"], re.IGNORECASE)
    end_match = re.search(r"Subscription End Date:\s*(.*?)\n", data["text"], re.IGNORECASE)
    
    start_val = start_match.group(1).strip() if start_match and start_match.group(1).strip() else ""
    end_val = end_match.group(1).strip() if end_match and end_match.group(1).strip() else ""
    
    # Fallback if standard regex fails due to line breaks
    if not start_val or not end_val:
        # Matches standard Coupa date formats (e.g. 31 Mar, 2026)
        dates_found = re.findall(r"(\d{1,2}\s+[A-Za-z]{3}\.?\s*,?\s*\d{4}|\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{2,4})", clean_text)
        if len(dates_found) >= 2:
            if not start_val: start_val = dates_found[0]
            if not end_val: end_val = dates_found[1]

    data["dates"] = {
        "start": start_val if start_val else None,
        "end": end_val if end_val else None
    }

    # Billing & Contact Information Extraction (Universal)
    ap_contact_match = re.search(r"Accounts Payable Contact:\s*(.*)", data["text"], re.IGNORECASE)
    ap_email_match = re.search(r"Accounts Payable Email:\s*(.*)", data["text"], re.IGNORECASE)
    
    ap_contact = ap_contact_match.group(1).strip() if ap_contact_match else ""
    ap_email = ap_email_match.group(1).strip() if ap_email_match else ""
    
    data["billing_info"] = {
        "ap_contact": ap_contact if ap_contact else None,
        "ap_email": ap_email if ap_email else None,
        "has_billing": bool(re.search(r"Customer Billing Information", data["text"], re.IGNORECASE)),
        "has_shipping": bool(re.search(r"Customer Shipping Information", data["text"], re.IGNORECASE))
    }

    # 1. MSA Logic (Strictly anchor to known boilerplate starts to avoid capturing product descriptions)
    msa_pattern = r"((?:The Coupa subscriptions ordered above|The subscriptions in this Order Form|This Order Form is).*?governed by.*?Privacy Terms[\"']?\.?)"
    msa_match = re.search(msa_pattern, clean_text, re.IGNORECASE)
    
    if msa_match:
        data["msa_comment"] = msa_match.group(1).strip()
    else:
        # Fallback if the paragraph doesn't end with "Privacy Terms" but starts correctly
        fallback_match = re.search(r"((?:The Coupa subscriptions ordered above|The subscriptions in this Order Form|This Order Form is).*?governed by.*?\.)", clean_text, re.IGNORECASE)
        data["msa_comment"] = fallback_match.group(1).strip() if fallback_match else ""

    # 2. Year-by-Year Schedule (For New Business / Renewals with Ramps)
    yearly_fees = re.findall(r"Total Year (\d+) Fee:\s*USD\s*([\d,.]+)", data["text"])
    data["yearly_schedule"] = {int(y): float(v.replace(',', '')) for y, v in yearly_fees}

    # 3. Flat Fees & Prorations (For Add-Ons / Baseline checks)
    fee_patterns = {
        "prorated": r"Total Year 1 Prorated Fee:\s*USD\s*([\d,.]+)",
        "total": r"Total Fee:\s*USD\s*([\d,.]+)",
        "annual": r"Annual Subscription Fee:\s*USD\s*([\d,.]+)"
    }
    for key, pattern in fee_patterns.items():
        match = re.search(pattern, data["text"])
        if match:
            data["fees"][key] = float(match.group(1).replace(',', ''))
            
    return data

def process_pasted_sfdc(text):
    """Parses messy Salesforce copy-paste text and converts it into tabular data."""
    parsed_rows = []
    sfdc_data = {"names": [], "products": {}, "latest_end_date": None}
    
    # Split text by SUB- to isolate each subscription line item
    blocks = re.split(r'(?=SUB-\d+)', text)
    
    for block in blocks:
        if not block.strip() or not block.startswith('SUB-'): 
            continue
            
        # Extract SUB ID
        sub_id_match = re.search(r'(SUB-\d+)', block)
        sub_id = sub_id_match.group(1) if sub_id_match else "N/A"
        
        # Extract QL
        ql_match = re.search(r'(QL-\d+)', block)
        quote_line = ql_match.group(1) if ql_match else "N/A"
        
        # Find everything after QL (or SUB if no QL)
        if quote_line != "N/A":
            payload = block.split(quote_line, 1)[-1]
        else:
            payload = block.split(sub_id, 1)[-1]
            
        # Clean payload: replace newlines with spaces to handle both squished and separated formats
        payload_clean = payload.replace('\n', ' ').strip()
        
        prod_name = "N/A"
        start_date = "N/A"
        end_date = "N/A"
        qty = 1.0
        net_price = "N/A"
        opportunity = "N/A"
        
        # Regex to untangle: ProductName StartDate EndDate Qty Currency NetPrice Opportunity
        # e.g. "P2P (Procurement + Invoicing) 3/31/2023 3/30/2026 150.00 USD 103,125.00 LaserAway - P2P, Smash"
        match = re.search(r'^(.*?)\s*(\d{1,2}/\d{1,2}/\d{2,4})\s*(\d{1,2}/\d{1,2}/\d{2,4})\s*([\d,.]+)\s*([A-Z]{3})\s*([\d,.]+)\s*(.*)$', payload_clean)
        
        if match:
            prod_name = match.group(1).strip()
            start_date = match.group(2).strip()
            end_date = match.group(3).strip()
            qty_str = match.group(4).strip()
            currency = match.group(5).strip()
            price_str = match.group(6).strip()
            opportunity = match.group(7).strip()
            
            try:
                qty = float(qty_str.replace(',', ''))
            except ValueError: pass
            
            net_price = f"{currency} {price_str}"
            sfdc_data["latest_end_date"] = end_date
        else:
            # Fallback: Just look for Dates and Qty if the line is missing Net Price/Opportunity
            fallback_match = re.search(r'^(.*?)\s*(\d{1,2}/\d{1,2}/\d{2,4})\s*(\d{1,2}/\d{1,2}/\d{2,4})\s*([\d,.]+)', payload_clean)
            if fallback_match:
                prod_name = fallback_match.group(1).strip()
                start_date = fallback_match.group(2).strip()
                end_date = fallback_match.group(3).strip()
                qty_str = fallback_match.group(4).strip()
                try:
                    qty = float(qty_str.replace(',', ''))
                except ValueError: pass
                sfdc_data["latest_end_date"] = end_date
            else:
                # Absolute fallback
                lines = [line.strip() for line in payload.strip().split('\n') if line.strip()]
                if lines:
                    prod_name = lines[0]
        
        parsed_rows.append({
            "Subscription #": sub_id,
            "QUOTE line": quote_line,
            "Product name": prod_name,
            "Start Date": start_date,
            "End Date": end_date,
            "Sub qty": qty,
            "Net Price": net_price,
            "Opportunity": opportunity
        })
        
        if prod_name and prod_name != "N/A":
            norm_name = normalize_name(prod_name)
            if norm_name not in sfdc_data["products"]:
                sfdc_data["products"][norm_name] = {"name": prod_name, "qty": 0.0}
                sfdc_data["names"].append(prod_name)
                
            sfdc_data["products"][norm_name]["qty"] += qty
            
    sfdc_data["parsed_table"] = parsed_rows
    return sfdc_data

# --- UI Layout & Configuration ---
st.set_page_config(page_title="FinOps Master Audit", layout="wide")
st.title("🛡️ FinOps Automation: OF Master Auditor")

# --- Session State Management ---
if 'app_key' not in st.session_state:
    st.session_state.app_key = 0
if 'run_audit' not in st.session_state:
    st.session_state.run_audit = False

def reset_app():
    """Callback to clear all inputs and hide audit findings."""
    st.session_state.app_key += 1
    st.session_state.run_audit = False

def trigger_audit():
    """Callback to unhide the audit findings."""
    st.session_state.run_audit = True

# Dynamic Sidebar Selector
opp_type = st.sidebar.selectbox(
    "Select Opportunity Type:", 
    ["Renewal", "Add-On (AO)", "New Business"],
    on_change=reset_app
)

col1, col2 = st.columns(2)

contract_end_date = None

with col1:
    st.header("Step 1: Reference Data")
    ref_data = None    
    sfdc_data = None   
    
    if opp_type == "Renewal":
        prev_file = st.file_uploader("Upload PREVIOUSLY Signed OF (PDF) [For MSA & Date Checks]", type="pdf", key=f"prev_file_{st.session_state.app_key}")
        if prev_file:
            ref_data = extract_master_data(prev_file)
            st.success("Previous OF processed successfully.")
            
        past_sub = st.text_area("Paste Salesforce Subscriptions (Raw Text):", height=150, help="Copy the raw text from the Salesforce Subscriptions list.", key=f"past_sub_ren_{st.session_state.app_key}")
        if past_sub:
            sfdc_data = process_pasted_sfdc(past_sub)
            if sfdc_data and sfdc_data["parsed_table"]:
                st.success("Data successfully parsed into tabular format:")
                st.dataframe(pd.DataFrame(sfdc_data["parsed_table"]), use_container_width=True)
            
    elif opp_type == "Add-On (AO)":
        past_sub = st.text_area("Paste Salesforce Subscriptions (Raw Text):", height=150, help="Copy the raw text from the Salesforce Subscriptions list.", key=f"past_sub_ao_{st.session_state.app_key}")
        if past_sub:
            sfdc_data = process_pasted_sfdc(past_sub)
            if sfdc_data and sfdc_data["parsed_table"]:
                st.success("Data successfully parsed into tabular format:")
                st.dataframe(pd.DataFrame(sfdc_data["parsed_table"]), use_container_width=True)
            
        auto_date = sfdc_data.get("latest_end_date", "") if sfdc_data else ""
        contract_end_date = st.text_input("Contract End Date (Auto-detected or enter manually):", value=auto_date, key=f"ao_date_{st.session_state.app_key}")
            
    else:
        st.info("New Business Mode: Auditing internal math, RPI ramps, and MSA boilerplate only.")

with col2:
    st.header("Step 2: Current OF")
    curr_file = st.file_uploader("Upload NEW OF to Audit (PDF)", type="pdf", key=f"curr_file_{st.session_state.app_key}")
    if curr_file:
        curr_data = extract_master_data(curr_file)

st.divider()

# --- Action Buttons ---
btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 2])
with btn_col1:
    st.button("🚀 Run Validation", use_container_width=True, on_click=trigger_audit, type="primary")
with btn_col2:
    st.button("🗑️ Clear Uploaded Data", use_container_width=True, on_click=reset_app)

# --- Master Audit Engine ---
if curr_file and st.session_state.run_audit:
    st.header("🚨 Audit Findings")

    # 1. 📅 Term & Date Validation
    st.subheader("📅 Term & Date Validation")
    curr_start = curr_data["dates"].get("start")
    curr_end = curr_data["dates"].get("end")
    
    col_d1, col_d2 = st.columns(2)
    with col_d1: st.write(f"**Current OF Start Date:** {curr_start or 'Not Found'}")
    with col_d2: st.write(f"**Current OF End Date:** {curr_end or 'Not Found'}")
    
    if opp_type == "Renewal" and ref_data:
        prev_end = ref_data["dates"].get("end")
        st.write(f"**Previous OF End Date:** {prev_end or 'Not Found'}")
        
        if curr_start and prev_end:
            try:
                prev_end_dt = pd.to_datetime(prev_end).date()
                curr_start_dt = pd.to_datetime(curr_start).date()
                delta = (curr_start_dt - prev_end_dt).days
                
                if delta == 1:
                    st.success(f"✅ **Continuous Term:** Renewal starts exactly 1 day after previous contract ends.")
                elif delta == 0:
                    st.warning(f"⚠️ **Date Overlap:** Renewal starts on the exact same day the previous contract ends ({prev_end}).")
                else:
                    st.error(f"❌ **DATE GAP ERROR:** There is a gap/overlap of {delta - 1} days between the old end date and new start date.")
            except Exception:
                st.info("ℹ️ Could not parse dates for automatic math. Please verify continuity manually.")

    elif opp_type == "Add-On (AO)":
        if contract_end_date and curr_end:
            try:
                curr_end_dt = pd.to_datetime(curr_end).date()
                contract_end_dt = pd.to_datetime(contract_end_date).date()
                
                if curr_end_dt == contract_end_dt:
                    st.success(f"✅ **Coterminous Alignment:** Add-On end date perfectly matches the SFDC Contract end date.")
                else:
                    st.error(f"❌ **COTERM ERROR:** Add-On ends on **{curr_end}**, but SFDC contract ends on **{contract_end_date}**.")
            except Exception:
                if str(curr_end).strip().lower() == str(contract_end_date).strip().lower():
                    st.success(f"✅ **Coterminous Alignment:** Add-On end date matches the SFDC Contract end date.")
                else:
                    st.error(f"❌ **COTERM ERROR:** Add-On ends on {curr_end}, but SFDC contract ends on {contract_end_date}.")
        else:
            st.warning("⚠️ **Missing Data:** Enter the Contract End Date from Salesforce above to validate coterminus dates.")

    st.divider()

    # 2. Universal Math Check: Prorated vs Total
    if "prorated" in curr_data["fees"] and "total" in curr_data["fees"]:
        if abs(curr_data["fees"]["prorated"] - curr_data["fees"]["total"]) > 0.01:
            diff = curr_data["fees"]["total"] - curr_data["fees"]["prorated"]
            st.error(f"❌ TOTAL FEE ERROR: Prorated and Total Fees mismatch by \${diff:,.2f}")

    # 3. NB & RENEWAL: Multi-Year Math & ACV Increase (Opp Amount)
    if curr_data["yearly_schedule"]:
        st.subheader("💰 YoY RPI & ACV Increase (Opp Amounts)")
        rows = []
        calc_total = 0
        prev_f = None
        
        for y, f in sorted(curr_data["yearly_schedule"].items()):
            calc_total += f
            rpi = ((f - prev_f) / prev_f * 100) if prev_f else 0
            acv_increase = (f - prev_f) if prev_f else 0 
            
            rows.append({
                "Year": f"Year {y}",
                "Fee (USD)": f"${f:,.2f}",
                "ACV Increase (Opp Amount)": f"${acv_increase:,.2f}" if prev_f else "Baseline",
                "YoY RPI %": f"{rpi:.2f}%" if prev_f else "N/A"
            })
            prev_f = f
            
        st.table(pd.DataFrame(rows))
        
        actual_total = curr_data["fees"].get("total", 0)
        if abs(calc_total - actual_total) > 0.01:
            st.error(f"❌ SUM ERROR: Sum of years is \${calc_total:,.2f}, but Total listed is \${actual_total:,.2f}")
        else:
            st.success(f"✅ Total Fee validated: \${actual_total:,.2f}")

    # Helper function to extract dicts for Foundational Checks & Storytelling
    curr_prods_dict = {}
    for p in curr_data.get('products', []):
        nk = normalize_name(p['name'])
        if nk not in curr_prods_dict:
            curr_prods_dict[nk] = {"name": p['name'], "qty": 0.0}
        curr_prods_dict[nk]["qty"] += p['qty']

    # 4. FOUNDATIONAL TIER CHECK (Support & Platform)
    if opp_type in ["Add-On (AO)", "Renewal"] and sfdc_data:
        st.subheader("🏗️ Foundational Tier Changes (Support & Platform)")
        sfdc_prods_dict = sfdc_data['products']
        
        def check_foundational(keyword):
            sfdc_match = {k: v['name'] for k, v in sfdc_prods_dict.items() if keyword in v['name'].lower()}
            curr_match = {k: v['name'] for k, v in curr_prods_dict.items() if keyword in v['name'].lower()}
            
            if not sfdc_match and not curr_match:
                st.info(f"ℹ️ No **{keyword.title()}** products found in either document.")
                return
                
            if set(sfdc_match.keys()) == set(curr_match.keys()):
                st.success(f"✅ **{keyword.title()} Unchanged:** {', '.join(set(curr_match.values()))}")
            else:
                old_str = ", ".join(set(sfdc_match.values())) if sfdc_match else "None (Not in SFDC paste)"
                new_str = ", ".join(set(curr_match.values())) if curr_match else "None (Missing from new OF)"
                st.error(f"⚠️ **{keyword.title().upper()} CHANGED:** Previous: `{old_str}` ➡️ New OF: `{new_str}`")
                
        check_foundational("support")
        check_foundational("platform")

    # 5. PRODUCT COMPARISONS
    if opp_type == "Add-On (AO)" and sfdc_data:
        st.subheader("📦 Product Expansion Validation")
        sfdc_prods_norm = set(sfdc_data['products'].keys())
        
        for prod in curr_data["products"]:
            norm_prod_name = normalize_name(prod['name'])
            # Skip support/platform as they are handled in foundational check
            if "support" in norm_prod_name or "platform" in norm_prod_name:
                continue
                
            if norm_prod_name in sfdc_prods_norm:
                st.info(f"✅ **{prod['name']}**: Existing SKU matched. Validated as expansion.")
            else:
                st.warning(f"⚠️ **{prod['name']}**: NEW MODULE detected. Confirm if this is authorized.")

    elif opp_type == "Renewal" and sfdc_data:
        st.subheader("📦 Subscription Comparison (Storytelling)")
        st.write("Comparing Sum of Active Contract (Salesforce) against the Current Renewal OF.")
        
        comparison_rows = []
        all_norm_keys = sorted(set(sfdc_prods_dict.keys()).union(set(curr_prods_dict.keys())))
        
        for norm_key in all_norm_keys:
            sfdc_prod = sfdc_prods_dict.get(norm_key)
            curr_prod = curr_prods_dict.get(norm_key)
            
            display_name = curr_prod['name'] if curr_prod else sfdc_prod['name']
            
            # Highlight support/platform in the table
            if "support" in display_name.lower() or "platform" in display_name.lower():
                display_name = f"🏗️ {display_name}"
            
            sfdc_qty = sfdc_prod['qty'] if sfdc_prod else 0.0
            curr_qty = curr_prod['qty'] if curr_prod else 0.0
            
            if sfdc_prod and curr_prod:
                diff = curr_qty - sfdc_qty
                if diff > 0:
                    status = f" 🟠 Increased by {diff:g}"
                elif diff < 0:
                    status = f"🟡 Decreased by {abs(diff):g}"
                else:
                    status = "🟢 Maintained (Same Qty)"
            elif not sfdc_prod and curr_prod:
                status = "🔵 Net New (Added)"
            elif sfdc_prod and not curr_prod:
                status = "🔴 Dropped (Removed)"
            else:
                status = "Unknown"
                
            comparison_rows.append({
                "Product Name": display_name,
                "SFDC Contract Qty": f"{sfdc_qty:g}" if sfdc_prod else "0",
                "Renewal OF Qty": f"{curr_qty:g}" if curr_prod else "0",
                "Status": status
            })
            
        st.table(pd.DataFrame(comparison_rows))

    # Helper for extracting MSA type classification
    def get_msa_type(comment_str):
        c_lower = comment_str.lower()
        if "agreed to by the parties" in c_lower or "between customer and coupa" in c_lower:
            return "Signed"
        elif "www.coupa.com" in c_lower or "online" in c_lower:
            return "Online"
        return "Unknown"

    # 6. MSA Audit: Comment 1 Governance Check
    st.subheader("⚖️ MSA Governance Check")
    if curr_data["msa_comment"]:
        st.info(f"**Current Comment:** {curr_data['msa_comment']}")
        
        curr_msa_type = get_msa_type(curr_data["msa_comment"])
        
        if curr_msa_type == "Signed":
            st.warning("⚠️ **SIGNED MSA DETECTED:** This document references a signed agreement. The document is missing, please look for it before proceeding.")
        elif curr_msa_type == "Online":
            st.success("✅ **ONLINE MSA DETECTED:** You can proceed with signature.")
        else:
            st.info("ℹ️ **MANUAL REVIEW:** Please read the comment to determine if it requires a signed MSA.")
        
        # Only compare the TYPE of the MSA (Signed vs Online), ignore minor text differences
        if opp_type == "Renewal" and ref_data and ref_data.get("msa_comment"):
            prev_msa_type = get_msa_type(ref_data["msa_comment"])
            
            if curr_msa_type != prev_msa_type:
                st.error(f"⚠️ **MSA TYPE CHANGED:** Previous OF was **{prev_msa_type}**, but Current OF is **{curr_msa_type}**!")
    else:
        st.error("❌ **MSA Comment:** Could not be automatically found in the document. Please verify manually.")

    st.divider()

    # 7. Billing, Shipping & AP Information Check
    st.subheader("🧾 Billing, Shipping & AP Information (Universal Check)")
    billing = curr_data.get("billing_info", {})
    
    col_b1, col_b2 = st.columns(2)
    
    with col_b1:
        if billing.get("ap_contact"):
            st.success(f"✅ **AP Contact:** {billing['ap_contact']}")
        else:
            st.error("❌ **AP Contact:** Missing from document")
            
        if billing.get("ap_email"):
            st.success(f"✅ **AP Email:** {billing['ap_email']}")
        else:
            st.error("❌ **AP Email:** Missing from document")
            
    with col_b2:
        if billing.get("has_billing"):
            st.success("✅ **Billing Information:** Block detected")
        else:
            st.error("❌ **Billing Information:** Block missing")
            
        if billing.get("has_shipping"):
            st.success("✅ **Shipping Information:** Block detected")
        else:
            st.error("❌ **Shipping Information:** Block missing")