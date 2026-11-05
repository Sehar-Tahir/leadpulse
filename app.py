import sys
import os
import asyncio

if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import urllib.parse
import re
from playwright.sync_api import sync_playwright
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(
    page_title="LeadPulse | B2B Intelligence Engine",
    page_icon="⚜️",
    layout="wide",
    initial_sidebar_state="expanded"
)

VAULT_FILE = "leadpulse_vault.csv"

# ---------------- LUXURY OBSIDIAN & GOLD UI/UX STYLING ----------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

h1, h2, h3, h4, .brand-title {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
}

.stApp {
    background: radial-gradient(circle at top right, #131824 0%, #07090E 60%);
    color: #F8FAFC;
}

section[data-testid="stSidebar"] {
    background-color: #0B0E17 !important;
    border-right: 1px solid rgba(212, 175, 55, 0.18);
}

.hero-card {
    background: linear-gradient(135deg, rgba(17, 21, 31, 0.95) 0%, rgba(11, 14, 23, 0.95) 100%);
    border: 1px solid rgba(212, 175, 55, 0.3);
    border-radius: 14px;
    padding: 26px 30px;
    margin-bottom: 24px;
    box-shadow: 0 12px 32px rgba(0, 0, 0, 0.45);
}

.brand-badge {
    display: inline-block;
    background: rgba(212, 175, 55, 0.12);
    color: #E5C158;
    border: 1px solid rgba(212, 175, 55, 0.35);
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1.5px;
    padding: 4px 10px;
    border-radius: 6px;
    text-transform: uppercase;
    margin-bottom: 8px;
}

.brand-title {
    font-size: 32px;
    font-weight: 800;
    margin: 0;
    background: linear-gradient(90deg, #FFFFFF 0%, #E5C158 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}

.brand-sub {
    color: #94A3B8;
    font-size: 14px;
    margin-top: 4px;
}

.query-bar {
    background: #11151F;
    border-left: 4px solid #D4AF37;
    border: 1px solid #1E2536;
    padding: 14px 18px;
    border-radius: 8px;
    font-size: 14px;
    color: #E2E8F0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.2);
}

.kpi-card {
    background: #11151F;
    border: 1px solid rgba(255, 255, 255, 0.07);
    border-top: 2px solid #D4AF37;
    border-radius: 12px;
    padding: 18px 20px;
    box-shadow: 0 6px 20px rgba(0,0,0,0.3);
}
.kpi-title {
    color: #94A3B8;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.8px;
}
.kpi-value {
    color: #F8FAFC;
    font-size: 26px;
    font-weight: 800;
    margin-top: 4px;
    font-family: 'Plus Jakarta Sans', sans-serif;
}
.kpi-sub {
    color: #E5C158;
    font-size: 11px;
    margin-top: 2px;
}

/* Footer Creator Card */
.footer-card {
    background: linear-gradient(135deg, rgba(17, 21, 31, 0.9) 0%, rgba(11, 14, 23, 0.9) 100%);
    border: 1px solid rgba(212, 175, 55, 0.25);
    border-radius: 12px;
    padding: 20px;
    text-align: center;
    margin-top: 50px;
    margin-bottom: 20px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.4);
}
.footer-name {
    color: #E5C158;
    font-weight: 800;
    font-size: 30px;
    letter-spacing: 0.5px;
    font-family: 'Plus Jakarta Sans', sans-serif;
}
.footer-role {
    color: #94A3B8;
    font-size: 16px;
    margin-top: 4px;
}

div.stButton > button:first-child {
    background: linear-gradient(135deg, #E5C158 0%, #D4AF37 50%, #AA7C11 100%);
    color: #07090E !important;
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-weight: 800;
    font-size: 14px;
    border: none;
    padding: 0.8rem 1.3rem;
    border-radius: 8px;
    width: 100%;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    box-shadow: 0 4px 20px rgba(212, 175, 55, 0.28);
    transition: all 0.25s ease;
}
div.stButton > button:first-child:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 25px rgba(212, 175, 55, 0.45);
}

div.stDownloadButton > button:first-child {
    background: #141A29;
    color: #E5C158 !important;
    border: 1px solid #D4AF37;
    font-weight: 700;
    padding: 0.7rem 1.1rem;
    border-radius: 8px;
    width: 100%;
    letter-spacing: 0.4px;
}

div[data-testid="stAlert"] {
    margin-top: 18px !important;
    margin-bottom: 12px !important;
    background-color: #11151F !important;
    border: 1px solid rgba(212, 175, 55, 0.35) !important;
    border-left: 4px solid #D4AF37 !important;
    border-radius: 10px !important;
    color: #F8FAFC !important;
}
</style>
""", unsafe_allow_html=True)

# ---------------- DATA DICTIONARIES (EXPANDED FOR DIGITAL ERA) ----------------
CATEGORIES = [
    # Real Estate & Construction
    "Real Estate Agency", "Property Dealer", "Construction Company", "Architect Firm", "Interior Designer",
    # Food & Hospitality
    "Restaurant", "Cafe & Coffee Shop", "Bakery & Sweets", "Boutique Hotel", "Catering Service",
    # Health & Beauty
    "Beauty Salon", "Hair Salon", "Bridal Makeup Studio", "Dental Clinic", "Aesthetic & Skin Clinic", "Physiotherapy Center", "Diagnostic Lab", "Pharmacy",
    # Fitness & Sports
    "Gym & Fitness Center", "CrossFit Box", "Sports Academy",
    # Retail & Ecommerce
    "Clothing Boutique", "Jewelry Store", "Furniture Showroom", "Optical Store", "Electronics Store", "E-commerce Brand",
    # Professional & Corporate Services
    "Digital Marketing Agency", "Software House", "Law Firm", "Tax Consultant", "Travel Agency", "Event Planner", "Marquee & Wedding Hall", "Photography Studio", "Car Dealership", "Auto Detailing Studio",
    # Education & Training
    "Private School", "Coaching Academy", "IELTS Institute", "Coding Bootcamp",
    # Energy & Industry
    "Solar Energy Company", "Veterinary Clinic", "Security Systems Company"
]

CITIES = {
    "Bahawalpur": [
        "All Areas", "Model Town A", "Model Town B", "Model Town C",
        "Satellite Town", "DHA Bahawalpur", "Commercial Area", "Cantt",
        "Farid Gate", "Dubai Chowk", "University Chowk", "Shadab Colony"
    ],
    "Islamabad": [
        "All Areas", "F-6 Markaz", "F-7 Markaz", "F-8 Markaz", "F-10 Markaz",
        "F-11 Markaz", "G-8 Markaz", "G-9 Markaz", "G-10 Markaz", "G-11 Markaz",
        "I-8 Markaz", "Blue Area", "DHA Phase 2", "Bahria Town"
    ],
    "Lahore": ["All Areas", "Gulberg", "DHA Phase 5", "DHA Phase 6", "Johar Town", "Model Town", "Bahria Town"],
    "Rawalpindi": ["All Areas", "Saddar", "Satellite Town", "Commercial Market", "Bahria Town Phase 4", "Bahria Town Phase 7"],
    "Multan": ["All Areas", "Gulgasht Colony", "Cantt", "Model Town", "Bosan Road", "DHA Multan"],
    "Faisalabad": ["All Areas", "D Ground", "Peoples Colony", "Madina Town", "Canal Road"],
    "Karachi": ["All Areas", "DHA", "Clifton", "Gulshan-e-Iqbal", "PECHS", "Tariq Road"]
}

SOCIAL_DOMAINS = ["facebook.com", "instagram.com", "fb.com", "wa.me", "tiktok.com", "linktr.ee", "youtube.com"]

# ---------------- VAULT & HISTORY FUNCTIONS ----------------
def load_vault():
    if os.path.exists(VAULT_FILE):
        try:
            return pd.read_csv(VAULT_FILE)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def save_to_vault(new_leads_list):
    if not new_leads_list:
        return
    new_df = pd.DataFrame(new_leads_list)
    existing_df = load_vault()
    if not existing_df.empty:
        combined = pd.concat([new_df, existing_df], ignore_index=True)
        combined = combined.drop_duplicates(subset=["Business Name", "Phone Number"], keep="first")
    else:
        combined = new_df
    combined.to_csv(VAULT_FILE, index=False)

def delete_batch_from_vault(batch_name):
    df = load_vault()
    if not df.empty and "Batch ID" in df.columns:
        filtered_df = df[df["Batch ID"] != batch_name]
        if filtered_df.empty:
            if os.path.exists(VAULT_FILE):
                os.remove(VAULT_FILE)
        else:
            filtered_df.to_csv(VAULT_FILE, index=False)

def get_existing_keys(skip_saved):
    if not skip_saved:
        return set(), set()
    df = load_vault()
    if df.empty:
        return set(), set()
    urls = set(df["Google Maps Link"].dropna().astype(str).tolist())
    names = set(df["Business Name"].dropna().astype(str).str.lower().str.strip().tolist())
    return urls, names

# ---------------- HELPER FUNCTIONS ----------------
def verify_no_website(website_url, allow_social_only):
    if not website_url:
        return True, "No Website Listed"
    if allow_social_only:
        for dom in SOCIAL_DOMAINS:
            if dom in website_url.lower():
                return True, f"Social Profile ({dom})"
    return False, "Has Active Website"

def score_lead_potential(rating, reviews, phone, web_status):
    pts = 0
    if reviews >= 25: pts += 40
    elif reviews >= 8: pts += 25
    elif reviews >= 2: pts += 10

    if rating >= 4.2: pts += 30
    elif rating >= 3.6: pts += 15

    if phone and ("03" in phone or "+92 3" in phone or "+923" in phone):
        pts += 30
    elif phone:
        pts += 15

    if pts >= 75:
        return "🔥 High Ticket (Web + Booking)"
    elif pts >= 45:
        return "⚜️ Strong Fit (Pitch Business Web)"
    return "❄️ Standard Opportunity"

def get_niche_value_prop(category):
    cat_lower = category.lower()
    if "real estate" in cat_lower or "property" in cat_lower:
        return "Ager aap ki apni website ho, to aap apne clients ko sari properties ki detailed listings, high-resolution photos aur direct inquiry form aik hi professional portal par dikha sakte hain, jis se clients bar bar phone par details poonchne ki bajaye direct online properties explore kar sakein."
    elif "restaurant" in cat_lower or "cafe" in cat_lower or "bakery" in cat_lower or "sweet" in cat_lower:
        return "Aaj kal log kisi bhi spot par jane se pehle ya home delivery ke liye online menu zaroor dekhte hain. Ager aap ki apni website ho, to aap ke customers digital menu dekh kar direct table reservation ya home delivery orders place kar sakein ge, jis se aap ki daily sales mazeed barh jaye gi."
    elif "salon" in cat_lower or "makeup" in cat_lower or "clinic" in cat_lower or "dental" in cat_lower:
        return "Salon ya clinic par clients hamesha pichhla kaam (portfolio), packages aur pricing dekhna chahte hain. Ager aap ki apni website ho, to clients aap ka work dekh kar apni pasand ki appointment khud online book kar sakein ge, aur aap ko bar bar WhatsApp par rates ya pictures bhejne ki zaroorat nahi paregi."
    elif "gym" in cat_lower or "fitness" in cat_lower:
        return "Gym join karne wale log sab se pehle packages, trainer ki details aur timetable check karte hain. Ager aap ki apni website ho, to members saari membership plans aur workout schedules dekh kar online sign-up ya inquiry generate kar sakein ge."
    elif "software" in cat_lower or "digital marketing" in cat_lower or "agency" in cat_lower:
        return "Aik tech ya digital agency ke liye online digital presence sab se bara trust factor hoti hai. Ager aap ki apni advanced website ho, to aap apne case studies aur client portfolios international ya local clients ko professional tarike se show kar sakein ge."
    elif "school" in cat_lower or "academy" in cat_lower or "institute" in cat_lower:
        return "Parents admission se pehle courses, fee structures aur faculty ki details online check karte hain. Ager aap ki apni website ho, to students aur parents online fee inquiries aur admissions ke liye direct apply kar sakein ge."
    else:
        return "Ager aap ki apni website ho, to aap ke local customers ko aap ki services, pricing aur contact details aik professional digital storefront par mil jayein gi, jis se aap ka brand trust mazeed strong ho ga."

def create_whatsapp_link(phone, biz_name, category, city, website_url, cta_choice):
    if not phone:
        return ""
    digits = "".join(filter(str.isdigit, phone))
    if digits.startswith("03") and len(digits) == 11:
        digits = "92" + digits[1:]
    elif digits.startswith("3") and len(digits) == 10:
        digits = "92" + digits
    elif not digits.startswith("923"):
        return ""

    value_prop = get_niche_value_prop(category)

    if cta_choice == "Option 1: Free Mockup Preview":
        cta_text = "Kya hum aap ke business ke liye aik chota sa free custom website mockup bana kar bhej sakte hain? Agar aap ko wo pasand aaye, tab hum aap ke liye complete website ready kar sakte hain."
    elif cta_choice == "Option 2: 10-Min Strategy Call":
        cta_text = "Kya hum is week sirf 10 minutes ki quick call par discuss kar sakte hain ke website hone se aap ko kitna faida ho sakta hai? Agar aap ko layout pasand aaye, tab hum complete website bana kar de sakte hain."
    else:
        cta_text = "Kya hum aap ke business ke liye aik chota sa demo layout design kar ke bhej sakte hain? Agar aap ko pasand aaye, tab hum complete professional website bana kar de sakte hain."

    msg = (
        f"Assalam-o-Alaikum {biz_name} Team!\n\n"
        f"Main ne Google Maps par {city} mein aap ka {category} profile dekha. "
        f"Aap ke customer reviews bohot achay hain, lekin aap ke profile par official website link mojood nahi hai. "
        f"{value_prop} "
        f"Hum local businesses ke liye professional websites design karte hain. Ye hamari website hai: {website_url}. "
        f"{cta_text}"
    )

    return f"https://wa.me/{digits}?text={urllib.parse.quote(msg)}"

def build_styled_excel(df):
    output = BytesIO()
    safe_df = df.fillna("").astype(str)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_df.to_excel(writer, index=False, sheet_name="LeadPulse_Verified")
        ws = writer.sheets["LeadPulse_Verified"]

        header_fill = PatternFill(start_color="11151F", end_color="11151F", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="D4AF37")
        cell_font = Font(name="Calibri", size=10, color="111111")
        thin_border = Border(
            left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0")
        )

        ws.row_dimensions[1].height = 28
        for col_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="left", vertical="center")
            
            col_series = safe_df[col_name]
            max_len = max(col_series.map(len).max() if not col_series.empty else 0, len(str(col_name))) + 4
            col_letter = ws.cell(row=1, column=col_idx).column_letter
            ws.column_dimensions[col_letter].width = min(max_len, 42)

        for row in ws.iter_rows(min_row=2, max_row=len(df) + 1, min_col=1, max_col=len(df.columns)):
            ws.row_dimensions[row[0].row].height = 22
            for cell in row:
                cell.font = cell_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

    return output.getvalue()

# ---------------- PLAYWRIGHT SCRAPER ENGINE ----------------
def extract_maps_leads(query, target_count, min_rating, min_reviews, require_phone, mobile_only, social_as_no_web, skip_saved, batch_label, website_url, cta_choice, status_ui, prog_bar):
    leads = []
    saved_urls, saved_names = get_existing_keys(skip_saved)
    skipped_dups = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        context = browser.new_context(
            locale="en-US",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        encoded_q = urllib.parse.quote(query)
        maps_url = f"https://www.google.com/maps/search/{encoded_q}?hl=en"
        status_ui.info(f"🌐 Connecting to Google Maps directory for `{query}`...")
        page.goto(maps_url, timeout=60000)

        try:
            page.wait_for_selector('div[role="feed"]', timeout=15000)
        except Exception:
            status_ui.error("Listing feed not found for this sector. Try setting Area to 'All Areas' in the sidebar.")
            browser.close()
            return []

        feed = page.locator('div[role="feed"]')
        scroll_cycles = max(10, int(target_count * 1.6))
        for _ in range(scroll_cycles):
            feed.evaluate("el => el.scrollBy(0, 2600)")
            page.wait_for_timeout(850)
            found_links = page.locator('a[href*="/maps/place/"]').count()
            status_ui.info(f"🔍 Scanning directory... Found {found_links} listings. Filtering website status...")
            if found_links >= (target_count * 3) + len(saved_names):
                break

        elements = page.locator('a[href*="/maps/place/"]').all()
        seen = set()
        place_urls = []
        for el in elements:
            href = el.get_attribute("href")
            if href and href not in seen:
                seen.add(href)
                place_urls.append(href)

        city_label = query.split(" in ")[-1].replace(", Pakistan", "")
        searched_cat = query.split(" in ")[0]

        for place_url in place_urls:
            if len(leads) >= target_count:
                break

            if skip_saved and place_url in saved_urls:
                skipped_dups += 1
                continue

            try:
                page.goto(place_url, timeout=30000)
                page.wait_for_selector("h1.DUwDvf", timeout=8000)
                page.wait_for_timeout(450)

                biz_name = page.locator("h1.DUwDvf").first.inner_text().strip()

                if skip_saved and biz_name.lower() in saved_names:
                    skipped_dups += 1
                    continue

                website_url_maps = ""
                web_el = page.locator('a[data-item-id="authority"]')
                if web_el.count() > 0:
                    website_url_maps = web_el.first.get_attribute("href") or ""

                is_no_web, web_badge = verify_no_website(website_url_maps, social_as_no_web)
                if not is_no_web:
                    continue

                rating = 0.0
                reviews = 0
                r_el = page.locator('div.F7nice span[aria-hidden="true"]')
                if r_el.count() > 0:
                    try:
                        rating = float(r_el.first.inner_text().strip().replace(",", "."))
                    except ValueError:
                        rating = 0.0

                rev_el = page.locator('div.F7nice span[aria-label*="review"]')
                if rev_el.count() > 0:
                    rev_str = rev_el.first.get_attribute("aria-label") or ""
                    nums = re.findall(r"[\d,]+", rev_str)
                    if nums:
                        reviews = int(nums[0].replace(",", ""))

                if rating < min_rating or reviews < min_reviews:
                    continue

                phone = ""
                p_el = page.locator('button[data-item-id^="phone:tel:"]')
                if p_el.count() > 0:
                    raw_p = p_el.first.get_attribute("data-item-id") or ""
                    phone = raw_p.replace("phone:tel:", "").strip()

                if require_phone and not phone:
                    continue

                if mobile_only:
                    clean_digits = "".join(filter(str.isdigit, phone))
                    if not (clean_digits.startswith("03") or clean_digits.startswith("923")):
                        continue

                primary_cat = searched_cat
                cat_el = page.locator('button[jsaction*="category"]')
                if cat_el.count() > 0:
                    primary_cat = cat_el.first.inner_text().strip()
                secondary_cat = f"{searched_cat} • {city_label}"

                address = "Not Listed"
                addr_el = page.locator('button[data-item-id="address"]')
                if addr_el.count() > 0:
                    raw_addr = addr_el.first.get_attribute("aria-label") or ""
                    address = raw_addr.replace("Address:", "").strip()

                leads.append({
                    "Batch ID": batch_label,
                    "Business Name": biz_name,
                    "Google Maps Link": place_url,
                    "Category": primary_cat,
                    "Secondary Category": secondary_cat,
                    "Website Status": web_badge,
                    "Service Offer Potential": score_lead_potential(rating, reviews, phone, web_badge),
                    "Rating": rating,
                    "Reviews Count": reviews,
                    "Phone Number": phone if phone else "Not Listed",
                    "WhatsApp Outreach": create_whatsapp_link(phone, biz_name, primary_cat, city_label, website_url, cta_choice),
                    "Social Link (If Any)": website_url_maps if website_url_maps else "None",
                    "Full Address": address,
                    "Target City": city_label
                })

                prog_bar.progress(min(len(leads) / target_count, 1.0))
                status_ui.success(
                    f"✅ Verified Lead [{len(leads)}/{target_count}]: {biz_name} - *{web_badge}* "
                    f"(Skipped {skipped_dups} duplicates)"
                )

            except Exception:
                continue

        browser.close()
    return leads

# ---------------- SIDEBAR CONTROLS ----------------
vault_df = load_vault()
vault_count = len(vault_df)

with st.sidebar:
    st.markdown(f"""
        <div style="padding: 6px 0 16px 0; border-bottom: 1px solid rgba(212,175,55,0.2); margin-bottom: 16px;">
            <div style="display: flex; align-items: center; gap: 8px;">
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#D4AF37" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon></svg>
                <span style="color:#D4AF37; font-weight:800; font-size:22px; letter-spacing:0.5px;">LeadPulse</span>
            </div>
            <div style="color:#94A3B8; font-size:12px; margin-top:2px;">B2B Intelligence Suite</div>
            <div style="margin-top:10px; font-size:11px; color:#34D399; font-weight:600; display:flex; align-items:center; gap:5px;">
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#34D399" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                Master Vault: {vault_count} Saved Leads
            </div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown("#### 💼 Outreach Configuration")
    website_url = st.text_input("Hamari Website Link", value="https://techaurcode.com")
    cta_choice = st.selectbox("WhatsApp CTA Option", [
        "Option 1: Free Mockup Preview",
        "Option 2: 10-Min Strategy Call",
        "Option 3: Direct Portfolio & Pricing Share"
    ])

    st.markdown("#### 📍 Market & Niche")
    city = st.selectbox("Target City", list(CITIES.keys()))
    area = st.selectbox("Sector / Area Filter", CITIES[city])
    category = st.selectbox("Industry Category", CATEGORIES)
    custom_niche = st.text_input("Custom Category Override", placeholder="e.g., Montessori School")

    st.markdown("#### 🎯 Volume & Strict Filters")
    lead_limit = st.slider("Leads to Generate (1 - 50)", min_value=1, max_value=50, value=15)
    skip_saved = st.checkbox("🛡️ Skip Already Extracted Leads", value=True)
    social_as_no_web = st.checkbox("Count Social Profiles as No Website", value=True)
    require_phone = st.checkbox("Must Have Phone Number", value=True)
    mobile_only = st.checkbox("Only Pakistani Mobile (03xx)", value=False)

    col_r, col_v = st.columns(2)
    with col_r:
        min_rating = st.number_input("Min Rating", 0.0, 5.0, 3.5, 0.5)
    with col_v:
        min_reviews = st.number_input("Min Reviews", 0, 500, 2, 1)

# ---------------- HERO HEADER ----------------
st.markdown("""
<div class="hero-card">
    <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap;">
        <div>
            <div class="brand-badge">⚜️ STRICT NO-WEBSITE VERIFICATION ENGINE</div>
            <div class="brand-title">LeadPulse Prospecting Suite</div>
            <div class="brand-sub">Automated Google Maps Discovery • High-Ticket Website & SEO Opportunities for Local Pakistani Markets</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

active_cat = custom_niche.strip() if custom_niche.strip() else category
area_part = "" if area == "All Areas" else f"{area}, "
search_query = f"{active_cat} in {area_part}{city}, Pakistan"

bar_col, btn_col = st.columns([3.2, 1.2])
with bar_col:
    st.markdown(f"""
    <div class="query-bar">
        🎯 <b>Target:</b> <span style="color:#E5C158;">{search_query}</span>  | 
        🛡️ <b>Condition:</b> <span style="color:#34D399;">No Website</span>  | 
        📊 <b>Limit:</b> {lead_limit} Leads
    </div>
    """, unsafe_allow_html=True)
with btn_col:
    run_engine = st.button("⚜️ Generate Leads")

# ---------------- EXECUTION ----------------
if run_engine:
    st.markdown("<div style='margin-top: 18px;'></div>", unsafe_allow_html=True)
    status_box = st.empty()
    progress = st.progress(0)

    short_batch_label = f"{active_cat} @ {city} ({datetime.now().strftime('%d %b, %H:%M')})"

    results = extract_maps_leads(
        query=search_query,
        target_count=lead_limit,
        min_rating=min_rating,
        min_reviews=min_reviews,
        require_phone=require_phone,
        mobile_only=mobile_only,
        social_as_no_web=social_as_no_web,
        skip_saved=skip_saved,
        batch_label=short_batch_label,
        website_url=website_url,
        cta_choice=cta_choice,
        status_ui=status_box,
        prog_bar=progress
    )

    if results:
        save_to_vault(results)
        st.session_state["lp_leads"] = results
        st.session_state["lp_tag"] = f"{active_cat}_{city}".replace(" ", "_")

# ---------------- TABS: CURRENT BATCH VS MASTER VAULT ----------------
st.markdown("<br>", unsafe_allow_html=True)
tab_current, tab_vault = st.tabs([
    "✨ Current Generated Batch",
    f"🗄️ Master Vault & History ({len(load_vault())} Saved Leads)"
])

with tab_current:
    if "lp_leads" in st.session_state and st.session_state["lp_leads"]:
        df = pd.DataFrame(st.session_state["lp_leads"])

        total_found = len(df)
        high_ticket = len(df[df["Service Offer Potential"].str.contains("High Ticket")])
        wa_ready = len(df[df["WhatsApp Outreach"] != ""])
        avg_rev = int(df["Reviews Count"].mean()) if total_found > 0 else 0

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f'<div class="kpi-card"><div class="kpi-title">Verified Leads</div><div class="kpi-value">{total_found}</div><div class="kpi-sub">100% No Custom Domain</div></div>', unsafe_allow_html=True)
        with k2:
            st.markdown(f'<div class="kpi-card"><div class="kpi-title">High Priority</div><div class="kpi-value">{high_ticket}</div><div class="kpi-sub">Strong Reviews + Active</div></div>', unsafe_allow_html=True)
        with k3:
            st.markdown(f'<div class="kpi-card"><div class="kpi-title">WhatsApp Ready</div><div class="kpi-value">{wa_ready}</div><div class="kpi-sub">Direct Pitch Link</div></div>', unsafe_allow_html=True)
        with k4:
            st.markdown(f'<div class="kpi-card"><div class="kpi-title">Avg Reviews</div><div class="kpi-value">{avg_rev}</div><div class="kpi-sub">Foot Traffic Metric</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        excel_bytes = build_styled_excel(df)
        time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"LeadPulse_{st.session_state['lp_tag']}_{time_str}.xlsx"

        dl_col1, _ = st.columns([2, 2])
        with dl_col1:
            st.download_button(
                label=f"📥 Download Current Batch Excel ({excel_filename})",
                data=excel_bytes,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_current_batch"
            )

        st.dataframe(
            df,
            column_config={
                "Google Maps Link": st.column_config.LinkColumn("Verify on Maps", display_text="Open Google Maps ↗"),
                "WhatsApp Outreach": st.column_config.LinkColumn("Direct Conversion", display_text="Send WhatsApp Pitch 💬"),
            },
            width="stretch",
            hide_index=True
        )
    else:
        st.markdown("""
        <div style="background:#11151F; border:1px solid rgba(212,175,55,0.25); padding:40px; border-radius:12px; text-align:center; margin-top:12px; box-shadow:0 8px 25px rgba(0,0,0,0.3);">
            <div style="font-size:18px; font-weight:700; color:#F8FAFC; margin-bottom:8px;">No Active Leads Generated Yet</div>
            <div style="font-size:13px; color:#94A3B8; max-width:450px; margin:0 auto; line-height:1.5;">Select your target city, sector, and industry category from the sidebar, then click the <b>'Generate Leads'</b> button above to start prospecting.</div>
        </div>
        """, unsafe_allow_html=True)

with tab_vault:
    master_df = load_vault()
    if not master_df.empty:
        batch_options = list(master_df["Batch ID"].dropna().unique())
        
        f_col1, f_col2, f_col3 = st.columns([2.0, 1.5, 1.5])
        with f_col1:
            selected_batch = st.selectbox("📂 Select Search Batch:", ["All Saved Batches (Complete History)"] + batch_options)
        
        with f_col2:
            st.markdown("<div style='margin-top: 26px;'></div>", unsafe_allow_html=True)
            batch_to_delete = st.selectbox("🗑️ Select Batch to Delete:", ["-- Choose Batch to Delete --"] + batch_options)
            if batch_to_delete != "-- Choose Batch to Delete --":
                if st.button("⚠️ Confirm Delete Batch"):
                    delete_batch_from_vault(batch_to_delete)
                    st.success(f"Batch '{batch_to_delete}' deleted successfully! Ab yeh leads dobara generate ki ja sakti hain.")
                    st.rerun()

        if selected_batch == "All Saved Batches (Complete History)":
            filtered_vault = master_df
        else:
            filtered_vault = master_df[master_df["Batch ID"] == selected_batch]

        vault_excel = build_styled_excel(filtered_vault)
        v_time = datetime.now().strftime("%Y%m%d_%H%M")
        v_filename = f"LeadPulse_Vault_{v_time}.xlsx"

        with f_col3:
            st.markdown("<div style='margin-top: 26px;'></div>", unsafe_allow_html=True)
            st.download_button(
                label=f"📥 Export Selected History ({len(filtered_vault)} Leads)",
                data=vault_excel,
                file_name=v_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_vault_batch"
            )

        st.dataframe(
            filtered_vault,
            column_config={
                "Google Maps Link": st.column_config.LinkColumn("Verify on Maps", display_text="Open Google Maps ↗"),
                "WhatsApp Outreach": st.column_config.LinkColumn("Direct Conversion", display_text="Send WhatsApp Pitch 💬"),
            },
            width="stretch",
            hide_index=True
        )
    else:
        st.markdown("""
        <div style="background:#11151F; border:1px solid rgba(212,175,55,0.25); padding:40px; border-radius:12px; text-align:center; margin-top:12px; box-shadow:0 8px 25px rgba(0,0,0,0.3);">
            <div style="font-size:18px; font-weight:700; color:#F8FAFC; margin-bottom:8px;">Master Vault is Currently Empty</div>
            <div style="font-size:13px; color:#94A3B8; max-width:450px; margin:0 auto; line-height:1.5;">Every batch of leads you generate will be automatically saved here for lifetime tracking, deduplication, and export.</div>
        </div>
        """, unsafe_allow_html=True)

# ---------------- LUXURY CREATOR FOOTER ----------------
st.markdown("""
<div class="footer-card">
    <div style="color: #94A3B8; font-size: 11px; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px;">Engineered & Developed By</div>
    <div class="footer-name">Sehar Tahir</div>
    <div class="footer-role">Full Stack Web Developer | GoHighLevel (GHL) CRM Automation Specialist</div>
</div>
""", unsafe_allow_html=True)