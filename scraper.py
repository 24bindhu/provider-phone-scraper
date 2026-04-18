from playwright.sync_api import sync_playwright
import pandas as pd
import re

INPUT_FILE = "providers.xlsx"
OUTPUT_FILE = "providers_with_phones.xlsx"

SEARCH_URL = "https://vadbhdsv7prod.glsuite.us/GLSuiteWeb/Clients/VADBHDS/Public/ProviderSearch/ProviderSearchSearch.aspx"

# Load Excel
df = pd.read_excel(INPUT_FILE, engine="openpyxl")

# Ensure correct column name
PHONE_COL = "Phone number"

def extract_phone(page):
    html = page.content()
    match = re.search(r"\(\d{3}\)\s*\d{3}-\d{4}", html)
    return match.group(0) if match else ""

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, slow_mo=100)
    page = browser.new_page()

    # =========================
    # STEP 1: Go to search page
    # =========================
    page.goto(SEARCH_URL)
    page.wait_for_load_state("networkidle")

    # =========================
    # STEP 2: Select Active
    # =========================
    page.select_option("#ContentPlaceHolder1_ddlStatus", label="Active")

    # =========================
    # STEP 3: Submit
    # =========================
    page.click("input[value='Submit']")
    page.wait_for_timeout(5000)

    # =========================
    # STEP 4: Wait for results
    # =========================
    page.wait_for_selector("input.linkButton")

    # =========================
    # STEP 5: Loop through Excel rows
    # =========================
    for i, row in df.iterrows():
        provider_name = str(row["Provider Name"]).strip()

        try:
            print(f"\nProcessing {i+1}: {provider_name}")

            providers = page.locator("input.linkButton")
            count = providers.count()

            found = False

            for j in range(count):
                btn = providers.nth(j)
                name = (btn.get_attribute("value") or "").strip()

                if provider_name.lower() == name.lower():

                    print("Matched:", name)

                    btn.click()
                    page.wait_for_timeout(2000)

                    phone = extract_phone(page)

                    # ✅ WRITE TO EXISTING COLUMN
                    df.at[i, PHONE_COL] = phone

                    print("Phone:", phone)

                    # go back
                    page.go_back()
                    page.wait_for_timeout(3000)

                    found = True
                    break

            if not found:
                print("Not found")
                df.at[i, PHONE_COL] = ""

        except Exception as e:
            print("Error:", provider_name, e)
            df.at[i, PHONE_COL] = ""

            # recovery (go back to results page)
            try:
                page.goto(SEARCH_URL)
                page.wait_for_timeout(3000)
                page.select_option("#ContentPlaceHolder1_ddlStatus", label="Active")
                page.click("input[value='Submit']")
                page.wait_for_timeout(5000)
            except:
                pass

    browser.close()

# =========================
# SAVE FINAL FILE
# =========================
df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")

print("\n✅ DONE! File saved as:", OUTPUT_FILE)
