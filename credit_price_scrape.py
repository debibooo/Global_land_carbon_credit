from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
import re
from rapidfuzz import process, fuzz


def clean_price(text):
    match = re.search(r"\d+\.\d+", text)
    return float(match.group()) if match else None


def scrape_all_pages_selenium():
    driver = webdriver.Chrome()
    results = []
    page = 1

    try:
        while True:
            if page == 1:
                url = "https://www.carbmetric.com/projects"
            else:
                url = f"https://www.carbmetric.com/projects?page={page}"

            print(f"Scraping page {page}: {url}")
            driver.get(url)
            time.sleep(5)  # 等 JS 加载

            cards = driver.find_elements(By.CLASS_NAME, "project-card")

            if len(cards) == 0:
                print(f"No projects found on page {page}, stop scraping.")
                break

            print(f"Found {len(cards)} projects on page {page}")

            for card in cards:
                title = card.find_element(By.CLASS_NAME, "project-card-title").text

                price = None
                items = card.find_elements(By.CLASS_NAME, "info-item")

                for item in items:
                    label = item.find_element(By.CLASS_NAME, "info-label").text.lower()
                    value = item.find_element(By.CLASS_NAME, "info-value").text

                    if "price" in label:
                        price = clean_price(value)
                        break

                results.append({
                    "Project Name": title,
                    "Credit Price": price
                })

            page += 1

    finally:
        driver.quit()

    return pd.DataFrame(results, columns=["Project Name", "Credit Price"])


def match_projects(scraped_df, excel_df):
    base_names = excel_df["Project Name"].dropna().tolist()
    matches = []

    for _, row in scraped_df.iterrows():
        name = row["Project Name"]

        result = process.extractOne(
            name,
            base_names,
            scorer=fuzz.token_sort_ratio
        )

        if result:
            match, score, idx = result

            if score > 85:
                matches.append({
                    "Project Name": name,
                    "matched_name": match,
                    "match_score": score,
                    "Credit Price": row["Credit Price"]
                })

    return pd.DataFrame(matches, columns=[
        "Project Name", "matched_name", "match_score", "Credit Price"
    ])


def update_excel(excel_df, matched_df):
    if matched_df.empty:
        excel_df = excel_df.copy()
        excel_df["matched_name"] = pd.NA
        excel_df["match_score"] = pd.NA
        excel_df["Credit Price"] = pd.NA
        return excel_df

    merged = excel_df.merge(
        matched_df,
        left_on="Project Name",
        right_on="matched_name",
        how="left"
    )

    return merged


excel_df = pd.read_excel(
    "/Users/shirley/Downloads/Voluntary-Registry-Offsets-Database--v2025-12-year-end.xlsx",
    sheet_name="PROJECTS",
    header=3
)

scraped_df = scrape_all_pages_selenium()
matched_df = match_projects(scraped_df, excel_df)
final_df = update_excel(excel_df, matched_df)

final_df.to_excel(
    "/Users/shirley/Downloads/Voluntary-Registry-Offsets-Database--v2025-12-year-end-with-price.xlsx",
    index=False
)

print("✅ Done! Voluntary Projects With Price.xlsx has been created.")