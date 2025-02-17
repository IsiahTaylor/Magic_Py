import pandas as pd
import requests
import time

# File path for the Excel sheet
EXCEL_FILE = "mtg_collection.xlsx"
CARD_SHEET = "Cards"
PRICE_SHEET = "Prices"

# Base URLs for Scryfall API
SCRYFALL_API_URL = "https://api.scryfall.com/cards/named"
SCRYFALL_SEARCH_URL = "https://api.scryfall.com/cards/search"

# Read Excel file
df = pd.read_excel(EXCEL_FILE, sheet_name=CARD_SHEET, usecols="A:C", dtype=str)
df.columns = ["Name", "Set", "Quantity"]

# Function to get price from Scryfall
def get_card_price(name, set_code=None):
    name = name.strip()  # Remove extra spaces
    params = {"exact": name}  # Default to exact name search
    if set_code:
        params["set"] = set_code.lower().replace(" ", "")

    # 1️⃣ Try exact match first
    response = requests.get(SCRYFALL_API_URL, params=params)
    
    print(f"Attempting exact match for '{name}' ({set_code if set_code else 'Any Set'})")
    print(f"Response: {response.status_code} - {response.json() if response.status_code != 200 else 'Success'}")

    if response.status_code == 200:
        data = response.json()
        return data.get("prices", {}).get("usd", "N/A")  # USD price

    # 2️⃣ If exact match fails, try fuzzy search
    print(f"Exact match failed for '{name}', trying fuzzy search...")
    fuzzy_response = requests.get(f"{SCRYFALL_SEARCH_URL}?q={name.replace(' ', '+')}")
    
    print(f"Fuzzy search response: {fuzzy_response.status_code}")
    
    if fuzzy_response.status_code == 200:
        fuzzy_data = fuzzy_response.json().get("data", [])
        if fuzzy_data:
            print(f"Fuzzy search found '{fuzzy_data[0]['name']}' for '{name}', using its price.")
            return fuzzy_data[0].get("prices", {}).get("usd", "N/A")

    print(f"No results found for '{name}'.")
    return "Not Found"

# Fetch prices
price_data = []
for _, row in df.iterrows():
    card_name, set_name = row["Name"], row["Set"]

    # Format set code properly
    set_code = set_name.strip().lower().replace(" ", "") if pd.notna(set_name) else None

    price = get_card_price(card_name, set_code)
    price_data.append([card_name, price])

    # Respect API rate limits
    time.sleep(0.1)  # Wait 100ms between requests

# Convert to DataFrame
df_prices = pd.DataFrame(price_data, columns=["Name", "Price"])

# Ensure the Excel file is closed before proceeding
print("Ensure the Excel file is closed before proceeding.")

# Save prices to a new sheet in the same Excel file
with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_prices.to_excel(writer, sheet_name=PRICE_SHEET, index=False)

print("Updated prices written to the Excel file.")
input("\nPress Enter to exit...")

