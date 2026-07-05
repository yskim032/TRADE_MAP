import pandas as pd

excel_file = r"C:\Users\user\Downloads\Port Codes (14 May 2026 05-09-41).xlsx"
df_excel = pd.read_excel(excel_file, sheet_name='ExportedData')
df_excel['Code'] = df_excel['Code'].dropna().astype(str).str.strip().str.upper()

# Let's count how many ports we have for each country
country_counts = df_excel['Country'].value_counts()
print("Country ports counts:")
print(country_counts.head(20))

# Let's identify the countries that might span multiple regions:
# UNITED STATES, CANADA, MEXICO, COLOMBIA, SAUDI ARABIA, SPAIN, FRANCE, ITALY, RUSSIA, INDIA, AUSTRALIA, EGYPT
multi_region_countries = [
    "UNITED STATES", "CANADA", "MEXICO", "COLOMBIA", "SAUDI ARABIA", 
    "SPAIN", "FRANCE", "ITALY", "RUSSIAN FEDERATION", "INDIA", "AUSTRALIA", "EGYPT"
]

print("\nPorts for multi-region countries in dataset:")
for c in multi_region_countries:
    subset = df_excel[df_excel['Country'] == c]
    print(f"\nCountry: {c} ({len(subset)} ports)")
    for idx, row in subset.iterrows():
        print(f"  {row['Code']} - {row['Name']}")
