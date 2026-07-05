import pandas as pd
import numpy as np
import json
import os

excel_file = r"C:\Users\user\Downloads\Port Codes (14 May 2026 05-09-41).xlsx"
df_excel = pd.read_excel(excel_file, sheet_name='ExportedData')
df_excel['Code'] = df_excel['Code'].dropna().astype(str).str.strip().str.upper()
df_excel['Name'] = df_excel['Name'].dropna().astype(str).str.strip()
df_excel['Country'] = df_excel['Country'].dropna().astype(str).str.strip().str.upper()
df_excel['OpenCommerce'] = df_excel['Open for commerce'].fillna('').astype(str).str.strip()
df_excel['OpenCargo'] = df_excel['Open for cargo operation'].fillna('').astype(str).str.strip() if 'Open for cargo operation' in df_excel.columns else df_excel['Open for Cargo operation'].fillna('').astype(str).str.strip()

# Reference coordinates data
ref_file = "code-list-improved.csv"
ref_df = pd.read_csv(ref_file, encoding='utf-8')
ref_df['RefCode'] = ref_df['Country'].astype(str).str.strip().str.upper() + ref_df['Location'].astype(str).str.strip().str.upper()

# Create lookup for coordinates
coord_lookup = {}
for idx, row in ref_df.iterrows():
    code = row['RefCode']
    coords_dec = row['CoordinatesDecimal']
    if pd.notna(coords_dec) and str(coords_dec).strip() != '':
        try:
            lat, lon = [float(x.strip()) for x in str(coords_dec).split(',')]
            coord_lookup[code] = (lat, lon)
        except Exception:
            pass

# Manual coordinate additions for the 4 missing ports
coord_lookup['ALMBM'] = (41.3667, 19.4289)
coord_lookup['IDPMB'] = (-6.244, 107.901)
coord_lookup['JPSTS'] = (31.8133, 130.3039)
coord_lookup['QAMES'] = (24.9880, 51.5450)

# Define country to region mapping
country_regions = {
    "ALBANIA": "ADR/CEMED",
    "ALGERIA": "NAF",
    "AMERICAN SAMOA": "Fiji",
    "ANGOLA": "West Africa",
    "ARGENTINA": "SAEC",
    "BAHAMAS": "Caribbean",
    "BAHRAIN": "Middle East",
    "BANGLADESH": "South India+Sri Lanka",
    "BARBADOS": "Caribbean",
    "BELGIUM": "NWC",
    "BELIZE": "Central America",
    "BENIN": "West Africa",
    "BRAZIL": "SAEC",
    "BRUNEI DARUSSALAM": "Intra Asia",
    "BULGARIA": "B/SEA",
    "CAMBODIA": "Intra Asia",
    "CAMEROON": "West Africa",
    "CHILE": "SAWC",
    "CHINA": "Intra Asia",
    "COMOROS": "East Africa",
    "CONGO": "West Africa",
    "CONGO, THE DEMOCRATIC REPUBLIC OF THE": "West Africa",
    "COOK ISLANDS": "New Zealand",
    "COSTA RICA": "Central America",
    "COTE D'IVOIRE": "West Africa",
    "CROATIA": "ADR/CEMED",
    "CYPRUS": "EMED",
    "DENMARK": "SCAN",
    "DJIBOUTI": "East Africa",
    "DOMINICA": "Caribbean",
    "DOMINICAN REPUBLIC": "Caribbean",
    "ECUADOR": "SAWC",
    "EL SALVADOR": "Central America",
    "FIJI": "Fiji",
    "FINLAND": "SCAN",
    "FRENCH POLYNESIA": "Noumea",
    "GABON": "West Africa",
    "GAMBIA": "West Africa",
    "GEORGIA": "B/SEA",
    "GERMANY": "NWC",
    "GREECE": "EMED",
    "GRENADA": "Caribbean",
    "GUATEMALA": "Central America",
    "GUINEA": "West Africa",
    "GUINEA-BISSAU": "West Africa",
    "GUYANA": "Caribbean",
    "HAITI": "Caribbean",
    "HONDURAS": "Central America",
    "HONG KONG": "Intra Asia",
    "ICELAND": "SCAN",
    "INDONESIA": "Intra Asia",
    "IRELAND": "NWC",
    "ISRAEL": "EMED",
    "JAMAICA": "Caribbean",
    "JAPAN": "Intra Asia",
    "JORDAN": "Red Sea",
    "KENYA": "East Africa",
    "KOREA, REPUBLIC OF": "Intra Asia",
    "KUWAIT": "Middle East",
    "LATVIA": "SCAN",
    "LEBANON": "EMED",
    "LIBERIA": "West Africa",
    "LIBYA": "NAF",
    "LITHUANIA": "SCAN",
    "MADAGASCAR": "East Africa",
    "MALAYSIA": "Intra Asia",
    "MALDIVES": "Indian Ocean",
    "MALTA": "WMED",
    "MAURITANIA": "West Africa",
    "MAURITIUS": "Indian Ocean",
    "MAYOTTE": "East Africa",
    "MONTENEGRO": "ADR/CEMED",
    "MOROCCO": "NAF",
    "MOZAMBIQUE": "East Africa",
    "NETHERLANDS": "NWC",
    "NEW CALEDONIA": "Noumea",
    "NEW ZEALAND": "New Zealand",
    "NICARAGUA": "Central America",
    "NIGERIA": "West Africa",
    "NORWAY": "SCAN",
    "OMAN": "Middle East",
    "PAKISTAN": "North India+Pakistan",
    "PANAMA": "Central America",
    "PARAGUAY": "SAEC",
    "PERU": "SAWC",
    "PHILIPPINES": "Intra Asia",
    "POLAND": "SCAN",
    "PORTUGAL": "Por+Biscay",
    "PUERTO RICO": "Caribbean",
    "QATAR": "Middle East",
    "REUNION": "Indian Ocean",
    "ROMANIA": "B/SEA",
    "SAINT LUCIA": "Caribbean",
    "SAINT VINCENT AND THE GRENADINES": "Caribbean",
    "SAMOA": "Fiji",
    "SENEGAL": "West Africa",
    "SIERRA LEONE": "West Africa",
    "SINGAPORE": "Intra Asia",
    "SINT MAARTEN (DUTCH PART)": "Caribbean",
    "SLOVENIA": "ADR/CEMED",
    "SOLOMON ISLANDS": "Fiji",
    "SOMALIA": "East Africa",
    "SOUTH AFRICA": "South Africa",
    "SRI LANKA": "South India+Sri Lanka",
    "SUDAN": "Red Sea",
    "SURINAME": "Caribbean",
    "SWEDEN": "SCAN",
    "SYRIAN ARAB REPUBLIC": "EMED",
    "TAIWAN": "Intra Asia",
    "TANZANIA, UNITED REPUBLIC OF": "East Africa",
    "THAILAND": "Intra Asia",
    "TIMOR-LESTE": "Intra Asia",
    "TOGO": "West Africa",
    "TONGA": "Fiji",
    "TRINIDAD AND TOBAGO": "Caribbean",
    "TURKIYE": "EMED",
    "TURKS AND CAICOS ISLANDS": "Caribbean",
    "UKRAINE": "B/SEA",
    "UNITED ARAB EMIRATES": "Middle East",
    "UNITED KINGDOM": "NWC",
    "URUGUAY": "SAEC",
    "VENEZUELA": "Caribbean",
    "VIET NAM": "Intra Asia",
    "YEMEN": "Red Sea",
    "CAPE VERDE": "West Africa"
}

# Override function for cities/ports in multi-region countries
def get_region(row):
    code = row['Code']
    country = row['Country']
    name = row['Name'].upper()
    
    # 1. UNITED STATES
    if country == "UNITED STATES":
        if code in ["USLAX", "USLGB", "USOAK", "USPDX", "USSEA", "USNTD"] or "LOS ANGELES" in name or "LONG BEACH" in name or "OAKLAND" in name or "PORTLAND" in name or "SEATTLE" in name or "HUENEME" in name:
            return "USWC"
        else:
            return "USEC"
            
    # 2. CANADA
    if country == "CANADA":
        if code in ["CAPRR"] or "PRINCE RUPERT" in name or "VANCOUVER" in name:
            return "Canada WC"
        else:
            return "Canada EC"
            
    # 3. MEXICO
    if country == "MEXICO":
        if code in ["MXATM", "MXPGO", "MXPMS"] or "ALTAMIRA" in name or "PROGRESO" in name or "MORELOS" in name:
            return "MEXICO EC"
        else:
            return "MEXICO WC"
            
    # 4. COLOMBIA
    if country == "COLOMBIA":
        if code in ["COBUN"] or "BUENAVENTURA" in name:
            return "SAWC"
        else:
            return "Caribbean"
            
    # 5. SAUDI ARABIA
    if country == "SAUDI ARABIA":
        if code in ["SAJED", "SAKAC", "SANEO"] or "JEDDAH" in name or "KING ABDULLAH" in name or "NEOM" in name:
            return "Red Sea"
        else:
            return "Middle East"
            
    # 6. SPAIN
    if country == "SPAIN":
        if code in ["ESBIO", "ESGIJ"] or "BILBAO" in name or "GIJON" in name:
            return "Por+Biscay"
        elif code in ["ESACE", "ESFUE", "ESLPA", "ESSCT"] or "LAS PALMAS" in name or "TENERIFE" in name or "ARRECIFE" in name or "FUERTEVENTURA" in name:
            return "West Africa"  # Canary Islands grouped under West Africa
        else:
            return "WMED"
            
    # 7. FRANCE
    if country == "FRANCE":
        if code in ["FRFOS", "FRSET"] or "FOS" in name or "SETE" in name:
            return "WMED"
        else:
            return "NWC"
            
    # 8. ITALY
    if country == "ITALY":
        if code in ["ITAOI", "ITBRI", "ITRAN"] or "ANCONA" in name or "BARI" in name or "RAVENNA" in name:
            return "ADR/CEMED"
        else:
            return "WMED"
            
    # 9. RUSSIAN FEDERATION
    if country == "RUSSIAN FEDERATION":
        if code in ["RULED"] or "PETERSBURG" in name:
            return "SCAN"
        elif code in ["RUNVS"] or "NOVOROSSIYSK" in name:
            return "B/SEA"
        else:
            return "Intra Asia"
            
    # 10. INDIA
    if country == "INDIA":
        if code in ["INHZA", "INMRM", "INMUN", "INNML", "INNSA", "INPAV"] or "HAZIRA" in name or "MUNDRA" in name or "NHAVA" in name or "PIP" in name or "MANGALORE" in name or "Marmugao" in name:
            return "North India+Pakistan"
        else:
            return "South India+Sri Lanka"
            
    # 11. AUSTRALIA
    if country == "AUSTRALIA":
        if code in ["AUFRE", "AUEPR"] or "FREMANTLE" in name or "ESPERANCE" in name:
            return "West Australia"
        else:
            return "East Australia"
            
    # 12. EGYPT
    if country == "EGYPT":
        if code in ["EGSOK"] or "SOKHNA" in name or "ADABIYA" in name:
            return "Red Sea"
        else:
            return "NAF"

    # Default country region
    return country_regions.get(country, "Other")

# Process each port
processed_ports = []
for idx, row in df_excel.iterrows():
    code = row['Code']
    if pd.isna(code):
        continue
    
    lat_lon = coord_lookup.get(code, (None, None))
    region = get_region(row)
    
    port_dict = {
        "code": code,
        "name": row['Name'],
        "country": row['Country'],
        "lat": lat_lon[0],
        "lng": lat_lon[1],
        "region": region,
        "open_commerce": row['OpenCommerce'],
        "open_cargo": row['OpenCargo']
    }
    processed_ports.append(port_dict)

# Print missing coordinates
missing_coords = [p for p in processed_ports if p['lat'] is None]
print(f"\nPorts missing coordinates: {len(missing_coords)}")
for p in missing_coords:
    print(f"  {p['code']} - {p['name']} - {p['country']}")

# Print region breakdown
df_processed = pd.DataFrame(processed_ports)
print("\nRegion Breakdown:")
print(df_processed['region'].value_counts())

# Save to JSON
with open('ports_data.json', 'w', encoding='utf-8') as f:
    json.dump(processed_ports, f, ensure_ascii=False, indent=2)

print("\nSaved ports_data.json successfully.")
