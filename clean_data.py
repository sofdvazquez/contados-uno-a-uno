"""
Animals killed in Spain's slaughterhouses: data cleaning script
================================================================
Source: Ministerio de Agricultura, Pesca y Alimentación (MAPA),
"Encuesta de sacrificio de ganado".
https://www.mapa.gob.es/es/estadistica/temas/estadisticas-agrarias/ganaderia/encuestas-sacrificio-ganado

What this script does:
  1. Reads the yearly files (2004-2025) and takes the summary sheet "Nº CABEZAS"
     (number of animals killed, by province/region and species).
  2. Reads the August 2026 bulletin for monthly numbers (2022-2026).
  3. Fixes units: until 2024, birds and rabbits are given in THOUSANDS.
  4. Marks confidential cells ("DC") instead of treating them as zero.
  5. Saves three tidy tables (CSV) plus an Excel file with all of them.

How to run it:
  python clean_data.py <folder_with_excel_files>[,<another_folder>] <output_folder>

Needs: pandas, openpyxl, xlrd   (pip install pandas openpyxl xlrd)
"""

import re
import sys
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

# Species columns in the order they appear in the MAPA tables (columns 1-7).
SPECIES = ["BOVINO", "OVINO", "CAPRINO", "PORCINO", "EQUINO", "AVES", "CONEJOS"]

# English names, for the web page.
SPECIES_EN = {
    "BOVINO": "Cows and calves",
    "OVINO": "Sheep and lambs",
    "CAPRINO": "Goats and kids",
    "PORCINO": "Pigs",
    "EQUINO": "Horses",
    "AVES": "Birds (chickens, turkeys, ducks...)",
    "CONEJOS": "Rabbits",
}

# The 17 autonomous communities. MAPA spells them differently across years
# (e.g. 'CAST. Y LEÓN', 'CASTILLA-LEÓN', 'Castilla y León'), so we compare
# names in UPPER CASE and map every spelling to one clean name.
REGIONS = {
    "GALICIA": "Galicia",
    "P. ASTURIAS": "Asturias", "P. DE ASTURIAS": "Asturias",
    "CANTABRIA": "Cantabria",
    "PAÍS VASCO": "País Vasco",
    "NAVARRA": "Navarra",
    "LA RIOJA": "La Rioja",
    "ARAGÓN": "Aragón",
    "CATALUÑA": "Catalunya",
    "BALEARES": "Illes Balears",
    "CAST. Y LEÓN": "Castilla y León", "CASTILLA-LEÓN": "Castilla y León", "CASTILLA Y LEÓN": "Castilla y León",
    "MADRID": "Madrid",
    "C. LA MANCHA": "Castilla-La Mancha", "CASTILLA-LA MANCHA": "Castilla-La Mancha", "CASTILLA LA MANCHA": "Castilla-La Mancha",
    "C. VALENCIANA": "Comunitat Valenciana",
    "R. DE MURCIA": "Región de Murcia",
    "EXTREMADURA": "Extremadura",
    "ANDALUCÍA": "Andalucía",
    "CANARIAS": "Canarias",
}

MONTHS = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
          "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def find_count_sheet(path: Path) -> str:
    """The sheet with the NUMBER of animals (not meat weight, 'peso'/'carne').
    Its name changes over the years: 'Nº CABEZAS', 'Nº Cabezas por especie',
    'Total nº cabezas', 'CABEZAS TOTAL PROVINCIAS'..."""
    for name in pd.ExcelFile(path).sheet_names:
        n = name.upper()
        if "CABEZ" in n and (n.startswith("Nº") or "TOTAL" in n or "ESPECIE" in n):
            return name
    raise ValueError(f"No animal-count sheet found in {path.name}")


def find_year(sheet: pd.DataFrame) -> int:
    """The year is in the first rows: in the title ('... MATADEROS 2016')
    or in a separate cell ('AÑO | 2004')."""
    for value in sheet.head(6).values.flatten():
        match = re.search(r"\b(20[0-3]\d)\b", str(value))
        if match:
            return int(match.group(1))
    raise ValueError("Could not find the year")


def find_columns(sheet: pd.DataFrame):
    """Find the header row (the one that says BOVINO) and which column
    holds each species. Column positions change between years."""
    for i in range(12):
        cells = [str(v).upper() for v in sheet.iloc[i]]
        if any(c.startswith("BOVINO") for c in cells):
            return {sp: next(j for j, c in enumerate(cells) if c.startswith(sp))
                    for sp in SPECIES}
    raise ValueError("Header row not found")


def parse_cell(value):
    """
    Turn one cell into (number, confidential_code).
      123456   -> (123456, None)
      'DC (2)' -> (None, 'DC (2)')   confidential, NOT zero
      0        -> (0, None)
    """
    if pd.isna(value):
        return None, None
    text = str(value).strip()
    if text.upper().startswith("DC"):
        return None, re.sub(r"\s+", " ", text.replace("DC(", "DC ("))
    return float(value), None


# ---------------------------------------------------------------------------
# Step 1: yearly files (2016-2025)
# ---------------------------------------------------------------------------

def read_yearly_file(path: Path):
    sheet = pd.read_excel(path, sheet_name=find_count_sheet(path), header=None)
    year = find_year(sheet)
    cols = find_columns(sheet)
    labels = sheet[0].astype(str).str.strip().str.upper()

    # Units: birds and rabbits are in THOUSANDS until 2024. Instead of trusting
    # the header text (it isn't always there), we check the size: Spain kills
    # hundreds of millions of birds, so a total below 100 million means thousands.
    total_row = labels[labels == "TOTAL"].index[0]
    in_thousands = float(sheet.iat[total_row, cols["AVES"]]) < 1e8

    def value(i, species):
        number, dc = parse_cell(sheet.iat[i, cols[species]])
        if number is not None and in_thousands and species in ("AVES", "CONEJOS"):
            number *= 1000
        return number, dc

    region_rows, national_rows = [], []
    for i, label in labels.items():
        if label in REGIONS:
            for species in SPECIES:
                number, dc = value(i, species)
                region_rows.append({
                    "year": year, "region": REGIONS[label],
                    "species_es": species, "species_en": SPECIES_EN[species],
                    "animals": None if number is None else round(number),
                    "confidential": dc is not None, "confidential_code": dc,
                })
        # "TOTAL" = animals killed in slaughterhouses (measured the same way
        # every year). "TOTAL ESPAÑA"/"ESPAÑA" also adds MAPA's ESTIMATE of
        # other slaughter, which covered many species until 2007 and only
        # rabbits later -> not comparable over time. We keep it for reference.
        if label in ("TOTAL", "TOTAL ESPAÑA", "ESPAÑA"):
            kind = "slaughterhouses" if label == "TOTAL" else "total_spain"
            for species in SPECIES:
                number, _ = value(i, species)
                national_rows.append({"year": year, "species_es": species,
                                      "kind": kind, "animals": round(number)})

    national = (pd.DataFrame(national_rows)
                .pivot_table(index=["year", "species_es"], columns="kind",
                             values="animals")
                .reset_index())
    national["species_en"] = national["species_es"].map(SPECIES_EN)
    found = {r["region"] for r in region_rows}
    assert len(found) == 17, f"{year}: found {len(found)} regions"
    return pd.DataFrame(region_rows), national


# ---------------------------------------------------------------------------
# Step 2: monthly bulletin (2022-2026)
# ---------------------------------------------------------------------------

def read_monthly_bulletin(path: Path) -> pd.DataFrame:
    sheet = pd.read_excel(path, sheet_name="2022-2026", header=None)
    rows = []
    species = None
    for i in range(len(sheet)):
        first = str(sheet.iat[i, 0]).strip()
        # The second half of the sheet is meat weight (tonnes): stop there.
        if "Peso canal" in " ".join(str(v) for v in sheet.iloc[i]):
            break
        if first and first != "nan":
            name = first.split("(")[0].strip().upper()   # 'Aves (miles)' -> 'AVES'
            if name in SPECIES:
                species = name
        year = sheet.iat[i, 1]
        if species is None or pd.isna(year):
            continue
        thousands = "miles" in first.lower() or species in ("AVES", "CONEJOS")
        for m, month in enumerate(MONTHS):
            value = sheet.iat[i, 2 + m]
            if pd.isna(value):
                continue  # month not published yet
            animals = float(value) * (1000 if thousands else 1)
            rows.append({
                "year": int(year),
                "month": m + 1,
                "month_name": month,
                "species_es": species,
                "species_en": SPECIES_EN[species],
                "animals": round(animals),
                "provisional": int(year) == 2026,
            })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Run everything
# ---------------------------------------------------------------------------

def main(data_folder: str, output_folder: str):
    folders, out = [Path(f) for f in data_folder.split(",")], Path(output_folder)
    out.mkdir(parents=True, exist_ok=True)

    regions, national = [], []
    monthly = pd.DataFrame()
    for path in sorted(p for f in folders for p in f.glob("*.xls*")):
        if "boletin" in path.name.lower():
            monthly = read_monthly_bulletin(path)
            print(f"Monthly bulletin: {path.name} -> {len(monthly)} rows")
        else:
            r, n = read_yearly_file(path)
            regions.append(r)
            national.append(n)
            print(f"Year {int(n.year.iloc[0])}: {path.name}")
    # (folders: pass several, e.g. python clean_data.py data,data_old output)

    regions = pd.concat(regions).sort_values(["year", "region", "species_es"])
    national = pd.concat(national).sort_values(["year", "species_es"])
    national = national[["year", "species_es", "species_en",
                         "slaughterhouses", "total_spain"]]

    regions.to_csv(out / "animals_killed_by_region.csv", index=False)
    national.to_csv(out / "animals_killed_spain_by_year.csv", index=False)
    monthly.to_csv(out / "animals_killed_spain_by_month.csv", index=False)
    with pd.ExcelWriter(out / "animals_killed_spain_clean.xlsx") as xl:
        national.to_excel(xl, sheet_name="Spain by year", index=False)
        monthly.to_excel(xl, sheet_name="Spain by month", index=False)
        regions.to_excel(xl, sheet_name="By region", index=False)
    print("Saved to", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
