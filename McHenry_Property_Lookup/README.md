# McHenry Property Lookup

A local, self-contained search tool for McHenry County, IL's public GIS
open-data export. Search parcels by owner, site address, or PIN and see
the full property profile (owner/mailing info, tax code, all taxing
districts, land use classification, acreage, legal description). Also
includes searchable browse tools for subdivisions, road centerlines, and
address points.

No hosted API, no external service, no internet connection required at
run time -- everything runs off a local SQLite database built from the
county's CSV export. Backend is pure Python standard library (no `pip
install` needed); frontend is plain HTML/CSS/JS.

## 1. Get the data

You need McHenry County's GIS open-data CSV export, containing:

- `McHenry_County_TaxParcels.csv`
- `Planning_Landuse.csv`
- `Address_Points.csv`
- `Subdivisions.csv`
- `Road_Centerlines.csv`

Unzip it somewhere, e.g. into `McHenry_Property_Lookup/data/Mchenry/`
(this folder is git-ignored -- the raw CSVs are large and are not
committed to this repo).

## 2. Build the database

```
cd McHenry_Property_Lookup
python3 build_database.py --data-dir data/Mchenry
```

This reads the CSVs and writes `mchenry_property.db` (a SQLite file,
~115 MB, also git-ignored) with indexes for fast search by PIN, owner
name, and site address. Takes under a minute.

If your CSVs live somewhere else, pass that path instead:

```
python3 build_database.py --data-dir /path/to/Mchenry
```

## 3. Run the app

```
python3 server.py
```

Then open **http://localhost:8000/** in a browser.

Options:

```
python3 server.py --port 8080          # different port
python3 server.py --db /path/to/other.db
```

## Using it

- **Parcels tab** -- search by owner name, site address, or PIN (e.g.
  `14-21-301-003`). Click a result to see the full parcel profile:
  mailing vs. site address (flagged when they differ -- useful for
  spotting absentee owners), property class, tax code, every taxing
  district (school, fire, park, library, road, sanitary, drainage,
  forest preserve, etc.), land use classification, acreage, a map link,
  and the legal description.
- **Subdivisions tab** -- search platted subdivisions by name or subcode,
  with plat book/page references.
- **Roads tab** -- search road centerlines by name or municipality;
  segments are grouped by name/jurisdiction and show functional
  classification and route numbers (US/state/county).
- **Address Points tab** -- search individual addressed points by
  address or street name.

## Data notes

- Land use is joined to parcels by PIN. A parcel can have more than one
  land-use record if it spans multiple use classifications.
- Subdivisions and road centerlines have no direct PIN linkage in this
  export (no geometry column), so they're separate lookup tables rather
  than joined into the parcel detail view.
- "Mailing address differs from site address" is a simple string
  comparison of `SiteAddress` vs `MailToAddress1` -- a useful signal for
  absentee-owner research, not a legal determination.

## Rebuilding after a new county data export

Re-run `build_database.py` -- it drops and recreates the database from
scratch each time, so it's safe to re-run whenever the county publishes
an updated extract.
