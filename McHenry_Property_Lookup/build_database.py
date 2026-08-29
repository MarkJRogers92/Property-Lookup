#!/usr/bin/env python3
"""Build the McHenry Property Lookup SQLite database from the county's
open-data CSV export.

Usage:
    python3 build_database.py [--data-dir DATA_DIR] [--out OUT_DB]

Expects the following files inside DATA_DIR (default: ./data/Mchenry):
    McHenry_County_TaxParcels.csv
    Planning_Landuse.csv
    Address_Points.csv
    Subdivisions.csv
    Road_Centerlines.csv

These are the raw files from the county's GIS open-data extract. They are
not included in this repo (too large for git) -- unzip the county export
and point --data-dir at the folder that contains the CSVs.
"""
import argparse
import csv
import os
import sqlite3
import sys

csv.field_size_limit(10_000_000)

DEFAULT_DATA_DIR = os.path.join(os.path.dirname(__file__), "data", "Mchenry")
DEFAULT_OUT_DB = os.path.join(os.path.dirname(__file__), "mchenry_property.db")


def normalize_pin(value):
    """Normalize a parcel PIN for join/search: uppercase, strip whitespace."""
    if value is None:
        return ""
    return value.strip().upper()


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            yield row


def get(row, key, default=""):
    v = row.get(key)
    return v.strip() if isinstance(v, str) else (v if v is not None else default)


def to_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def build_parcels(conn, data_dir):
    path = os.path.join(data_dir, "McHenry_County_TaxParcels.csv")
    print(f"Loading parcels from {path} ...")
    conn.execute("""
        CREATE TABLE parcels (
            objectid INTEGER PRIMARY KEY,
            parcel_number TEXT,
            parcel_number_norm TEXT,
            township TEXT,
            tax_code TEXT,
            tax_status TEXT,
            property_class TEXT,
            owner TEXT,
            mail_address1 TEXT,
            mail_address2 TEXT,
            mail_city TEXT,
            mail_state TEXT,
            mail_zip TEXT,
            site_address TEXT,
            site_city TEXT,
            site_state TEXT,
            site_zip TEXT,
            site_house_number TEXT,
            site_prefix_dir TEXT,
            site_street_name TEXT,
            site_street_suffix TEXT,
            site_post_dir TEXT,
            site_unit TEXT,
            airport TEXT,
            community_college TEXT,
            county TEXT,
            drainage TEXT,
            forest_preserve TEXT,
            fire_district TEXT,
            grade_school TEXT,
            high_school TEXT,
            non_high_school TEXT,
            unit_school TEXT,
            hospital_district TEXT,
            library_district TEXT,
            library TEXT,
            multi_township_district TEXT,
            park_district TEXT,
            road_district TEXT,
            sanitary_district TEXT,
            special_district TEXT,
            street_light_district TEXT,
            tif_district TEXT,
            legal_description TEXT,
            latitude REAL,
            longitude REAL,
            parcel_area REAL,
            global_id TEXT
        )
    """)
    rows = []
    for row in read_csv(path):
        rows.append((
            int(get(row, "OBJECTID", 0) or 0),
            get(row, "ParcelNumber"),
            normalize_pin(get(row, "ParcelNumber")),
            get(row, "Township"),
            get(row, "TaxCode"),
            get(row, "TaxStatus"),
            get(row, "PropertyClass"),
            get(row, "Owner"),
            get(row, "MailToAddress1"),
            get(row, "MailToAddress2"),
            get(row, "MailToCity"),
            get(row, "MailToState"),
            get(row, "MailToZip"),
            get(row, "SiteAddress"),
            get(row, "SiteCity"),
            get(row, "SiteState"),
            get(row, "SiteZip"),
            get(row, "SiteAddressHouseNumber"),
            get(row, "SiteAddressPrefixDirectional"),
            get(row, "SiteAddressStreetName"),
            get(row, "SiteAddressStreetSuffix"),
            get(row, "SiteAddressPostDirectional"),
            get(row, "SiteAddressUnitNumber"),
            get(row, "Airport"),
            get(row, "CommunityCollege"),
            get(row, "County"),
            get(row, "Drainage"),
            get(row, "ForestPreserve"),
            get(row, "FireDistrict"),
            get(row, "GradeSchool"),
            get(row, "HighSchool"),
            get(row, "NonHighSchool"),
            get(row, "UnitSchool"),
            get(row, "HospitalDistrict"),
            get(row, "LibraryDistrict"),
            get(row, "Library"),
            get(row, "MultiTownshipDistrict"),
            get(row, "ParkDistrict"),
            get(row, "RoadDistrict"),
            get(row, "SanitaryDistrict"),
            get(row, "SpecialDistrict"),
            get(row, "StreetLightDistrict"),
            get(row, "TifDistrict"),
            get(row, "LegalDescription"),
            to_float(get(row, "Latitude")),
            to_float(get(row, "Longitude")),
            to_float(get(row, "ParcelArea")),
            get(row, "GlobalID"),
        ))
    conn.executemany(f"INSERT OR REPLACE INTO parcels VALUES ({','.join('?' * 48)})", rows)
    print(f"  {len(rows):,} parcels loaded")

    conn.execute("CREATE INDEX idx_parcels_pin ON parcels(parcel_number_norm)")
    conn.execute("CREATE INDEX idx_parcels_owner ON parcels(owner COLLATE NOCASE)")
    conn.execute("CREATE INDEX idx_parcels_site_address ON parcels(site_address COLLATE NOCASE)")
    conn.execute("CREATE INDEX idx_parcels_site_street ON parcels(site_street_name COLLATE NOCASE)")
    conn.execute("CREATE INDEX idx_parcels_property_class ON parcels(property_class)")


def build_landuse(conn, data_dir):
    path = os.path.join(data_dir, "Planning_Landuse.csv")
    print(f"Loading land use from {path} ...")
    conn.execute("""
        CREATE TABLE landuse (
            objectid INTEGER PRIMARY KEY,
            pin_norm TEXT,
            description TEXT,
            shape_area REAL,
            shape_length REAL
        )
    """)
    rows = []
    for row in read_csv(path):
        rows.append((
            int(get(row, "OBJECTID", 0) or 0),
            normalize_pin(get(row, "PIN")),
            get(row, "DESCRIPTION"),
            to_float(get(row, "Shape__Area")),
            to_float(get(row, "Shape__Length")),
        ))
    conn.executemany("INSERT OR REPLACE INTO landuse VALUES (?,?,?,?,?)", rows)
    print(f"  {len(rows):,} land use records loaded")
    conn.execute("CREATE INDEX idx_landuse_pin ON landuse(pin_norm)")


def build_subdivisions(conn, data_dir):
    path = os.path.join(data_dir, "Subdivisions.csv")
    print(f"Loading subdivisions from {path} ...")
    conn.execute("""
        CREATE TABLE subdivisions (
            fid INTEGER PRIMARY KEY,
            subcode TEXT,
            name TEXT,
            pages TEXT,
            global_id TEXT,
            createdon TEXT,
            lastupdate TEXT,
            shape_area REAL,
            shape_length REAL
        )
    """)
    rows = []
    for row in read_csv(path):
        rows.append((
            int(get(row, "FID", 0) or 0),
            get(row, "SUBCODE"),
            get(row, "NAME"),
            get(row, "PAGES"),
            get(row, "GlobalID"),
            get(row, "CREATEDON"),
            get(row, "LASTUPDATE"),
            to_float(get(row, "Shape__Area")),
            to_float(get(row, "Shape__Length")),
        ))
    conn.executemany("INSERT OR REPLACE INTO subdivisions VALUES (?,?,?,?,?,?,?,?,?)", rows)
    print(f"  {len(rows):,} subdivisions loaded")
    conn.execute("CREATE INDEX idx_subdivisions_name ON subdivisions(name COLLATE NOCASE)")
    conn.execute("CREATE INDEX idx_subdivisions_subcode ON subdivisions(subcode)")


def build_roads(conn, data_dir):
    path = os.path.join(data_dir, "Road_Centerlines.csv")
    print(f"Loading road centerlines from {path} ...")
    conn.execute("""
        CREATE TABLE roads (
            objectid INTEGER PRIMARY KEY,
            jurisdiction_name TEXT,
            name TEXT,
            alt_name TEXT,
            us_route1 TEXT,
            us_route2 TEXT,
            state_route1 TEXT,
            state_route2 TEXT,
            county_route1 TEXT,
            county_route2 TEXT,
            jurisdiction TEXT,
            status TEXT,
            funct_class TEXT,
            funct_class_name TEXT,
            shape_length REAL
        )
    """)
    rows = []
    for row in read_csv(path):
        rows.append((
            int(get(row, "OBJECTID", 0) or 0),
            get(row, "JURISDICTION_NAME"),
            get(row, "NAME"),
            get(row, "ALT"),
            get(row, "US1"),
            get(row, "US2"),
            get(row, "ST1"),
            get(row, "ST2"),
            get(row, "CS1"),
            get(row, "CS2"),
            get(row, "JURISDICTION"),
            get(row, "STATUS"),
            get(row, "IDOT_FUNCT_CLASS"),
            get(row, "IDOT_FUNCT_CLASS_NAME"),
            to_float(get(row, "Shape__Length")),
        ))
    conn.executemany(f"INSERT OR REPLACE INTO roads VALUES ({','.join('?' * 15)})", rows)
    print(f"  {len(rows):,} road segments loaded")
    conn.execute("CREATE INDEX idx_roads_name ON roads(name COLLATE NOCASE)")
    conn.execute("CREATE INDEX idx_roads_jurisdiction ON roads(jurisdiction_name COLLATE NOCASE)")


def build_address_points(conn, data_dir):
    path = os.path.join(data_dir, "Address_Points.csv")
    print(f"Loading address points from {path} ...")
    conn.execute("""
        CREATE TABLE address_points (
            objectid INTEGER PRIMARY KEY,
            full_address TEXT,
            add_number TEXT,
            street_name TEXT,
            street_pretype TEXT,
            street_predir TEXT,
            street_type TEXT,
            street_posdir TEXT,
            unit TEXT,
            municipality TEXT,
            unincorporated_community TEXT,
            postal_code TEXT,
            latitude REAL,
            longitude REAL
        )
    """)
    rows = []
    for row in read_csv(path):
        rows.append((
            int(get(row, "OBJECTID", 0) or 0),
            get(row, "gcFullAdr"),
            get(row, "addNumComb") or get(row, "Add_Number"),
            get(row, "St_Name"),
            get(row, "St_PreTyp"),
            get(row, "St_PreDir"),
            get(row, "St_PosTyp") or get(row, "St_Type"),
            get(row, "St_PosDir"),
            get(row, "Unit"),
            get(row, "Inc_Muni"),
            get(row, "Uninc_Comm"),
            get(row, "Post_Code"),
            to_float(get(row, "Lat")),
            to_float(get(row, "Long")),
        ))
    conn.executemany(f"INSERT OR REPLACE INTO address_points VALUES ({','.join('?' * 14)})", rows)
    print(f"  {len(rows):,} address points loaded")
    conn.execute("CREATE INDEX idx_addrpts_full ON address_points(full_address COLLATE NOCASE)")
    conn.execute("CREATE INDEX idx_addrpts_street ON address_points(street_name COLLATE NOCASE)")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-dir", default=DEFAULT_DATA_DIR)
    parser.add_argument("--out", default=DEFAULT_OUT_DB)
    args = parser.parse_args()

    required = [
        "McHenry_County_TaxParcels.csv",
        "Planning_Landuse.csv",
        "Address_Points.csv",
        "Subdivisions.csv",
        "Road_Centerlines.csv",
    ]
    missing = [f for f in required if not os.path.isfile(os.path.join(args.data_dir, f))]
    if missing:
        print(f"ERROR: missing CSV(s) in {args.data_dir}:", file=sys.stderr)
        for f in missing:
            print(f"  - {f}", file=sys.stderr)
        print(
            "\nUnzip the McHenry County GIS open-data export and pass its folder "
            "with --data-dir, e.g.:\n"
            "  python3 build_database.py --data-dir /path/to/Mchenry",
            file=sys.stderr,
        )
        sys.exit(1)

    if os.path.exists(args.out):
        os.remove(args.out)

    conn = sqlite3.connect(args.out)
    conn.execute("PRAGMA journal_mode=OFF")
    conn.execute("PRAGMA synchronous=OFF")

    build_parcels(conn, args.data_dir)
    build_landuse(conn, args.data_dir)
    build_subdivisions(conn, args.data_dir)
    build_roads(conn, args.data_dir)
    build_address_points(conn, args.data_dir)

    conn.commit()
    conn.execute("VACUUM")
    conn.close()

    size_mb = os.path.getsize(args.out) / (1024 * 1024)
    print(f"\nDone. Wrote {args.out} ({size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
