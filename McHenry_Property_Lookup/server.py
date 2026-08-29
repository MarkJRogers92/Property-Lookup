#!/usr/bin/env python3
"""McHenry Property Lookup -- local web server.

Zero third-party dependencies: uses only the Python standard library
(http.server + sqlite3). Serves the static frontend in static/ and a
small JSON search API backed by the SQLite database built by
build_database.py.

Usage:
    python3 server.py [--db mchenry_property.db] [--port 8000]

Then open http://localhost:8000/ in a browser.
"""
import argparse
import json
import os
import re
import sqlite3
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
DEFAULT_DB = os.path.join(os.path.dirname(__file__), "mchenry_property.db")

DB_PATH = DEFAULT_DB
SEARCH_LIMIT_MAX = 200


def get_conn():
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def normalize_pin(value):
    return value.strip().upper()


def rows_to_list(rows):
    return [dict(r) for r in rows]


def clamp_limit(raw, default=50):
    try:
        n = int(raw)
    except (TypeError, ValueError):
        return default
    return max(1, min(n, SEARCH_LIMIT_MAX))


def escape_like(s):
    return re.sub(r"([%_\\])", r"\\\1", s)


def wildcard_pattern(q):
    """Build a LIKE pattern that matches regardless of how many spaces
    separate words. McHenry's own export is inconsistent about this (e.g.
    site addresses have a double space between house number and street
    name), so a plain single-spaced user query would otherwise match
    nothing."""
    tokens = q.strip().upper().split()
    return "%" + "%".join(escape_like(t) for t in tokens) + "%"


def search_parcels(conn, q, limit):
    q = q.strip()
    pin_pattern = f"%{escape_like(normalize_pin(q))}%"
    text_pattern = wildcard_pattern(q)
    rows = conn.execute(
        """
        SELECT objectid, parcel_number, owner, site_address, site_city,
               property_class, tax_status, township, parcel_area
        FROM parcels
        WHERE parcel_number_norm LIKE ? ESCAPE '\\'
           OR owner LIKE ? ESCAPE '\\' COLLATE NOCASE
           OR site_address LIKE ? ESCAPE '\\' COLLATE NOCASE
        ORDER BY
            CASE WHEN parcel_number_norm LIKE ? ESCAPE '\\' THEN 0 ELSE 1 END,
            owner
        LIMIT ?
        """,
        (pin_pattern, text_pattern, text_pattern, pin_pattern, limit),
    ).fetchall()
    return rows_to_list(rows)


def get_parcel(conn, parcel_number):
    pin_norm = normalize_pin(parcel_number)
    row = conn.execute(
        "SELECT * FROM parcels WHERE parcel_number_norm = ?", (pin_norm,)
    ).fetchone()
    if row is None:
        return None
    parcel = dict(row)
    landuse = conn.execute(
        "SELECT description, shape_area FROM landuse WHERE pin_norm = ?",
        (pin_norm,),
    ).fetchall()
    parcel["landuse"] = rows_to_list(landuse)
    # McHenry County's live tax/assessment records (assessed value, sale
    # history, payment status) live on the county's own wEDGE portal, not
    # in this GIS export. That portal's per-parcel URLs require a tax year
    # segment we can't reliably predict here, so we link to its search
    # home page instead and hand over the PIN digits to paste in.
    parcel["county_pin_digits"] = re.sub(r"\D", "", parcel_number)
    parcel["county_portal_url"] = "https://mchenryil.devnetwedge.com/"
    return parcel


def search_subdivisions(conn, q, limit):
    pattern = wildcard_pattern(q)
    rows = conn.execute(
        """
        SELECT fid, subcode, name, pages, createdon, lastupdate
        FROM subdivisions
        WHERE name LIKE ? ESCAPE '\\' COLLATE NOCASE OR subcode LIKE ? ESCAPE '\\'
        ORDER BY name
        LIMIT ?
        """,
        (pattern, pattern, limit),
    ).fetchall()
    return rows_to_list(rows)


def search_roads(conn, q, limit):
    pattern = wildcard_pattern(q)
    rows = conn.execute(
        """
        SELECT name, jurisdiction_name, jurisdiction, funct_class_name,
               us_route1, state_route1, county_route1,
               COUNT(*) AS segment_count, SUM(shape_length) AS total_length_ft
        FROM roads
        WHERE name LIKE ? ESCAPE '\\' COLLATE NOCASE OR jurisdiction_name LIKE ? ESCAPE '\\' COLLATE NOCASE
        GROUP BY name, jurisdiction_name
        ORDER BY name
        LIMIT ?
        """,
        (pattern, pattern, limit),
    ).fetchall()
    return rows_to_list(rows)


def search_addresses(conn, q, limit):
    pattern = wildcard_pattern(q)
    rows = conn.execute(
        """
        SELECT full_address, add_number, street_name, street_type, unit,
               municipality, postal_code, latitude, longitude
        FROM address_points
        WHERE full_address LIKE ? ESCAPE '\\' COLLATE NOCASE OR street_name LIKE ? ESCAPE '\\' COLLATE NOCASE
        ORDER BY full_address
        LIMIT ?
        """,
        (pattern, pattern, limit),
    ).fetchall()
    return rows_to_list(rows)


class Handler(BaseHTTPRequestHandler):
    server_version = "McHenryPropertyLookup/1.0"

    def log_message(self, fmt, *args):
        pass

    def send_json(self, payload, status=200):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_static(self, rel_path):
        safe_path = os.path.normpath(rel_path).lstrip(os.sep)
        full_path = os.path.join(STATIC_DIR, safe_path)
        if not full_path.startswith(os.path.abspath(STATIC_DIR) + os.sep) and full_path != os.path.abspath(STATIC_DIR):
            self.send_error(403)
            return
        if not os.path.isfile(full_path):
            self.send_error(404)
            return
        ctype = {
            ".html": "text/html; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".css": "text/css; charset=utf-8",
        }.get(os.path.splitext(full_path)[1], "application/octet-stream")
        with open(full_path, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        params = parse_qs(parsed.query)

        try:
            if path == "/" or path == "/index.html":
                self.send_static("index.html")
                return
            if path in ("/app.js", "/style.css"):
                self.send_static(path.lstrip("/"))
                return

            if path == "/api/search/parcels":
                q = (params.get("q") or [""])[0]
                limit = clamp_limit((params.get("limit") or [None])[0])
                if not q.strip():
                    self.send_json({"results": []})
                    return
                with get_conn() as conn:
                    self.send_json({"results": search_parcels(conn, q, limit)})
                return

            if path.startswith("/api/parcels/"):
                parcel_number = path[len("/api/parcels/"):]
                from urllib.parse import unquote
                parcel_number = unquote(parcel_number)
                with get_conn() as conn:
                    parcel = get_parcel(conn, parcel_number)
                if parcel is None:
                    self.send_json({"error": "not found"}, status=404)
                else:
                    self.send_json(parcel)
                return

            if path == "/api/search/subdivisions":
                q = (params.get("q") or [""])[0]
                limit = clamp_limit((params.get("limit") or [None])[0])
                if not q.strip():
                    self.send_json({"results": []})
                    return
                with get_conn() as conn:
                    self.send_json({"results": search_subdivisions(conn, q, limit)})
                return

            if path == "/api/search/roads":
                q = (params.get("q") or [""])[0]
                limit = clamp_limit((params.get("limit") or [None])[0])
                if not q.strip():
                    self.send_json({"results": []})
                    return
                with get_conn() as conn:
                    self.send_json({"results": search_roads(conn, q, limit)})
                return

            if path == "/api/search/addresses":
                q = (params.get("q") or [""])[0]
                limit = clamp_limit((params.get("limit") or [None])[0])
                if not q.strip():
                    self.send_json({"results": []})
                    return
                with get_conn() as conn:
                    self.send_json({"results": search_addresses(conn, q, limit)})
                return

            self.send_error(404)
        except sqlite3.OperationalError as e:
            self.send_json({"error": str(e)}, status=500)


def main():
    global DB_PATH
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--db", default=DEFAULT_DB)
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()
    DB_PATH = args.db

    if not os.path.isfile(DB_PATH):
        raise SystemExit(
            f"Database not found at {DB_PATH}.\n"
            f"Run build_database.py first (see README.md)."
        )

    server = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"McHenry Property Lookup running at http://{args.host}:{args.port}/")
    print("Press Ctrl+C to stop.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
