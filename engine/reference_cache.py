import sqlite3
import hashlib
from pathlib import Path
import pandas as pd

from engine.normalizer import (
    normalize_person_name,
    normalize_entity_name,
    normalize_zip,
    normalize_dob
)

CACHE_DIR = Path.home() / "ExclusionAppData" / "reference_cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)


def file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(8192):
            h.update(chunk)
    return h.hexdigest()


def get_cache_path(month: str):
    return CACHE_DIR / f"reference_{month}.sqlite"


def cache_exists(month: str):
    return get_cache_path(month).exists()


def build_reference_cache(month: str, oig_path: str, sam_path: str):
    db_path = get_cache_path(month)

    if db_path.exists():
        raise Exception(f"Cache for {month} already exists.")

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    create_tables(cur)

    print("Loading OIG...")
    load_oig(cur, oig_path)

    print("Loading SAM (this may take a while)...")
    load_sam(cur, sam_path)

    print("Creating indexes...")
    create_indexes(cur)

    conn.commit()
    conn.close()

    print(f"Reference cache built for {month}")


def create_tables(cur):
    cur.execute("""
        CREATE TABLE oig_people (
            first TEXT,
            last TEXT,
            dob TEXT,
            dob_compact TEXT,
            exclusion_date TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE oig_entities (
            name TEXT,
            exclusion_date TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE sam_people (
            first TEXT,
            last TEXT,
            exclusion_date TEXT,
            city TEXT,
            state TEXT,
            zip TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE sam_entities (
            name TEXT,
            exclusion_date TEXT,
            city TEXT,
            state TEXT,
            zip TEXT
        )
    """)


def create_indexes(cur):
    cur.execute("CREATE INDEX idx_oig_people ON oig_people(last, first, dob_compact)")
    cur.execute("CREATE INDEX idx_oig_entities ON oig_entities(name)")
    cur.execute("CREATE INDEX idx_sam_people ON sam_people(last, first)")
    cur.execute("CREATE INDEX idx_sam_entities ON sam_entities(name)")


def load_oig(cur, oig_path):
    df = pd.read_csv(oig_path, dtype=str).fillna("")

    for _, row in df.iterrows():
        exclusion_date = row.get("EXCLDATE", "")

        # Person
        if row.get("FIRSTNAME") and row.get("LASTNAME"):
            first, last, _, _ = normalize_person_name(
                row.get("FIRSTNAME"),
                row.get("LASTNAME")
            )
            dob_iso, dob_compact = normalize_dob(row.get("DOB"))
            cur.execute("""
                INSERT INTO oig_people VALUES (?, ?, ?, ?, ?)
            """, (first, last, dob_iso, dob_compact, exclusion_date))

        # Entity
        if row.get("BUSNAME"):
            name = normalize_entity_name(row.get("BUSNAME"))
            cur.execute("""
                INSERT INTO oig_entities VALUES (?, ?)
            """, (name, exclusion_date))


def load_sam(cur, sam_path):
    chunk_size = 50000

    for chunk in pd.read_csv(sam_path, dtype=str, chunksize=chunk_size).fillna(""):
        for _, row in chunk.iterrows():
            exclusion_date = row.get("Exclusion Date", "")

            # Person
            if row.get("First") and row.get("Last"):
                first, last, _, _ = normalize_person_name(
                    row.get("First"),
                    row.get("Last")
                )
                city = row.get("City", "").upper()
                state = row.get("State", "").upper()
                zip_code = normalize_zip(row.get("Zip"))
                cur.execute("""
                    INSERT INTO sam_people VALUES (?, ?, ?, ?, ?, ?)
                """, (first, last, exclusion_date, city, state, zip_code))

            # Entity
            if row.get("Name"):
                name = normalize_entity_name(row.get("Name"))
                city = row.get("City", "").upper()
                state = row.get("State", "").upper()
                zip_code = normalize_zip(row.get("Zip"))
                cur.execute("""
                    INSERT INTO sam_entities VALUES (?, ?, ?, ?, ?)
                """, (name, exclusion_date, city, state, zip_code))
