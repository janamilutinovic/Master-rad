# Sqlite3 model created for storing SrpWn data, translated by ai models, based on Princeton WordNet

import sqlite3
import json
from pathlib import Path
from datetime import datetime, timezone
import pandas as pd

# Fields that come from the parsed JSON translation (mirrors the SrpWN row shape).
TRANSLATED_FIELDS = [
    "literals",
    "def",
    "usage",
    "hypernymID",
    "hypernym_literals",
    "hypernym_def",
    "hypernym_usage",
]

SCHEMA = """
CREATE TABLE IF NOT EXISTS translations (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    synset_id       TEXT NOT NULL,
    domain          TEXT,
    provider        TEXT,
    model           TEXT NOT NULL,
    strategy        TEXT NOT NULL,
    status          TEXT NOT NULL,
    literals        TEXT,
    def_            TEXT,
    usage           TEXT,
    hypernymID      TEXT,
    hypernym_literals TEXT,
    hypernym_def    TEXT,
    hypernym_usage  TEXT,
    raw_response    TEXT,
    error_message   TEXT,
    created_at      TEXT NOT NULL,
    UNIQUE(synset_id, model, strategy)
);

"""

# Initialize db and create table translations where we will store all translated synsets by specific ai version model and information about prompt strategy used 
def init_db(path: str) -> sqlite3.Connection:
    folder = Path(path).parent
    if folder and not folder.exists():
        folder.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.executescript(SCHEMA)
    conn.commit()
    return conn

# Function returns True if we already translated sysnset_id successfully (so we don't call api multiple times) 
# or False if not yet translated or model returned some error (then repeat translation) 
def already_done(conn: sqlite3.Connection, synset_id: str, model: str, strategy: str) -> bool:
    """True if we already have a *successful* result for this combination.
    Errors are not counted as done, so a failed call gets retried on the next run."""
    row = conn.execute(
        "SELECT status FROM translations WHERE synset_id = ? AND model = ? AND strategy = ?",
        (synset_id, model, strategy),
    ).fetchone()
    return row is not None and row[0] == "success"


def upsert_translation(conn: sqlite3.Connection, record: dict) -> None:
    """Insert or update a single (synset, model, strategy) result.

    record keys:
      synset_id, domain, provider, model, strategy, status, raw_response,
      parsed (dict or None), error_message (str or None)
    """
    parsed = record.get("parsed") or {}
    if not isinstance(parsed, dict):
        # model returned a JSON-encoded string instead of a JSON object
        # (e.g. double-encoded JSON) -- don't crash, just store nothing for the fields.
        parsed = {}
    values = {
        "synset_id": record["synset_id"],
        "domain": record.get("domain"),
        "provider": record.get("provider"),
        "model": record["model"],
        "strategy": record["strategy"],
        "status": record["status"],
        "literals": parsed.get("literals"),
        "def_": parsed.get("def"),
        "usage": parsed.get("usage"),
        "hypernymID": parsed.get("hypernymID"),
        "hypernym_literals": parsed.get("hypernym_literals"),
        "hypernym_def": parsed.get("hypernym_def"),
        "hypernym_usage": parsed.get("hypernym_usage"),
        "raw_response": record.get("raw_response"),
        "error_message": record.get("error_message"),
        "created_at": datetime.now(timezone.utc).isoformat(),
    }
    conn.execute(
        """
        INSERT INTO translations (
            synset_id, domain, provider, model, strategy, status,
            literals, def_, usage, hypernymID, hypernym_literals, hypernym_def, hypernym_usage,
            raw_response, error_message, created_at
        ) VALUES (
            :synset_id, :domain, :provider, :model, :strategy, :status,
            :literals, :def_, :usage, :hypernymID, :hypernym_literals, :hypernym_def, :hypernym_usage,
            :raw_response, :error_message, :created_at
        )
        ON CONFLICT(synset_id, model, strategy) DO UPDATE SET
            domain=excluded.domain,
            provider=excluded.provider,
            status=excluded.status,
            literals=excluded.literals,
            def_=excluded.def_,
            usage=excluded.usage,
            hypernymID=excluded.hypernymID,
            hypernym_literals=excluded.hypernym_literals,
            hypernym_def=excluded.hypernym_def,
            hypernym_usage=excluded.hypernym_usage,
            raw_response=excluded.raw_response,
            error_message=excluded.error_message,
            created_at=excluded.created_at
        """,
        values,
    )
    conn.commit()

# insert or update translations table for that sysnet_id
def safe_parse_and_store(conn: sqlite3.Connection, synset_id: str, domain: str, api_result: dict) -> None:
    if api_result["status"] != "success":
        upsert_translation(conn, {
            "synset_id": synset_id,
            "domain": domain,
            "provider": api_result.get("provider"),
            "model": api_result["model"],
            "strategy": api_result["strategy"],
            "status": "error",
            "raw_response": api_result["text"],
            "parsed": None,
            "error_message": api_result["text"],
        })
        return

    clean_text = api_result["text"].replace("```json", "").replace("```", "").strip()
    try:
        parsed = json.loads(clean_text)
        if isinstance(parsed, str):
            # some models double-encode: the first decode just unwraps a JSON string
            # that itself contains the real JSON object -- decode once more.
            parsed = json.loads(parsed)
        if not isinstance(parsed, dict):
            raise ValueError(f"expected a JSON object, got {type(parsed).__name__}")
        status = "success"
        error_message = None
    except (json.JSONDecodeError, ValueError) as e:
        parsed = None
        status = "parse_error"
        error_message = f"{e}"

    upsert_translation(conn, {
        "synset_id": synset_id,
        "domain": domain,
        "provider": api_result.get("provider"),
        "model": api_result["model"],
        "strategy": api_result["strategy"],
        "status": status,
        "raw_response": api_result["text"],
        "parsed": parsed,
        "error_message": error_message,
    })


def load_dataframe(conn: sqlite3.Connection, only_success: bool = False) -> pd.DataFrame:
    query = "SELECT * FROM translations"
    if only_success:
        query += " WHERE status = 'success'"
    df = pd.read_sql_query(query, conn)
    return df.rename(columns={"def_": "def"})


def progress_summary(conn: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql_query(
        "SELECT model, strategy, status, COUNT(*) as n FROM translations GROUP BY model, strategy, status ORDER BY model, strategy",
        conn,
    )

# Export table to excel
def export_to_excel(conn: sqlite3.Connection, path: str, only_success: bool = True) -> None:
    df = load_dataframe(conn, only_success=only_success)
    folder = Path(path).parent
    if folder and not folder.exists():
        folder.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="all_results", index=False)
        for (model, strategy), group in df.groupby(["model", "strategy"]):
            sheet_name = f"{model}_{strategy}"[:31]  # Excel sheet name limit - 31 is some default limit 
            group.to_excel(writer, sheet_name=sheet_name, index=False)

# Selects just parse errors
def show_parse_errors(conn: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql_query(
        "SELECT * FROM translations WHERE status = 'parse_error'",
        conn,
    )
