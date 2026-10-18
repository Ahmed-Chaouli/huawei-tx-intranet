
# ========================================"
# =============================================================================

import streamlit as st
import pandas as pd
import sqlite3
import os
import io
import re
import csv
import uuid
import base64
import hashlib
import hmac
import secrets
import string
import logging
import datetime as _dt
from logging.handlers import RotatingFileHandler
from pathlib import Path
from contextlib import contextmanager
from datetime import datetime, timedelta

def _adapt_ts(v):
    return None if v is None or pd.isna(v) else v.strftime("%Y-%m-%d %H:%M:%S")

def _adapt_dt(v):
    if v is None: return None
    if isinstance(v, _dt.datetime): return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, _dt.date):     return v.strftime("%Y-%m-%d")
    return str(v)

sqlite3.register_adapter(_dt.datetime, _adapt_dt)
sqlite3.register_adapter(_dt.date,     _adapt_dt)
sqlite3.register_adapter(pd.Timestamp, _adapt_ts)

def sanitize_val_for_sqlite(v):
    if v is None: return None
    try:
        if pd.isna(v) is True: return None
    except (TypeError, ValueError): pass
    if isinstance(v, pd.Timestamp): return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, _dt.datetime): return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, _dt.date):     return v.strftime("%Y-%m-%d")
    if isinstance(v, (list, dict, set, tuple)): return str(v)
    return v

def sanitize_df_for_sqlite(df):
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%Y-%m-%d %H:%M:%S")
        elif out[col].dtype == "object":
            out[col] = out[col].map(sanitize_val_for_sqlite)
    return out

def qi(name):
    return '"' + str(name).replace('"', '""') + '"'

def _sql_insert(cur, table, cols, vals):
    cols_sql = ", ".join(qi(c) for c in cols)
    ph = ", ".join(["?"] * len(cols))
    cur.execute(f"INSERT INTO {qi(table)} ({cols_sql}) VALUES ({ph})", vals)

def _neutralize_formula(val):
    if isinstance(val, str) and val and val[0] in ("=", "+", "@", "\t", "\r"):
        return "'" + val
    return val

def _sanitize_excel_formulas(workbook):
    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    cell.data_type = "s"
                if isinstance(cell.value, str) and cell.value and cell.value[0] in ("=", "+", "@", "\t", "\r"):
                    cell.value = "'" + cell.value

def _write_csv_safe(path, header, rows, mode="x"):
    with open(path, mode, encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        for r in rows:
            writer.writerow([_neutralize_formula(v) if isinstance(v, str) else v for v in r])

# --- Comparison helpers ---
def normalize_val(v):
    vc = sanitize_val_for_sqlite(v)
    return "" if vc is None else str(vc).strip()

def ci_norm(v):
    vc = sanitize_val_for_sqlite(v)
    return "" if vc is None else str(vc).strip().casefold()

def _smart_equal(a, b):
    """FIX P1-A: numeric if both look numeric, else case-insensitive text."""
    av = sanitize_val_for_sqlite(a)
    bv = sanitize_val_for_sqlite(b)
    if av is None and bv is None: return True
    if av is None or bv is None: return False
    try:
        af = float(av); bf = float(bv)
        return abs(af - bf) < 1e-9
    except (ValueError, TypeError):
        pass
    return str(av).strip().casefold() == str(bv).strip().casefold()

def ci_unique(values):
    seen = {}
    for v in values:
        s = sanitize_val_for_sqlite(v)
        if s is None: continue
        s = str(s).strip()
        if not s: continue
        k = s.casefold()
        if k not in seen:
            seen[k] = s
    return sorted(seen.values(), key=str.casefold)

def _ci_filter(df, col, sel):
    if sel == "All" or col not in df.columns:
        return df
    key = str(sel).strip().casefold()
    mask = df[col].fillna("").astype(str).str.strip().str.casefold() == key
    return df[mask]

def _ci_group_series(series):
    canon = {}
    result = []
    for v in series:
        s = "N/A" if pd.isna(v) else str(v).strip()
        if not s: s = "N/A"
        k = s.casefold()
        if k not in canon:
            canon[k] = s
        result.append(canon[k])
    return pd.Series(result, index=series.index)

# --- Logging ---
BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "logs"; LOG_DIR.mkdir(exist_ok=True)
BACKUP_DIR = BASE_DIR / "backups"; BACKUP_DIR.mkdir(exist_ok=True)

logger = logging.getLogger("HuaweiTXIntranet")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh = RotatingFileHandler(LOG_DIR / "app.log", maxBytes=5*1024*1024, backupCount=5, encoding="utf-8")
    fh.setFormatter(fmt); sh = logging.StreamHandler(); sh.setFormatter(fmt)
    logger.addHandler(fh); logger.addHandler(sh)

# --- Page config ---
st.set_page_config(page_title="Huawei TX Intranet - Enterprise v12.5", layout="wide")
st.markdown("""
<style>
  html, body, [class*="css"] { direction: ltr; text-align: left; }
  [data-testid="stSidebar"] { direction: ltr; text-align: left; }
  .stDataFrame, [data-testid="stDataFrame"], table { direction: ltr !important; text-align: left !important; }
  input, textarea, td, th { unicode-bidi: plaintext; }
  div[data-testid="stMetricValue"] { direction: ltr; text-align: left; }
  code { direction: ltr; display: inline-block; }
</style>
""", unsafe_allow_html=True)

def _vtuple(v):
    try: return tuple(int(x) for x in v.split(".")[:2])
    except Exception: return (0, 0)

if _vtuple(st.__version__) < (1, 30):
    st.error(f"❌ Streamlit {st.__version__} < 1.30. Run: `pip install -U streamlit`")
    st.stop()

_HAS_CONTEXT_COOKIES = (
    hasattr(st, "context")
    and hasattr(getattr(st, "context", None), "cookies")
    and _vtuple(st.__version__) >= (1, 37)
)

# --- Constants ---
DB_PATH         = str(BASE_DIR / "huawei_tx_database.db")
INITIAL_EXCEL   = str(BASE_DIR / "Copie de Situation HUAWEI(1).xlsx")
BOOTSTRAP_TOKEN = os.environ.get("HUAWEI_BOOTSTRAP_TOKEN", "").strip()
DEBUG           = os.environ.get("HX_DEBUG", "0") == "1"

SESSION_COOKIE_NAME     = "hx_sid"
SESSION_COOKIE_SECURE   = os.environ.get("HX_COOKIE_SECURE", "0") == "1"
SESSION_COOKIE_SAMESITE = "Strict"
HX_ALLOW_URL_SID        = os.environ.get("HX_ALLOW_URL_SID", "1") == "1"
HX_PRODUCTION           = os.environ.get("HX_PRODUCTION", "0") == "1"

SESSION_TIMEOUT_MINUTES = 30
SESSION_TTL_HOURS       = 8

PBKDF2_ITERATIONS  = 600_000
MAX_LOGIN_ATTEMPTS = 5
LOCKOUT_MINUTES    = 5

BACKUP_RETENTION_DAYS = 30
AUDIT_RETENTION_DAYS  = 180
SCHEMA_VERSION        = 12

# FIX P1-B: internal columns never editable
INTERNAL_COLS = {"id", "record_id", "revision", "last_modified", "modified_by"}

RSL_COMMENT  = "UOP Comment"
CONG_COMMENT = "DTV / UOP feedback"

ROLE_DISPLAY = {"BOTX": "BO TX"}
def role_label(r): return ROLE_DISPLAY.get(r, r)

ASSIGNABLE_ROLES = ["UOP", "DTV", "BOTX", "ADMIN"]

EDITABLE_MATRIX = {
    ("UOP",  "rsl"):        [RSL_COMMENT],
    ("DTV",  "rsl"):        [RSL_COMMENT],
    ("UOP",  "congestion"): [CONG_COMMENT],
    ("DTV",  "congestion"): [CONG_COMMENT],
    ("BOTX", "rsl"):        ["BO TX feedback", "Status", RSL_COMMENT],
    ("BOTX", "congestion"): ["BO TX feedback", "status", CONG_COMMENT],
}

NUMERIC_HINTS = {"rsl": {"RSL nearend", "RSL farend"}, "congestion": set()}

COMMON_PASSWORDS_BLOCKLIST = {
    "password1234","admin1234567","huawei123456","123456789012",
    "qwerty123456","welcome12345","mobilis12345","algeria12345"
}

def password_policy(pwd):
    if len(pwd) < 12: return False, "Must be at least 12 characters."
    if not re.search(r"[A-Z]", pwd): return False, "Needs an uppercase letter."
    if not re.search(r"[a-z]", pwd): return False, "Needs a lowercase letter."
    if not re.search(r"\d", pwd):    return False, "Needs a digit."
    if not re.search(r"[!@#$%^&*()_+\-=\[\]{};:,.<>?]", pwd):
        return False, "Needs a special character."
    if pwd.lower() in COMMON_PASSWORDS_BLOCKLIST:
        return False, "Common password; choose another."
    return True, ""

def generate_strong_password(length=14):
    alphabet = string.ascii_letters + string.digits + "!@#$%^&*()-_=+"
    while True:
        p = "".join(secrets.choice(alphabet) for _ in range(length))
        if password_policy(p)[0]: return p

# --- FIX P0-1: production guard ---
if HX_PRODUCTION:
    problems = []
    if HX_ALLOW_URL_SID:
        problems.append("HX_ALLOW_URL_SID must be 0 in production")
    if not SESSION_COOKIE_SECURE:
        problems.append("HX_COOKIE_SECURE must be 1 in production (HTTPS required)")
    if problems:
        st.error(
            "❌ **Production guard tripped (HX_PRODUCTION=1):**\n\n"
            + "\n".join(f"- {p}" for p in problems)
            + "\n\nEither fix the environment variables, or unset HX_PRODUCTION."
        )
        st.stop()

# --- DB connection ---
def _set_wal():
    try:
        c = sqlite3.connect(DB_PATH, timeout=30)
        c.execute("PRAGMA journal_mode=WAL;")
        c.close()
    except Exception as e:
        logger.warning(f"WAL setup failed: {e}")

@contextmanager
def get_db(immediate=False):
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False, isolation_level=None)
    conn.execute("PRAGMA busy_timeout=30000;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    committed_or_rolled_back = False
    try:
        if immediate: conn.execute("BEGIN IMMEDIATE;")
        yield conn
        if immediate and conn.in_transaction:
            try:
                conn.execute("COMMIT;")
                committed_or_rolled_back = True
            except Exception as ce:
                try: conn.execute("ROLLBACK;")
                except Exception: pass
                committed_or_rolled_back = True
                logger.error(f"COMMIT failed: {ce}")
                raise
    except Exception as e:
        if immediate and conn.in_transaction and not committed_or_rolled_back:
            try: conn.execute("ROLLBACK;")
            except Exception: pass
        logger.error(f"DB error: {e}")
        raise
    finally:
        conn.close()

# --- Password hashing ---
def hash_password(password, salt=None):
    password = password.strip()
    if salt is None: salt = os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return base64.b64encode(digest).decode("ascii"), base64.b64encode(salt).decode("ascii")

def verify_password(password, stored_hash, stored_salt):
    try:
        salt = base64.b64decode(stored_salt.encode("ascii"))
        new_hash, _ = hash_password(password, salt)
        return hmac.compare_digest(new_hash, stored_hash)
    except Exception:
        return False

@st.cache_resource(show_spinner=False)
def _get_dummy_credentials():
    return hash_password("dummy_constant_time_protection_v12_5!")

DUMMY_HASH, DUMMY_SALT = _get_dummy_credentials()

# --- Audit ---
def log_audit(cursor, username, role, table_name, target, action, col, old_val, new_val):
    cursor.execute("""
        INSERT INTO audit_log (timestamp,username,role,table_name,link_name,action,column_changed,old_value,new_value)
        VALUES (?,?,?,?,?,?,?,?,?)
    """, (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), username, role,
          table_name, str(target), action, col, str(old_val), str(new_val)))
    logger.info(f"AUDIT user={username} action={action} table={table_name} target={target}")

def require_admin(role, action):
    if role != "ADMIN":
        raise PermissionError(f"Permission denied — {action} requires ADMIN role.")

# --- Migrations ---
REQUIRED_TRACKING = {
    "record_id": "TEXT", "revision": "INTEGER DEFAULT 1",
    "last_modified": "TEXT", "modified_by": "TEXT",
}

def _migrate_table_safely(conn, table_name):
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({qi(table_name)})")
    cols = {c[1] for c in cur.fetchall()}
    if not cols: return False
    for name, typ in REQUIRED_TRACKING.items():
        if name not in cols:
            try:
                cur.execute(f"ALTER TABLE {qi(table_name)} ADD COLUMN {qi(name)} {typ}")
            except Exception as e:
                raise RuntimeError(f"Migration failed on {table_name}.{name}: {e}")
    cur.execute(f"SELECT rowid FROM {qi(table_name)} WHERE record_id IS NULL OR record_id = ''")
    for (rid,) in cur.fetchall():
        cur.execute(f"UPDATE {qi(table_name)} SET record_id = ? WHERE rowid = ?", (str(uuid.uuid4()), rid))
    cur.execute(f"UPDATE {qi(table_name)} SET revision = 1 WHERE revision IS NULL OR revision < 1")
    try:
        cur.execute(f"CREATE UNIQUE INDEX IF NOT EXISTS idx_{table_name}_record_id ON {qi(table_name)}(record_id)")
    except sqlite3.IntegrityError:
        cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{table_name}_record_id ON {qi(table_name)}(record_id)")
    return True

def _ensure_skeleton(conn, table_name, link_col, comment_col):
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
    if cur.fetchone(): return
    cur.execute(f"""
        CREATE TABLE {qi(table_name)} (
            record_id TEXT PRIMARY KEY, revision INTEGER DEFAULT 1,
            last_modified TEXT, modified_by TEXT,
            {qi(link_col)} TEXT, {qi(comment_col)} TEXT
        )
    """)

def infer_sql_type(series, col_name, table_name):
    if col_name in NUMERIC_HINTS.get(table_name, set()): return "REAL"
    if pd.api.types.is_integer_dtype(series):    return "INTEGER"
    if pd.api.types.is_float_dtype(series):      return "REAL"
    if pd.api.types.is_bool_dtype(series):       return "INTEGER"
    if pd.api.types.is_datetime64_any_dtype(series): return "TEXT"
    return "TEXT"

def ensure_columns_from_df(conn, table_name, df):
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({qi(table_name)})")
    existing_lower = {c[1].strip().lower() for c in cur.fetchall()}
    for col in df.columns:
        col_clean = str(col).strip()
        if col_clean in REQUIRED_TRACKING or col_clean in INTERNAL_COLS: continue
        if col_clean.lower() in existing_lower: continue
        col_type = infer_sql_type(df[col], col_clean, table_name)
        try:
            cur.execute(f"ALTER TABLE {qi(table_name)} ADD COLUMN {qi(col_clean)} {col_type}")
            existing_lower.add(col_clean.lower())
        except Exception as e:
            logger.warning(f"add column {col_clean} failed: {e}")

def _normalize_columns_against_db(conn, table_name, df):
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({qi(table_name)})")
    db_map = {c[1].strip().lower(): c[1] for c in cur.fetchall()}
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    rename = {}
    for col in df.columns:
        key = col.lower()
        if key in db_map and db_map[key] != col:
            rename[col] = db_map[key]
    if rename:
        df = df.rename(columns=rename)
    return df

def init_db():
    _set_wal()
    with get_db(immediate=True) as conn:
        cur = conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS schema_version (version INTEGER PRIMARY KEY, applied_at TEXT)")
        cur.execute("CREATE TABLE IF NOT EXISTS app_meta (k TEXT PRIMARY KEY, v TEXT)")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY, password_hash TEXT NOT NULL,
                salt TEXT NOT NULL, role TEXT NOT NULL, full_name TEXT NOT NULL,
                must_change_pwd INTEGER DEFAULT 0, session_version INTEGER DEFAULT 1, last_seen TEXT
            )
        """)
        cur.execute("PRAGMA table_info(users)")
        u_cols = {c[1] for c in cur.fetchall()}
        for name, typ in {"salt":"TEXT","must_change_pwd":"INTEGER DEFAULT 0",
                          "session_version":"INTEGER DEFAULT 1","last_seen":"TEXT"}.items():
            if name not in u_cols:
                cur.execute(f"ALTER TABLE users ADD COLUMN {qi(name)} {typ}")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS login_attempts (
                username TEXT PRIMARY KEY, failed_count INTEGER DEFAULT 0,
                lockout_until TEXT, last_attempt TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT, username TEXT, role TEXT, table_name TEXT,
                link_name TEXT, action TEXT, column_changed TEXT,
                old_value TEXT, new_value TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                token TEXT PRIMARY KEY, username TEXT NOT NULL,
                session_version INTEGER DEFAULT 1,
                created_at TEXT NOT NULL, expires_at TEXT NOT NULL,
                last_activity TEXT NOT NULL,
                FOREIGN KEY (username) REFERENCES users(username) ON DELETE CASCADE
            )
        """)
        cur.execute("PRAGMA table_info(sessions)")
        s_cols = {c[1] for c in cur.fetchall()}
        if "session_version" not in s_cols:
            cur.execute('ALTER TABLE sessions ADD COLUMN "session_version" INTEGER DEFAULT 1')
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sessions_user ON sessions(username)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sessions_exp ON sessions(expires_at)")
        cur.execute("DELETE FROM sessions WHERE expires_at < ?",
                    (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),))

        cutoff = (datetime.now() - timedelta(days=AUDIT_RETENTION_DAYS)).strftime("%Y-%m-%d %H:%M:%S")
        cur.execute(
            "DELETE FROM audit_log WHERE timestamp < ? AND (action IS NULL OR action != 'AUDIT_DELETE')",
            (cutoff,),
        )

        _ensure_skeleton(conn, "rsl",        "Link Name", RSL_COMMENT)
        _ensure_skeleton(conn, "congestion", "Link name", CONG_COMMENT)

        cur.execute("SELECT v FROM app_meta WHERE k='initial_import_done'")
        already_imported = cur.fetchone() is not None
        if not already_imported:
            existing_rows = (
                cur.execute("SELECT COUNT(*) FROM rsl").fetchone()[0]
                + cur.execute("SELECT COUNT(*) FROM congestion").fetchone()[0]
            )
            if existing_rows > 0:
                cur.execute(
                    "INSERT OR REPLACE INTO app_meta (k, v) VALUES ('initial_import_done', ?)",
                    (f"migrated_existing_db_{datetime.now().strftime('%Y%m%d_%H%M%S')}",)
                )
                already_imported = True

        if not already_imported and os.path.exists(INITIAL_EXCEL):
            try:
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for sheet, tbl, lcol in [("RSL","rsl","Link Name"),("Congestion","congestion","Link name")]:
                    df = sanitize_df_for_sqlite(pd.read_excel(INITIAL_EXCEL, sheet_name=sheet))
                    if lcol in df.columns:
                        df = df[df[lcol].notna() & (df[lcol].astype(str).str.strip() != "")]
                    ensure_columns_from_df(conn, tbl, df)
                    df["record_id"]     = [str(uuid.uuid4()) for _ in range(len(df))]
                    df["revision"]      = 1
                    df["last_modified"] = now_str
                    df["modified_by"]   = "System Init"
                    for _, row in df.iterrows():
                        _sql_insert(cur, tbl, list(df.columns),
                                    [sanitize_val_for_sqlite(row[c]) for c in df.columns])
                cur.execute(
                    "INSERT OR REPLACE INTO app_meta (k, v) VALUES ('initial_import_done', ?)",
                    (now_str,)
                )
            except Exception as e:
                logger.error(f"Excel import failed: {e}")

        for tbl in ["rsl", "congestion"]:
            _migrate_table_safely(conn, tbl)

        cur.execute("INSERT OR IGNORE INTO schema_version (version, applied_at) VALUES (?, ?)",
                    (SCHEMA_VERSION, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

@st.cache_resource(show_spinner=False)
def _ensure_schema_once():
    init_db(); return True

try:
    _ensure_schema_once()
except Exception as e:
    st.error(f"❌ Database initialization failed: {e}")
    st.stop()

# --- Session management ---
def _read_cookie(name):
    if not _HAS_CONTEXT_COOKIES: return None
    try: return st.context.cookies.get(name)
    except Exception: return None

def _queue_write_cookie(name, value, ttl_seconds):
    st.session_state["_pending_cookie_op"] = ("set", name, value, int(ttl_seconds))

def _queue_delete_cookie(name):
    st.session_state["_pending_cookie_op"] = ("del", name, "", 0)

def _flush_pending_cookie_op():
    op = st.session_state.pop("_pending_cookie_op", None)
    if not op: return
    kind, name, value, ttl_seconds = op
    secure = "; Secure" if SESSION_COOKIE_SECURE else ""
    if kind == "set":
        safe_val = str(value).replace('"', "").replace(";", "").replace("\n", "").replace("\r", "")
        html = f"""<script>
        (function() {{
            try {{
                document.cookie = "{name}={safe_val}; path=/; max-age={ttl_seconds}; SameSite={SESSION_COOKIE_SAMESITE}{secure}";
            }} catch(e) {{ console.error(e); }}
        }})();
        </script>"""
    else:
        html = f"""<script>
        (function() {{
            try {{
                document.cookie = "{name}=; path=/; max-age=0; SameSite={SESSION_COOKIE_SAMESITE}{secure}";
            }} catch(e) {{ console.error(e); }}
        }})();
        </script>"""
    try: st.components.v1.html(html, height=0, scrolling=False)
    except Exception as e: logger.warning(f"cookie flush failed: {e}")

def _get_url_sid():
    if not HX_ALLOW_URL_SID: return None
    try:
        v = st.query_params.get("sid")
        return v if v else None
    except Exception: return None

def _set_url_sid(token):
    if not HX_ALLOW_URL_SID: return
    try: st.query_params["sid"] = token
    except Exception as e: logger.warning(f"URL sid write failed: {e}")

def _del_url_sid():
    if not HX_ALLOW_URL_SID: return
    try:
        if "sid" in st.query_params: del st.query_params["sid"]
    except Exception: pass

def _create_session(username, session_version=1):
    token = secrets.token_urlsafe(32)
    now = datetime.now()
    now_s = now.strftime("%Y-%m-%d %H:%M:%S")
    exp_s = (now + timedelta(hours=SESSION_TTL_HOURS)).strftime("%Y-%m-%d %H:%M:%S")
    with get_db(immediate=True) as conn:
        conn.execute(
            "INSERT INTO sessions (token,username,session_version,created_at,expires_at,last_activity) VALUES (?,?,?,?,?,?)",
            (token, username, int(session_version), now_s, exp_s, now_s))
    return token

def _revoke_session(token):
    if not token: return
    try:
        with get_db(immediate=True) as conn:
            conn.execute("DELETE FROM sessions WHERE token = ?", (token,))
    except Exception as e: logger.warning(f"revoke failed: {e}")

def _validate_session(token):
    if not token: return None
    try:
        with get_db() as conn:
            row = conn.execute("""
                SELECT s.username, u.role, u.full_name, u.session_version,
                       u.must_change_pwd, s.last_activity, s.expires_at, s.session_version
                FROM sessions s JOIN users u ON s.username = u.username
                WHERE s.token = ?
            """, (token,)).fetchone()
        if not row: return None
        if row[7] is not None and int(row[7]) != int(row[3]):
            _revoke_session(token); return None
        now = datetime.now()
        try:
            exp = datetime.strptime(row[6], "%Y-%m-%d %H:%M:%S")
            if now > exp: _revoke_session(token); return None
        except Exception: pass
        try:
            last = datetime.strptime(row[5], "%Y-%m-%d %H:%M:%S")
            if (now - last).total_seconds() > SESSION_TIMEOUT_MINUTES * 60:
                _revoke_session(token); return None
        except Exception: pass
        with get_db(immediate=True) as conn:
            conn.execute("UPDATE sessions SET last_activity = ? WHERE token = ?",
                        (now.strftime("%Y-%m-%d %H:%M:%S"), token))
        return row
    except Exception as e:
        logger.warning(f"session validate failed: {e}")
        return None

def _revoke_all_user_sessions(username, except_token=None):
    try:
        with get_db(immediate=True) as conn:
            if except_token:
                conn.execute("DELETE FROM sessions WHERE username=? AND token != ?", (username, except_token))
            else:
                conn.execute("DELETE FROM sessions WHERE username=?", (username,))
    except Exception as e: logger.warning(f"bulk revoke failed: {e}")

def _touch_session(token):
    if not token: return
    try:
        with get_db(immediate=True) as conn:
            conn.execute("UPDATE sessions SET last_activity = ? WHERE token = ?",
                        (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), token))
    except Exception: pass

# --- Lockout ---
def check_db_lockout(username):
    if not username: return False, 0
    with get_db() as conn:
        row = conn.execute("SELECT lockout_until FROM login_attempts WHERE username = ?",
                          (username.strip().lower(),)).fetchone()
        if row and row[0]:
            try:
                lt = datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S")
                if datetime.now() < lt:
                    return True, int((lt - datetime.now()).total_seconds())
            except Exception: pass
    return False, 0

def record_login_result(username, success, user_exists=True):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_db(immediate=True) as conn:
        cur = conn.cursor()
        if success:
            cur.execute("DELETE FROM login_attempts WHERE username = ?", (username,))
            cur.execute("UPDATE users SET last_seen = ? WHERE username = ?", (now_str, username))
            log_audit(cur, username, "AUTH", "users", username, "LOGIN_SUCCESS", "last_seen", "", now_str)
            return 0
        cur.execute("""
            INSERT INTO login_attempts (username, failed_count, last_attempt) VALUES (?,1,?)
            ON CONFLICT(username) DO UPDATE SET
                failed_count = CASE WHEN lockout_until IS NOT NULL AND lockout_until < ?
                                    THEN 1 ELSE failed_count + 1 END,
                lockout_until = CASE WHEN lockout_until IS NOT NULL AND lockout_until < ?
                                     THEN NULL ELSE lockout_until END,
                last_attempt = excluded.last_attempt
        """, (username, now_str, now_str, now_str))
        count = cur.execute("SELECT failed_count FROM login_attempts WHERE username = ?", (username,)).fetchone()[0]
        if count >= MAX_LOGIN_ATTEMPTS:
            lock_until = (datetime.now() + timedelta(minutes=LOCKOUT_MINUTES)).strftime("%Y-%m-%d %H:%M:%S")
            cur.execute("UPDATE login_attempts SET lockout_until = ? WHERE username = ?", (lock_until, username))
            if user_exists:
                log_audit(cur, username, "AUTH", "login_attempts", username, "LOCKOUT", "lockout_until", str(count), lock_until)
        elif user_exists:
            log_audit(cur, username, "AUTH", "login_attempts", username, "LOGIN_FAILED", "failed_count", str(count-1), str(count))
        return count

# --- FIX P1-backup: atomic backup with integrity check ---
def prune_backups(keep_days=BACKUP_RETENTION_DAYS):
    cutoff = datetime.now() - timedelta(days=keep_days)
    for f in BACKUP_DIR.glob("backup_huawei_tx_*.db"):
        try:
            stem = f.stem.replace("backup_huawei_tx_", "")
            # New format: YYYYMMDD_HHMMSS_ffffff_<uuid8>
            parts = stem.split("_")
            if len(parts) >= 3 and len(parts[0]) == 8 and parts[0].isdigit():
                ts = datetime.strptime(f"{parts[0]}_{parts[1]}_{parts[2]}", "%Y%m%d_%H%M%S_%f")
            else:
                # Old format: YYYY-MM-DD_HH-MM-SS
                ts = datetime.strptime(stem, "%Y-%m-%d_%H-%M-%S")
            if ts < cutoff:
                f.unlink(); logger.info(f"Pruned {f.name}")
        except Exception as e:
            logger.warning(f"prune skip {f.name}: {e}")

def create_atomic_backup():
    BACKUP_DIR.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    uid = uuid.uuid4().hex[:8]
    final_name = f"backup_huawei_tx_{ts}_{uid}.db"
    final_path = BACKUP_DIR / final_name
    tmp_path = BACKUP_DIR / f".tmp_{ts}_{uid}.db"

    src = None
    dst = None
    chk = None
    try:
        src = sqlite3.connect(DB_PATH)
        dst = sqlite3.connect(str(tmp_path))
        with dst:
            src.backup(dst)
        dst.close(); dst = None
        src.close(); src = None

        chk = sqlite3.connect(str(tmp_path))
        ok = chk.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
        chk.close(); chk = None
        if not ok:
            raise RuntimeError("Backup integrity check failed")

        os.replace(str(tmp_path), str(final_path))
        prune_backups()
        logger.info(f"Backup OK: {final_path}")
        return str(final_path)
    except Exception:
        try:
            if tmp_path and tmp_path.exists(): tmp_path.unlink()
        except Exception: pass
        raise
    finally:
        for c in (src, dst, chk):
            if c is not None:
                try: c.close()
                except Exception: pass

# --- Load / Export ---
@st.cache_data(ttl=15, show_spinner=False)
def load_table(table_name):
    try:
        with get_db() as conn:
            df = pd.read_sql(f"SELECT * FROM {qi(table_name)}", conn)
        for num_col in NUMERIC_HINTS.get(table_name, set()):
            if num_col in df.columns:
                df[num_col] = pd.to_numeric(df[num_col], errors="coerce")
        return df
    except Exception as e:
        logger.error(f"load {table_name}: {e}")
        return pd.DataFrame()

def get_db_version_token():
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*), COALESCE(MAX(last_modified),'') FROM rsl")
            r = cur.fetchone()
            cur.execute("SELECT COUNT(*), COALESCE(MAX(last_modified),'') FROM congestion")
            c = cur.fetchone()
            return f"{r}_{c}"
    except Exception: return "v0"

def build_excel_bytes(df_rsl=None, df_cong=None, skip_empty=True):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        wrote = False
        if df_rsl is not None and (not skip_empty or not df_rsl.empty):
            df_rsl.drop(columns=["id","record_id","revision"], errors="ignore").to_excel(w, sheet_name="RSL", index=False)
            wrote = True
        if df_cong is not None and (not skip_empty or not df_cong.empty):
            df_cong.drop(columns=["id","record_id","revision"], errors="ignore").to_excel(w, sheet_name="Congestion", index=False)
            wrote = True
        if not wrote:
            pd.DataFrame({"Info": ["No Data"]}).to_excel(w, sheet_name="Empty", index=False)
        _sanitize_excel_formulas(w.book)
    out.seek(0); return out.getvalue()

@st.cache_data(ttl=60, show_spinner=False)
def export_full_excel_cached(version_token):
    return build_excel_bytes(load_table("rsl"), load_table("congestion"), skip_empty=False)

# --- Save edits (FIX P1-B: INTERNAL_COLS) ---
def save_edits_atomic(table_name, edited_df, original_df, role, username, link_col, allow_delete=False):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    upd = ins = dele = 0
    conflicts = []
    with get_db(immediate=True) as conn:
        cur = conn.cursor()
        data_cols = [c for c in original_df.columns if c not in INTERNAL_COLS]
        if role == "ADMIN" and allow_delete:
            orig_ids = set(original_df["record_id"].dropna().astype(str))
            edit_ids = set(edited_df["record_id"].dropna().astype(str))
            for del_id in (orig_ids - edit_ids):
                orig = original_df[original_df["record_id"] == del_id].iloc[0]
                orig_rev = int(orig["revision"]) if not pd.isna(orig["revision"]) else 1
                cur.execute(f"DELETE FROM {qi(table_name)} WHERE record_id=? AND revision=?",
                            (del_id, orig_rev))
                if cur.rowcount == 0:
                    conflicts.append(f"Cannot delete **{orig.get(link_col, del_id)}**: row modified by another user.")
                else:
                    log_audit(cur, username, role, table_name, orig.get(link_col, del_id), "DELETE", "ALL", "existed", "deleted")
                    dele += 1
        for _, row in edited_df.iterrows():
            rec_id = row.get("record_id")
            if pd.isna(rec_id) or str(rec_id).strip() == "":
                if role != "ADMIN": continue
                if normalize_val(row.get(link_col)) == "": continue
                cols = ["record_id","revision","last_modified","modified_by"] + data_cols
                vals = [str(uuid.uuid4()), 1, now_str, username] + \
                       [sanitize_val_for_sqlite(row.get(c)) for c in data_cols]
                _sql_insert(cur, table_name, cols, vals)
                log_audit(cur, username, role, table_name, row.get(link_col, "New"),
                          "INSERT", "ALL", "", "added")
                ins += 1
                continue
            rec_id = str(rec_id)
            m = original_df[original_df["record_id"] == rec_id]
            if m.empty: continue
            orig = m.iloc[0]
            check_cols = data_cols if role == "ADMIN" else EDITABLE_MATRIX.get((role, table_name), [])
            changed = [(c, orig[c], row[c]) for c in check_cols
                       if c in edited_df.columns and normalize_val(row[c]) != normalize_val(orig[c])]
            if not changed: continue
            orig_rev = int(orig["revision"]) if not pd.isna(orig["revision"]) else 1
            set_parts = [f"{qi(c)} = ?" for c,_,_ in changed] + [
                '"revision" = "revision" + 1',
                '"last_modified" = ?', '"modified_by" = ?']
            params = [sanitize_val_for_sqlite(v) for _,_,v in changed] + \
                     [now_str, username, rec_id, orig_rev]
            cur.execute(f"UPDATE {qi(table_name)} SET {', '.join(set_parts)} "
                        f"WHERE record_id=? AND revision=?", params)
            if cur.rowcount == 0:
                meta = cur.execute(f"SELECT modified_by,last_modified FROM {qi(table_name)} WHERE record_id=?",
                                   (rec_id,)).fetchone()
                mod_by = meta[0] if meta else "another user"
                mod_at = meta[1] if meta else "now"
                conflicts.append(f"Link **{row.get(link_col, rec_id)}** modified by `{mod_by}` ({mod_at}). Refresh and retry.")
                continue
            for c, ov, nv in changed:
                log_audit(cur, username, role, table_name, row.get(link_col, rec_id), "UPDATE", c, ov, nv)
            upd += 1
    load_table.clear()
    return upd, ins, dele, conflicts

# --- Smart Merge (FIX P1-A: _smart_equal) ---
BLANK_POLICIES = {
    "Preserve existing value (recommended)": "preserve",
    "Clear existing value (blank wins)":      "clear",
    "Skip row if any non-comment field is blank": "skip_row",
}

def smart_merge_excel(uploaded_file, merge_mode, username, blank_policy="preserve", role="ADMIN"):
    require_admin(role, "smart_merge_excel")
    create_atomic_backup()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    xls = pd.ExcelFile(uploaded_file)
    report = []
    if "RSL" not in xls.sheet_names and "Congestion" not in xls.sheet_names:
        return ["⚠️ File must contain 'RSL' or 'Congestion' sheet."]
    with get_db(immediate=True) as conn:
        cur = conn.cursor()
        for sheet, tbl, link_col, comment_col in [
            ("RSL", "rsl", "Link Name", RSL_COMMENT),
            ("Congestion", "congestion", "Link name", CONG_COMMENT),
        ]:
            if sheet not in xls.sheet_names: continue
            new_df = sanitize_df_for_sqlite(pd.read_excel(uploaded_file, sheet_name=sheet))
            new_df = _normalize_columns_against_db(conn, tbl, new_df)
            if link_col not in new_df.columns:
                ci = {c.lower(): c for c in new_df.columns}
                if link_col.lower() in ci:
                    new_df = new_df.rename(columns={ci[link_col.lower()]: link_col})
                else:
                    report.append(f"❌ Sheet '{sheet}' missing column '{link_col}'.")
                    continue
            new_df = new_df[new_df[link_col].notna() & (new_df[link_col].astype(str).str.strip() != "")]
            ensure_columns_from_df(conn, tbl, new_df)
            dup = new_df[link_col].duplicated().sum()
            if dup > 0:
                report.append(f"⚠️ '{sheet}': {dup} duplicate link names — keeping latest.")
                new_df = new_df.drop_duplicates(subset=[link_col], keep="last")
            cur.execute(f"PRAGMA table_info({qi(tbl)})")
            db_cols = [c[1] for c in cur.fetchall()]
            data_cols = [c for c in db_cols if c not in INTERNAL_COLS]
            if merge_mode == "Full Replace":
                cur.execute(f"DELETE FROM {qi(tbl)}")
                for _, row in new_df.iterrows():
                    cols = ["record_id","revision","last_modified","modified_by"] + data_cols
                    vals = [str(uuid.uuid4()), 1, now_str, username] + \
                           [sanitize_val_for_sqlite(row.get(c)) if c in new_df.columns else None
                            for c in data_cols]
                    _sql_insert(cur, tbl, cols, vals)
                log_audit(cur, username, "ADMIN", tbl, "ALL", "REPLACE_ALL", "ALL", "", f"{len(new_df)}")
                report.append(f"✅ '{sheet}': replaced with {len(new_df)} rows.")
            else:
                added = updated = unchanged = skipped = 0
                for _, nrow in new_df.iterrows():
                    link = nrow[link_col]
                    ex = cur.execute(f"SELECT record_id FROM {qi(tbl)} WHERE {qi(link_col)} = ? LIMIT 1",
                                     (link,)).fetchone()
                    if ex:
                        rec_id = ex[0]
                        cur.execute(f"SELECT * FROM {qi(tbl)} WHERE record_id = ?", (rec_id,))
                        rowobj = cur.fetchone()
                        old = dict(zip([d[0] for d in cur.description], rowobj)) if rowobj else {}
                        final_vals = {}; skip = False
                        for c in data_cols:
                            new_raw = nrow.get(c) if c in new_df.columns else None
                            is_blank = normalize_val(new_raw) == ""
                            if is_blank:
                                if c == comment_col or blank_policy == "preserve":
                                    final_vals[c] = old.get(c)
                                elif blank_policy == "clear":
                                    final_vals[c] = None
                                elif blank_policy == "skip_row":
                                    if c != comment_col: skip = True; break
                                    final_vals[c] = old.get(c)
                            else:
                                final_vals[c] = sanitize_val_for_sqlite(new_raw)
                        if skip: skipped += 1; continue
                        # FIX P1-A: numeric + CI smart comparison
                        changed = any(not _smart_equal(final_vals.get(c), old.get(c)) for c in data_cols)
                        if changed:
                            set_parts, params = [], []
                            for c in data_cols:
                                set_parts.append(f"{qi(c)} = ?"); params.append(final_vals.get(c))
                            set_parts += ['"revision" = "revision" + 1',
                                          '"last_modified" = ?', '"modified_by" = ?']
                            params += [now_str, username, rec_id]
                            cur.execute(f"UPDATE {qi(tbl)} SET {', '.join(set_parts)} WHERE record_id = ?", params)
                            updated += 1
                        else:
                            unchanged += 1
                    else:
                        if blank_policy == "skip_row" and any(
                            normalize_val(nrow.get(c)) == "" for c in data_cols if c != comment_col
                        ):
                            skipped += 1; continue
                        cols = ["record_id","revision","last_modified","modified_by"] + data_cols
                        vals = [str(uuid.uuid4()), 1, now_str, username] + \
                               [sanitize_val_for_sqlite(nrow.get(c)) if c in new_df.columns else None
                                for c in data_cols]
                        _sql_insert(cur, tbl, cols, vals)
                        added += 1
                log_audit(cur, username, "ADMIN", tbl, "ALL", "SMART_MERGE", "ALL", "",
                          f"add={added} upd={updated} same={unchanged} skip={skipped}")
                report.append(f"✅ '{sheet}' [{blank_policy}]: {added} added, {updated} updated, "
                              f"{unchanged} unchanged, {skipped} skipped.")
    load_table.clear()
    return report

# --- FIX P0-3: Restore exports audit first ---
def restore_backup(sel_bkp, pwd_confirm, username, role):
    require_admin(role, "restore_backup")
    safe_name = Path(sel_bkp).name
    if safe_name != sel_bkp or ".." in safe_name or "/" in safe_name or "\\" in safe_name:
        raise ValueError(f"Invalid backup filename: {sel_bkp}")
    p = (BACKUP_DIR / safe_name).resolve()
    if p.parent != BACKUP_DIR.resolve():
        raise ValueError(f"Path traversal attempt blocked: {sel_bkp}")
    if not p.exists():
        raise FileNotFoundError(f"Backup not found: {sel_bkp}")
    chk = sqlite3.connect(str(p))
    try:
        if chk.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise RuntimeError("Backup integrity check failed.")
        u = chk.execute("SELECT password_hash,salt FROM users WHERE username=?", (username,)).fetchone()
        if not u: raise RuntimeError("Your account is NOT in the backup. Restore would lock you out.")
        if not verify_password(pwd_confirm, u[0], u[1]):
            raise RuntimeError("Password does not match account in backup.")
        if chk.execute("SELECT COUNT(*) FROM users WHERE role='ADMIN'").fetchone()[0] == 0:
            raise RuntimeError("Backup has no admin. Restore would lock everyone out.")
    finally:
        chk.close()

    # FIX P0-3: export current audit BEFORE restoring
    audit_pre_name = "none"
    try:
        with get_db() as conn:
            a_df = pd.read_sql("SELECT * FROM audit_log ORDER BY id", conn)
        if not a_df.empty:
            ts_str = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            uid = uuid.uuid4().hex[:8]
            audit_pre_path = LOG_DIR / f"audit_pre_restore_{ts_str}_{uid}.csv"
            with open(audit_pre_path, "x", encoding="utf-8", newline="") as f:
                w = csv.writer(f)
                w.writerow(a_df.columns.tolist())
                for _, row in a_df.iterrows():
                    w.writerow([_neutralize_formula(str(v)) if pd.notna(v) else "" for v in row])
            audit_pre_name = audit_pre_path.name
    except Exception as e:
        raise RuntimeError(f"Cannot export audit before restore — aborting: {e}")

    create_atomic_backup()
    src = sqlite3.connect(str(p)); dst = sqlite3.connect(DB_PATH)
    with dst: src.backup(dst)
    dst.close(); src.close()
    init_db()
    with get_db(immediate=True) as conn:
        conn.execute("DELETE FROM sessions")
        log_audit(conn.cursor(), username, "ADMIN", "SYSTEM", sel_bkp, "DB_RESTORE", "ALL",
                  f"audit_pre_export={audit_pre_name}", "restored")
    load_table.clear(); st.cache_data.clear()

# --- Audit deletion ---
def delete_audit_rows(ids_to_delete, username, role, password):
    require_admin(role, "delete_audit_rows")
    if not ids_to_delete:
        return 0, "No entries selected."
    try:
        requested_ids = sorted({int(x) for x in ids_to_delete if x is not None})
    except (TypeError, ValueError):
        return 0, "Invalid entry IDs."
    if not requested_ids:
        return 0, "No entries selected."

    with get_db() as conn:
        u = conn.execute("SELECT password_hash, salt FROM users WHERE username=?",
                         (username,)).fetchone()
    if not u or not verify_password(password, u[0], u[1]):
        return 0, "Password incorrect."

    req_placeholders = ",".join(["?"] * len(requested_ids))
    forensic_path = None

    try:
        with get_db(immediate=True) as conn:
            cur = conn.cursor()
            cur.execute(f"""
                SELECT id, timestamp, username, role, table_name, link_name,
                       action, column_changed, old_value, new_value
                FROM audit_log WHERE id IN ({req_placeholders}) ORDER BY id
            """, requested_ids)
            rows_to_delete = cur.fetchall()

            if not rows_to_delete:
                conn.execute("ROLLBACK;")
                return 0, "No matching entries found."

            protected = [r for r in rows_to_delete if r[6] == "AUDIT_DELETE"]
            if protected:
                conn.execute("ROLLBACK;")
                protected_ids = [r[0] for r in protected]
                preview = protected_ids[:10]
                return 0, (f"Cannot delete {len(protected)} meta-audit entries "
                           f"(IDs: {preview}{'...' if len(protected_ids) > 10 else ''}). "
                           f"AUDIT_DELETE entries are protected.")

            actual_ids = [int(r[0]) for r in rows_to_delete]
            actual_placeholders = ",".join(["?"] * len(actual_ids))

            ts_str = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            unique_suffix = uuid.uuid4().hex[:8]
            forensic_path = LOG_DIR / f"audit_deleted_{ts_str}_{unique_suffix}.csv"

            try:
                _write_csv_safe(
                    forensic_path,
                    ["id","timestamp","username","role","table_name",
                     "link_name","action","column_changed","old_value","new_value"],
                    rows_to_delete,
                    mode="x",
                )
            except FileExistsError:
                conn.execute("ROLLBACK;")
                return 0, "Forensic backup collision — please retry."
            except Exception as e:
                try:
                    if forensic_path and forensic_path.exists(): forensic_path.unlink()
                except Exception: pass
                conn.execute("ROLLBACK;")
                return 0, f"Cannot delete: forensic backup failed ({e}). Deletion aborted."

            meta_summary = (f"deleted_ids={','.join(map(str, actual_ids[:50]))}"
                            + (f"...(+{len(actual_ids)-50})" if len(actual_ids) > 50 else ""))
            cur.execute("""
                INSERT INTO audit_log (timestamp, username, role, table_name, link_name,
                                       action, column_changed, old_value, new_value)
                VALUES (?, ?, ?, 'audit_log', ?, 'AUDIT_DELETE', 'rows', ?, ?)
            """, (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                  username, role,
                  f"deleted_{len(rows_to_delete)}_entries",
                  meta_summary,
                  f"forensic_backup={forensic_path.name}"))

            cur.execute(f"DELETE FROM audit_log WHERE id IN ({actual_placeholders})", actual_ids)
            deleted_count = cur.rowcount

        return deleted_count, None
    except Exception:
        try:
            if forensic_path and forensic_path.exists(): forensic_path.unlink()
        except Exception: pass
        raise

# =============================================================================
# SESSION STATE + RESTORE
# =============================================================================
for k, d in [("logged_in", False), ("last_activity", datetime.now()),
             ("flash_msg", None), ("flash_conflict", []),
             ("generated_pwd", ""), ("db_for_download", None),
             ("session_token", None), ("_last_db_touch", datetime.now())]:
    if k not in st.session_state:
        st.session_state[k] = d

_flush_pending_cookie_op()

if not st.session_state.logged_in:
    _cookie_sid = _read_cookie(SESSION_COOKIE_NAME)
    _url_sid    = _get_url_sid()
    _candidate  = _cookie_sid or _url_sid
    if _candidate:
        _row = _validate_session(_candidate)
        if _row:
            st.session_state.logged_in       = True
            st.session_state.username        = _row[0]
            st.session_state.role            = _row[1]
            st.session_state.name            = _row[2]
            st.session_state.session_version = _row[3]
            st.session_state.must_change_pwd = bool(_row[4])
            st.session_state.last_activity   = datetime.now()
            st.session_state.session_token   = _candidate
            st.session_state._last_db_touch  = datetime.now()
            if _cookie_sid and _url_sid:
                _del_url_sid()
        else:
            _queue_delete_cookie(SESSION_COOKIE_NAME)
            _del_url_sid()

# --- Bootstrap ---
with get_db() as conn:
    user_count = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]

if user_count == 0:
    st.markdown("<h2 style='text-align:center;'>🛡️ First-Run Protected Setup</h2>", unsafe_allow_html=True)
    if not BOOTSTRAP_TOKEN:
        st.error(
            "🔐 **Set `HUAWEI_BOOTSTRAP_TOKEN` env variable first, then restart.**\n\n"
            "**PowerShell:**\n"
            "```powershell\n"
            "$env:HUAWEI_BOOTSTRAP_TOKEN=\"MySecret2026\"\n"
            "streamlit run app.py\n"
            "```\n\n"
            "**Or double-click `run.bat`** (in the same folder as `app.py`)."
        )
        st.stop()
    c1, c2, c3 = st.columns([1, 1.4, 1])
    with c2:
        with st.form("bootstrap"):
            tok   = st.text_input("Setup Token", type="password")
            adm_u = st.text_input("Admin username", value="admin")
            adm_n = st.text_input("Full name", value="BO TX Admin")
            p1    = st.text_input("Password (12+, mixed case, digit, special)", type="password")
            p2    = st.text_input("Confirm", type="password")
            if st.form_submit_button("Create Admin & Activate", use_container_width=True):
                if not hmac.compare_digest(tok.strip().encode("utf-8"), BOOTSTRAP_TOKEN.encode("utf-8")):
                    st.error("❌ Invalid token.")
                else:
                    ok, msg = password_policy(p1)
                    if not ok:     st.error(f"❌ {msg}")
                    elif p1 != p2: st.error("❌ Passwords do not match.")
                    else:
                        ph, ps = hash_password(p1)
                        cu = adm_u.strip().lower()
                        with get_db(immediate=True) as conn:
                            cur = conn.cursor()
                            cur.execute("""
                                INSERT INTO users (username,password_hash,salt,role,full_name,must_change_pwd,session_version)
                                SELECT ?,?,?,'ADMIN',?,0,1
                                WHERE NOT EXISTS (SELECT 1 FROM users WHERE role='ADMIN')
                            """, (cu, ph, ps, adm_n.strip()))
                            if cur.rowcount == 0:
                                st.error("❌ Admin already created. Refresh and login."); st.stop()
                            log_audit(cur, cu, "ADMIN", "users", cu, "BOOTSTRAP_ADMIN", "ALL", "", "created")
                        st.session_state.flash_msg = "✅ Admin created. Remove HUAWEI_BOOTSTRAP_TOKEN."
                        st.rerun()
    st.stop()

# --- Validate session ---
if st.session_state.logged_in:
    now = datetime.now()
    if now - st.session_state.last_activity > timedelta(minutes=SESSION_TIMEOUT_MINUTES):
        _revoke_session(st.session_state.get("session_token"))
        _queue_delete_cookie(SESSION_COOKIE_NAME); _del_url_sid()
        st.session_state.logged_in = False
        st.warning("⏰ Session expired (30 min idle). Please log in again.")
        st.stop()
    st.session_state.last_activity = now
    with get_db() as conn:
        cu = conn.execute("SELECT role,session_version,must_change_pwd FROM users WHERE username=?",
                          (st.session_state.username,)).fetchone()
    if not cu or cu[1] != st.session_state.get("session_version", 1) or cu[0] != st.session_state.get("role"):
        _revoke_session(st.session_state.get("session_token"))
        _queue_delete_cookie(SESSION_COOKIE_NAME); _del_url_sid()
        st.session_state.logged_in = False
        st.warning("🔒 Session invalidated (password/role changed). Please log in again.")
        st.stop()
    if now - st.session_state.get("_last_db_touch", now - timedelta(hours=1)) > timedelta(seconds=60):
        _touch_session(st.session_state.get("session_token"))
        st.session_state._last_db_touch = now

# --- Login ---
if not st.session_state.logged_in:
    if st.session_state.flash_msg:
        st.success(st.session_state.flash_msg); st.session_state.flash_msg = None
    st.markdown("<h2 style='text-align:center;'>🌐 Huawei TX Intranet</h2>", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        last = st.session_state.get("last_tried_user", "")
        if last:
            locked, rem = check_db_lockout(last)
            if locked:
                st.warning(f"🔒 Account `{last}` locked. Retry in **{rem}s**.")
        with st.form("login"):
            u_in = st.text_input("Username")
            p_in = st.text_input("Password", type="password")
            if st.form_submit_button("Log in", use_container_width=True):
                cu = u_in.strip().lower()
                st.session_state.last_tried_user = cu
                locked, rem = check_db_lockout(cu)
                if locked:
                    st.error(f"🔒 Locked. Wait {rem}s.")
                else:
                    with get_db() as conn:
                        row = conn.execute(
                            "SELECT password_hash,salt,role,full_name,must_change_pwd,session_version "
                            "FROM users WHERE username=?", (cu,)).fetchone()
                    th = row[0] if row else DUMMY_HASH
                    ts = row[1] if row else DUMMY_SALT
                    ok = verify_password(p_in, th, ts)
                    if ok and row:
                        record_login_result(cu, True, True)
                        sess_token = _create_session(cu, session_version=row[5])
                        st.session_state.logged_in       = True
                        st.session_state.username        = cu
                        st.session_state.role            = row[2]
                        st.session_state.name            = row[3]
                        st.session_state.must_change_pwd = bool(row[4])
                        st.session_state.session_version = row[5]
                        st.session_state.last_activity   = datetime.now()
                        st.session_state.session_token   = sess_token
                        st.session_state._last_db_touch  = datetime.now()
                        _queue_write_cookie(SESSION_COOKIE_NAME, sess_token, SESSION_TTL_HOURS * 3600)
                        _set_url_sid(sess_token)
                        st.rerun()
                    else:
                        fails = record_login_result(cu, False, user_exists=(row is not None))
                        rem = max(0, MAX_LOGIN_ATTEMPTS - fails)
                        if rem == 0:
                            st.error(f"❌ Too many attempts. Locked for {LOCKOUT_MINUTES} min.")
                        else:
                            st.error(f"❌ Invalid credentials. Attempts left: {rem}")
    st.stop()

# --- Force password change ---
if st.session_state.get("must_change_pwd", False):
    st.warning("🔑 You must change your temporary password.")
    with st.form("force_pwd"):
        np1 = st.text_input("New password", type="password")
        np2 = st.text_input("Confirm", type="password")
        if st.form_submit_button("Change & Continue", type="primary"):
            ok, msg = password_policy(np1)
            if not ok:      st.error(f"❌ {msg}")
            elif np1 != np2: st.error("❌ Do not match.")
            else:
                with get_db() as conn:
                    old = conn.execute("SELECT password_hash,salt FROM users WHERE username=?",
                                       (st.session_state.username,)).fetchone()
                if old and verify_password(np1, old[0], old[1]):
                    st.error("❌ New password must differ from current.")
                else:
                    nh, ns = hash_password(np1)
                    with get_db(immediate=True) as conn:
                        cur = conn.cursor()
                        cur.execute("UPDATE users SET password_hash=?,salt=?,must_change_pwd=0,"
                                    "session_version=session_version+1 WHERE username=?",
                                    (nh, ns, st.session_state.username))
                        new_ver = cur.execute("SELECT session_version FROM users WHERE username=?",
                                              (st.session_state.username,)).fetchone()[0]
                        if st.session_state.get("session_token"):
                            cur.execute("UPDATE sessions SET session_version=? WHERE token=?",
                                        (new_ver, st.session_state.session_token))
                        log_audit(cur, st.session_state.username, st.session_state.role,
                                  "users", st.session_state.username, "FORCE_PWD_CHANGE",
                                  "password_hash", "", "updated")
                    _revoke_all_user_sessions(st.session_state.username,
                                              except_token=st.session_state.session_token)
                    st.session_state.must_change_pwd = False
                    st.session_state.session_version = new_ver
                    st.session_state.flash_msg = "✅ Password updated."
                    st.rerun()
    st.stop()

# --- Sidebar ---
role = st.session_state.role
username = st.session_state.username

if st.session_state.flash_msg:
    st.toast(st.session_state.flash_msg, icon="✅"); st.session_state.flash_msg = None
if st.session_state.flash_conflict:
    for m in st.session_state.flash_conflict:
        st.warning(f"⚠️ Conflict: {m}")
    st.session_state.flash_conflict = []

if DEBUG:
    with st.sidebar.expander("🐛 Debug", expanded=False):
        st.write(f"Streamlit: {st.__version__}")
        st.write(f"Cookie API: {'yes' if _HAS_CONTEXT_COOKIES else 'no'}")
        st.write(f"URL sid: {'allowed' if HX_ALLOW_URL_SID else 'disabled'}")
        st.write(f"Production: {'yes' if HX_PRODUCTION else 'no'}")
        st.write(f"Role: {role}")

st.sidebar.title(f"👋 {st.session_state.name}")
st.sidebar.caption(f"`{username}` | `{role_label(role)}`")
if st.sidebar.button("🔄 Refresh Data", use_container_width=True):
    load_table.clear(); st.rerun()

pages = ["📡 RSL", "🚦 Congestion", "📊 Dashboard"]
if role == "ADMIN": pages.append("⚙️ Admin & Audit")
page = st.sidebar.radio("Menu:", pages)

with st.sidebar.expander("🔑 Change my password"):
    with st.form("self_pwd"):
        op  = st.text_input("Current", type="password")
        np1 = st.text_input("New (12+)", type="password")
        np2 = st.text_input("Confirm", type="password")
        if st.form_submit_button("Change"):
            with get_db() as conn:
                u = conn.execute("SELECT password_hash,salt FROM users WHERE username=?",
                                 (username,)).fetchone()
            if not u or not verify_password(op, u[0], u[1]):
                st.error("❌ Current password incorrect.")
            elif np1 == op:
                st.error("❌ New password must differ from current.")
            elif not password_policy(np1)[0]:
                st.error(f"❌ {password_policy(np1)[1]}")
            elif np1 != np2:
                st.error("❌ Do not match.")
            else:
                nh, ns = hash_password(np1)
                with get_db(immediate=True) as conn:
                    cur = conn.cursor()
                    cur.execute("UPDATE users SET password_hash=?,salt=?,session_version=session_version+1 "
                                "WHERE username=?", (nh, ns, username))
                    v = cur.execute("SELECT session_version FROM users WHERE username=?",
                                    (username,)).fetchone()[0]
                    if st.session_state.get("session_token"):
                        cur.execute("UPDATE sessions SET session_version=? WHERE token=?",
                                    (v, st.session_state.session_token))
                    log_audit(cur, username, role, "users", username,
                              "SELF_PWD_CHANGE", "password_hash", "", "updated")
                _revoke_all_user_sessions(username,
                                          except_token=st.session_state.session_token)
                st.session_state.session_version = v
                st.session_state.flash_msg = "✅ Password changed; other sessions invalidated."
                st.rerun()

st.sidebar.markdown("---")
st.sidebar.download_button(
    "📥 Export full Excel",
    data=export_full_excel_cached(get_db_version_token()),
    file_name=f"Huawei_TX_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True)

if st.sidebar.button("🚪 Log out", use_container_width=True):
    _revoke_session(st.session_state.get("session_token"))
    _queue_delete_cookie(SESSION_COOKIE_NAME); _del_url_sid()
    st.session_state.logged_in = False
    st.session_state.session_token = None
    st.rerun()

# =============================================================================
# DATA SHEET — v12.5
# =============================================================================
def render_data_sheet(table_name, title, link_col, status_col):
    st.header(title)
    df = load_table(table_name)
    if df.empty:
        st.warning("No data yet. Upload Excel from Admin panel.")
        return

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        wil = ["All"] + ci_unique(df["WILAYA"].dropna().tolist()) if "WILAYA" in df.columns else ["All"]
        sel_w = st.selectbox("WILAYA:", wil, key=f"{table_name}_w")
    with c2:
        rs = ["All"] + ci_unique(df["Responsibility"].dropna().tolist()) if "Responsibility" in df.columns else ["All"]
        sel_r = st.selectbox("Responsibility:", rs, key=f"{table_name}_r")
    with c3:
        ow = ["All"] + ci_unique(df["Owner"].dropna().tolist()) if "Owner" in df.columns else ["All"]
        sel_o = st.selectbox("Owner:", ow, key=f"{table_name}_o")
    with c4:
        sts = ["All"] + ci_unique(df[status_col].dropna().tolist()) if status_col in df.columns else ["All"]
        sel_s = st.selectbox(f"{status_col}:", sts, key=f"{table_name}_s")
    with c5:
        q = st.text_input(f"🔎 Search (or by {link_col}):", key=f"{table_name}_q")

    f = df.copy()
    f = _ci_filter(f, "WILAYA", sel_w)
    f = _ci_filter(f, "Responsibility", sel_r)
    f = _ci_filter(f, "Owner", sel_o)
    f = _ci_filter(f, status_col, sel_s)

    if q:
        s_cols = [c for c in f.columns if c not in INTERNAL_COLS]
        ql = q.strip().lower()
        mask = f[s_cols].fillna("").astype(str).apply(lambda r: ql in " ".join(r).lower(), axis=1)
        f = f[mask]

    if role == "ADMIN":
        # FIX P1-B: disable id as well
        disabled = ["id", "record_id", "last_modified", "modified_by"]
        allow_del = st.checkbox("🔓 Enable add/delete rows", value=False, key=f"{table_name}_del")
        col_cfg = {"id": None, "record_id": None, "revision": None}
    else:
        allowed = EDITABLE_MATRIX.get((role, table_name), [])
        disabled = [c for c in f.columns if c not in allowed]
        allow_del = False
        col_cfg = {"id": None, "record_id": None, "revision": None,
                   "last_modified": None, "modified_by": None}

    for numcol, label in [("RSL nearend", "RSL nearend (dBm)"), ("RSL farend", "RSL farend (dBm)")]:
        if numcol in f.columns:
            col_cfg[numcol] = st.column_config.NumberColumn(label, format="%.1f")

    if status_col in f.columns:
        base = ["Pending", "Cleared"]
        existing = ci_unique(f[status_col].dropna().tolist())
        merged, seen = [], set()
        for o in base + existing:
            k = o.casefold()
            if k not in seen:
                seen.add(k); merged.append(o)
        col_cfg[status_col] = st.column_config.SelectboxColumn(
            status_col, options=merged,
            help="Select Pending or Cleared. Existing values shown as-is. "
                 "Filters ignore case; selecting a value is intentional and will be saved.",
        )

    st.caption(f"Showing **{len(f)}** of **{len(df)}** | 💡 Filters are case-insensitive.")
    if len(f) < len(df) and not f.empty:
        st.download_button(
            "📥 Export filtered",
            data=build_excel_bytes(
                f if table_name=="rsl" else None,
                f if table_name=="congestion" else None,
                True),
            file_name=f"Filtered_{table_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    sheet_filter_sig = hashlib.sha256(
        repr((table_name, sel_w, sel_r, sel_o, sel_s, q, allow_del)).encode("utf-8")
    ).hexdigest()[:12]

    edited = st.data_editor(
        f, disabled=disabled, hide_index=True,
        column_config=col_cfg, use_container_width=True,
        num_rows="dynamic" if (role == "ADMIN" and allow_del) else "fixed",
        key=f"{table_name}_editor__{sheet_filter_sig}")

    if st.button(f"💾 Save {table_name.upper()}", type="primary"):
        try:
            u, i, d, conf = save_edits_atomic(table_name, edited, f, role, username, link_col, allow_delete=allow_del)
            st.session_state.flash_msg = f"✅ Saved (updates: {u} | inserts: {i} | deletes: {d})"
            st.session_state.flash_conflict = conf
            st.rerun()
        except Exception as e:
            st.error(f"❌ Save failed: {e}")

# --- Pages ---
if page == "📡 RSL":
    render_data_sheet("rsl", "📡 RSL Degraded", "Link Name", "Status")
elif page == "🚦 Congestion":
    render_data_sheet("congestion", "🚦 Congestion", "Link name", "status")
elif page == "📊 Dashboard":
    st.header("📊 Live Dashboard")
    rsl_df  = load_table("rsl"); cong_df = load_table("congestion")

    # FIX P1-C: strip whitespace before matching
    rsl_status = rsl_df["Status"].fillna("").astype(str).str.strip().str.casefold() if not rsl_df.empty else pd.Series(dtype=str)
    cong_status = cong_df["status"].fillna("").astype(str).str.strip().str.casefold() if not cong_df.empty else pd.Series(dtype=str)
    rp = int((rsl_status == "pending").sum())
    cp = int((cong_status == "pending").sum())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("RSL total", len(rsl_df)); k2.metric("RSL pending", rp)
    k3.metric("Congestion total", len(cong_df)); k4.metric("Congestion pending", cp)
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📡 RSL by Responsibility / Status")
        if not rsl_df.empty:
            view = rsl_df.copy()
            view["Responsibility"] = _ci_group_series(view["Responsibility"]) if "Responsibility" in view.columns else "N/A"
            view["Status"]         = _ci_group_series(view["Status"])         if "Status" in view.columns         else "N/A"
            ct = pd.crosstab(view["Responsibility"], view["Status"])
            st.dataframe(ct, use_container_width=True); st.bar_chart(ct)
    with col2:
        st.subheader("🚦 Congestion by Responsibility / Status")
        if not cong_df.empty:
            view = cong_df.copy()
            view["Responsibility"] = _ci_group_series(view["Responsibility"]) if "Responsibility" in view.columns else "N/A"
            view["status"]         = _ci_group_series(view["status"])         if "status" in view.columns         else "N/A"
            ct = pd.crosstab(view["Responsibility"], view["status"])
            st.dataframe(ct, use_container_width=True); st.bar_chart(ct)

elif page == "⚙️ Admin & Audit" and role == "ADMIN":
    st.header("⚙️ Admin & Audit")
    tab1, tab2, tab3, tab4 = st.tabs([
        "📤 Upload Excel", "👥 Users", "📝 Audit Log", "🗄️ Backup & Restore"])

    with tab1:
        st.subheader("📤 Merge Excel")
        mode = st.radio("Mode:", ["Smart Merge (preserves comments & timestamps)", "Full Replace"])
        confirm = True
        if mode == "Full Replace":
            st.warning("⚠️ Full Replace deletes all existing comments! Automatic backup taken first.")
            confirm = st.checkbox("I understand and accept")
        st.markdown("**Blank cell policy:**")
        policy_label = st.radio("Blank cells in uploaded Excel:",
                                list(BLANK_POLICIES.keys()), index=0)
        policy = BLANK_POLICIES[policy_label]
        up = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])
        if up and st.button("🚀 Run", type="primary", disabled=not confirm):
            with st.spinner("Backing up and merging..."):
                for r in smart_merge_excel(up, mode, username, blank_policy=policy, role=role):
                    st.write(r)

    with tab2:
        st.subheader("👥 Manage Users")
        with get_db() as conn:
            users_df = pd.read_sql("SELECT username,role,full_name,must_change_pwd,session_version,last_seen FROM users", conn)
            locks_df = pd.read_sql("SELECT * FROM login_attempts", conn)
        display_df = users_df.copy()
        if not display_df.empty:
            display_df["role"] = display_df["role"].map(lambda r: ROLE_DISPLAY.get(r, r))
        st.dataframe(display_df, use_container_width=True, hide_index=True)
        if not locks_df.empty:
            st.caption("Login attempts:")
            st.dataframe(locks_df, use_container_width=True, hide_index=True)
        ca, cd = st.columns(2)
        with ca:
            if st.button("🎲 Generate strong password"):
                st.session_state.generated_pwd = generate_strong_password()
            if st.session_state.generated_pwd:
                st.code(st.session_state.generated_pwd)
            with st.form("user_form"):
                nu  = st.text_input("Username")
                np_ = st.text_input("Password (12+)", type="password", value=st.session_state.generated_pwd)
                nr  = st.selectbox("Role", ASSIGNABLE_ROLES, format_func=role_label)
                nn  = st.text_input("Full name")
                fc  = st.checkbox("Force change at next login", value=True)
                ow  = st.checkbox("Overwrite existing account")
                if st.form_submit_button("Save user", type="primary"):
                    cu = nu.strip().lower()
                    existing = users_df["username"].tolist()
                    admins = len(users_df[users_df["role"]=="ADMIN"])
                    ok, msg = password_policy(np_)
                    if not cu or not nn.strip():   st.error("❌ Fill username and full name.")
                    elif not ok:                    st.error(f"❌ {msg}")
                    elif cu in existing and not ow: st.error("⚠️ User exists. Enable overwrite.")
                    elif cu == username and nr != "ADMIN" and admins <= 1:
                        st.error("🚫 Cannot remove last ADMIN.")
                    else:
                        ph, ps = hash_password(np_)
                        with get_db(immediate=True) as conn:
                            cur = conn.cursor()
                            cur.execute("""
                                INSERT INTO users (username,password_hash,salt,role,full_name,
                                                   must_change_pwd,session_version)
                                VALUES (?,?,?,?,?,?,1)
                                ON CONFLICT(username) DO UPDATE SET
                                    password_hash=excluded.password_hash, salt=excluded.salt,
                                    role=excluded.role, full_name=excluded.full_name,
                                    must_change_pwd=excluded.must_change_pwd,
                                    session_version=users.session_version + 1
                            """, (cu, ph, ps, nr, nn.strip(), int(fc)))
                            cur.execute("DELETE FROM login_attempts WHERE username=?", (cu,))
                            log_audit(cur, username, role, "users", cu,
                                      "USER_UPDATE" if cu in existing else "USER_CREATE",
                                      "role/pwd", "", f"role={nr}")
                        _revoke_all_user_sessions(cu)
                        st.session_state.generated_pwd = ""
                        if cu == username and nr != role:
                            _revoke_session(st.session_state.get("session_token"))
                            _queue_delete_cookie(SESSION_COOKIE_NAME); _del_url_sid()
                            st.session_state.logged_in = False
                            st.session_state.flash_msg = "⚠️ Your role changed; re-login."
                            st.rerun()
                        st.session_state.flash_msg = f"✅ Saved `{cu}` (sessions revoked)."
                        st.rerun()
        with cd:
            with st.form("del_form"):
                del_users = [u for u in users_df["username"].tolist() if u != username]
                du = st.selectbox("User:", [""] + del_users)
                b1, b2 = st.columns(2)
                if b1.form_submit_button("🔓 Unlock"):
                    if du:
                        with get_db(immediate=True) as conn:
                            cur = conn.cursor()
                            cur.execute("DELETE FROM login_attempts WHERE username=?", (du,))
                            log_audit(cur, username, role, "login_attempts", du,
                                      "USER_UNLOCK", "lockout_until", "locked", "unlocked")
                        st.session_state.flash_msg = f"✅ Unlocked `{du}`."
                        st.rerun()
                if b2.form_submit_button("🗑️ Delete"):
                    if du:
                        with get_db(immediate=True) as conn:
                            cur = conn.cursor()
                            cur.execute("DELETE FROM users WHERE username=?", (du,))
                            cur.execute("DELETE FROM login_attempts WHERE username=?", (du,))
                            log_audit(cur, username, role, "users", du, "USER_DELETE",
                                      "ALL", "existed", "deleted")
                        _revoke_all_user_sessions(du)
                        st.session_state.flash_msg = f"✅ Deleted `{du}` (sessions revoked)."
                        st.rerun()

    with tab3:
        st.subheader("📝 Audit Log")
        with get_db() as conn:
            a_df = pd.read_sql("SELECT * FROM audit_log ORDER BY id DESC LIMIT 1500", conn)
        if a_df.empty:
            st.info("No entries yet.")
        else:
            fa1, fa2, fa3, fa4 = st.columns(4)
            us  = ["All"] + sorted(a_df["username"].dropna().unique().tolist())
            ts  = ["All"] + sorted(a_df["table_name"].dropna().unique().tolist())
            acs = ["All"] + sorted(a_df["action"].dropna().unique().tolist())
            su  = fa1.selectbox("User:", us)
            st_ = fa2.selectbox("Table:", ts)
            sa  = fa3.selectbox("Action:", acs)
            sf  = fa4.text_input("Search:")
            ff = a_df.copy()
            if su != "All":  ff = ff[ff["username"]   == su]
            if st_ != "All": ff = ff[ff["table_name"] == st_]
            if sa != "All":  ff = ff[ff["action"]     == sa]
            if sf:
                sl = sf.strip().lower()
                ff = ff[ff.fillna("").astype(str).apply(lambda r: sl in " ".join(r).lower(), axis=1)]

            st.dataframe(ff, use_container_width=True, hide_index=True)

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                ff.to_excel(w, sheet_name="Audit", index=False)
                _sanitize_excel_formulas(w.book)
            st.download_button("📥 Export audit log", data=buf.getvalue(),
                file_name=f"audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.markdown("---")
            st.subheader("⚠️ Danger Zone — Delete Audit Entries")
            st.caption(
                "**Deleting audit entries breaks forensic integrity.** Every deletion writes a forensic "
                "CSV to `logs/` and records a protected `AUDIT_DELETE` meta-entry."
            )

            danger_on = st.checkbox("🔓 Enable delete mode (destructive — requires password)",
                                    value=False, key="audit_danger_mode")

            if danger_on:
                st.warning("⚠️ **AUDIT_DELETE meta-entries cannot be deleted**.")

                if ff.empty:
                    st.info("No rows match current filters.")
                else:
                    filter_sig = hashlib.sha256(repr((su, st_, sa, sf)).encode("utf-8")).hexdigest()[:12]
                    editor_key = f"audit_delete_editor__{filter_sig}"

                    ff_for_del = ff.copy()
                    ff_for_del.insert(0, "_delete", False)

                    audit_edit = st.data_editor(
                        ff_for_del, use_container_width=True, hide_index=True,
                        disabled=[c for c in ff.columns],
                        column_config={
                            "_delete": st.column_config.CheckboxColumn(
                                "🗑️ Delete", help="Check to mark this audit row for deletion",
                                default=False)
                        },
                        num_rows="fixed", key=editor_key)

                    checked_ids = set(
                        audit_edit.loc[audit_edit["_delete"] == True, "id"].dropna().astype(int).tolist())
                    orig_ids = set(ff["id"].dropna().astype(int).tolist())
                    remaining_ids = set(audit_edit["id"].dropna().astype(int).tolist())
                    to_delete = sorted(checked_ids | (orig_ids - remaining_ids))

                    if not to_delete:
                        st.info("👆 No rows marked yet.")
                    else:
                        st.error(f"🗑️ **{len(to_delete)} entries marked for deletion**")
                        with st.expander("👁️ Preview entries to be deleted", expanded=True):
                            st.dataframe(ff[ff["id"].isin(to_delete)],
                                         use_container_width=True, hide_index=True)

                        cc1, cc2 = st.columns(2)
                        with cc1:
                            typed = st.text_input('Type exactly **DELETE** (uppercase)',
                                                  key=f"audit_del_type_{filter_sig}",
                                                  placeholder="DELETE")
                        with cc2:
                            pwd = st.text_input("Your password", type="password",
                                                key=f"audit_del_pwd_{filter_sig}")

                        confirm_ok = (typed == "DELETE") and bool(pwd)

                        if st.button(f"🗑️ Permanently delete {len(to_delete)} entries",
                                     type="primary", disabled=not confirm_ok,
                                     use_container_width=True, key=f"audit_del_btn_{filter_sig}"):
                            try:
                                count, err = delete_audit_rows(to_delete, username, role, pwd)
                                if err:
                                    st.error(f"❌ {err}")
                                else:
                                    st.session_state.flash_msg = (
                                        f"✅ Deleted {count} audit entries.")
                                    st.rerun()
                            except Exception as e:
                                st.error(f"❌ Deletion failed: {e}")

    with tab4:
        st.subheader(f"🗄️ Backup & Restore (retention: {BACKUP_RETENTION_DAYS} days)")
        cb1, cb2 = st.columns(2)
        with cb1:
            if st.button("➕ Create backup now"):
                try:
                    p = create_atomic_backup()
                    st.success(f"✅ `{p}`")
                except Exception as e:
                    st.error(f"❌ {e}")
            if st.button("📦 Prepare DB snapshot for download"):
                snap = create_atomic_backup()
                with open(snap, "rb") as f:
                    st.session_state.db_for_download = f.read()
            if st.session_state.get("db_for_download"):
                st.download_button("💾 Download DB snapshot",
                    data=st.session_state.db_for_download,
                    file_name=f"snapshot_{datetime.now().strftime('%Y%m%d_%H%M')}.db",
                    mime="application/octet-stream")
        with cb2:
            bkp_files = sorted([f.name for f in BACKUP_DIR.glob("backup_huawei_tx_*.db")], reverse=True)
            if bkp_files:
                sel = st.selectbox("Restore from backup:", bkp_files)
                st.caption("ℹ️ Current audit log will be exported to `logs/audit_pre_restore_*.csv` before restore.")
                confirm_r = st.checkbox("I confirm restoring this backup")
                pwd_confirm = st.text_input("Your password (final confirmation)", type="password")
                if st.button("⚠️ Restore now", disabled=not confirm_r, use_container_width=True):
                    if not pwd_confirm:
                        st.error("❌ Password required.")
                    else:
                        try:
                            restore_backup(sel, pwd_confirm, username, role)
                            _queue_delete_cookie(SESSION_COOKIE_NAME); _del_url_sid()
                            st.session_state.logged_in = False
                            st.session_state.session_token = None
                            st.session_state.flash_msg = f"✅ Restored `{sel}`. Audit exported. Please log in again."
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ {e}")
            else:
                st.info("No backups yet.")