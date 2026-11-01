#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Import Excel -> SQL Server (web UI chạy local).

Luồng: Kết nối server -> chọn DB + bảng đích -> upload Excel -> map cột -> import.
Import chạy trong 1 transaction: lỗi là rollback sạch.

Chạy:  python app.py   rồi mở http://127.0.0.1:5000
"""

import io
import threading
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

import pandas as pd
from flask import Flask, jsonify, render_template, request

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB

# Tool chạy local cho 1 người dùng -> giữ state trong RAM, không ghi password ra đĩa.
STATE = {"conn": None, "xlsx": None, "df": None, "table": None, "cols": None}
LOCK = threading.Lock()

CHAR_T = {"char", "varchar", "nchar", "nvarchar", "text", "ntext"}
INT_T = {"tinyint", "smallint", "int", "bigint"}
DEC_T = {"decimal", "numeric", "money", "smallmoney"}
FLT_T = {"float", "real"}
DT_T = {"date", "datetime", "datetime2", "smalldatetime"}
SKIP_T = {"timestamp", "rowversion"}  # không insert được
MAX_ERRORS = 30


# ----------------------------------------------------------------------------
# DB
# ----------------------------------------------------------------------------
def pick_driver():
    import pyodbc
    names = [d for d in pyodbc.drivers() if "SQL Server" in d]
    for pref in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"):
        if pref in names:
            return pref
    if names:
        return names[-1]
    raise RuntimeError("Máy chưa cài ODBC Driver 17/18 for SQL Server.")


def connect(database=None, autocommit=True):
    import pyodbc
    c = STATE["conn"]
    parts = [f"DRIVER={{{pick_driver()}}}", f"SERVER={c['server']}",
             f"DATABASE={database or c.get('database') or 'master'}",
             "TrustServerCertificate=yes"]
    if c["user"]:
        parts += [f"UID={c['user']}", f"PWD={c['password']}"]
    else:
        parts.append("Trusted_Connection=yes")
    conn = pyodbc.connect(";".join(parts), timeout=15, autocommit=autocommit)
    conn.setdecoding(pyodbc.SQL_CHAR, encoding="utf-8")
    conn.setdecoding(pyodbc.SQL_WCHAR, encoding="utf-8")
    return conn


def split_table(name):
    """'dbo.CDXE' / 'CDXE' -> (schema, table)"""
    if "." in name:
        s, t = name.split(".", 1)
        return s.strip("[]"), t.strip("[]")
    return "dbo", name.strip("[]")


def fetch_columns(conn, schema, table):
    """Metadata cột đích. Bỏ qua cột computed/rowversion (không insert được)."""
    cur = conn.cursor()
    cur.execute("""
        SELECT c.name, t.name, c.max_length, c.precision, c.scale,
               c.is_nullable, c.is_identity, c.is_computed,
               CASE WHEN c.default_object_id <> 0 THEN 1 ELSE 0 END
        FROM sys.columns c
        JOIN sys.types t ON t.user_type_id = c.system_type_id
        WHERE c.object_id = OBJECT_ID(?)
        ORDER BY c.column_id""", f"[{schema}].[{table}]")
    cols = []
    for r in cur.fetchall():
        name, typ, ml, prec, scale, nullable, ident, computed, has_def = r
        if computed or typ in SKIP_T:
            continue
        if typ in ("nchar", "nvarchar") and ml > 0:
            ml //= 2  # max_length tính bằng byte
        cols.append({"name": name, "type": typ, "max_len": ml, "precision": prec,
                     "scale": scale, "nullable": bool(nullable), "identity": bool(ident),
                     "has_default": bool(has_def)})
    if not cols:
        raise ValueError(f"Bảng [{schema}].[{table}] không tồn tại hoặc không có cột.")
    return cols


# ----------------------------------------------------------------------------
# Convert giá trị Excel -> kiểu cột đích (xử lý các bẫy trong DEV_GUIDE mục 5)
# ----------------------------------------------------------------------------
def is_blank(v):
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(v, str) and v.strip() == ""


def make_converter(col):
    typ, ml = col["type"], col["max_len"]

    def conv(v):
        if is_blank(v):
            # ô trống: cột char NOT NULL -> '' (kiểu Cyber hay dùng), còn lại NULL
            return "" if typ in CHAR_T and not col["nullable"] else None
        if typ in CHAR_T:
            if isinstance(v, float) and v.is_integer():
                v = int(v)  # 7.0 -> "7"
            s = str(v)
            if ml and ml > 0 and len(s) > ml:
                raise ValueError(f"dài {len(s)} > max {ml}")
            return s
        if typ in INT_T:
            d = Decimal(str(v).strip())
            if d != d.to_integral_value():
                raise ValueError(f"'{v}' không phải số nguyên")
            return int(d)
        if typ in DEC_T:
            return Decimal(str(v).strip())  # tránh lệch float nhị phân
        if typ in FLT_T:
            return float(str(v).strip())
        if typ == "bit":
            s = str(v).strip().lower()
            if s in ("1", "true", "x", "yes", "có", "co"):
                return True
            if s in ("0", "false", "no", "không", "khong"):
                return False
            raise ValueError(f"'{v}' không phải bit")
        if typ in DT_T:
            if isinstance(v, pd.Timestamp):
                d = v.to_pydatetime()
            elif isinstance(v, datetime):
                d = v
            elif isinstance(v, date):
                d = datetime(v.year, v.month, v.day)
            else:
                d = pd.to_datetime(str(v).strip(), dayfirst=True).to_pydatetime()
            if typ == "smalldatetime" and not (datetime(1900, 1, 1) <= d <= datetime(2079, 6, 6)):
                raise ValueError(f"{d:%d/%m/%Y} ngoài range smalldatetime (1900..2079-06-06)")
            if typ == "datetime" and d.year < 1753:
                raise ValueError(f"{d:%d/%m/%Y} ngoài range datetime (>=1753)")
            return d.date() if typ == "date" else d
        return v  # kiểu khác (uniqueidentifier, time...) để SQL tự convert

    return conv


def build_rows(df, mapping, cols_by_name, header_row):
    """Convert toàn bộ data. Trả (rows, errors). errors = [{row, col, msg}]."""
    names = list(mapping.keys())
    convs = [make_converter(cols_by_name[n]) for n in names]
    idxs = [mapping[n] for n in names]
    required = [not cols_by_name[n]["nullable"] for n in names]
    rows, errors = [], []
    for i, rec in enumerate(df.itertuples(index=False, name=None)):
        out = []
        for n, f, ix, req in zip(names, convs, idxs, required):
            try:
                v = f(rec[ix])
                if v is None and req:
                    raise ValueError("để trống nhưng cột NOT NULL")
            except (ValueError, InvalidOperation, OverflowError) as e:
                msg = f"'{rec[ix]}' sai định dạng" if isinstance(e, InvalidOperation) else str(e)
                if len(errors) < MAX_ERRORS:
                    errors.append({"row": i + header_row + 1, "col": n, "msg": msg})
                v = None
            out.append(v)
        rows.append(tuple(out))
    return rows, errors


# ----------------------------------------------------------------------------
# Insert
# ----------------------------------------------------------------------------
def insert_sql(schema, table, names):
    cl = ", ".join(f"[{n}]" for n in names)
    return f"INSERT INTO [{schema}].[{table}] ({cl}) VALUES ({', '.join('?' * len(names))})"


def run_insert(conn, schema, table, names, rows, truncate, ident_on, batch, fast):
    cur = conn.cursor()
    cur.fast_executemany = fast
    full = f"[{schema}].[{table}]"
    try:
        if truncate:
            cur.execute(f"DELETE FROM {full}")
        if ident_on:
            cur.execute(f"SET IDENTITY_INSERT {full} ON")
        sql = insert_sql(schema, table, names)
        for i in range(0, len(rows), batch):
            cur.executemany(sql, rows[i:i + batch])
        if ident_on:
            cur.execute(f"SET IDENTITY_INSERT {full} OFF")
        conn.commit()  # chỉ commit 1 lần ở cuối
    except Exception:
        conn.rollback()
        raise


def locate_bad_row(schema, table, names, rows, truncate, ident_on, header_row):
    """Chạy lại từng dòng (rồi rollback) để tìm dòng gây lỗi đầu tiên."""
    conn = connect(STATE["conn"]["database"], autocommit=False)
    try:
        cur = conn.cursor()
        full = f"[{schema}].[{table}]"
        if truncate:
            cur.execute(f"DELETE FROM {full}")
        if ident_on:
            cur.execute(f"SET IDENTITY_INSERT {full} ON")
        sql = insert_sql(schema, table, names)
        for i, r in enumerate(rows):
            try:
                cur.execute(sql, r)
            except Exception as e:
                return {"row": i + header_row + 1, "msg": str(e)}
        return None
    finally:
        conn.rollback()
        conn.close()


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
def load_df(sheet, header_row):
    df = pd.read_excel(io.BytesIO(STATE["xlsx"]), sheet_name=sheet,
                       header=header_row - 1, dtype=object)  # dtype=object: giữ nguyên 007, biển số...
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")  # bỏ dòng trống hoàn toàn
    return df


def err(msg, code=400, **extra):
    return jsonify({"ok": False, "error": str(msg), **extra}), code


# ----------------------------------------------------------------------------
# API
# ----------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


@app.post("/api/connect")
def api_connect():
    d = request.json
    STATE["conn"] = {"server": d.get("server", "").strip(), "user": d.get("user", "").strip(),
                     "password": d.get("password", ""), "database": None}
    if not STATE["conn"]["server"]:
        return err("Chưa nhập Server Name")
    try:
        with connect("master") as conn:
            rows = conn.cursor().execute(
                "SELECT name FROM sys.databases WHERE state = 0 AND HAS_DBACCESS(name) = 1 "
                "AND database_id > 4 ORDER BY name").fetchall()
        return jsonify({"ok": True, "databases": [r[0] for r in rows]})
    except Exception as e:
        STATE["conn"] = None
        return err(f"Không kết nối được: {e}")


@app.post("/api/tables")
def api_tables():
    if not STATE["conn"]:
        return err("Chưa kết nối server")
    db = request.json.get("database")
    try:
        with connect(db) as conn:
            rows = conn.cursor().execute(
                "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES "
                "WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY 1, 2").fetchall()
        STATE["conn"]["database"] = db
        return jsonify({"ok": True, "tables": [f"{r[0]}.{r[1]}" for r in rows]})
    except Exception as e:
        return err(e)


@app.post("/api/upload")
def api_upload():
    """Upload file (lần đầu) hoặc đổi sheet/header (gửi lại không kèm file)."""
    f = request.files.get("file")
    if f:
        STATE["xlsx"] = f.read()
    if not STATE["xlsx"]:
        return err("Chưa có file")
    try:
        sheets = pd.ExcelFile(io.BytesIO(STATE["xlsx"])).sheet_names
        sheet = request.form.get("sheet") or sheets[0]
        header_row = max(1, int(request.form.get("header_row") or 1))
        STATE["df"] = load_df(sheet, header_row)
        STATE["header_row"] = header_row
    except Exception as e:
        return err(f"Đọc Excel lỗi (chỉ hỗ trợ .xlsx): {e}")
    df = STATE["df"]
    prev = [[None if is_blank(v) else str(v) for v in r] for r in df.head(5).itertuples(index=False, name=None)]
    return jsonify({"ok": True, "sheets": sheets, "sheet": sheet, "header_row": header_row,
                    "columns": list(df.columns), "n_rows": len(df), "preview": prev})


@app.post("/api/columns")
def api_columns():
    """Metadata bảng đích + gợi ý mapping tự động theo tên."""
    if not STATE["conn"] or STATE["df"] is None:
        return err("Thiếu kết nối hoặc file Excel")
    schema, table = split_table(request.json["table"])
    try:
        with connect(STATE["conn"]["database"]) as conn:
            cols = fetch_columns(conn, schema, table)
    except Exception as e:
        return err(e)
    STATE["table"], STATE["cols"] = (schema, table), cols
    lookup = {}
    for i, c in enumerate(STATE["df"].columns):
        lookup.setdefault(c.lower(), i)  # trùng tên header -> lấy cột đầu
    for c in cols:
        c["suggest"] = lookup.get(c["name"].lower())
    return jsonify({"ok": True, "columns": cols})


@app.post("/api/import")
def api_import():
    """body: {mapping: {table_col: excel_idx}, truncate, identity_insert, dry_run}"""
    d = request.json
    if STATE["df"] is None or not STATE["cols"]:
        return err("Chưa đủ dữ liệu (file / bảng đích)")
    if not LOCK.acquire(blocking=False):
        return err("Đang có 1 lần import chạy")
    try:
        schema, table = STATE["table"]
        cols_by_name = {c["name"]: c for c in STATE["cols"]}
        ident_on = bool(d.get("identity_insert"))
        mapping = {k: int(v) for k, v in d["mapping"].items()
                   if v not in (None, "") and k in cols_by_name
                   and (ident_on or not cols_by_name[k]["identity"])}  # mặc định bỏ cột identity
        if not mapping:
            return err("Chưa map cột nào")
        missing = [c["name"] for c in STATE["cols"]
                   if c["name"] not in mapping and not (c["nullable"] or c["has_default"] or c["identity"])]
        if missing:
            return err("Cột bắt buộc (NOT NULL, không default) chưa map: " + ", ".join(missing))
        ident_on = ident_on and any(cols_by_name[k]["identity"] for k in mapping)

        df, hr = STATE["df"], STATE["header_row"]
        rows, errors = build_rows(df, mapping, cols_by_name, hr)
        if errors:
            return err(f"Dữ liệu sai kiểu — chưa ghi gì (hiện tối đa {MAX_ERRORS} lỗi đầu)", errors=errors)
        if d.get("dry_run"):
            return jsonify({"ok": True, "dry_run": True, "rows": len(rows)})

        names = list(mapping.keys())
        truncate, batch = bool(d.get("truncate")), 1000
        try:
            conn = connect(STATE["conn"]["database"], autocommit=False)
        except Exception as e:
            return err(f"Không kết nối được: {e}")
        try:
            try:
                run_insert(conn, schema, table, names, rows, truncate, ident_on, batch, fast=True)
            except Exception:
                # fast_executemany hay lỗi truncation với nvarchar -> thử lại chế độ chậm mà chắc
                run_insert(conn, schema, table, names, rows, truncate, ident_on, batch, fast=False)
        except Exception as e:
            bad = None
            try:
                bad = locate_bad_row(schema, table, names, rows, truncate, ident_on, hr)
            except Exception:
                pass
            return err(f"Lỗi khi import — đã ROLLBACK, bảng không đổi. {e}", bad_row=bad)
        finally:
            conn.close()
        return jsonify({"ok": True, "rows": len(rows)})
    finally:
        LOCK.release()


if __name__ == "__main__":
    import webbrowser
    threading.Timer(1.0, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
