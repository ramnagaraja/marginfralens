#!/usr/bin/env python3
"""
cloud_cost_optimizer.py - Cloud Cost Optimization & Analysis Tool  (MARG ERP / MARG Books)

Reads three folders and nothing else:
    Azure Cost Logs/    one workbook per month (daily cost per resource)
    Webber Cost Logs/   one WebberStop bill per month (.xls) with per-server rows and shared lines
    Zabbix Logs/        Zabbix audit exports: vm_finops_audit_hourly_trends (CPU / memory / disk per host per hour),
                        failures / alert events, and the executive summary

Run:   streamlit run cloud_cost_optimizer.py
The first start installs any missing Python packages (needs internet once).
"""
from __future__ import annotations

import sys
import subprocess
import importlib.util


def _bootstrap():
    need = {"pandas": "pandas", "numpy": "numpy", "openpyxl": "openpyxl", "xlrd": "xlrd", "plotly": "plotly",
            "streamlit": "streamlit"}
    missing = [pip for mod, pip in need.items() if importlib.util.find_spec(mod) is None]
    if not missing:
        return
    print("First run: installing", ", ".join(missing), "...", flush=True)
    for extra in ([], ["--user"], ["--break-system-packages"]):
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", *extra, *missing])
            return
        except Exception:
            continue
    print("Could not install automatically. Run: pip install " + " ".join(missing), flush=True)


_bootstrap()

import os
import io
import warnings
import re
import glob
import shutil
import calendar
import html
import pickle
import hashlib
import threading
import datetime as dt
from concurrent.futures import ThreadPoolExecutor

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore", message="Could not infer format")

APP_TITLE = "Cloud Cost Optimization & Analysis"
FOLDER_NAMES = {"azure": "Azure Cost Logs", "webber": "Webber Cost Logs", "zabbix": "Zabbix Logs"}
DEFAULT_YEAR = 2026
DEFAULT_AZURE_DISCOUNT = 0.15     # partner discount; gross Azure files vs the old net sheet implied ~14.5%
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else ".", ".costtool_cache")


def disk_cached(key, fn):
    """Cache a pure function result on disk. key must contain file path, size and mtime."""
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        p = os.path.join(CACHE_DIR, hashlib.md5(repr(key).encode()).hexdigest() + ".pkl")
        if os.path.exists(p):
            with open(p, "rb") as f:
                return pickle.load(f)
    except Exception:
        p = None
    out = fn()
    try:
        if p:
            with open(p, "wb") as f:
                pickle.dump(out, f)
    except Exception:
        pass
    return out


def file_sig(path):
    s = os.stat(path)
    return (os.path.abspath(path), s.st_size, int(s.st_mtime))


def inr(v, short=True):
    """Indian-rupee formatting: lakh / crore for large numbers."""
    try:
        v = float(v)
    except Exception:
        return "-"
    if v != v:
        return "-"
    a = abs(v)
    sgn = "-" if v < 0 else ""
    if not short:
        return f"{sgn}₹{a:,.0f}"
    if a >= 1e7:
        return f"{sgn}₹{a / 1e7:.2f} Cr"
    if a >= 1e5:
        return f"{sgn}₹{a / 1e5:.2f} L"
    if a >= 1e3:
        return f"{sgn}₹{a / 1e3:.1f} K"
    return f"{sgn}₹{a:,.0f}"


def find_dir(base, name):
    """Find a sub-folder by name ignoring case; base itself if it already is that folder."""
    if not base or not os.path.isdir(base):
        return None
    if os.path.basename(os.path.normpath(base)).lower() == name.lower():
        return base
    for d in sorted(os.listdir(base)):
        p = os.path.join(base, d)
        if os.path.isdir(p) and d.lower() == name.lower():
            return p
    return None


def month_label(p):
    """'2026-04' -> 'Apr 26'"""
    y, m = str(p).split("-")
    return f"{MONTHS[int(m) - 1]} {y[2:]}"
# ==============================================================================
# 2. COST LOGS  (Azure daily resource cost, WebberStop monthly per-server bills)
# ==============================================================================
_IP_RE = re.compile(r"\b\d{1,3}\.\d{1,3}\.\d{1,3}(?:\.\d{1,3})?(?::\d+)?\b")


def host_keys(name) -> set:
    """Several normalised spellings of a server name so 'WEB-MCLOUDS-ERP12', 'WEB-MCLOUDS-12' and
    'ADC-SRV-192.168.4.235' meet their log / bill spellings. Exact matches only - no prefix guessing."""
    h = _IP_RE.sub(" ", str(name).lower().strip())
    a = re.sub(r"\(.*?(\)|$)|\[.*?(\]|$)", " ", h)
    b = re.sub(r"[()\[\]]", " ", h)
    out = set()
    for x in (a, b):
        t = re.findall(r"[a-z0-9]+", x)
        while t and t[-1] in ("w", "l"):
            t = t[:-1]
        s = "".join(t).replace("mcloudserp", "mclouds")
        if s:
            out.add(s)
            out.add("|".join(sorted(t)))
    return out


def load_host_aliases(folder) -> dict:
    """Optional host_aliases.csv (cost_name,zabbix_name) in the Zabbix Logs folder for names that cannot be matched."""
    try:
        p = os.path.join(folder, "host_aliases.csv")
        if os.path.exists(p):
            a = pd.read_csv(p)
            c1, c2 = a.columns[:2]
            return {str(r[c1]).strip().lower(): str(r[c2]).strip() for _, r in a.dropna().iterrows()}
    except Exception:
        pass
    return {}


_MON = {m.lower(): i + 1 for i, m in enumerate(MONTHS)}


def _azure_date(v, year):
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        return pd.Timestamp(v)
    m = re.match(r"^\s*(\d{1,2})[-/ ]([A-Za-z]{3})", str(v))
    if m and m.group(2).lower() in _MON:
        return pd.Timestamp(year, _MON[m.group(2).lower()], int(m.group(1)))
    return pd.NaT


def _read_azure_file(path, year):
    xl = pd.ExcelFile(path)
    best = None
    for sh in xl.sheet_names:
        head = xl.parse(sh, nrows=3)
        cols = {str(c).strip().lower(): c for c in head.columns}
        if "resource" in cols and "cost" in cols and "date" in cols:
            best = sh
            break
    if best is None:
        return pd.DataFrame()
    df = xl.parse(best)
    cols = {str(c).strip().lower(): c for c in df.columns}
    g = lambda *n: next((cols[x] for x in n if x in cols), None)
    out = pd.DataFrame({
        "date": df[g("date")].map(lambda v: _azure_date(v, year)),
        "resource": df[g("resource")].astype(str).str.strip(),
        "rtype": df[g("resource type")].astype(str).str.strip().str.lower() if g("resource type") else "unknown",
        "rg": df[g("resource group name", "resource group")].astype(str).str.strip().str.lower() if g("resource group name", "resource group") else "",
        "product": df[g("product")].astype(str).str.strip() if g("product") else "",
        "group": df[g("grouping")].astype(str).str.strip() if g("grouping") else "",
        "cost": pd.to_numeric(df[g("cost")], errors="coerce"),
    })
    out = out.dropna(subset=["date", "cost"])
    out["group"] = out["group"].replace({"nan": "", "None": ""})
    return out.groupby(["date", "resource", "rtype", "rg", "product", "group"], as_index=False)["cost"].sum()


def parse_azure_folder(folder, year=DEFAULT_YEAR, progress=None):
    """-> (daily DataFrame, notes). Each workbook is read by header names; the cost-group column is
    missing in some months, so it is filled from the same resource's group in other months."""
    files = sorted(glob.glob(os.path.join(folder, "*.xls*")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    parts, notes = [], []
    for i, f in enumerate(files):
        try:
            d = disk_cached(("azure", year) + file_sig(f), lambda f=f: _read_azure_file(f, year))
            if len(d):
                parts.append(d)
            else:
                notes.append(f"{os.path.basename(f)}: no sheet with Date / Resource / Cost columns - skipped.")
        except Exception as e:
            notes.append(f"{os.path.basename(f)}: could not be read ({type(e).__name__}).")
        if progress:
            progress((i + 1) / max(len(files), 1))
    if not parts:
        return pd.DataFrame(columns=["date", "month", "resource", "rtype", "rg", "product", "group", "cost"]), notes + ["No Azure cost workbooks found."]
    df = pd.concat(parts, ignore_index=True)
    known = df[df["group"] != ""].groupby("resource")["group"].agg(lambda s: s.value_counts().index[0])
    miss = df["group"] == ""
    df.loc[miss, "group"] = df.loc[miss, "resource"].map(known)
    for col in ("product", "rg"):                       # still unknown: use the group most often seen with the same product / resource group
        kn = df[(df["group"] != "") & df["group"].notna() & (df[col] != "")].groupby(col)["group"].agg(lambda s: s.value_counts().index[0])
        m2 = df["group"].isna() | (df["group"] == "")
        df.loc[m2, "group"] = df.loc[m2, col].map(kn)
    df["group"] = df["group"].fillna("UNMAPPED")
    df["month"] = df["date"].dt.strftime("%Y-%m")
    df = df.groupby(["date", "month", "resource", "rtype", "rg", "product", "group"], as_index=False)["cost"].sum()
    return df, notes


# ------------------------------------------------------------------------------ WebberStop
def _w_kind(item: str, cat: str):
    s = item.lower()
    n = re.search(r"\((\d[\d,]*)\s*(vms?|endpoints?|gb)", s)
    cnt = float(n.group(1).replace(",", "")) if n else np.nan
    if "cloudstack" in s or cat.upper() == "K8":
        return "Kubernetes (CloudStack)", cnt
    if "backup storage" in s:
        return "Backup storage", cnt
    if "moc backup" in s:
        return "Backup - MOC", cnt
    if "standard backup" in s:
        return "Backup - standard", cnt
    if "antivirus" in s:
        return "Antivirus", cnt
    if "firewall" in s or "fortinet" in s:
        return "Firewall", np.nan
    if "object storage" in s:
        return "Object storage", np.nan
    if "licen" in s:
        return "Licence", np.nan
    if "additional backup" in s:
        return "Backup - additional", np.nan
    if re.search(r"\dcore", s):
        return "Ad-hoc server line", np.nan
    return "Other shared", np.nan


def _read_webber_file(path):
    import xlrd
    b = xlrd.open_workbook(path)
    s = b.sheet_by_index(0)
    hdr = [str(x).strip().lower() for x in s.row_values(0)]
    ix = {h: i for i, h in enumerate(hdr)}
    di = ix.get("40% discount on invoice")
    need = ["category", "server_name", "ip", "hostname", "from_date", "to_date", "cpu", "ram", "hdd", "cost", "cost*60%"]
    miss = [n for n in need if n not in ix]
    if miss:
        raise ValueError("missing columns " + ", ".join(miss))

    def dte(v):
        try:
            return xlrd.xldate_as_datetime(float(v), b.datemode)
        except Exception:
            return pd.NaT

    def num(v):
        try:
            return float(v)
        except Exception:
            return np.nan

    srv, sh = [], []
    for r in range(1, s.nrows):
        v = s.row_values(r)
        if str(v[ix["hdd"]]).strip().upper() == "TOTAL" or not any(str(x).strip() for x in v):
            continue
        f, t = dte(v[ix["from_date"]]), dte(v[ix["to_date"]])
        gross, net = num(v[ix["cost"]]), num(v[ix["cost*60%"]])
        cpu = num(v[ix["cpu"]])
        base = {"category": str(v[ix["category"]]).strip().upper(), "start": f, "end": t, "gross": gross, "net": net,
                "disc": str(v[di]).strip() if di is not None else ""}
        if cpu == cpu and cpu > 0:
            srv.append({**base, "server_name": str(v[ix["server_name"]]).strip(), "ip": str(v[ix["ip"]]).strip(),
                        "host": str(v[ix["hostname"]]).strip(), "vcpu": cpu, "ram": num(v[ix["ram"]]), "hdd": num(v[ix["hdd"]])})
        else:
            item = str(v[ix["server_name"]]).strip()
            kind, cnt = _w_kind(item, base["category"])
            if cpu == 0:
                kind = "Public IP / add-on"
            sh.append({**base, "item": item, "kind": kind, "count": cnt, "ip": str(v[ix["ip"]]).strip()})
    sv, sd = pd.DataFrame(srv), pd.DataFrame(sh)
    for d in (sv, sd):
        if len(d):
            d["month"] = d["end"].dt.strftime("%Y-%m")
            d["days"] = (d["end"] - d["start"]).dt.days + 1
    return sv, sd


def parse_webber_folder(folder, progress=None):
    """-> (servers, shared, notes). servers = one row per bill segment (a mid-month resize makes two rows);
    net = the 'cost*60%' column = what Marg actually pays. TOTAL rows are excluded."""
    files = [f for f in sorted(glob.glob(os.path.join(folder, "*.xls*"))) if not os.path.basename(f).startswith("~$")]
    sv, sd, notes = [], [], []
    for i, f in enumerate(files):
        try:
            a, b = disk_cached(("webber",) + file_sig(f), lambda f=f: _read_webber_file(f))
            if len(a):
                sv.append(a)
            if len(b):
                sd.append(b)
        except Exception as e:
            notes.append(f"{os.path.basename(f)}: could not be read ({type(e).__name__}: {e}).")
        if progress:
            progress((i + 1) / max(len(files), 1))
    S = pd.concat(sv, ignore_index=True) if sv else pd.DataFrame()
    D = pd.concat(sd, ignore_index=True) if sd else pd.DataFrame()
    if not len(S):
        notes.append("No WebberStop bills found.")
    else:
        S["rate"] = S["net"] / S["days"].clip(lower=1) * 30      # monthly-equivalent net rate of the segment
        S["keys"] = S["host"].map(lambda h: set().union(*[host_keys(x) for x in re.split(r"\s*/\s*", str(h))]))
    return S, D, notes
# ==============================================================================
# 2b. SAVED DATA STORE  (one clean copy of all cost and Zabbix files; every save replaces the previous one)
# ==============================================================================
STORE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else ".", "data_store")
STORE_EXT = {"azure": (".xlsx", ".xlsm", ".xls"), "webber": (".xls", ".xlsx"), "zabbix": (".csv",)}
STORE_META = os.path.join(STORE_DIR, "saved.json")


def store_path(kind):
    return os.path.join(STORE_DIR, FOLDER_NAMES[kind])


def _store_files(kind):
    d = store_path(kind)
    return [os.path.join(r, f) for r, _, fs in os.walk(d) for f in fs] if os.path.isdir(d) else []


def store_info():
    """What is saved: file counts, size, when and from where."""
    import json
    info = {"ready": True, "kinds": {}, "saved_at": None, "source": ""}
    for k in FOLDER_NAMES:
        fl = _store_files(k)
        info["kinds"][k] = {"files": len(fl), "mb": sum(os.path.getsize(f) for f in fl) / 1e6}
        info["ready"] = info["ready"] and len(fl) > 0
    try:
        with open(STORE_META, encoding="utf-8") as f:
            m = json.load(f)
        info["saved_at"], info["source"] = m.get("saved_at"), m.get("source", "")
    except Exception:
        pass
    return info


def active_base(input_base):
    """Saved data when it is complete, otherwise the input folder."""
    return STORE_DIR if store_info()["ready"] else input_base


def store_clear():
    shutil.rmtree(STORE_DIR, ignore_errors=True)
    shutil.rmtree(CACHE_DIR, ignore_errors=True)   # parsed copies of the removed files


def _store_meta(source):
    import json
    os.makedirs(STORE_DIR, exist_ok=True)
    with open(STORE_META, "w", encoding="utf-8") as f:
        json.dump({"saved_at": dt.datetime.now().strftime("%d %b %Y %H:%M"), "source": source}, f)


def _place(src, dst):
    """Hard-link when the folders share a drive (no extra disk space), otherwise copy."""
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    try:
        os.link(src, dst)
    except OSError:
        shutil.copy2(src, dst)


def store_from_folder(base):
    """Replace the saved data with the files of the input folder. Returns (counts, problems)."""
    if os.path.abspath(base) == os.path.abspath(STORE_DIR):
        return None, "The input folder is the saved store itself; point it at the folder holding your source files."
    found = {k: find_dir(base, n) for k, n in FOLDER_NAMES.items()}
    miss = [FOLDER_NAMES[k] for k, d in found.items() if not d]
    if miss:
        return None, "Folder not found: " + ", ".join(miss)
    pick = {}
    for k, d in found.items():
        pick[k] = [os.path.join(r, f) for r, _, fs in os.walk(d) for f in sorted(fs)
                   if f.lower().endswith(STORE_EXT[k]) and not f.startswith("~$")]
    if not all(pick.values()):
        return None, "No usable files in: " + ", ".join(FOLDER_NAMES[k] for k, v in pick.items() if not v)
    store_clear()
    for k, fl in pick.items():
        for f in fl:
            _place(f, os.path.join(store_path(k), os.path.basename(f)))
    _store_meta("input folder")
    return {k: len(v) for k, v in pick.items()}, None


def store_from_uploads(uploads):
    """uploads = {kind: [UploadedFile, ...]}. Replaces everything saved."""
    if not all(uploads.get(k) for k in FOLDER_NAMES):
        return None, "Add files for all three sources first."
    store_clear()
    for k, fl in uploads.items():
        os.makedirs(store_path(k), exist_ok=True)
        for u in fl:
            with open(os.path.join(store_path(k), os.path.basename(u.name)), "wb") as out:
                shutil.copyfileobj(u, out, 1 << 20)
    _store_meta("uploaded files")
    return {k: len(v) for k, v in uploads.items()}, None
# ==============================================================================
# 3. ZABBIX AUDIT EXPORTS  (hourly CPU / memory / disk per host, alert events, audit summary)
# ==============================================================================
_FS_RE = re.compile(r"^FS \[(.*)\]: Space: (Total|Used, in %)$")
_AUD_CPU = ("Linux: CPU utilization", "Windows: CPU utilization")
_AUD_RAM = ("Linux: Memory utilization", "Windows: Memory utilization")
_AUD_RAMTOT = ("Linux: Total memory", "Windows: Total memory")


def _fs_is_data(mount: str) -> bool:
    """Volumes that are billed storage. Boot partitions, tmpfs and the Azure temporary disk are excluded."""
    m = mount.lower().strip()
    if m.startswith(("/boot", "/run", "/dev", "/snap", "/sys", "/proc", "/var/lib/docker/overlay")):
        return False
    if m in ("/mnt", "/mnt/resource") or "temporary storage" in m:
        return False
    return True


def parse_audit_hourly(path: str) -> dict:
    """One hourly_trends CSV -> hourly CPU and memory series, RAM size, and daily disk usage per volume."""
    cols = ["Timestamp", "Host_Name", "Host_IP", "Category", "Metric_Name", "Avg_Usage", "Peak_Usage"]
    cpu, ram, tot, fs = [], [], [], []
    for ch in pd.read_csv(path, usecols=cols, chunksize=1_000_000, dtype={"Host_Name": str, "Host_IP": str, "Metric_Name": str, "Category": str}):
        m = ch["Metric_Name"]
        c = ch[m.isin(_AUD_CPU)]
        if len(c):
            cpu.append(c[["Timestamp", "Host_Name", "Host_IP", "Avg_Usage", "Peak_Usage"]])
        r = ch[m.isin(_AUD_RAM)]
        if len(r):
            ram.append(r[["Timestamp", "Host_Name", "Host_IP", "Avg_Usage", "Peak_Usage"]])
        t = ch[m.isin(_AUD_RAMTOT)]
        if len(t):
            tot.append(t[["Timestamp", "Host_Name", "Host_IP", "Avg_Usage"]])
        f = ch[(ch["Category"] == "Storage") & m.str.endswith(("Space: Total", "Space: Used, in %"))]
        if len(f):
            fs.append(f[["Timestamp", "Host_Name", "Host_IP", "Metric_Name", "Avg_Usage", "Peak_Usage"]])
    out = {}
    for k, parts, extra in (("cpu", cpu, None), ("ram", ram, None)):
        d = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=["Timestamp", "Host_Name", "Host_IP", "Avg_Usage", "Peak_Usage"])
        d = d.rename(columns={"Timestamp": "ts", "Host_Name": "host", "Host_IP": "ip", "Avg_Usage": "avg", "Peak_Usage": "peak"})
        d["ts"] = pd.to_datetime(d["ts"], errors="coerce")
        d["avg"] = d["avg"].astype("float32")
        d["peak"] = d["peak"].astype("float32")
        out[k] = d.dropna(subset=["ts"]).drop_duplicates(["ts", "host"])
    t = pd.concat(tot, ignore_index=True) if tot else pd.DataFrame(columns=["Timestamp", "Host_Name", "Host_IP", "Avg_Usage"])
    t["date"] = pd.to_datetime(t["Timestamp"], errors="coerce").dt.normalize()
    t = t.groupby(["Host_Name", "Host_IP", "date"], as_index=False)["Avg_Usage"].max()
    t["ram_gb"] = t["Avg_Usage"] / 2 ** 30
    out["ramtot"] = t.rename(columns={"Host_Name": "host", "Host_IP": "ip"})[["host", "ip", "date", "ram_gb"]]
    f = pd.concat(fs, ignore_index=True) if fs else pd.DataFrame(columns=["Timestamp", "Host_Name", "Host_IP", "Metric_Name", "Avg_Usage", "Peak_Usage"])
    if len(f):
        x = f["Metric_Name"].str.extract(_FS_RE)
        f["mount"], f["what"] = x[0], x[1]
        f = f.dropna(subset=["mount"])
        f = f[f["mount"].map(_fs_is_data)]
        f["date"] = pd.to_datetime(f["Timestamp"], errors="coerce").dt.normalize()
        tt = f[f["what"] == "Total"].groupby(["Host_Name", "Host_IP", "mount", "date"], as_index=False)["Avg_Usage"].max().rename(columns={"Avg_Usage": "total_bytes"})
        pp = f[f["what"] != "Total"].groupby(["Host_Name", "Host_IP", "mount", "date"], as_index=False).agg(pct_avg=("Avg_Usage", "mean"), pct_max=("Peak_Usage", "max"))
        g = tt.merge(pp, on=["Host_Name", "Host_IP", "mount", "date"], how="outer")
        g["total_gb"] = g["total_bytes"] / 2 ** 30
        out["fs"] = g.rename(columns={"Host_Name": "host", "Host_IP": "ip"})[["host", "ip", "mount", "date", "total_gb", "pct_avg", "pct_max"]]
    else:
        out["fs"] = pd.DataFrame(columns=["host", "ip", "mount", "date", "total_gb", "pct_avg", "pct_max"])
    return out


_EVT_CAT = [("cpu", r"cpu|load average|processor"), ("memory", r"memory|swap|pages/sec"), ("disk", r"disk|space|inode|filesystem|\bfs\b"),
            ("restart", r"restart|uptime"), ("unavailable", r"unavailable|not available|no data|unreachable|ping|icmp|agent"),
            ("service", r"service|process"), ("network", r"interface|network|bandwidth")]


def classify_event(d: str) -> str:
    s = str(d).lower()
    for k, rx in _EVT_CAT:
        if re.search(rx, s):
            return k
    return "other"


def parse_audit_events(path: str) -> pd.DataFrame:
    d = pd.read_csv(path)
    d = d.rename(columns={"Event_ID": "id", "Problem_Start": "start", "Problem_Resolved": "end", "Duration_Minutes": "minutes", "Status": "status",
                          "Host_Name": "host", "Host_IP": "ip", "Severity": "severity", "Failure_Description": "desc", "Acknowledged": "ack"})
    d["start"] = pd.to_datetime(d["start"], errors="coerce")
    d["end"] = pd.to_datetime(d["end"], errors="coerce")
    d["cat"] = d["desc"].map(classify_event)
    return d[["id", "start", "end", "minutes", "status", "host", "ip", "severity", "desc", "ack", "cat"]]


def parse_audit_summary(path: str) -> pd.DataFrame:
    d = pd.read_csv(path)
    return d.rename(columns={"Host_Name": "host", "Host_IP": "ip", "Sizing_Status": "audit_status", "Recommended_Action": "audit_action"})[["host", "ip", "audit_status", "audit_action"]]


def find_audit_files(folder: str) -> dict:
    out = {"hourly": [], "events": [], "summary": []}
    for p in sorted(glob.glob(os.path.join(folder, "**", "*.csv"), recursive=True)):
        b = os.path.basename(p).lower()
        if "hourly_trends" in b:
            out["hourly"].append(p)
        elif "failures" in b:
            out["events"].append(p)
        elif "executive_summary" in b:
            out["summary"].append(p)
    return out


def parse_audit_folder(folder: str, progress=None) -> dict:
    """Read every audit export in the folder (cached per file). Returns hourly cpu / ram, ram size, daily disk, events, latest audit summary."""
    res = {"cpu": pd.DataFrame(), "ram": pd.DataFrame(), "ramtot": pd.DataFrame(), "fs": pd.DataFrame(), "events": pd.DataFrame(),
           "summary": pd.DataFrame(), "files": [], "notes": [], "span": None}
    fl = find_audit_files(folder)
    if not fl["hourly"]:
        res["notes"].append("No vm_finops_audit_hourly_trends CSV files found in the Zabbix Logs folder - utilisation cannot be analysed.")
    parts = {"cpu": [], "ram": [], "ramtot": [], "fs": []}
    n = max(len(fl["hourly"]), 1)
    for i, p in enumerate(fl["hourly"]):
        try:
            r = disk_cached(("audit_h2",) + file_sig(p), lambda p=p: parse_audit_hourly(p))
            for k in parts:
                parts[k].append(r[k])
            res["files"].append({"file": os.path.basename(p), "first": r["cpu"]["ts"].min() if len(r["cpu"]) else pd.NaT,
                                 "last": r["cpu"]["ts"].max() if len(r["cpu"]) else pd.NaT, "hosts": int(r["cpu"]["host"].nunique()), "cpu_rows": len(r["cpu"])})
        except Exception as e:
            res["notes"].append(f"{os.path.basename(p)}: could not be read ({type(e).__name__}: {e}).")
        if progress:
            progress((i + 1) / n)
    for k, v in parts.items():
        v = [x for x in v if len(x)]
        if v:
            d = pd.concat(v, ignore_index=True)
            if k in ("cpu", "ram"):
                d = d.drop_duplicates(["ts", "host"]).sort_values("ts")
            elif k == "fs":
                d = d.drop_duplicates(["host", "mount", "date"], keep="last")
            else:
                d = d.drop_duplicates(["host", "date"], keep="last")
            res[k] = d.reset_index(drop=True)
    ev = []
    for p in fl["events"]:
        try:
            ev.append(disk_cached(("audit_e",) + file_sig(p), lambda p=p: parse_audit_events(p)))
        except Exception as e:
            res["notes"].append(f"{os.path.basename(p)}: could not be read ({type(e).__name__}).")
    if ev:
        res["events"] = pd.concat(ev, ignore_index=True).drop_duplicates("id")
    if fl["summary"]:
        try:
            res["summary"] = parse_audit_summary(fl["summary"][-1])
        except Exception:
            pass
    if len(res["cpu"]):
        res["span"] = (res["cpu"]["ts"].min(), res["cpu"]["ts"].max())
    return res
BASIS_ORDER = ["Bill-evidenced", "Utilisation-evidenced", "Verify first", "Assumption"]
COUNTED = ("Bill-evidenced", "Utilisation-evidenced")
WAVES = {
    1: ("Step 1", "Clean up what is clearly wasted", "This month", "Duplicates, orphaned items, missing discounts and servers that stopped reporting. Mostly a confirmation from an owner or from WebberStop."),
    2: ("Step 2", "Right-size and switch off", "Months 1-3", "Smaller sizes for servers whose measured load allows it, and switch-off for servers nobody uses. One server at a time, after the checks listed on each action."),
    3: ("Step 3", "Move Azure workloads to WebberStop", "Months 3-6", "WebberStop has no ingress or egress charges and lower Windows / SQL Server licence cost. Move the workloads that can move without compliance or re-platforming work."),
    4: ("Step 4", "Protect and fine-tune", "Ongoing", "Fix servers that are short of capacity, put monitoring on the blind spots, and tune what stays on Azure after the moves are decided."),
}
STEP_OF = {"IP_ORPHAN": 1, "OVERLAP": 1, "SILENT": 1, "DISCOUNT": 1, "RATE_CARD": 1, "BACKUP_VMS": 1, "SCHEDULE_W": 1,
           "IDLE": 2, "SCHEDULE": 2, "SQL_MEM": 2, "RS_NONPROD": 2, "RS_NONPROD_V": 2, "RS_PROD": 2, "RS_PROD_V": 2, "RS_DB": 2, "RS_DB_V": 2, "DISK": 2, "K8": 2,
           "MIG_IAAS": 3, "MIG_REVIEW": 3, "MIG_REPLAT": 3, "MIG_SIZES": 3,
           "CAPACITY": 4, "GPU": 2, "MON_GAPS": 4, "BACKUP_STORE": 4, "AZ_COMMIT": 4, "AZ_SNAP": 4}
VALIDATE = {
    "IP_ORPHAN": ["Open each IP in the WebberStop portal and confirm it is not attached to a server, load balancer or firewall rule.", "Ask the app owner whether a DNS record or customer whitelist still points at it."],
    "OVERLAP": ["Confirm the migration cut-over date and that the new side has run production for the agreed validation period.", "Check the last backup of the old server is stored and restorable.", "Check nothing still connects to the old server (Zabbix network traffic, SQL connections, DNS).", "Then stop (not delete) the old server for two weeks before removing it."],
    "SILENT": ["Ask the owner: decommissioned, moved, or running with a broken agent?", "Look at the WebberStop / Azure console: is the server powered on?", "If decommissioned: confirm a final backup, then request billing to stop."],
    "DISCOUNT": ["Compare the invoice lines: the 40% discount column is blank for these servers.", "Ask WebberStop in writing whether the discount applies; request a credit note for past months if it does."],
    "RATE_CARD": ["Compare the invoice lines of servers with identical vCPU, RAM and disk but different prices.", "Ask WebberStop to explain the difference or reprice to the lower rate."],
    "BACKUP_VMS": ["Compare backup-enabled servers on the invoice with the servers that exist today.", "Ask WebberStop to stop backup billing for servers that no longer exist."],
    "SCHEDULE_W": ["Ask WebberStop whether a stopped test server is billed at a lower rate.", "Only if yes: agree start/stop times with the test owners."],
    "IDLE": ["Ask the owner whether anyone or any job still uses the server.", "Check the last login / last scheduled job and the network traffic over the last 90 days.", "Snapshot, stop for two weeks, then delete. Keep the snapshot for the agreed retention."],
    "SCHEDULE": ["Confirm with the test owner that nothing runs overnight or on Sundays (batch, backup, sync).", "Turn on Azure auto-shutdown for one server first and check it restarts cleanly."],
    "SQL_MEM": ["DBA confirms 'max server memory' is currently unlimited.", "Set to about 80% of RAM (no restart) and watch page life expectancy for a week."],
    "RS_NONPROD": ["Owner confirms no load test or monthly run is planned that needs the full size.", "Check the server's peak in Zabbix for the last 30 days (shown in the list).", "Resize one server per window; watch CPU for a week; keep a rollback to the old size."],
    "RS_PROD": ["Owner confirms month-end / billing peaks fall inside the 30-day measurement window.", "Check licences that count by core (Windows / SQL) before the size change.", "Start with a standby or secondary server; resize in a maintenance window; watch CPU and memory for a week."],
    "RS_DB": ["DBA checks wait statistics, IO latency and batch requests per second (CPU % alone does not show SQL health).", "Confirm the licence model before cutting cores.", "Test on a secondary / non-production copy first."],
    "DISK": ["Confirm what fills the disk (data, logs, backups, temp) before shrinking.", "Check the growth trend: only shrink if it is flat or slow.", "Disks cannot shrink in place: plan a new smaller disk and copy."],
    "K8": ["Ask the platform team what runs in the Kubernetes clusters and who asked for them.", "Check node utilisation: are nodes mostly idle?"],
    "MIG_IAAS": ["Get a WebberStop quote for the listed sizes (the figure here is an estimate from current Webber prices).", "Confirm with legal / security that the data in each workload may sit with WebberStop (location, certifications, customer contracts).", "Check each application's licence terms (Windows, SQL Server) for moving to another provider.", "Plan a pilot with one non-production workload first; measure latency to the users and to the databases."],
    "MIG_SIZES": ["List the VM size (vCPU, RAM, disks) of each server from the Azure portal, or install the Zabbix agent.", "Re-run this tool: the workload then gets a WebberStop price."],
    "MIG_REVIEW": ["Legal / compliance sign-off on data location and customer contracts before anything else.", "Then treat as the other moves."],
    "MIG_REPLAT": ["Ask the app team how much work it is to run this on a Webber VM instead of the Azure service.", "Get a WebberStop quote for the VM plus SQL Server licence.", "Compare with the Azure service cost shown here."],
    "CAPACITY": ["Check the alerts listed for each server (CPU, memory, disk) and the last month-end.", "Clean logs and temp first, then expand disks, then scale up the worst offenders.", "Budget for what remains: this is spend to add, not a saving."],
    "GPU": ["Ask the AI / app owner whether the GPU servers are used.", "Check GPU usage in the server's own tools (not in this data)."],
    "MON_GAPS": ["Install the Zabbix agent on each server and register the host.", "Re-run this tool after a week: the server then gets a decision."],
    "BACKUP_STORE": ["List what is backed up and the retention per workload.", "Remove backups of retired servers; shorten retention where the business allows."],
    "AZ_COMMIT": ["Decide first which Azure servers will stay: do not commit for anything that will move.", "Check existing reservations for servers being retired (exchange or refund)."],
    "AZ_SNAP": ["Agree a retention period per workload with the owners.", "Delete snapshots older than that, oldest first."],
}
_DB_RE = re.compile(r"sql|\bdb|db\b|db[-_ ]|mssql|oracle|postgres|mysql|mongo|database", re.I)
_NONPROD_RE = re.compile(r"\b(uat|test|dev|demo|stag\w*|qa|sandbox|poc|migrat\w*)\b|[-_](uat|test|dev|demo|qa|stag\w*)\b|^(uat|test|dev|qa)[-_]", re.I)


def is_db(name):
    return bool(_DB_RE.search(str(name)))


def is_nonprod(name):
    return bool(_NONPROD_RE.search(str(name)))


# ------------------------------------------------------------------------------ cost views
def monthly_table(az, ws, wsh, disc):
    ms = sorted(set(az["month"]) | set(ws["month"]) | set(wsh["month"])) if len(az) or len(ws) else []
    rows = []
    for m in ms:
        a = az.loc[az["month"] == m, "cost"].sum()
        s = ws.loc[ws["month"] == m]
        sh = wsh.loc[wsh["month"] == m, "net"].sum() if len(wsh) else 0.0
        last = s.sort_values("end").groupby("ip").tail(1) if len(s) else s
        rows.append({"month": m, "azure_gross": a, "azure_net": a * (1 - disc), "webber_servers": s["net"].sum(),
                     "webber_shared": sh, "servers": s["ip"].nunique(), "vcpu": last["vcpu"].sum() if len(last) else 0,
                     "ram": last["ram"].sum() if len(last) else 0, "hdd": last["hdd"].sum() if len(last) else 0})
    t = pd.DataFrame(rows)
    if len(t):
        t["webber"] = t["webber_servers"] + t["webber_shared"]
        t["total"] = t["azure_net"] + t["webber"]
        t["cost_per_vcpu"] = t["webber_servers"] / t["vcpu"].replace(0, np.nan)
        t["label"] = t["month"].map(month_label)
    return t


def spend_lines(az, ws, wsh, disc):
    """Every cost line on one footing: provider - area, per month (net)."""
    parts = []
    if len(az):
        a = az.groupby(["group", "month"])["cost"].sum().reset_index()
        a["line"] = "Azure - " + a["group"].astype(str)
        a["value"] = a["cost"] * (1 - disc)
        a["provider"] = "Azure"
        parts.append(a[["line", "provider", "month", "value"]])
    if len(ws):
        w = ws.groupby(["category", "month"])["net"].sum().reset_index()
        w["line"] = "Webber - " + w["category"]
        w["provider"] = "Webber"
        w = w.rename(columns={"net": "value"})
        parts.append(w[["line", "provider", "month", "value"]])
    if len(wsh):
        w = wsh.groupby(["kind", "month"])["net"].sum().reset_index()
        w["line"] = "Webber - " + w["kind"]
        w["provider"] = "Webber"
        w = w.rename(columns={"net": "value"})
        parts.append(w[["line", "provider", "month", "value"]])
    if not parts:
        return pd.DataFrame(), pd.DataFrame()
    L = pd.concat(parts)
    P = L.pivot_table(index=["provider", "line"], columns="month", values="value", aggfunc="sum").fillna(0.0)
    ms = list(P.columns)
    mv = P.copy()
    mv["first"], mv["last"] = P[ms[0]], P[ms[-1]]
    mv["change"] = mv["last"] - mv["first"]
    mv["change_pct"] = np.where(mv["first"] > 0, mv["change"] / mv["first"], np.nan)
    mv["avg"] = P.mean(axis=1)
    x = np.arange(len(ms))
    mv["slope"] = [np.polyfit(x, r.values, 1)[0] if len(ms) > 2 else 0.0 for _, r in P[ms].iterrows()]
    return P, mv.reset_index().sort_values("change", ascending=False)


def azure_daily(az, disc):
    if not len(az):
        return pd.DataFrame()
    d = az.groupby("date")["cost"].sum().mul(1 - disc).rename("net").reset_index()
    d["ma7"] = d["net"].rolling(7, min_periods=3).mean()
    return d


def webber_exit(ws, latest, window_days=30):
    """Servers still billed at the end of the latest month, at their final segment's rate."""
    s = ws[ws["month"] == latest]
    if not len(s):
        return s
    end = s["end"].max()
    last = s.sort_values(["start", "end"]).groupby("ip").tail(1)
    return last[last["end"] >= end - pd.Timedelta(days=2)].copy()


def server_table(ws, latest):
    """One row per public IP across all months."""
    if not len(ws):
        return pd.DataFrame()
    ms = sorted(ws["month"].unique())
    pm = ws.pivot_table(index="ip", columns="month", values="net", aggfunc="sum").reindex(columns=ms)
    ex = webber_exit(ws, latest)
    lastseg = ws.sort_values(["end", "start"]).groupby("ip").tail(1).set_index("ip")
    t = lastseg[["category", "host", "server_name", "vcpu", "ram", "hdd", "disc", "keys"]].copy()
    t["exit_rate"] = ex.set_index("ip")["rate"].reindex(t.index).fillna(0.0)
    t["active"] = t.index.isin(ex["ip"])
    for m in ms:
        t[m] = pm[m]
    t["first_month"] = ws.groupby("ip")["month"].min()
    t["months_billed"] = ws.groupby("ip")["month"].nunique()
    t["resized"] = ws.groupby("ip")[["vcpu", "ram"]].nunique().max(axis=1) > 1
    return t.reset_index()


def peer_price(S, category, vcpu, ram, hdd, ip):
    """Median monthly rate of full-month servers with exactly this vCPU/RAM and similar storage (+/-30%)."""
    p = S[(S["vcpu"] == vcpu) & (S["ram"] == ram) & (S["ip"] != ip)]
    if not len(p):
        return np.nan, 0
    near = p[(p["hdd"] >= hdd * 0.7) & (p["hdd"] <= hdd * 1.3)]
    if len(near) >= 2:
        return float(near["rate"].median()), len(near)
    return np.nan, 0


def fit_price(S):
    d = S[(S["vcpu"] > 0) & S["ram"].notna() & S["hdd"].notna()]
    if len(d) < 20:
        return None
    X = np.c_[d["vcpu"], d["ram"], d["hdd"], np.ones(len(d))]
    co = np.linalg.lstsq(X, d["rate"].values.astype(float), rcond=None)[0]
    p = X @ co
    y = d["rate"].values.astype(float)
    r2 = 1 - ((y - p) ** 2).sum() / ((y - y.mean()) ** 2).sum()
    return {"vcpu": co[0], "ram": co[1], "hdd": co[2], "r2": r2}




# ------------------------------------------------------------------------------ utilisation from the Zabbix audit exports
DEFAULT_CFG = {"azure_discount": DEFAULT_AZURE_DISCOUNT, "window_days": 30, "cpu_p95_max": 35.0, "breach_share": 0.0025, "ram_halve_p99": 42.5,
               "db_floor": 8.0, "idle_pk99": 15.0, "idle_p95": 5.0, "idle_ram_p95": 40.0, "silent_days": 3,
               "ri_pct": 0.25, "snap_pct": 0.20, "backup_pct": 0.10, "platform_pct": 0.15}
CATEGORIES = ["Right-size compute", "Switch off / schedule", "Decommission idle", "Storage", "Commercial", "Duplicates & migration",
              "Backup & protection", "Platform services", "Capacity risk", "Monitoring"]
OFF_SHARE = 1 - 72 / 168        # share of the week outside Mon-Sat 08:00-20:00


def host_metrics(aud, cfg):
    """One row per monitored host: utilisation over the analysis window, RAM size, disk, alerts."""
    cpu, ram = aud["cpu"], aud["ram"]
    end = cpu["ts"].max()
    ws = end - pd.Timedelta(days=cfg["window_days"])
    g_all = cpu.groupby("host")
    H = pd.DataFrame({"ip": g_all["ip"].last(), "first_seen": g_all["ts"].min(), "last_seen": g_all["ts"].max()})
    cw = cpu[cpu["ts"] > ws]
    g = cw.groupby("host")
    H["hours"] = g["avg"].size()
    H["cpu_avg"] = g["avg"].mean()
    H["cpu_p95"] = g["avg"].quantile(0.95)
    H["cpu_pk99"] = g["peak"].quantile(0.99)
    H["cpu_max"] = g["peak"].max()
    H["cpu_gt50"] = cw.assign(b=cw["peak"] > 50).groupby("host")["b"].mean()
    hr, wd = cw["ts"].dt.hour, cw["ts"].dt.weekday
    off = cw[(hr < 8) | (hr >= 20) | (wd == 6)]
    H["off_p95"] = off.groupby("host")["avg"].quantile(0.95)
    rw = ram[ram["ts"] > ws]
    gr = rw.groupby("host")
    H["ram_avg"] = gr["avg"].mean()
    H["ram_p95"] = gr["avg"].quantile(0.95)
    H["ram_p99"] = gr["peak"].quantile(0.99)
    H["ram_max"] = gr["peak"].max()
    rt = aud["ramtot"]
    if len(rt):
        H["ram_gb"] = rt.sort_values("date").groupby("host")["ram_gb"].last()
    fs = aud["fs"]
    if len(fs):
        f = fs[fs["date"] > ws.normalize()]
        last = f.sort_values("date").groupby(["host", "mount"]).tail(1)
        agg = f.groupby(["host", "mount"]).agg(pct_avg=("pct_avg", "mean"), pct_max=("pct_max", "max")).reset_index().merge(last[["host", "mount", "total_gb"]], on=["host", "mount"])
        agg["used_gb"] = agg["total_gb"] * agg["pct_avg"] / 100
        agg["used_pk_gb"] = agg["total_gb"] * agg["pct_max"] / 100
        agg["target_gb"] = np.maximum(agg["used_pk_gb"] * 1.5, 20)
        agg["reclaim_gb"] = np.where((agg["pct_max"] < 50) & (agg["total_gb"] >= 100), np.maximum(agg["total_gb"] - agg["target_gb"], 0), 0.0)
        gg = agg.groupby("host")
        H["alloc_gb"] = gg["total_gb"].sum()
        H["used_gb"] = gg["used_gb"].sum()
        H["disk_max_pct"] = gg["pct_max"].max()
        H["reclaim_gb"] = gg["reclaim_gb"].sum()
        w = agg.sort_values("pct_max", ascending=False).groupby("host").head(1).set_index("host")
        H["worst_mount"] = w["mount"]
    ev = aud["events"]
    if len(ev):
        e = ev[ev["start"] > ws]
        for c in ("cpu", "memory", "disk", "unavailable", "restart"):
            H["ev_" + c] = e[e["cat"] == c].groupby("host").size()
        H["ev_high"] = e[e["severity"].isin(["High", "Disaster"])].groupby("host").size()
        H["ev_all"] = e.groupby("host").size()
    for c in [c for c in H.columns if c.startswith("ev_")]:
        H[c] = H[c].fillna(0).astype(int)
    sm = aud["summary"]
    if len(sm):
        H["audit_status"] = sm.drop_duplicates("host").set_index("host")["audit_status"]
    H = H.reset_index().rename(columns={"index": "host"})
    expected = cfg["window_days"] * 24
    H["coverage"] = (H["hours"].fillna(0) / expected).clip(upper=1.0)
    H["silent_days"] = (end - H["last_seen"]).dt.total_seconds() / 86400
    return H


def match_hosts(H, srv, az, cfg, aliases=None):
    """Link each monitored host to its bill line: Webber by public IP, then exact normalised name; Azure by name."""
    aliases = aliases or {}
    wip = srv.set_index("ip")
    wk = {}
    for _, r in srv.iterrows():
        for k in r["keys"]:
            wk.setdefault(k, set()).add(r["ip"])
    az_vm, az_cost = {}, {}
    if len(az):
        latest = az["month"].max()
        v = az[(az["rtype"] == "virtual machine") & (az["month"] == latest)].groupby("resource")["cost"].sum() * (1 - cfg["azure_discount"])
        for n, c in v.items():
            az_cost[n] = c
            for k in host_keys(n):
                az_vm.setdefault(k, set()).add(n)
    rows = []
    for _, h in H.iterrows():
        r = {"provider": "Unmatched", "bill_id": "", "bill_host": "", "category": "", "vcpu": np.nan, "ram": np.nan, "hdd": np.nan, "cost": np.nan, "active": False}
        name = h["host"]
        nm = host_keys(aliases.get(str(name).lower(), name)) | host_keys(name)
        ips = set()
        if h["ip"] in wip.index:
            ips = {h["ip"]}
        else:
            ips = set().union(*[wk.get(k, set()) for k in nm])
        if len(ips) > 1:
            act = [i for i in ips if bool(wip.loc[i, "active"])]
            ips = set(act) if len(act) == 1 else ips
        if len(ips) == 1:
            ip = next(iter(ips))
            b = wip.loc[ip]
            r.update(provider="Webber", bill_id=ip, bill_host=b["host"], category=b["category"], vcpu=b["vcpu"], ram=b["ram"], hdd=b["hdd"], cost=b["exit_rate"], active=bool(b["active"]))
        elif len(ips) > 1:
            r.update(provider="Ambiguous")
        else:
            azs = set().union(*[az_vm.get(k, set()) for k in nm])
            if len(azs) == 1:
                n = next(iter(azs))
                r.update(provider="Azure", bill_id=n, bill_host=n, category="Azure VM", cost=az_cost.get(n, np.nan), active=True)
        rows.append(r)
    return pd.concat([H.reset_index(drop=True), pd.DataFrame(rows)], axis=1)


def util_trend(aud, H):
    """Month-by-month utilisation of the whole monitored fleet and of each Webber category (from the hourly series)."""
    cpu, ram = aud["cpu"], aud["ram"]
    cat = H.set_index("host")["category"]
    vc = H.set_index("host")["vcpu"]
    c = cpu.assign(month=cpu["ts"].dt.strftime("%Y-%m"))
    hm = c.groupby(["host", "month"]).agg(cpu_avg=("avg", "mean"), cpu_p95=("avg", lambda s: s.quantile(0.95)), cpu_pk=("peak", "max"), hours=("avg", "size")).reset_index()
    r = ram.assign(month=ram["ts"].dt.strftime("%Y-%m")).groupby(["host", "month"]).agg(ram_avg=("avg", "mean"), ram_p95=("avg", lambda s: s.quantile(0.95))).reset_index()
    hm = hm.merge(r, on=["host", "month"], how="left")
    hm["category"] = hm["host"].map(cat).fillna("")
    hm["vcpu"] = hm["host"].map(vc)
    hm["used_vcpu"] = hm["cpu_avg"] / 100 * hm["vcpu"]
    return hm


def rightsize(H, ws, latest, cfg):
    """Decision per host, one action each. Gates: enough data, no pressure alerts, sustained load fits, rare spikes fit, databases stay >= floor."""
    S = ws[(ws["month"] == latest) & (ws["days"] >= 28)]
    fit = fit_price(S)
    out = []
    for _, h in H.iterrows():
        d = h.to_dict()
        d.update(action="", reason="", new_vcpu=np.nan, new_ram=np.nan, saving=0.0, price_basis="", risk="", basis="", is_db=is_db(h["host"]) or is_db(h.get("bill_host", "")),
                 nonprod=is_nonprod(h["host"]) or is_nonprod(h.get("bill_host", "")), at_stake=0.0)
        prov, cost = h["provider"], h["cost"]
        out.append(d)
        if h["coverage"] < 0.6 or h["cpu_p95"] != h["cpu_p95"]:
            d["action"], d["reason"] = ("Silent", f"No readings for {h['silent_days']:.0f} days") if h["silent_days"] > cfg["silent_days"] else ("Skip", "Less than 60% of the window has data")
            continue
        if prov not in ("Webber", "Azure") or not (cost == cost and cost > 0):
            d["action"], d["reason"] = "Skip", "Not on the Webber bill or Azure VM bill (nothing to save)" if prov in ("Unmatched", "Ambiguous") else "Not billed in the latest month"
            continue
        if h["cpu_p95"] >= 80 or (h["ram_p99"] == h["ram_p99"] and h["ram_p99"] >= 97 and not d["is_db"]) or (h.get("disk_max_pct", 0) >= 90):
            d["action"], d["reason"], d["risk"] = "Capacity risk", f"CPU p95 {h['cpu_p95']:.0f}%, memory peak {h['ram_p99']:.0f}%, fullest disk {h.get('disk_max_pct', 0):.0f}% - needs headroom, not a smaller size", "High"
            continue
        if (h["cpu_pk99"] < cfg["idle_pk99"] and h["cpu_p95"] < cfg["idle_p95"] and (h["ram_p95"] != h["ram_p95"] or h["ram_p95"] < cfg["idle_ram_p95"])
                and not d["is_db"] and h.get("ev_all", 0) < 200):
            d["action"], d["basis"], d["risk"] = "Idle", "Verify first", "Medium"
            d["at_stake"] = cost
            d["reason"] = f"Busiest 1% of hours peaked at {h['cpu_pk99']:.0f}% CPU, p95 average {h['cpu_p95']:.1f}%, memory p95 {h['ram_p95']:.0f}%: almost unused over {cfg['window_days']} days"
            continue
        if prov == "Azure" and d["nonprod"] and h["off_p95"] == h["off_p95"] and h["off_p95"] < 5:
            d["action"], d["basis"], d["risk"] = "Schedule", "Utilisation-evidenced", "Low"
            d["saving"] = cost * OFF_SHARE
            d["price_basis"] = "Azure bills VMs per hour"
            d["reason"] = f"Test/UAT server, CPU p95 outside Mon-Sat 08-20 is {h['off_p95']:.1f}%: auto-shutdown saves about {OFF_SHARE:.0%} of its compute cost"
            continue
        v = h["vcpu"]
        small = (prov == "Webber" and (v != v or v < 4)) or (prov == "Azure" and (h["ram_gb"] != h["ram_gb"] or h["ram_gb"] < 16))
        if small:
            d["action"], d["reason"] = "Keep", "Already a small size"
            continue
        if h["cpu_p95"] > cfg["cpu_p95_max"]:
            d["action"], d["reason"] = "Keep", f"Sustained CPU (p95 {h['cpu_p95']:.0f}%) is too high to halve"
            continue
        if h["cpu_gt50"] > cfg["breach_share"]:
            d["action"], d["reason"] = "Hold", f"{h['cpu_gt50'] * h['hours']:.0f} hours in {cfg['window_days']} days peaked above 50% (would pass 100% on half the vCPUs); busiest-hour peak {h['cpu_max']:.0f}%"
            continue
        if d["is_db"] and prov == "Webber" and v / 2 < cfg["db_floor"]:
            if h["ram_avg"] == h["ram_avg"] and h["ram_avg"] >= 70 and h["cpu_p95"] < 10:
                d["action"], d["reason"], d["risk"], d["basis"] = "SQL memory cap", "High memory with low CPU is SQL Server caching; cap 'max server memory' (online setting, no cost)", "Low", "Utilisation-evidenced"
            else:
                d["action"], d["reason"] = "Hold", f"Database below {cfg['db_floor']:.0f} vCPU is not proposed (DBA found 4 vCPU performed poorly)"
            continue
        halve_ram = (not d["is_db"]) and h["ram_p99"] == h["ram_p99"] and h["ram_p99"] <= cfg["ram_halve_p99"]
        if prov == "Azure" and not halve_ram:
            d["action"], d["reason"] = "Hold", "Azure sizes scale vCPU and memory together; memory would not fit at half the size"
            continue
        proj = h["cpu_p95"] * 2
        risk = "Low" if (h["cpu_pk99"] * 2 <= 70 and h["cpu_gt50"] == 0) else "Medium"
        d["risk"] = risk
        if prov == "Azure":
            d["new_vcpu"], d["new_ram"] = np.nan, h["ram_gb"] / 2
            d["saving"], d["price_basis"] = cost * 0.5, "Azure prices scale with size inside a VM series"
            d["action"] = "Right-size database" if d["is_db"] else "Right-size"
            d["basis"] = "Verify first" if d["is_db"] else "Utilisation-evidenced"
        else:
            nv, nr = v / 2, (h["ram"] / 2 if halve_ram else h["ram"])
            d["new_vcpu"], d["new_ram"] = nv, nr
            pp, n = peer_price(S, h["category"], nv, nr, h["hdd"], h["bill_id"])
            if pp == pp and pp < cost:
                d["saving"], d["price_basis"] = cost - pp, f"peer price ({n} servers)"
            elif fit:
                d["saving"], d["price_basis"] = max(0.0, fit["vcpu"] * (v - nv) + fit["ram"] * (h["ram"] - nr)), "estimated from price fit"
            d["action"] = "Right-size database" if d["is_db"] else "Right-size"
            d["basis"] = "Verify first" if (d["is_db"] or not d["price_basis"].startswith("peer") or d["saving"] <= 0) else "Utilisation-evidenced"
        d["reason"] = (f"CPU avg {h['cpu_avg']:.0f}%, p95 {h['cpu_p95']:.0f}% -> {proj:.0f}% at half size; busiest hours peak {h['cpu_pk99']:.0f}% -> {h['cpu_pk99'] * 2:.0f}%"
                       + ("; DBA must check waits / IO latency first" if d["is_db"] else ""))
    return pd.DataFrame(out)


# ------------------------------------------------------------------------------ opportunities
def _opp(id, title, category, services, wave, basis, risk, effort, owner, saving, at_stake, why, how, needs="", items=None, unit="servers"):
    return {"id": id, "title": title, "category": category, "services": services, "wave": wave, "basis": basis, "risk": risk, "effort": effort,
            "owner": owner, "saving": float(saving), "at_stake": float(at_stake), "why": why, "how": how, "needs": needs,
            "items": items if items is not None else pd.DataFrame(), "unit": unit}


_AZ_SERVICES = {
    "disk": ("Azure managed disks", "Check disk tiers (premium vs standard SSD) against real IO, delete disks of retired VMs, and recreate over-sized disks smaller (Azure disks cannot shrink in place)."),
    "storage account": ("Azure storage accounts", "Move cold blobs to cool / archive with lifecycle rules, delete stale containers, and review redundancy (GRS vs LRS) per account."),
    "app service plan": ("Azure App Service plans", "Scale plans in to match load, consolidate apps onto fewer plans, and move steady plans to reserved pricing."),
    "sql server": ("Azure SQL", "Right-size vCores / DTUs against measured load, pause or scale down non-production databases, and use reserved capacity for steady ones."),
    "nat gateway": ("Azure NAT gateways", "Each gateway bills hourly plus data: share one per virtual network / region and remove gateways of retired workloads."),
    "application gateway": ("Azure application gateways", "Lower minimum instance counts / capacity units, use autoscale, and consolidate gateways for low-traffic apps."),
    "public ip address": ("Azure public IPs", "Release addresses not attached to a running resource and consolidate behind gateways."),
    "log analytics workspace": ("Azure Log Analytics", "Cut ingestion of noisy tables, shorten retention, and use the cheaper basic-logs tier for high-volume logs."),
    "virtual machine scale set": ("Azure scale sets", "Lower the minimum instance count and add scale-in rules."),
    "connection": ("Azure VPN connections", "Remove tunnels to sites or clouds that no longer need them (check after any migration)."),
    "virtual network gateway": ("Azure VPN gateways", "Drop to a smaller gateway SKU or retire it once cross-cloud traffic ends."),
    "recovery services vault": ("Azure Backup vaults", "Review retention and remove backups of retired machines."),
}



_IAAS = {"virtual machine", "disk", "snapshot", "virtual machine scale set", "public ip address", "nat gateway", "application gateway", "connection",
         "virtual network gateway", "network security group", "private endpoint", "restore point collection", "recovery services vault", "load balancer", "network interface"}
_REPLAT = {"sql server", "app service plan", "app service", "azure database for postgresql flexible server", "service bus namespace", "api management service"}
_COMPL_RE = re.compile(r"uae|abha|abdm|health|pharma|clinic|hospital|bank|payment|kyc|gov", re.I)


def migration_assessment(M, cfg):
    """Azure workloads (the Azure 'product' tag) that could move to WebberStop, with an indicative Webber price built from Webber's own rates."""
    az, H, srv, latest = M["az"], M["hosts"], M["srv"], M["latest"]
    if not len(az) or not len(H) or not len(srv):
        return pd.DataFrame()
    azc = 1 - cfg["azure_discount"]
    S = M["ws"][(M["ws"]["month"] == latest) & (M["ws"]["days"] >= 28)]
    fit = fit_price(S)
    if not fit:
        return pd.DataFrame()
    c = az[az["month"] == latest]
    prod_of = c[c["rtype"] == "virtual machine"].groupby("resource")["product"].agg(lambda x: x.iloc[0]).to_dict()
    ah = H[H["provider"] == "Azure"].set_index("bill_id")
    rows = []
    for prod, g in c.groupby("product"):
        by = g.groupby("rtype")["cost"].sum() * azc
        vm_cost = by.get("virtual machine", 0.0)
        if vm_cost <= 0:
            continue
        iaas = sum(v for k, v in by.items() if k in _IAAS)
        replat = sum(v for k, v in by.items() if k in _REPLAT)
        other = by.sum() - iaas - replat
        vms = [r for r, p in prod_of.items() if p == prod]
        sized = [r for r in vms if r in ah.index and ah.loc[r, "ram_gb"] == ah.loc[r, "ram_gb"] and ah.loc[r, "alloc_gb"] == ah.loc[r, "alloc_gb"]]
        sized_cost = g[(g["rtype"] == "virtual machine") & g["resource"].isin(sized)]["cost"].sum() * azc
        est = 0.0
        for r in sized:
            ram = float(ah.loc[r, "ram_gb"])
            vc = float(min(max(2, round(ram / 4)), 32))
            est += max(fit["vcpu"] * vc + fit["ram"] * round(ram) + fit["hdd"] * float(ah.loc[r, "alloc_gb"]), 0)
        share = sized_cost / vm_cost if vm_cost else 0
        rows.append({"workload": prod, "vms": len(vms), "sized_vms": len(sized), "sized_share": share, "azure_in_scope": iaas, "azure_replat": replat, "azure_other": other,
                     "webber_est": est if share >= 0.8 else np.nan, "flag": bool(_COMPL_RE.search(str(prod)) or any(_COMPL_RE.search(r) for r in vms)),
                     "hosts": ", ".join(sorted(vms))})
    D = pd.DataFrame(rows)
    if not len(D):
        return D
    over = set()
    for o in M.get("_overlap_names", []):
        over.add(o)
    D["already_moving"] = D["hosts"].map(lambda h: any(x in over for x in h.split(", ")))
    D["saving"] = np.where(D["webber_est"].notna(), (D["azure_in_scope"] - D["webber_est"]).clip(lower=0), np.nan)
    def cls(r):
        if r["already_moving"]:
            return "Already moving"
        if r["sized_share"] < 0.8:
            return "Needs server sizes"
        if r["flag"]:
            return "Compliance check first"
        if r["saving"] <= 0:
            return "No saving"
        return "Can move"
    D["verdict"] = D.apply(cls, axis=1)
    return D.sort_values("azure_in_scope", ascending=False)


def find_opportunities(M, cfg):
    az, ws, wsh, srv, R, T = M["az"], M["ws"], M["wsh"], M["srv"], M["rs"], M["monthly"]
    latest = M["latest"]
    opps = []
    act = srv[srv["active"]] if len(srv) else srv
    azc = 1 - cfg["azure_discount"]

    # ---------------- Wave 0: facts
    if len(R) and len(act):
        seen = set(R[R["provider"] == "Webber"]["bill_id"])
        um = act[~act["ip"].isin(seen)].sort_values("exit_rate", ascending=False)
        M["unmonitored"] = um
        if len(um):
            c = um["exit_rate"].sum()
            opps.append(_opp("MON_GAPS", "Put monitoring on servers that are billed but never measured", "Monitoring", ["Webber compute"], 0, "Enabler", "Low", 1, "DevOps", 0, c,
                             f"{len(um)} billed Webber servers (about {inr(c)}/month) do not appear in the Zabbix audit at all, so nobody can say whether they are used or sized correctly.",
                             "Add the Zabbix agent (and register the host), then re-run this tool after a week of data.",
                             items=um[["host", "category", "vcpu", "ram", "hdd", "exit_rate"]].rename(columns={"exit_rate": "monthly_cost"})))
        sil = R[(R["action"] == "Silent") & (R["cost"] > 0)]
        if len(sil):
            opps.append(_opp("SILENT", "Confirm servers that stopped reporting", "Decommission idle", ["Webber compute", "Azure virtual machines"], 0, "Verify first", "Medium", 1, "DevOps + app owner", 0, sil["cost"].sum(),
                             f"{len(sil)} billed servers sent no readings for {cfg['silent_days']}+ days before the end of the data ({inr(sil['cost'].sum())}/month). Switched off but still billed is the cheapest saving; running but unmonitored is a risk.",
                             "Ask each owner: decommissioned (stop billing), moved (update the bill) or running with a broken agent (fix it).",
                             items=sil[["host", "provider", "category", "last_seen", "silent_days", "cost"]].rename(columns={"cost": "monthly_cost"})))
    if len(az) and len(act):
        v = az[(az["rtype"] == "virtual machine") & (az["month"] == latest)].groupby("resource")["cost"].sum() * azc
        wk = {}
        for _, r in act.iterrows():
            for k in r["keys"]:
                wk.setdefault(k, []).append(r)
        core = lambda s: re.sub(r"^(wvm|vm|web)+", "", re.sub(r"[^a-z0-9]", "", str(s).lower()))
        wcores = [(core(p), r) for _, r in act.iterrows() for p in re.split(r"\s*/\s*", str(r["host"])) if len(core(p)) >= 7]
        hu = R.set_index("bill_id") if len(R) else None
        rows = []
        for name, cst in v.items():
            ms = [r for k in host_keys(name) for r in wk.get(k, [])]
            how = "Same name"
            if not ms:
                ac = core(name)
                ms = [r for c, r in wcores if c in ac]
                how = "Similar name"
            if ms:
                r = ms[0]
                za = R[(R["bill_id"] == name)]["cpu_p95"]
                zw = R[(R["bill_id"] == r["ip"])]["cpu_p95"]
                rows.append({"azure_vm": name, "azure_monthly": cst, "azure_cpu_p95": za.iloc[0] if len(za) else np.nan, "webber_host": r["host"], "webber_monthly": r["exit_rate"],
                             "webber_cpu_p95": zw.iloc[0] if len(zw) else np.nan, "smaller": min(cst, r["exit_rate"]), "match": how})
        if rows:
            O = pd.DataFrame(rows).sort_values("smaller", ascending=False)
            M["_overlap_names"] = list(O["azure_vm"])
            M["_overlap_items"] = O
            opps.append(_opp("OVERLAP", "Retire the old side of servers billed on both clouds", "Duplicates & migration", ["Azure virtual machines", "Webber compute"], 0, "Verify first", "Medium", 1, "Infra lead", 0, O["smaller"].sum(),
                             f"{len(O)} server names appear on both the Azure and the Webber bill in {month_label(latest)}. After a migration cut-over the old side keeps billing until someone switches it off; the value shown is the cheaper side of each pair (conservative); if Azure is the old side, as in the Azure to WebberStop migration plan, stopping it saves the full Azure cost. CPU p95 of both sides is shown so the new server can be checked for size too.",
                             "Confirm per pair: production + DR is fine; a finished migration means decommissioning the old side after the validation period.", items=O, unit="pairs"))

    # ---------------- Utilisation-driven
    if len(R):
        idle = R[R["action"] == "Idle"]
        if len(idle):
            opps.append(_opp("IDLE", "Switch off or archive servers that are almost unused", "Decommission idle", ["Webber compute", "Azure virtual machines"], 1, "Verify first", "Medium", 2, "App owner + DevOps", 0, idle["at_stake"].sum(),
                             f"{len(idle)} servers show a peak under {cfg['idle_pk99']:.0f}% CPU in their busiest 1% of hours, p95 average under {cfg['idle_p95']:.0f}% and low memory over {cfg['window_days']} days. They cost {inr(idle['at_stake'].sum())}/month.",
                             "Ask the owner whether the server still serves anyone; if not, snapshot, stop and later delete. If it is needed but quiet, shrink it instead.",
                             items=idle[["host", "provider", "category", "cpu_avg", "cpu_p95", "cpu_pk99", "ram_p95", "cost", "reason"]].rename(columns={"cost": "monthly_cost"})))
        sch = R[R["action"] == "Schedule"]
        if len(sch):
            opps.append(_opp("SCHEDULE", "Auto-shutdown Azure test / UAT servers outside working hours", "Switch off / schedule", ["Azure virtual machines"], 1, "Utilisation-evidenced", "Low", 1, "DevOps", sch["saving"].sum(), sch["saving"].sum(),
                             f"{len(sch)} Azure test/UAT servers are idle outside Mon-Sat 08:00-20:00 (CPU p95 under 5%). Azure bills virtual machines per hour, so stopping them nights and Sundays saves about {OFF_SHARE:.0%} of the compute cost (disks keep billing).",
                             "Turn on Azure auto-shutdown / start-stop automation for these servers; exclude any that run overnight batch work.",
                             items=sch[["host", "off_p95", "cost", "saving"]].rename(columns={"cost": "monthly_cost"}), unit="VMs"))
        wn = R[(R["provider"] == "Webber") & R["nonprod"] & (R["off_p95"] < 5) & ~R["action"].isin(["Idle", "Capacity risk", "Silent", "Skip"])]
        if len(wn):
            opps.append(_opp("SCHEDULE_W", "Ask WebberStop whether stopped test servers stop billing", "Switch off / schedule", ["Webber compute"], 1, "Verify first", "Low", 1, "Procurement + DevOps", 0, wn["cost"].sum() * OFF_SHARE,
                             f"{len(wn)} Webber test/UAT servers are idle outside working hours. Webber bills a fixed monthly price, so a shutdown schedule only saves money if the contract bills stopped VMs less; the figure is the upper bound if it does.",
                             "Ask WebberStop for hourly or stop-start pricing for non-production servers; otherwise shrink them (see right-sizing).",
                             items=wn[["host", "category", "off_p95", "cost"]].rename(columns={"cost": "monthly_cost"})))
        sql = R[R["action"] == "SQL memory cap"]
        if len(sql):
            opps.append(_opp("SQL_MEM", "Cap SQL Server memory", "Right-size compute", ["Webber compute"], 1, "Utilisation-evidenced", "Low", 1, "DBA", 0, 0,
                             f"{len(sql)} SQL servers show very high memory with low CPU. SQL Server keeps memory it has read; this is caching, not load, and keeps memory alarms firing.",
                             "Set 'max server memory' to about 80% of RAM (online, no restart). No vCPU change.",
                             items=sql[["host", "cpu_avg", "ram_avg", "ram_p99"]]))
        rz = R[R["action"].isin(["Right-size", "Right-size database"])].copy()
        cols = ["host", "provider", "category", "vcpu", "new_vcpu", "ram", "new_ram", "cpu_avg", "cpu_p95", "cpu_pk99", "cost", "saving", "price_basis", "risk", "reason"]
        sets = [("RS_NONPROD", "Right-size test / UAT servers", 1, rz[(~rz["is_db"]) & rz["nonprod"]], "DevOps"),
                ("RS_PROD", "Right-size under-used production servers", 2, rz[(~rz["is_db"]) & ~rz["nonprod"]], "DevOps"),
                ("RS_DB", "Right-size over-sized database servers", 2, rz[rz["is_db"]], "DBA + DevOps")]
        for oid, ttl, wv, d, own in sets:
            if not len(d):
                continue
            counted, verify = d[d["basis"] == "Utilisation-evidenced"], d[d["basis"] != "Utilisation-evidenced"]
            if len(counted):
                opps.append(_opp(oid, ttl, "Right-size compute", ["Webber compute", "Azure virtual machines"], wv, "Utilisation-evidenced", "Low" if (counted["risk"] == "Low").all() else "Medium", 2, own,
                                 counted["saving"].sum(), counted["saving"].sum(),
                                 f"{len(counted)} servers: p95 of hourly CPU is under {cfg['cpu_p95_max']:.0f}% over {cfg['window_days']} days, no more than {cfg['breach_share']:.2%} of hours peaked above 50% (the point where half the vCPUs would saturate), and memory fits. "
                                 "Webber savings use the price of servers already on the bill with the same size; Azure savings assume price scales with size within a VM series.",
                                 "One server per maintenance window, standby/secondary first; watch CPU in Zabbix for a week before the next.", needs="MON_GAPS",
                                 items=counted[cols].rename(columns={"cost": "monthly_cost"})))
            if len(verify):
                opps.append(_opp(oid + "_V", ttl + " - verify first", "Right-size compute", ["Webber compute", "Azure virtual machines"], wv, "Verify first", "Unverified" if oid == "RS_DB" else "Medium", 2, own, 0, verify["saving"].sum(),
                                 ("Databases are never counted until the DBA team has checked wait statistics, IO latency and batch requests per second: CPU % alone does not show SQL health. " if oid == "RS_DB" else
                                  "These resizes pass the CPU test but rely on an estimated price, so they are listed but not counted. ") + f"{len(verify)} servers, up to {inr(verify['saving'].sum())}/month.",
                                 "Review with the owner; promote to the counted set once the check passes.", items=verify[cols].rename(columns={"cost": "monthly_cost"})))
        cap = R[R["action"] == "Capacity risk"]
        if len(cap):
            opps.append(_opp("CAPACITY", "Fix servers that are short of CPU, memory or disk", "Capacity risk", ["Webber compute", "Azure virtual machines"], 1, "Enabler", "High", 2, "DevOps + DBA", 0, 0,
                             f"{len(cap)} servers are running hot: CPU p95 at 80%+, memory above 97%, or a disk over 90% full. Shrinking is out of the question and some will need money added; the audit also records thousands of CPU, memory and disk alerts on them.",
                             "Fix before the next peak (month-end, billing runs): clean logs, expand disks, scale up the worst offenders. Budget for it in this plan.",
                             items=cap[["host", "provider", "category", "cpu_p95", "cpu_max", "ram_p99", "disk_max_pct", "worst_mount", "ev_cpu", "ev_memory", "ev_disk", "reason"]]))
        w = R[(R["provider"] == "Webber") & (R["reclaim_gb"] > 0) & ~R["action"].isin(["Capacity risk"])].copy()
        if len(w) and fit_price(ws[(ws["month"] == latest) & (ws["days"] >= 28)]):
            rate = max(fit_price(ws[(ws["month"] == latest) & (ws["days"] >= 28)])["hdd"], 0)
            w["saving_est"] = w["reclaim_gb"] * rate
            w = w.sort_values("saving_est", ascending=False)
            opps.append(_opp("DISK", "Right-size over-provisioned disks", "Storage", ["Webber compute"], 2, "Verify first", "Medium", 3, "DevOps", 0, w["saving_est"].sum(),
                             f"{len(w)} servers have large volumes (100 GB+) whose fullest point in {cfg['window_days']} days stayed under 50% used. Reclaimable space is {w['reclaim_gb'].sum():,.0f} GB, priced at the bill's own storage rate (about ₹{rate:.1f}/GB/month, fitted across all servers).",
                             "Volumes cannot shrink in place: create a smaller volume, copy, swap. Start with the largest and with data disks (not C: or /).",
                             items=w[["host", "category", "alloc_gb", "used_gb", "disk_max_pct", "reclaim_gb", "saving_est"]].rename(columns={"saving_est": "monthly_saving"}), unit="servers"))

    # ---------------- Commercial and bill-based
    if len(ws):
        s = ws[(ws["month"] == latest) & (ws["disc"] == "No")]
        if len(s):
            nd = s["net"].sum()
            opps.append(_opp("DISCOUNT", "Ask why some servers miss the 40% invoice discount", "Commercial", ["Webber compute"], 1, "Verify first", "Low", 1, "Procurement", 0, nd * 0.4,
                             f"{len(s)} server lines ({', '.join(sorted(s['category'].unique()))}) in the {month_label(latest)} bill are marked 'No' for the 40% discount and are charged at list price ({inr(nd)}). The other {int((ws[(ws['month']==latest)]['disc']=='Yes').sum())} server lines get it.",
                             "Ask WebberStop whether the discount applies under the contract; if yes, request the credit for the months billed.",
                             items=s[["category", "host", "vcpu", "ram", "hdd", "gross"]]))
        S = ws[(ws["month"] == latest) & (ws["days"] >= 28) & (ws["disc"] != "No")]
        rows = []
        for key, grp in S.groupby(["vcpu", "ram", "hdd"]):
            if len(grp) >= 2 and grp["rate"].max() > grp["rate"].median() * 1.05:
                med = grp["rate"].median()
                for _, r in grp.iterrows():
                    if r["rate"] > med * 1.05:
                        rows.append({"host": r["host"], "category": r["category"], "vcpu": key[0], "ram": key[1], "hdd": key[2], "rate": r["rate"], "median_same_spec": med, "excess": r["rate"] - med, "peers": len(grp)})
        if rows:
            Rc = pd.DataFrame(rows).sort_values("excess", ascending=False)
            opps.append(_opp("RATE_CARD", "Align servers of the same size to one price", "Commercial", ["Webber compute"], 1, "Bill-evidenced", "Medium", 1, "Procurement", Rc["excess"].sum(), Rc["excess"].sum(),
                             f"{len(Rc)} servers cost more than others with exactly the same vCPU, RAM and disk in {month_label(latest)}; the excess over the median is {inr(Rc['excess'].sum())}/month.",
                             "Send WebberStop the list and ask for the median rate; some differences may be an operating-system licence.", items=Rc))
        gp = ws[(ws["month"] == latest) & (ws["category"] == "GPU")]
        if len(gp):
            opps.append(_opp("GPU", "Confirm GPU servers are used", "Capacity risk", ["Webber compute"], 1, "Enabler", "Medium", 1, "AI/app owner", 0, 0,
                             f"{len(gp)} GPU servers cost {inr(gp['net'].sum())}/month (new since June, up from {inr(ws[(ws['month']=='2026-06') & (ws['category']=='GPU')]['net'].sum())}). The audit measures CPU and memory only, not GPU load, so GPU use cannot be verified from this data.",
                             "Ask the owners for GPU utilisation (nvidia-smi / Zabbix GPU template) and agree a stop-when-idle rule.", items=gp[["host", "vcpu", "ram", "hdd", "net"]].rename(columns={"net": "monthly_cost"})))
    if len(wsh):
        sh = wsh[wsh["month"] == latest]
        used = 0.0
        if len(M["hosts"]) and "used_gb" in M["hosts"]:
            used = M["hosts"][M["hosts"]["provider"] == "Webber"]["used_gb"].sum()
        bs = sh[sh["kind"] == "Backup storage"]
        if len(bs):
            cost = bs["net"].sum()
            gb = bs["count"].dropna().iloc[0] if bs["count"].notna().any() else np.nan
            ratio = (gb / used) if used and gb == gb else np.nan
            opps.append(_opp("BACKUP_STORE", "Review backup retention and what is backed up", "Backup & protection", ["Webber backup storage"], 3, "Assumption", "Medium", 3, "Infra lead", 0, cost * cfg["backup_pct"],
                             f"Backup storage is the largest shared charge ({inr(cost)}/month, {gb:,.0f} GB at INR 4/GB). " + (f"The monitored Webber servers hold about {used:,.0f} GB of live data, so backup is roughly {ratio:.1f}× the live data; " if ratio == ratio else "")
                             + "that points at long retention, duplicate copies or backups of retired servers.",
                             f"Check retention per workload and remove backups of retired servers; the figure assumes {cfg['backup_pct']:.0%}."))
        for kind, unit_name in (("Backup - standard", "backup"), ("Backup - MOC", "backup")):
            pass
        std = sh[sh["kind"].isin(["Backup - standard", "Backup - MOC"])]
        if len(std) and std["count"].notna().any() and len(act):
            billed = std["count"].sum()
            unit_cost = std["net"].sum() / billed
            extra = billed - len(act)
            if extra > 0:
                opps.append(_opp("BACKUP_VMS", "Trim per-VM backup charges to the servers that exist", "Backup & protection", ["Webber backup (per VM)"], 1, "Verify first", "Low", 1, "Infra lead", 0, extra * unit_cost,
                                 f"The bill charges backup for {billed:.0f} VMs ({inr(unit_cost, False)} each) but only {len(act)} Webber servers are billed in {month_label(latest)}: {extra:.0f} backups may belong to retired servers.",
                                 "Ask WebberStop for the list of backed-up VMs and remove the retired ones."))
        k8 = wsh[wsh["kind"] == "Kubernetes (CloudStack)"].groupby("month")["net"].sum()
        if len(k8) >= 2 and k8.iloc[-1] > k8.iloc[0] * 1.1:
            opps.append(_opp("K8", "Explain Kubernetes (CloudStack) growth", "Platform services", ["Webber Kubernetes"], 3, "Verify first", "Medium", 2, "DevOps", 0, k8.iloc[-1] - k8.iloc[0],
                             f"The CloudStack usage charge grew from {inr(k8.iloc[0])} in {month_label(k8.index[0])} to {inr(k8.iloc[-1])} in {month_label(k8.index[-1])}. The bill does not say which workloads drive it.",
                             "Ask WebberStop for the usage breakdown by namespace / node and set limits on idle workloads."))
        ip = wsh[(wsh["kind"] == "Public IP / add-on") & (wsh["month"] == latest)]
        if len(ip) and len(srv):
            orphan = ip[~ip["ip"].isin(set(act["ip"])) & (ip["ip"] != "")]
            if orphan["net"].sum() >= 1000:
                opps.append(_opp("IP_ORPHAN", "Release public IPs not attached to a billed server", "Platform services", ["Webber public IPs"], 1, "Bill-evidenced", "Low", 1, "DevOps", orphan["net"].sum(), orphan["net"].sum(),
                                 f"{len(orphan)} public-IP / add-on lines ({inr(orphan['net'].sum())}/month) refer to addresses that no billed server uses in {month_label(latest)}.",
                                 "Confirm the address is unused and release it.", items=orphan[["item", "ip", "net"]].rename(columns={"net": "monthly_cost"}), unit="lines"))

    # ---------------- Azure service by service
    if len(az):
        lm = az[az["month"] == latest]
        vm = lm[lm["rtype"] == "virtual machine"]["cost"].sum() * azc
        ri = lm[lm["group"] == "RI"]["cost"].sum() * azc
        if vm > 0:
            opps.append(_opp("AZ_COMMIT", "Cover steady Azure VMs with reservations (only for what stays)", "Commercial", ["Azure virtual machines"], 3, "Assumption", "Low", 2, "Cloud lead", 0, vm * cfg["ri_pct"],
                             f"Azure virtual machines cost {inr(vm)}/month and the bill already carries {inr(ri)} of reservation charges; the on-demand share is not in these files. If workloads are moving to Webber, new commitments and the reservations already held (which keep billing until term end) need care.",
                             f"Commit only for servers that will stay on Azure; exchange or refund reservations of servers being retired. Figure assumes {cfg['ri_pct']:.0%} on the on-demand part.", unit="VMs"))
        sn = lm[lm["rtype"] == "snapshot"]
        if len(sn):
            c = sn["cost"].sum() * azc
            opps.append(_opp("AZ_SNAP", "Tighten snapshot retention", "Storage", ["Azure snapshots"], 3, "Assumption", "Low", 2, "Cloud lead", 0, c * cfg["snap_pct"],
                             f"{sn['resource'].nunique():,} snapshots cost {inr(c)}/month. Almost all were created recently, so retention is rolling rather than hoarding, but it is a top-five storage line.",
                             f"Agree a retention period per workload; the figure assumes {cfg['snap_pct']:.0%}.", unit="snapshots"))
        for rt, (nm, how) in _AZ_SERVICES.items():
            x = lm[lm["rtype"] == rt]
            c = x["cost"].sum() * azc
            if c < 5000:
                continue
            first = az[(az["rtype"] == rt) & (az["month"] == M["months"][0])]["cost"].sum() * azc
            top = x.groupby("resource")["cost"].sum().sort_values(ascending=False).head(15) * azc
            cat = "Storage" if rt in ("disk", "storage account") else "Platform services"
            opps.append(_opp("AZ_" + rt.upper().replace(" ", "_"), f"Review {nm}", cat, [nm], 3, "Assumption", "Low", 2, "Cloud lead", 0, c * cfg["platform_pct"],
                             f"{nm} cost {inr(c)}/month ({x['resource'].nunique()} resources; {inr(first)} in {month_label(M['months'][0])}). {how}",
                             f"Review the top resources below; the figure assumes {cfg['platform_pct']:.0%}.",
                             items=top.rename("monthly_cost").reset_index().rename(columns={"resource": "resource"}), unit="resources"))
    D = migration_assessment(M, cfg)
    M["migration"] = D
    if len(D):
        mv = D[D["verdict"] == "Can move"]
        bt = ""
        OI = M.get("_overlap_items")
        if OI is not None and len(OI):
            H = M["hosts"]
            ah_ = H[H["provider"] == "Azure"].set_index("bill_id")
            S_ = M["ws"][(M["ws"]["month"] == latest) & (M["ws"]["days"] >= 28)]
            ft = fit_price(S_)
            for _, r_ in OI.sort_values("azure_monthly", ascending=False).iterrows():
                n_ = r_["azure_vm"]
                if ft and n_ in ah_.index and ah_.loc[n_, "ram_gb"] == ah_.loc[n_, "ram_gb"] and r_["match"] == "Similar name" and r_["azure_monthly"] > 1e5:
                    ram_ = float(ah_.loc[n_, "ram_gb"])
                    e_ = ft["vcpu"] * min(max(2, round(ram_ / 4)), 32) + ft["ram"] * round(ram_) + ft["hdd"] * float(ah_.loc[n_, "alloc_gb"])
                    bt = f" Check on a real pair: the same method prices {n_} at {inr(e_)}/month and WebberStop actually bills {inr(r_['webber_monthly'])}/month for its replacement (Azure: {inr(r_['azure_monthly'])})."
                    break
        if len(mv):
            tot = mv["saving"].sum()
            opps.append(_opp("MIG_IAAS", "Move Azure workloads to WebberStop", "Move to WebberStop", ["Azure virtual machines", "Azure managed disks", "Azure snapshots", "Azure NAT gateways", "Azure application gateways", "Azure public IPs"], 3, "Assumption", "Medium", 3, "Cloud lead + app owners", 0, tot,
                             f"{len(mv)} Azure workloads cost {inr(mv['azure_in_scope'].sum())}/month for servers, disks, snapshots and their network items. The same sizes at WebberStop's current prices come to about {inr(mv['webber_est'].sum())}/month. "
                             "Not counted: Azure data-transfer charges and the Windows / SQL Server licence difference, because the Azure export has no meter detail. WebberStop has no ingress or egress charges, so the real saving is likely higher." + bt,
                             "Get a WebberStop quote, clear the checks below, pilot with one non-production workload, then move in order of saving.",
                             items=mv[["workload", "vms", "azure_in_scope", "webber_est", "saving", "azure_replat", "azure_other"]].rename(columns={"azure_in_scope": "azure_monthly", "azure_replat": "azure_paas_left", "azure_other": "azure_other_left"}), unit="workloads"))
        cc = D[D["verdict"] == "Compliance check first"]
        if len(cc):
            opps.append(_opp("MIG_REVIEW", "Workloads to move only after a compliance check", "Move to WebberStop", ["Azure virtual machines"], 3, "Verify first", "High", 3, "Legal + Security", 0, cc["saving"].clip(lower=0).sum(),
                             f"{len(cc)} workloads look cheaper at WebberStop but their names suggest regulated or location-bound data (health, UAE, payments). The saving is real only if legal and security approve the move.",
                             "Get sign-off on data location, certifications and customer contracts before any pilot.", items=cc[["workload", "vms", "azure_in_scope", "webber_est", "saving"]].rename(columns={"azure_in_scope": "azure_monthly"}), unit="workloads"))
        rp = D[D["azure_replat"] > 0]
        if len(rp):
            opps.append(_opp("MIG_REPLAT", "Azure database and app services that need re-platforming to move", "Move to WebberStop", ["Azure SQL", "Azure App Service plans"], 3, "Enabler", "Medium", 3, "App owners + DBA", 0, 0,
                             f"{inr(rp['azure_replat'].sum())}/month is on Azure SQL, App Service and similar services that cannot be copied to WebberStop as they are. They must be rebuilt as servers (with SQL Server licences) first. Worth a quote if SQL or Windows licences dominate the Azure cost.",
                             "Ask each app team for the effort, and WebberStop for the server plus licence price.",
                             items=rp[["workload", "azure_replat"]].rename(columns={"azure_replat": "azure_monthly"}), unit="workloads"))
        nz = D[D["verdict"] == "Needs server sizes"]
        if len(nz):
            opps.append(_opp("MIG_SIZES", "Get the sizes of Azure servers that are not in Zabbix", "Monitoring", ["Azure virtual machines"], 3, "Enabler", "Low", 1, "DevOps", 0, 0,
                             f"{len(nz)} Azure workloads ({inr(nz['azure_in_scope'].sum())}/month) have servers that are not in the Zabbix audit, so their size is unknown and they cannot be priced at WebberStop.",
                             "Install the Zabbix agent, or list the VM sizes from the Azure portal.", items=nz[["workload", "vms", "sized_vms", "azure_in_scope"]].rename(columns={"azure_in_scope": "azure_monthly"}), unit="workloads"))
    for o in opps:
        o["wave"] = STEP_OF.get(o["id"], 4 if o["id"].startswith("AZ_") else o["wave"])
        o["validate"] = VALIDATE.get(o["id"], VALIDATE.get(o["id"].replace("_V", ""), ["Ask the owner to confirm before acting."]))
    for o in opps:
        o["score"] = (o["saving"] if o["saving"] else o["at_stake"] * 0.5) / o["effort"]
    return sorted(opps, key=lambda o: (o["wave"], -o["score"]))


def savings_summary(opps):
    s = {b: 0.0 for b in BASIS_ORDER + ["Enabler"]}
    for o in opps:
        v = o["saving"] if o["basis"] in COUNTED else o["at_stake"]
        s[o["basis"]] = s.get(o["basis"], 0.0) + v
    s["counted"] = s["Bill-evidenced"] + s["Utilisation-evidenced"]
    by = {}
    for o in opps:
        if o["basis"] == "Enabler":
            continue
        k = o["category"]
        d = by.setdefault(k, {"Proven": 0.0, "Verify first": 0.0, "Assumption": 0.0})
        if o["basis"] in COUNTED:
            d["Proven"] += o["saving"]
        elif o["basis"] == "Verify first":
            d["Verify first"] += o["at_stake"]
        else:
            d["Assumption"] += o["at_stake"]
    s["by_category"] = by
    return s


def service_matrix(M):
    """Every service on the two bills with its cost, trend and the levers found."""
    az, ws, wsh, azc = M["az"], M["ws"], M["wsh"], 1 - M["cfg"]["azure_discount"]
    ms, latest = M["months"], M["latest"]
    first = ms[0]
    rows = []
    if len(az):
        sv = az.groupby(["rtype", "month"])["cost"].sum().unstack().fillna(0) * azc
        for rt, r in sv.iterrows():
            nm = _AZ_SERVICES.get(rt, (None,))[0] or {"virtual machine": "Azure virtual machines", "snapshot": "Azure snapshots"}.get(rt, ("Azure: unclassified" if str(rt).strip("- ")=="" else f"Azure: {rt}"))
            rows.append({"provider": "Azure", "service": nm, "first": r.get(first, 0), "latest": r.get(latest, 0), "avg": r.mean()})
    if len(ws):
        wc = ws.groupby("month")["net"].sum()
        rows.append({"provider": "Webber", "service": "Webber compute", "first": wc.get(first, 0), "latest": wc.get(latest, 0), "avg": wc.mean()})
    if len(wsh):
        sk = wsh.groupby(["kind", "month"])["net"].sum().unstack().fillna(0)
        for k, r in sk.iterrows():
            nm = {"Backup storage": "Webber backup storage", "Backup - standard": "Webber backup (per VM)", "Backup - MOC": "Webber backup (per VM)", "Kubernetes (CloudStack)": "Webber Kubernetes",
                  "Public IP / add-on": "Webber public IPs"}.get(k, f"Webber {k.lower()}")
            rows.append({"provider": "Webber", "service": nm, "first": r.get(first, 0), "latest": r.get(latest, 0), "avg": r.mean()})
    T = pd.DataFrame(rows).groupby(["provider", "service"], as_index=False).sum()
    T["change_pct"] = np.where(T["first"] > 0, (T["latest"] / T["first"] - 1) * 100, np.nan)
    out = []
    lat = T.set_index("service")["latest"].to_dict()
    for _, r in T.iterrows():
        pr = vf = asm = 0.0
        mine = []
        for o in M["opps"]:
            if r["service"] not in o["services"]:
                continue
            mine.append(o)
            tot = sum(lat.get(sv, 0) for sv in o["services"]) or 1
            w = lat.get(r["service"], 0) / tot  # an action spanning several services is split by their cost, never counted twice
            pr += o["saving"] * w if o["basis"] in COUNTED else 0
            vf += o["at_stake"] * w if o["basis"] == "Verify first" else 0
            asm += o["at_stake"] * w if o["basis"] == "Assumption" else 0
        out.append({**r.to_dict(), "levers": "; ".join(sorted({o["category"] for o in mine})), "proven": pr, "verify": vf, "assumption": asm})
    return pd.DataFrame(out).sort_values("latest", ascending=False)


def build_model(base, cfg, year=DEFAULT_YEAR, progress=None):
    """Read the three folders and produce every table the app shows."""
    M = {"base": base, "cfg": cfg, "notes": [], "dirs": {}}
    for k, n in FOLDER_NAMES.items():
        M["dirs"][k] = find_dir(base, n)
    if not M["dirs"]["azure"] or not M["dirs"]["webber"]:
        M["error"] = "Folder not found: " + ", ".join(FOLDER_NAMES[k] for k in ("azure", "webber") if not M["dirs"][k]) + f" (looked in {base})"
        return M
    M["az"], n1 = parse_azure_folder(M["dirs"]["azure"], year)
    M["ws"], M["wsh"], n2 = parse_webber_folder(M["dirs"]["webber"])
    M["notes"] += n1 + n2
    if not len(M["az"]) and not len(M["ws"]):
        M["error"] = "No cost data could be read from the Azure and Webber folders."
        return M
    M["months"] = sorted(set(M["az"]["month"]) | set(M["ws"]["month"]))
    M["latest"] = M["months"][-1]
    M["monthly"] = monthly_table(M["az"], M["ws"], M["wsh"], cfg["azure_discount"])
    M["lines"], M["movers"] = spend_lines(M["az"], M["ws"], M["wsh"], cfg["azure_discount"])
    M["az_daily"] = azure_daily(M["az"], cfg["azure_discount"])
    M["srv"] = server_table(M["ws"], M["latest"])
    zd = M["dirs"]["zabbix"]
    M["aud"] = parse_audit_folder(zd, progress) if zd else {"cpu": pd.DataFrame(), "notes": ["Zabbix Logs folder not found."]}
    M["notes"] += M["aud"].get("notes", [])
    M["unmonitored"] = pd.DataFrame()
    M["hosts"], M["rs"], M["trend"] = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    if len(M["aud"].get("cpu", [])):
        H = host_metrics(M["aud"], cfg)
        H = match_hosts(H, M["srv"], M["az"], cfg, load_host_aliases(zd))
        M["hosts"] = H
        M["rs"] = rightsize(H, M["ws"], M["latest"], cfg)
        M["trend"] = util_trend(M["aud"], H)
    M["opps"] = find_opportunities(M, cfg)
    M["summary"] = savings_summary(M["opps"])
    M["services"] = service_matrix(M)
    return M
# ==============================================================================
# 5b. MONTH-ON-MONTH COMPARISON  (what changed between two billed months, and why)
# ==============================================================================
def _dims(r):
    return f"{r['vcpu']:.0f} vCPU · {r['ram']:.0f} GB RAM · {r['hdd']:.0f} GB disk"


def _dim_change(a, b):
    """'vCPU 8→16 · RAM 16→32 GB' for the dimensions that differ, plus +1/-1/0 signs."""
    out, ups, downs = [], 0, 0
    for k, lab, unit in (("vcpu", "vCPU", ""), ("ram", "RAM", " GB"), ("hdd", "Disk", " GB")):
        x, y = a[k], b[k]
        if x == x and y == y and abs(x - y) > 0.5:
            out.append(f"{lab} {x:.0f}→{y:.0f}{unit}")
            ups += y > x
            downs += y < x
    return " · ".join(out), ups, downs


def _zbx_cfg(aud, host, month):
    """(RAM GB, disk GB) Zabbix last saw for a host in a month."""
    end = pd.Period(month).end_time
    ram = aud["ramtot"][(aud["ramtot"]["host"] == host) & (aud["ramtot"]["date"] <= end) & (aud["ramtot"]["date"] >= end - pd.Timedelta(days=45))] if len(aud.get("ramtot", [])) else None
    fs = aud["fs"][(aud["fs"]["host"] == host) & (aud["fs"]["date"] <= end) & (aud["fs"]["date"] >= end - pd.Timedelta(days=45))] if len(aud.get("fs", [])) else None
    r = float(ram.sort_values("date")["ram_gb"].iloc[-1]) if ram is not None and len(ram) else np.nan
    d = np.nan
    if fs is not None and len(fs):
        last = fs["date"].max()
        d = float(fs[fs["date"] == last].drop_duplicates("mount")["total_gb"].sum())
    return r, d


def month_compare(M, m0, m1):
    """Everything that changed between billed months m0 and m1, per server, with a driver (reason) for every rupee."""
    ws, wsh, az, cfg = M["ws"], M["wsh"], M["az"], M["cfg"]
    azc = 1 - cfg["azure_discount"]
    drivers = []      # (provider, driver, effect, good_when_negative)
    web_rows, az_rows, sh_rows = [], [], []

    # ---------------- Webber servers
    def snap(m):
        d = ws[ws["month"] == m].sort_values("end")
        if not len(d):
            return pd.DataFrame()
        last = d.groupby("ip").tail(1).set_index("ip")[["host", "category", "vcpu", "ram", "hdd", "rate", "disc"]]
        last["cost"] = d.groupby("ip")["net"].sum()
        last["days"] = d.groupby("ip")["days"].sum().clip(upper=31)
        return last
    A, B = snap(m0), snap(m1)
    for ip in sorted(set(A.index) | set(B.index)):
        a = A.loc[ip] if ip in A.index else None
        b = B.loc[ip] if ip in B.index else None
        if a is None:
            part = f" Billed {b['days']:.0f} days in {month_label(m1)}; full month is {inr(b['rate'])}." if b["days"] < 28 else ""
            web_rows.append({"server": b["host"], "line": b["category"], "driver": "New server", "before": 0.0, "after": b["cost"], "change": b["cost"],
                             "what": f"Added: {_dims(b)}.{part}"})
            drivers.append(("Webber", "New servers", b["cost"], True))
        elif b is None:
            part = f" It was billed {a['days']:.0f} days in {month_label(m0)}." if a["days"] < 28 else ""
            web_rows.append({"server": a["host"], "line": a["category"], "driver": "Removed server", "before": a["cost"], "after": 0.0, "change": -a["cost"],
                             "what": f"No longer billed ({_dims(a)}).{part}"})
            drivers.append(("Webber", "Servers removed", -a["cost"], True))
        else:
            delta = b["cost"] - a["cost"]
            if abs(delta) < 1:
                continue
            conf = b["rate"] - a["rate"]               # change in full-month price at month end
            tim = delta - conf                          # part-month billing left over
            txt, ups, downs = _dim_change(a, b)
            ren = f"Server renamed or replaced ({a['host']} → {b['host']}). " if str(a["host"]).strip().lower() != str(b["host"]).strip().lower() else ""
            if abs(conf) < 1:
                drv, name = "Part-month billing", "Timing (part-month billing)"
                what = f"Same size and price; billed {a['days']:.0f} days in {month_label(m0)} and {b['days']:.0f} in {month_label(m1)}."
                tim, conf = delta, 0.0
            elif txt and ups and not downs and conf > 0:
                drv, name = "Capacity upgrade", "Capacity upgrades (RAM / CPU / storage)"
                what = f"Upgraded: {txt}. Full-month price {inr(a['rate'])} → {inr(b['rate'])}."
            elif txt and downs and not ups and conf < 0:
                drv, name = "Downsized", "Downsizing / optimisation"
                what = f"Reduced: {txt}. Full-month price {inr(a['rate'])} → {inr(b['rate'])}."
            elif txt:
                drv, name = "Reconfigured", "Capacity upgrades (RAM / CPU / storage)" if conf > 0 else "Downsizing / optimisation"
                what = f"Changed: {txt}. Full-month price {inr(a['rate'])} → {inr(b['rate'])}."
            elif str(a["disc"]) != str(b["disc"]):
                drv, name = "Discount change", "Price / discount changes"
                what = f"Same size; 40% discount {'applied' if str(b['disc']).lower().startswith('y') else 'removed'}. Full-month price {inr(a['rate'])} → {inr(b['rate'])}."
            else:
                drv, name = "Price change", "Price / discount changes"
                what = f"Same size; full-month price {inr(a['rate'])} → {inr(b['rate'])}."
            if abs(tim) >= 1 and conf != 0:
                what += f" Part-month billing adds {inr(tim)} this month."
            web_rows.append({"server": b["host"], "line": b["category"], "driver": drv, "before": a["cost"], "after": b["cost"], "change": delta, "what": ren + what})
            if conf:
                drivers.append(("Webber", name, conf, True))
            if tim and abs(tim) >= 1:
                drivers.append(("Webber", "Timing (part-month billing)", tim, True))

    # ---------------- Webber shared lines (backup, Kubernetes, firewall, IPs ...)
    if len(wsh):
        def sk(m):
            d = wsh[wsh["month"] == m]
            g = d.groupby("kind").agg(cost=("net", "sum"), count=("count", "max"), n=("item", "size"))
            return g
        SA, SB = sk(m0), sk(m1)
        for k in sorted(set(SA.index) | set(SB.index)):
            a = SA.loc[k] if k in SA.index else None
            b = SB.loc[k] if k in SB.index else None
            c0 = a["cost"] if a is not None else 0.0
            c1 = b["cost"] if b is not None else 0.0
            if abs(c1 - c0) < 1:
                continue
            if a is None:
                what, drv = f"New charge in {month_label(m1)}.", "New service line"
            elif b is None:
                what, drv = f"No longer billed in {month_label(m1)}.", "Service line stopped"
            else:
                q0, q1 = a["count"], b["count"]
                if q0 == q0 and q1 == q1 and q0 != q1:
                    what, drv = f"Quantity {q0:,.0f} → {q1:,.0f}.", "Quantity change"
                elif a["n"] != b["n"]:
                    what, drv = f"{a['n']} → {b['n']} billed items.", "Quantity change"
                else:
                    what, drv = "Same quantity, different price or billing days.", "Price change"
            sh_rows.append({"server": k, "line": "Shared", "driver": drv, "before": c0, "after": c1, "change": c1 - c0, "what": what})
            drivers.append(("Webber", "Shared services (backup, Kubernetes, firewall, IPs)", c1 - c0, True))

    # ---------------- Azure resources
    zh = {}
    H = M["hosts"]
    if len(H):
        for _, h in H[H["provider"] == "Azure"].iterrows():
            zh[h["bill_id"]] = h["host"]
    c = az[az["month"].isin([m0, m1])].copy()
    # snapshots carry a timestamp in the name and disks a random id: group them per family so a rolling backup is not "new" and "removed" every day
    c["resource"] = c["resource"].astype(str).str.replace(r"_at_\d{4}-\d{2}-\d{2}.*$", "", regex=True).str.replace(r"_[0-9a-f]{32}$", "", regex=True)
    g = c.groupby(["resource", "rtype", "product", "month"]).agg(cost=("cost", "sum"), days=("date", "nunique"), first=("date", "min"), last=("date", "max")).reset_index()
    g["cost"] *= azc
    P = g.pivot_table(index=["resource", "rtype", "product"], columns="month", values=["cost", "days"], aggfunc="sum").fillna(0)
    ed = {m: pd.Period(m).end_time.normalize() for m in (m0, m1)}
    st_ = {m: pd.Period(m).start_time for m in (m0, m1)}
    fl = g.set_index(["resource", "rtype", "product", "month"])[["first", "last"]]
    for (res, rt, prod), r in P.iterrows():
        c0, c1, d0, d1 = r[("cost", m0)], r[("cost", m1)], r[("days", m0)], r[("days", m1)]
        if abs(c1 - c0) < 1:
            continue
        lab = f"{rt}, {prod}"
        if c0 == 0:
            az_rows.append({"server": res, "line": lab, "driver": "New resource", "before": 0.0, "after": c1, "change": c1, "what": f"New {rt} in {month_label(m1)} ({d1:.0f} days billed)."})
            drivers.append(("Azure", "New Azure resources", c1, True))
        elif c1 == 0:
            az_rows.append({"server": res, "line": lab, "driver": "Removed resource", "before": c0, "after": 0.0, "change": -c0, "what": f"No longer billed ({rt})."})
            drivers.append(("Azure", "Azure resources removed", -c0, True))
        else:
            a0, a1 = c0 / max(d0, 1), c1 / max(d1, 1)
            rate_eff = (a1 - a0) * d0
            day_eff = a1 * (d1 - d0)
            what = []
            nm_day = None
            if abs(a1 / a0 - 1) >= 0.05:
                z = ""
                zh_ = zh.get(res)
                if zh_:
                    r0, k0 = _zbx_cfg(M["aud"], zh_, m0)
                    r1, k1 = _zbx_cfg(M["aud"], zh_, m1)
                    parts = []
                    if r0 == r0 and r1 == r1 and abs(r1 - r0) / max(r0, 1) > 0.05:
                        parts.append(f"RAM {r0:.0f}→{r1:.0f} GB")
                    if k0 == k0 and k1 == k1 and abs(k1 - k0) / max(k0, 1) > 0.05:
                        parts.append(f"disk {k0:.0f}→{k1:.0f} GB")
                    if parts:
                        z = " Zabbix shows " + ", ".join(parts) + "."
                verb = "Higher" if a1 > a0 else "Lower"
                what.append(f"{verb} cost per day ({inr(a0)} → {inr(a1)}): usage, size or price change.{z}")
                drivers.append(("Azure", "Azure usage / size increases" if rate_eff > 0 else "Azure usage / size reductions (optimisation)", rate_eff, True))
            else:
                day_eff = c1 - c0
            if abs(day_eff) >= 1:
                f1, l1 = fl.loc[(res, rt, prod, m1)] if (res, rt, prod, m1) in fl.index else (None, None)
                if l1 is not None and day_eff < 0 and l1 < ed[m1] - pd.Timedelta(days=1):
                    what.append(f"Stopped billing after {l1:%d %b} ({d1:.0f} days in {month_label(m1)} vs {d0:.0f}): switched off or deleted part-way.")
                    nm = "Azure resources stopped part-way through the month"
                elif f1 is not None and day_eff > 0 and f1 > st_[m1] + pd.Timedelta(days=1):
                    what.append(f"Started on {f1:%d %b} ({d1:.0f} days in {month_label(m1)} vs {d0:.0f}).")
                    nm = "Azure resources started part-way through the month"
                else:
                    what.append(f"Billed {d0:.0f} days in {month_label(m0)} vs {d1:.0f} in {month_label(m1)}.")
                    nm = "Azure timing (days in the month)"
                drivers.append(("Azure", nm, day_eff, True))
                nm_day = nm
            drv = "Usage / size change" if abs(a1 / a0 - 1) >= 0.05 else ("Days in month" if abs(d1 - d0) <= 1 else "Part-month billing")
            if nm_day and abs(day_eff) > abs(rate_eff):
                drv = "Stopped part-way" if "stopped" in nm_day else ("Started part-way" if "started" in nm_day else "Days in month")
            az_rows.append({"server": res, "line": lab, "driver": drv, "before": c0, "after": c1, "change": c1 - c0, "what": " ".join(what)})

    D = pd.DataFrame(drivers, columns=["provider", "driver", "effect", "x"]).groupby(["provider", "driver"], as_index=False).agg(effect=("effect", "sum"), n=("effect", "size"))
    D["effect"] = D["effect"].round(2)
    W = pd.DataFrame(web_rows + sh_rows)
    Z = pd.DataFrame(az_rows)
    for d in (W, Z):
        if len(d):
            d.sort_values("change", key=lambda s: -s.abs(), inplace=True)
    T = M["monthly"].set_index("month")
    return {"m0": m0, "m1": m1, "drivers": D, "webber": W.reset_index(drop=True), "azure": Z.reset_index(drop=True),
            "azure0": T.loc[m0, "azure_net"], "azure1": T.loc[m1, "azure_net"], "webber0": T.loc[m0, "webber"], "webber1": T.loc[m1, "webber"]}
# ==============================================================================
# 6. LOOK & FEEL  (white, slate + dark maroon)
# ==============================================================================
LIGHT_THEME = {"base": "light", "primaryColor": "#7a1c1c", "backgroundColor": "#ffffff",
               "secondaryBackgroundColor": "#f5f7fa", "textColor": "#1d2733", "font": "sans serif"}


def ensure_theme_file():
    """Light theme with a dark-maroon accent, applied to the running server and saved to .streamlit/config.toml.
    Streamlit reads the file at start-up, so the very first launch may need one browser refresh."""
    path = os.path.join(".streamlit", "config.toml")
    try:
        old = open(path, encoding="utf-8").read() if os.path.exists(path) else ""
        block = "[theme]\n" + "".join(f'{k} = "{v}"\n' for k, v in LIGHT_THEME.items())
        if "[theme]" not in old:
            os.makedirs(".streamlit", exist_ok=True)
            with open(path, "a", encoding="utf-8") as f:
                f.write("\n" + block)
        elif "#1b232e" in old or 'base = "dark"' in old:
            with open(path, "w", encoding="utf-8") as f:
                f.write(re.sub(r"\[theme\].*?(?=\n\[|\Z)", block, old, flags=re.S))
    except OSError:
        pass
    try:
        for k, v in LIGHT_THEME.items():
            st._config.set_option(f"theme.{k}", v)
    except Exception:
        pass


PAL = {"bg": "#ffffff", "panel": "#f5f7fa", "panel2": "#e8ecf1", "line": "#cfd6df", "text": "#1d2733",
       "muted": "#556273", "slate": "#53687e", "slate_d": "#33475b", "maroon": "#7a1c1c", "maroon_d": "#5a1313",
       "good": "#3d7a58", "grey": "#8892a0", "soft": "#a9b6c4"}
BASIS_COLOR = {"Bill-evidenced": PAL["good"], "Utilisation-evidenced": PAL["good"], "Verify first": PAL["slate"],
               "Assumption": PAL["grey"], "Enabler": PAL["slate_d"]}
CSS = f"""
<style>
  html, body, [class*="css"] {{ font-size: 16px; }}
  .block-container {{ padding-top: 3.2rem; padding-bottom: 3rem; max-width: 1480px; }}
  h1, h2, h3, h4 {{ letter-spacing: -0.01em; }}
  .ct-title {{ font-size: 1.9rem; font-weight: 760; margin: 0; }}
  .ct-sub {{ color: {PAL['muted']}; font-size: .97rem; margin: .1rem 0 .8rem 0; }}
  .ct-banner {{ background: {PAL['panel']}; border: 1px solid {PAL['line']}; border-left: 6px solid {PAL['maroon']};
      border-radius: 10px; padding: 16px 22px; margin: .4rem 0 1rem 0; font-size: 1.07rem; line-height: 1.6; color: {PAL['text']}; }}
  .ct-banner b {{ color: {PAL['maroon_d']}; }}
  .ct-kpi {{ background: {PAL['panel']}; border: 1px solid {PAL['line']}; border-top: 4px solid {PAL['slate_d']};
      border-radius: 10px; padding: 14px 16px; min-height: 150px; box-shadow: 0 1px 2px rgba(29,39,51,.06); color: {PAL['text']}; }}
  .ct-kpi.accent {{ border-top-color: {PAL['maroon']}; }}
  .ct-kpi.good {{ border-top-color: {PAL['good']}; }}
  .ct-l {{ font-size: .74rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: {PAL['muted']}; }}
  .ct-v {{ font-size: 1.85rem; font-weight: 780; line-height: 1.2; margin: .25rem 0 .15rem 0; }}
  .ct-s {{ font-size: .92rem; }}
  .ct-h {{ font-size: .82rem; color: {PAL['muted']}; margin-top: .4rem; line-height: 1.4; }}
  .ct-card {{ background: {PAL['panel']}; border: 1px solid {PAL['line']}; border-radius: 10px; padding: 14px 18px; margin-bottom: 12px; color: {PAL['text']}; }}
  .ct-card-t {{ font-weight: 700; font-size: 1.04rem; }}
  .ct-card-m {{ color: {PAL['muted']}; font-size: .88rem; margin: .15rem 0 .4rem 0; }}
  .ct-card-d {{ font-size: .96rem; line-height: 1.55; }}
  .ct-pill {{ display: inline-block; border-radius: 6px; padding: 2px 9px; font-size: .72rem; font-weight: 700; letter-spacing: .05em;
      color: #fff; margin-right: 8px; }}
  .ct-wave {{ font-size: 1.3rem; font-weight: 740; margin: 1.2rem 0 .05rem 0; }}
  .ct-wave-s {{ color: {PAL['muted']}; margin-bottom: .6rem; }}
  .ct-note {{ color: {PAL['muted']}; font-size: .9rem; line-height: 1.5; }}
  .ct-find {{ border-left: 4px solid {PAL['slate']}; padding: 4px 0 4px 14px; margin: 0 0 10px 0; line-height: 1.55; }}
  .stApp {{ background: {PAL['bg']}; color: {PAL['text']}; }}
  section[data-testid="stSidebar"] {{ background: {PAL['panel']}; border-right: 1px solid {PAL['line']}; }}
  div[data-testid="stExpander"] {{ border: 1px solid {PAL['line']}; border-radius: 10px; background: {PAL['bg']}; }}
  button[data-baseweb="tab"] {{ font-size: 1rem; font-weight: 600; padding: 10px 16px; }}
  div[data-testid="stExpander"] summary p {{ font-size: 1.02rem; font-weight: 600; }}
</style>
"""
_ST_VER = tuple(int(x) for x in re.findall(r"\d+", st.__version__)[:2])
STRETCH = {"width": "stretch"} if _ST_VER >= (1, 50) else {"use_container_width": True}


def esc(x):
    return html.escape(str(x))


def kpi(label, value, sub, hint, cls=""):
    return (f'<div class="ct-kpi {cls}"><div class="ct-l">{esc(label)}</div><div class="ct-v">{value}</div>'
            f'<div class="ct-s">{sub}</div><div class="ct-h">{hint}</div></div>')


def pill(text, color):
    return f'<span class="ct-pill" style="background:{color}">{esc(text)}</span>'


def style_fig(fig, height=340, legend=True, title=None):
    fig.update_layout(height=height, margin=dict(l=8, r=8, t=44 if title else 18, b=8), paper_bgcolor="rgba(0,0,0,0)",
                      plot_bgcolor="rgba(0,0,0,0)", font=dict(color=PAL["text"], size=13),
                      title=dict(text=title, font=dict(size=15), x=0.0) if title else None,
                      legend=dict(orientation="h", y=-0.2, x=0, font=dict(size=12)) if legend else dict(visible=False),
                      colorway=[PAL["slate_d"], PAL["maroon"], PAL["soft"], PAL["good"], PAL["grey"], PAL["slate"]])
    fig.update_xaxes(gridcolor=PAL["line"], zeroline=False, linecolor=PAL["line"])
    fig.update_yaxes(gridcolor=PAL["line"], zeroline=False, linecolor=PAL["line"])
    return fig


def money(label, fmt="₹%d"):
    return st.column_config.NumberColumn(label, format=fmt)


def rupee_axis(fig, vmax, vmin=0.0, axis="y", n=5):
    """Axis ticks in rupees / lakh / crore instead of plotly's 'M' and 'k'."""
    step = (vmax - vmin) / max(n - 1, 1) or 1
    mag = 10 ** int(np.floor(np.log10(step)))
    step = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= step)
    vals = list(np.arange(np.floor(vmin / step) * step, vmax + step, step))
    upd = dict(tickvals=vals, ticktext=[inr(v) if v else "0" for v in vals], range=[vals[0], vals[-1]])
    (fig.update_yaxes if axis == "y" else fig.update_xaxes)(**upd)


def table(df, money_cols=(), pct_cols=(), num=None, height=None, names=None):
    """A readable dataframe: rupee columns (in lakh when large), % columns, friendly headers."""
    if df is None or not len(df):
        st.caption("Nothing to show.")
        return
    df = df.copy()
    cfg = {}
    names = names or {}
    for c in df.columns:
        lab = names.get(c, c.replace("_", " ").capitalize())
        if c in money_cols:
            col = pd.to_numeric(df[c], errors="coerce")
            if col.abs().max() >= 1e5:
                df[c] = col / 1e5
                cfg[c] = money(lab, "₹%.2f L")
            elif col.abs().max() >= 1e3:
                df[c] = col / 1e3
                cfg[c] = money(lab, "₹%.1f K")
            else:
                cfg[c] = money(lab, "₹%.0f")
        elif c in pct_cols:
            cfg[c] = st.column_config.NumberColumn(lab, format="%.0f%%")
        elif pd.api.types.is_float_dtype(df[c]):
            cfg[c] = st.column_config.NumberColumn(lab, format="%.1f")
        elif pd.api.types.is_datetime64_any_dtype(df[c]):
            cfg[c] = st.column_config.DatetimeColumn(lab, format="DD MMM YYYY HH:mm")
        else:
            cfg[c] = st.column_config.Column(lab)
    kw = {"height": height} if height else {}
    st.dataframe(df, hide_index=True, column_config=cfg, **kw, **STRETCH)


# ------------------------------------------------------------------------------ model cache
def folder_signature(base):
    sig = []
    for k, n in FOLDER_NAMES.items():
        d = find_dir(base, n)
        if d:
            for r, _, fs in os.walk(d):
                for f in sorted(fs):
                    p = os.path.join(r, f)
                    try:
                        s = os.stat(p)
                        sig.append((p, s.st_size, int(s.st_mtime)))
                    except OSError:
                        pass
    return tuple(sig)


@st.cache_resource(show_spinner=False, max_entries=3)
def get_model(base, sig, cfg_items, year):
    return build_model(base, dict(cfg_items), year)


def default_base():
    here = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
    for c in [here, os.path.dirname(here), os.getcwd(), os.path.dirname(os.getcwd())]:
        if find_dir(c, FOLDER_NAMES["azure"]) and find_dir(c, FOLDER_NAMES["webber"]):
            return c
    return here


def run_rates(M):
    """What the latest month's bill understates: servers resized or added late in the month are billed for only part of it."""
    srv, wsh, latest = M["srv"], M["wsh"], M["latest"]
    web_billed = M["monthly"].set_index("month").loc[latest, "webber"]
    web_exit = srv.loc[srv["active"], "exit_rate"].sum() + (wsh[wsh["month"] == latest]["net"].sum() if len(wsh) else 0)
    d = M["az_daily"]
    az_billed = M["monthly"].set_index("month").loc[latest, "azure_net"]
    az_exit = d.tail(14)["net"].mean() * 30 if len(d) else 0
    return {"web_billed": web_billed, "web_exit": web_exit, "az_billed": az_billed, "az_exit": az_exit}



CSS2 = f"""
<style>
  .ct-i {{ display:inline-block; cursor: help; color: {PAL['slate']}; font-size: .95rem; font-weight: 600; margin-left: 4px; }}
  .ct-kpi {{ min-height: 238px; }}
  .ct-up {{ color: {PAL['maroon']}; font-weight: 800; white-space: nowrap; }}
  .ct-down {{ color: {PAL['good']}; font-weight: 800; white-space: nowrap; }}
  .ct-flat {{ color: {PAL['grey']}; }}
  table.ct-plan td.what {{ color: {PAL['muted']}; font-size: .88rem; }}
  .ct-delta {{ font-size: 1.05rem; font-weight: 700; margin-top: .25rem; }}
  .ct-m {{ margin-top: .55rem; padding-top: .45rem; border-top: 1px dashed {PAL['line']}; font-size: .86rem; line-height: 1.4; color: {PAL['text']}; }}
  table.ct-plan {{ width: 100%; border-collapse: collapse; font-size: .93rem; margin: .3rem 0 .6rem 0; }}
  table.ct-plan th {{ text-align: left; color: {PAL['muted']}; font-weight: 600; font-size: .8rem; padding: 8px 10px; border-bottom: 2px solid {PAL['line']}; }}
  table.ct-plan td {{ padding: 9px 10px; border-bottom: 1px solid {PAL['panel2']}; vertical-align: top; line-height: 1.4; }}
  table.ct-plan .n {{ text-align: right; white-space: nowrap; }}
  table.ct-plan .chk {{ color: {PAL['muted']}; font-size: .86rem; }}
  table.ct-plan tr.ct-step td {{ background: {PAL['slate_d']}; color: #fff; padding: 8px 10px; font-size: .95rem; }}
  table.ct-plan tr.ct-step td span {{ color: #cfd8e3; font-size: .85rem; }}
  table.ct-plan .ct-pill {{ margin-right: 0; }}
</style>
"""


def fleet_use(M):
    """Paid vs used vCPU on the Webber servers that are monitored (window average)."""
    H = M["hosts"]
    if not len(H):
        return None
    d = H[(H["provider"] == "Webber") & (H["coverage"] >= 0.6) & H["vcpu"].notna()]
    if not len(d):
        return None
    paid = d["vcpu"].sum()
    return {"n": len(d), "paid": paid, "used_avg": (d["cpu_avg"] / 100 * d["vcpu"]).sum(), "used_p95": (d["cpu_p95"] / 100 * d["vcpu"]).sum(), "cost": d["cost"].sum()}


# ------------------------------------------------------------------------------ plan structure
# Each section groups related actions; the same grouping drives the Overview table and the Action plan tab.
SECTIONS = [
    (1, "OLD_SIDE", "Switch off the old Azure side of servers already moved to WebberStop", ["OVERLAP"]),
    (1, "SILENT", "Confirm servers that stopped reporting, then stop paying for them", ["SILENT"]),
    (1, "COMMERCIAL", "Ask WebberStop for the missing 40% discount and one price for same-size servers", ["DISCOUNT", "RATE_CARD"]),
    (1, "SMALL", "Release unattached public IPs and backup / test-server charges that no longer apply", ["IP_ORPHAN", "BACKUP_VMS", "SCHEDULE_W"]),
    (2, "IDLE", "Switch off servers that are almost unused", ["IDLE", "GPU"]),
    (2, "RIGHTSIZE", "Right-size servers whose measured load allows a smaller size", ["RS_NONPROD", "RS_NONPROD_V", "RS_PROD", "RS_PROD_V", "RS_DB", "RS_DB_V", "SQL_MEM", "SCHEDULE"]),
    (2, "DISKS", "Shrink over-provisioned disks", ["DISK"]),
    (2, "K8", "Explain the growth in Kubernetes (CloudStack) charges", ["K8"]),
    (3, "MOVE", "Move Azure workloads to WebberStop", ["MIG_IAAS"]),
    (3, "MOVE_CHECK", "Workloads to move only after a compliance check", ["MIG_REVIEW"]),
    (3, "REPLAT", "Azure SQL and App Service that must be rebuilt as servers before they can move", ["MIG_REPLAT"]),
    (3, "SIZES", "Get the sizes of Azure servers that are not in Zabbix, so they can be priced", ["MIG_SIZES"]),
    (4, "CAPACITY", "Fix servers that are short of CPU, memory or disk", ["CAPACITY"]),
    (4, "MONITOR", "Put monitoring on billed servers nobody measures", ["MON_GAPS"]),
    (4, "TUNE", "Tune what stays on Azure and backup retention", ["AZ_*", "BACKUP_STORE"]),
]


def _match(o, pats):
    return any((o["id"].startswith(p[:-1]) if p.endswith("*") else o["id"] == p) for p in pats)


def sections(M):
    out = []
    used = set()
    for step, key, title, pats in SECTIONS:
        os_ = [o for o in M["opps"] if _match(o, pats)]
        used |= {o["id"] for o in os_}
        if not os_:
            continue
        conf = sum(o["saving"] for o in os_ if o["basis"] in COUNTED)
        tbc = sum(o["at_stake"] for o in os_ if o["basis"] in ("Verify first", "Assumption"))
        n = sum(len(o["items"]) for o in os_)
        unit = next((o["unit"] for o in os_ if len(o["items"])), "")
        risk = max((("Medium" if o["risk"] == "Unverified" else o["risk"]) for o in os_), key=lambda r: {"Low": 0, "Medium": 1, "High": 3}.get(r, 1))
        out.append({"step": step, "key": key, "title": title, "opps": os_, "confirmed": conf, "tbc": tbc, "n": n, "unit": unit, "risk": risk,
                    "kind": "estimate" if any(o["basis"] == "Assumption" for o in os_) and not conf else ("saving" if conf or tbc else "spend"),
                    "check": os_[0]["validate"][0] if os_[0].get("validate") else ""})
    return out


def _risk_color(r):
    return {"Low": PAL["good"], "Medium": PAL["slate"], "Unverified": PAL["grey"], "High": PAL["maroon"]}.get(r, PAL["slate"])


# ------------------------------------------------------------------------------ item tables
_ITEM_VIEW = {
    "OVERLAP": [("azure_vm", "Azure server"), ("webber_host", "WebberStop server"), ("azure_monthly", "Azure cost /mo"), ("webber_monthly", "WebberStop cost /mo"), ("smaller", "Saving /mo (cheaper side, conservative)"), ("azure_cpu_p95", "Azure CPU p95 %"), ("webber_cpu_p95", "WebberStop CPU p95 %")],
    "SILENT": [("host", "Server"), ("provider", "Where"), ("last_seen", "Last reading"), ("silent_days", "Days silent"), ("monthly_cost", "Saving /mo if decommissioned")],
    "DISCOUNT": [("host", "Server"), ("category", "Line"), ("vcpu", "vCPU"), ("ram", "RAM GB"), ("gross", "Billed /mo"), ("_disc", "Saving /mo if discounted")],
    "RATE_CARD": [("host", "Server"), ("category", "Line"), ("monthly_cost", "Cost /mo"), ("excess", "Saving /mo")],
    "IP_ORPHAN": [("item", "Item"), ("ip", "IP"), ("monthly_cost", "Saving /mo")],
    "SCHEDULE_W": [("host", "Server"), ("category", "Line"), ("off_p95", "CPU p95 off-hours %"), ("monthly_cost", "Cost /mo")],
    "SCHEDULE": [("host", "Server"), ("off_p95", "CPU p95 off-hours %"), ("monthly_cost", "Cost /mo"), ("saving", "Saving /mo")],
    "IDLE": [("host", "Server"), ("provider", "Where"), ("cpu_p95", "CPU p95 %"), ("cpu_pk99", "Busiest-1% peak %"), ("ram_p95", "RAM p95 %"), ("monthly_cost", "Saving /mo if switched off")],
    "RS": [("host", "Server"), ("provider", "Where"), ("vcpu", "vCPU now"), ("new_vcpu", "vCPU after"), ("cpu_p95", "CPU p95 %"), ("cpu_pk99", "Busiest-1% peak %"), ("monthly_cost", "Cost now /mo"), ("saving", "Saving /mo")],
    "SQL_MEM": [("host", "Server"), ("cpu_avg", "CPU avg %"), ("ram_avg", "RAM avg %"), ("ram_p99", "RAM p99 %")],
    "DISK": [("host", "Server"), ("alloc_gb", "Disk GB"), ("used_gb", "Used GB"), ("disk_max_pct", "Fullest %"), ("reclaim_gb", "Reclaimable GB"), ("monthly_saving", "Saving /mo (estimate)")],
    "MIG": [("workload", "Azure workload"), ("vms", "Servers"), ("azure_monthly", "Azure cost /mo"), ("webber_est", "WebberStop estimate /mo"), ("saving", "Saving /mo"), ("azure_paas_left", "PaaS staying on Azure /mo")],
    "MIG_REVIEW": [("workload", "Azure workload"), ("vms", "Servers"), ("azure_monthly", "Azure cost /mo"), ("webber_est", "WebberStop estimate /mo"), ("saving", "Saving /mo")],
    "MIG_REPLAT": [("workload", "Azure workload"), ("azure_monthly", "Azure service cost /mo")],
    "MIG_SIZES": [("workload", "Azure workload"), ("vms", "Servers"), ("sized_vms", "Sized"), ("azure_monthly", "Azure cost /mo")],
    "CAPACITY": [("host", "Server"), ("provider", "Where"), ("cpu_p95", "CPU p95 %"), ("ram_p99", "RAM p99 %"), ("disk_max_pct", "Fullest disk %"), ("worst_mount", "Volume"), ("reason", "What is short")],
    "GPU": [("host", "Server"), ("vcpu", "vCPU"), ("ram", "RAM GB"), ("monthly_cost", "Cost /mo")],
    "MON_GAPS": [("host", "Server"), ("category", "Line"), ("vcpu", "vCPU"), ("ram", "RAM GB"), ("monthly_cost", "Cost /mo with no monitoring")],
    "AZ": [("resource", "Azure resource"), ("monthly_cost", "Cost /mo")],
}
_MONEY = re.compile(r"/mo|cost|saving|billed|estimate", re.I)


def item_view(o):
    """(DataFrame, money columns) with plain headers, at most 8 columns, for one action's affected servers / resources."""
    d = o["items"]
    if d is None or not len(d):
        return None, ()
    d = d.copy()
    oid = o["id"]
    key = "RS" if oid.startswith("RS_") else ("MIG" if oid == "MIG_IAAS" else ("AZ" if oid.startswith("AZ_") else oid))
    spec = _ITEM_VIEW.get(key)
    if oid == "DISCOUNT" and "gross" in d:
        d["_disc"] = d["gross"] * 0.4
    if not spec:
        spec = [(c, c.replace("_", " ").capitalize()) for c in d.columns[:7]]
    spec = [(c, n) for c, n in spec if c in d.columns]
    d = d[[c for c, _ in spec]]
    d.columns = [n for _, n in spec]
    for c in d.columns:
        if pd.api.types.is_datetime64_any_dtype(d[c]):
            d[c] = d[c].dt.strftime("%d %b %Y")
    sort = next((c for c in d.columns if re.search(r"Saving|cost", c, re.I) and pd.api.types.is_numeric_dtype(d[c])), None)
    if sort:
        d = d.sort_values(sort, ascending=False)
    return d, tuple(c for c in d.columns if _MONEY.search(c) and pd.api.types.is_numeric_dtype(d[c]))


def tune_items(S):
    """One combined list for the many small 'tune what stays on Azure' actions."""
    parts = []
    for o in S["opps"]:
        if len(o["items"]) and "resource" in o["items"]:
            q = o["items"].sort_values("monthly_cost", ascending=False).head(5).copy()
            q.insert(0, "Service", o["title"].replace("Review ", ""))
            parts.append(q)
    return pd.concat(parts) if parts else pd.DataFrame()


# ------------------------------------------------------------------------------ headline pieces
def tile(label, value, meaning, move, tip, cls=""):
    return (f'<div class="ct-kpi {cls}"><div class="ct-l">{esc(label)} <span class="ct-i" title="{esc(tip)}">&#9432;</span></div><div class="ct-v">{value}</div>'
            f'<div class="ct-s">{meaning}</div><div class="ct-m"><b>Your move:</b> {move}</div></div>')


def totals(M):
    S = M["summary"]
    by = lambda i: next((o for o in M["opps"] if o["id"] == i), None)
    mig = by("MIG_IAAS")
    rev = by("MIG_REVIEW")
    confirm = S["Verify first"] - (rev["at_stake"] if rev else 0)
    tune = sum(o["at_stake"] for o in M["opps"] if o["basis"] == "Assumption" and not o["id"].startswith("MIG"))
    return {"conf": S["counted"], "confirm": confirm, "mig": mig["at_stake"] if mig else 0, "mig_n": len(mig["items"]) if mig else 0,
            "rev": rev["at_stake"] if rev else 0, "tune": tune, "unmon": S["Enabler"]}


def headline(M, rr):
    T, cfg = M["monthly"], M["cfg"]
    a, b = T.iloc[0], T.iloc[-1]
    out = []
    chg = b["total"] / a["total"] - 1
    out.append(f"<b>Spend is {'up' if chg > 0 else 'down'} {abs(chg):.0%}</b> from {a['label']} to {b['label']} ({inr(a['total'])} → {inr(b['total'])} a month). "
               f"WebberStop {b['webber'] / a['webber'] - 1:+.0%}, Azure {b['azure_net'] / a['azure_net'] - 1:+.0%}.")
    fu = fleet_use(M)
    if fu:
        out.append(f"<b>Most paid capacity is idle.</b> On {fu['n']} monitored WebberStop servers the average CPU in use is {fu['used_avg'] / fu['paid']:.0%} of the vCPU paid for (busy hours: {fu['used_p95'] / fu['paid']:.0%}). "
                   "Averages hide spikes, so only servers that also pass the spike test are proposed for a smaller size.")
    R = M["rs"]
    if len(R):
        cap = R[R["action"] == "Capacity risk"]
        out.append(f"<b>{len(cap)} servers need more headroom, not less</b>, so some spend must be added (Step 4).")
    if rr["web_exit"] > rr["web_billed"] * 1.02:
        out.append(f"<b>The latest bill understates the run-rate:</b> at closing size WebberStop costs about {inr(rr['web_exit'])} a month, not {inr(rr['web_billed'])}.")
    return out


def util_by_category(M):
    """Like-for-like CPU trend by product line (servers monitored in every month shown)."""
    TR = M["trend"]
    if not len(TR):
        return None
    TR = TR[TR["category"].astype(str) != ""]
    months = sorted(TR["month"].unique())
    need = months[1:] if len(months) > 2 else months
    cnt = TR[TR["month"].isin(need)].groupby("host")["month"].nunique()
    d = TR[TR["host"].isin(cnt[cnt == len(need)].index) & TR["month"].isin(need)]
    return d.groupby(["category", "month"]).agg(cpu=("cpu_avg", "mean"), servers=("host", "nunique")).reset_index()


# ------------------------------------------------------------------------------ report export
def build_report(M):
    buf = io.BytesIO()
    secs = sections(M)
    rows = []
    for s in secs:
        for o in s["opps"]:
            rows.append({"Step": WAVES[s["step"]][0] + " - " + WAVES[s["step"]][1], "Section": s["title"], "Action": o["title"], "Evidence": o["basis"], "Risk": o["risk"], "Owner": o["owner"],
                         "Confirmed saving /mo": o["saving"] if o["basis"] in COUNTED else 0, "To confirm or estimate /mo": o["at_stake"] if o["basis"] in ("Verify first", "Assumption") else 0,
                         "Why": o["why"], "What to do": o["how"], "Validate before acting": " | ".join(o.get("validate", [])), "Applies to (count)": len(o["items"])})
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(rows).to_excel(xw, sheet_name="Action plan", index=False)
        if len(M.get("migration", [])):
            M["migration"].drop(columns=["flag"], errors="ignore").to_excel(xw, sheet_name="Azure to WebberStop", index=False)
        for s in secs:
            for o in s["opps"]:
                d, _ = item_view(o)
                if d is not None and len(d):
                    d.to_excel(xw, sheet_name=o["id"][:31], index=False)
        M["monthly"].drop(columns=["label"]).to_excel(xw, sheet_name="Monthly cost", index=False)
        M["services"].to_excel(xw, sheet_name="Cost by service", index=False)
        if len(M["rs"]):
            M["rs"].drop(columns=[c for c in M["rs"].columns if M["rs"][c].map(lambda v: isinstance(v, (set, frozenset, pd.DataFrame))).any()]).to_excel(xw, sheet_name="Server decisions", index=False)
    return buf.getvalue()


# ------------------------------------------------------------------------------ tabs
def plan_html(M):
    rows = ['<table class="ct-plan"><thead><tr><th>What to do</th><th>Applies to</th><th class="n">Confirmed saving /month</th><th class="n">To confirm or estimate /month</th><th>Risk</th><th>First thing to check</th></tr></thead><tbody>']
    cur = 0
    for s in sections(M):
        if s["step"] != cur:
            cur = s["step"]
            nm, ttl, when, goal = WAVES[cur]
            rows.append(f'<tr class="ct-step"><td colspan="6"><b>{nm}: {esc(ttl)}</b> <span>({when})</span></td></tr>')
        conf = inr(s["confirmed"]) if s["confirmed"] else "&ndash;"
        tbc = inr(s["tbc"]) if s["tbc"] else ("spend to add" if s["key"] == "CAPACITY" else "&ndash;")
        if s["key"] == "MONITOR":
            tbc = f'{inr(M["summary"]["Enabler"])} unmeasured'
        n = f'{s["n"]} {s["unit"][:-1] if s["n"] == 1 else s["unit"]}' if s["n"] else "&ndash;"
        if s["tbc"] and not s["confirmed"] and all(o["basis"] == "Assumption" for o in s["opps"]):
            tbc += " <i>(estimate)</i>"
        rows.append(f'<tr><td><b>{esc(s["title"])}</b></td><td>{n}</td><td class="n">{conf}</td><td class="n">{tbc}</td>'
                    f'<td>{pill(s["risk"], _risk_color(s["risk"]))}</td><td class="chk">{esc(s["check"])}</td></tr>')
    rows.append("</tbody></table>")
    return "".join(rows)


def tab_overview(M):
    T, rr, S, cfg = M["monthly"], run_rates(M), M["summary"], M["cfg"]
    a, b = T.iloc[0], T.iloc[-1]
    t = totals(M)
    R = M["rs"]
    ncap = int((R["action"] == "Capacity risk").sum()) if len(R) else 0
    mig = M.get("migration", pd.DataFrame())
    nsz = int((mig["verdict"] == "Needs server sizes").sum()) if len(mig) else 0
    st.markdown(f'<div class="ct-banner">Cloud spend is <b>{inr(b["total"])}</b> a month ({b["label"]}), {b["total"] / a["total"] - 1:+.0%} since {a["label"]}. '
                f'Of that, <b>{inr(t["conf"])}</b> a month can be saved on evidence you can already check, <b>{inr(t["confirm"])}</b> more if the owners confirm, '
                f'and an estimated <b>{inr(t["mig"])}</b> by moving suitable Azure workloads to WebberStop.</div>', unsafe_allow_html=True)
    c = st.columns(5)
    c[0].markdown(tile(f"{b['label']} spend", inr(b["total"]), f"Azure {inr(b['azure_net'])} · WebberStop {inr(b['webber'])}",
                       "see where it grew in Evidence → Cost by service", "Azure net of the assumed 15% partner discount; WebberStop = amount paid on the bill (after its 40% discount)."), unsafe_allow_html=True)
    c[1].markdown(tile("Confirmed saving", inr(t["conf"]), "per month, backed by bills and measured load",
                       "do the Step 1 and Step 2 actions marked confirmed", "Counted only where the bills or 30 days of hourly utilisation prove it. Nothing here depends on a percentage.", "good"), unsafe_allow_html=True)
    c[2].markdown(tile("To confirm with owners", inr(t["confirm"]), "per month, real but needs a yes from an owner",
                       "send the checks in Step 1 and 2 to the owners", "Idle servers, silent servers, the Azure/WebberStop overlap, missing discounts and disks. Not counted until someone confirms.", ""), unsafe_allow_html=True)
    c[3].markdown(tile("Azure → WebberStop", inr(t["mig"]), f"per month estimate, {t['mig_n']} workloads" + (f" · {nsz} more need server sizes" if nsz else ""),
                       "request a WebberStop quote and start the compliance check", "Azure cost of the servers, disks and network items minus the same sizes at WebberStop's current prices. Excludes Azure data-transfer and licence costs (not in the export), so likely higher.", "accent"), unsafe_allow_html=True)
    c[4].markdown(tile("Short of capacity", f"{ncap} servers", "running hot on CPU, memory or disk",
                       "fix these before resizing anything else", "CPU p95 at 80%+, memory above 97%, or a disk over 90% full in the last 30 days. Spend may have to be added here.", "accent"), unsafe_allow_html=True)
    st.markdown("&nbsp;")
    l, r = st.columns([3, 2])
    with l:
        fig = go.Figure()
        for col, nm, colr in (("azure_net", "Azure", PAL["slate_d"]), ("webber_servers", "WebberStop servers", PAL["maroon"]), ("webber_shared", "WebberStop shared (backup, K8, firewall)", PAL["soft"])):
            fig.add_bar(x=T["label"], y=T[col], name=nm, marker_color=colr)
        fig.update_layout(barmode="stack")
        fig.add_scatter(x=T["label"], y=T["total"], mode="text", text=[inr(v) for v in T["total"]], textposition="top center", showlegend=False)
        fig = style_fig(fig, 320, title="Cost trend: spend per month")
        rupee_axis(fig, T["total"].max() * 1.15)
        st.plotly_chart(fig, **STRETCH)
        U = util_by_category(M)
        if U is not None and len(U):
            fig = go.Figure()
            for k, g in U.groupby("category"):
                fig.add_scatter(x=g["month"].map(month_label), y=g["cpu"], name=f"{k} ({int(g['servers'].max())})", mode="lines+markers")
            fig.update_yaxes(title="Average CPU %", rangemode="tozero")
            st.plotly_chart(style_fig(fig, 300, title="Utilisation trend: average CPU by product line (servers monitored in every month; Oct is partial)"), **STRETCH)
            fu = fleet_use(M)
            if fu:
                st.caption(f"Reading it: lines this low mean the servers have room to shrink or be shared, but only the ones that also pass the spike test are proposed (Step 2). "
                           f"Across {fu['n']} monitored WebberStop servers, {fu['used_avg'] / fu['paid']:.0%} of the vCPU paid for is used on average.")
    with r:
        st.markdown("##### What the data says")
        for f in headline(M, rr):
            st.markdown(f'<div class="ct-find">{f}</div>', unsafe_allow_html=True)
    st.markdown("##### The plan, in order")
    st.markdown(plan_html(M), unsafe_allow_html=True)
    st.caption("Confirmed = proven by bills or measured load. To confirm = real but an owner must say yes. Estimates are not added to confirmed savings, and moving a server makes its right-sizing saving irrelevant, so do not add the columns together. Open the Action plan tab for the servers behind each line and what to check.")


def tab_plan(M):
    st.markdown('<div class="ct-banner">Each section below says <b>what to do</b>, <b>which servers it applies to with the saving per server</b>, '
                'and <b>what to verify before acting</b>. Work through the steps in order: cheap confirmations first, size changes after the checks, '
                'moves to WebberStop once the quote and the compliance review are done.</div>', unsafe_allow_html=True)
    st.download_button("Download the plan with all server lists (Excel)", build_report(M), file_name="cost_optimization_plan.xlsx", key="dl_plan")
    cur = 0
    for s in sections(M):
        if s["step"] != cur:
            cur = s["step"]
            nm, ttl, when, goal = WAVES[cur]
            ss = [x for x in sections(M) if x["step"] == cur]
            st.markdown(f'<div class="ct-wave">{nm} · {esc(ttl)} <span class="ct-note">({when})</span></div><div class="ct-wave-s">{esc(goal)}</div>', unsafe_allow_html=True)
        val = []
        if s["confirmed"]:
            val.append(f"confirmed {inr(s['confirmed'])}")
        if s["tbc"]:
            val.append(f"{'estimate' if s['kind'] == 'estimate' else 'to confirm'} {inr(s['tbc'])}")
        lab = " · ".join(val) if val else ("spend to add, not a saving" if s["key"] in ("CAPACITY",) else "enables other savings")
        with st.expander(f"{s['title']}  —  {lab} per month", expanded=(s["step"] == 1)):
            for o in s["opps"]:
                if len(s["opps"]) > 1 and s["key"] != "TUNE":
                    st.markdown(f"**{o['title']}**")
                st.markdown(f'{pill(o["basis"], BASIS_COLOR.get(o["basis"], PAL["grey"]))}{pill("Risk: " + o["risk"], _risk_color(o["risk"]))}'
                            f'<span class="ct-note">Owner: {esc(o["owner"])}'
                            + (f' · saving {inr(o["saving"])}/month' if o["basis"] in COUNTED and o["saving"] else (f' · {inr(o["at_stake"])}/month at stake' if o["at_stake"] and s["key"] != "TUNE" else "")) + "</span>", unsafe_allow_html=True)
                if s["key"] != "TUNE" or o is s["opps"][0]:
                    st.markdown(f"**Why:** {o['why']}")
                    st.markdown(f"**What to do:** {o['how']}")
                    st.markdown("**Verify before acting:**\n" + "\n".join(f"- {v}" for v in o.get("validate", [])))
                if s["key"] != "TUNE":
                    d, mc = item_view(o)
                    if d is not None:
                        table(d, money_cols=mc, names={c: c for c in d.columns})
            if s["key"] == "TUNE":
                d = tune_items(s)
                if len(d):
                    st.markdown("**Largest resources per service** (the saving per service is an assumption, shown in the plan):")
                    table(d.rename(columns={"resource": "Azure resource", "monthly_cost": "Cost /mo"}), money_cols=("Cost /mo",))


def tab_evidence(M):
    cfg, H, aud, R = M["cfg"], M["hosts"], M["aud"], M["rs"]
    T = M["monthly"]
    st.caption("The numbers behind the plan. Nothing here needs to be read to act on the plan; it is here to check or challenge it.")
    with st.expander("Cost by service: what each service costs and how it moved", expanded=True):
        S = M["services"]
        S = S[(S["latest"] >= 500) | (S["avg"] >= 500)]
        d = S[["provider", "service", "first", "latest", "change_pct"]].rename(columns={"first": month_label(M["months"][0]), "latest": month_label(M["latest"]), "change_pct": "Change"})
        table(d, money_cols=(month_label(M["months"][0]), month_label(M["latest"])), pct_cols=("Change",), names={"provider": "Provider", "service": "Service"})
        st.caption("'Change' is blank for services that did not exist in the first month.")
    with st.expander("Biggest changes in spend, first month to latest"):
        mv = M["movers"]
        sel = pd.concat([mv.head(8), mv.tail(6)]).drop_duplicates("line").sort_values("change")
        fig = go.Figure(go.Bar(x=sel["change"], y=sel["line"], orientation="h", marker_color=[PAL["maroon"] if v > 0 else PAL["good"] for v in sel["change"]],
                               text=[inr(v) for v in sel["change"]], textposition="outside"))
        fig = style_fig(fig, 420, False)
        fig.update_traces(cliponaxis=False)
        rupee_axis(fig, sel["change"].max() * 1.35, min(sel["change"].min() * 1.35, -1), axis="x")
        st.plotly_chart(fig, **STRETCH)
    with st.expander("Servers: what each one uses and the decision taken"):
        if not len(R):
            st.info("No Zabbix utilisation data was found, so server decisions cannot be shown.")
        else:
            st.markdown("**How to read the decision:** *Right-size* = halve vCPUs, passes every check · *Idle* = almost unused · *Hold* = peaks or load do not fit a smaller size · "
                        "*Capacity risk* = needs more headroom · *Silent* = no readings for days · *Keep* = no change · *Skip* = not enough data.")
            opts = sorted(R["action"].unique())
            pick = st.multiselect("Show decisions", opts, default=[o for o in ("Idle", "Right-size", "Right-size database", "Capacity risk", "Silent") if o in opts], key="rs_filter")
            q = R[R["action"].isin(pick)][["host", "provider", "category", "vcpu", "cpu_avg", "cpu_p95", "cpu_pk99", "ram_p95", "disk_max_pct", "cost", "action", "reason"]]
            table(q.sort_values("cost", ascending=False), money_cols=("cost",),
                  names={"host": "Server", "provider": "Where", "category": "Line", "vcpu": "vCPU", "cpu_avg": "CPU avg %", "cpu_p95": "CPU p95 %", "cpu_pk99": "Busiest-1% peak %",
                         "ram_p95": "RAM p95 %", "disk_max_pct": "Fullest disk %", "cost": "Cost /mo", "action": "Decision", "reason": "Why"})
    with st.expander("One server in detail"):
        if len(H):
            hl = H.sort_values("cost", ascending=False)["host"].tolist()
            hsel = st.selectbox("Server", hl, key="host_sel")
            cs = aud["cpu"][aud["cpu"]["host"] == hsel].set_index("ts")
            rsr = aud["ram"][aud["ram"]["host"] == hsel].set_index("ts")
            if len(cs):
                dd, pk = cs["avg"].resample("D").mean(), cs["peak"].resample("D").max()
                fig = go.Figure()
                fig.add_scatter(x=pk.index, y=pk.values, name="CPU daily peak", line=dict(color=PAL["soft"]), fill="tozeroy", fillcolor="rgba(169,182,196,.35)")
                fig.add_scatter(x=dd.index, y=dd.values, name="CPU daily average", line=dict(color=PAL["maroon"], width=2.5))
                if len(rsr):
                    rr_ = rsr["avg"].resample("D").mean()
                    fig.add_scatter(x=rr_.index, y=rr_.values, name="Memory daily average", line=dict(color=PAL["slate_d"], width=2.5))
                fig.update_yaxes(title="%", range=[0, 105])
                st.plotly_chart(style_fig(fig, 300, title=hsel), **STRETCH)
                row = H[H["host"] == hsel].iloc[0]
                dec = R[R["host"] == hsel]
                st.caption(f"{row['provider']} · {row['category']} · {inr(row['cost']) if row['cost'] == row['cost'] else 'cost unknown'}/month" + (f" · decision: {dec['action'].iloc[0]} — {dec['reason'].iloc[0]}" if len(dec) else ""))
    with st.expander("Azure workloads priced at WebberStop (basis of the move estimate)"):
        D = M.get("migration", pd.DataFrame())
        if len(D):
            q = D[["workload", "vms", "azure_in_scope", "webber_est", "saving", "verdict"]].rename(columns={"workload": "Azure workload", "vms": "Servers", "azure_in_scope": "Azure cost /mo", "webber_est": "WebberStop estimate /mo", "saving": "Saving /mo", "verdict": "Verdict"})
            table(q, money_cols=("Azure cost /mo", "WebberStop estimate /mo", "Saving /mo"))
            st.caption("WebberStop estimate = vCPU, RAM and disk of each server (RAM and disk from Zabbix; vCPU assumed RAM ÷ 4) priced at the rates WebberStop charges on the current bill. "
                       "Azure cost includes the servers' disks, snapshots and network items. 'Needs server sizes' = servers not in Zabbix. Verdicts: Already moving = counted under Step 1; Compliance check first = names suggest regulated or location-bound data.")
    with st.expander("Data used and checks"):
        rows = []
        for k, n in FOLDER_NAMES.items():
            d = M["dirs"][k]
            fl = [os.path.join(r, f) for r, _, fs in os.walk(d) for f in fs] if d else []
            rows.append({"Folder": n, "Files": len(fl), "Size (MB)": round(sum(os.path.getsize(f) for f in fl) / 1e6)})
        table(pd.DataFrame(rows))
        zs = aud.get("span")
        if zs:
            st.write(f"Zabbix hourly utilisation: {zs[0]:%d %b} – {zs[1]:%d %b %Y}, {len(aud['files'])} files, {H['host'].nunique()} hosts. Cost: {T.iloc[0]['label']} – {T.iloc[-1]['label']}.")
        if len(H):
            vc = H["provider"].value_counts()
            st.write(f"Zabbix hosts matched to the bills (by public IP, then exact name): {vc.get('Webber', 0)} WebberStop, {vc.get('Azure', 0)} Azure, "
                     f"{vc.get('Unmatched', 0) + vc.get('Ambiguous', 0)} on neither bill. {len(M['unmonitored'])} billed WebberStop servers have no monitoring.")
            um = H[H["provider"].isin(["Unmatched", "Ambiguous"])]
            if len(um):
                st.markdown("**Monitored hosts that match neither bill** (another provider, or named differently; add host_aliases.csv with cost_name,zabbix_name to the Zabbix Logs folder to link them):")
                table(um[["host", "ip", "cpu_avg", "ram_gb"]].rename(columns={"host": "Host", "ip": "IP", "cpu_avg": "CPU avg %", "ram_gb": "RAM GB"}))
        rec = T[["label", "azure_gross", "azure_net", "webber_servers", "webber_shared", "total"]]
        st.markdown("**Reconciliation to the bills**")
        table(rec, money_cols=tuple(rec.columns[1:]), names={"label": "Month", "azure_gross": "Azure billed", "azure_net": "Azure net", "webber_servers": "WebberStop servers", "webber_shared": "WebberStop shared", "total": "Total"})
        for n in M["notes"]:
            st.warning(n)
    with st.expander("Limits and assumptions used (fixed from the data, not settings)"):
        st.markdown(f"- **Window:** the last {cfg['window_days']} days of hourly readings; a server needs 60% coverage to be judged.\n"
                    f"- **Why these limits:** halving vCPUs doubles CPU %, so a server is proposed for a smaller size only if its busy-hour (p95) load is ≤ {cfg['cpu_p95_max']:.0f}% "
                    f"(→ ≤ 70% after) **and** no more than {cfg['breach_share']:.2%} of hours peaked above 50% (those hours would saturate). Databases keep at least {cfg['db_floor']:.0f} vCPU and are 'verify first'.\n"
                    f"- **Idle:** busiest 1% of hours under {cfg['idle_pk99']:.0f}% CPU, p95 under {cfg['idle_p95']:.0f}%, memory p95 under {cfg['idle_ram_p95']:.0f}%.\n"
                    f"- **Capacity risk:** CPU p95 ≥ 80%, memory p99 ≥ 97% (non-database) or a disk ≥ 90% full.\n"
                    f"- **Azure partner discount {cfg['azure_discount']:.0%}** is an assumption (the Azure files are gross).\n"
                    f"- **Percentages that are assumptions, never counted as confirmed:** reservations {cfg['ri_pct']:.0%}, snapshot retention {cfg['snap_pct']:.0%}, backup storage {cfg['backup_pct']:.0%}, other Azure services {cfg['platform_pct']:.0%}.\n"
                    f"- **Not in the data:** Azure data-transfer and licence meters, WebberStop quotes for new servers, vCPU counts of Azure servers, compliance status of any workload.")



# ------------------------------------------------------------------------------ Trends tab
def _icon(v):
    if abs(v) < 1:
        return '<span class="ct-flat">&#9679;</span>'
    return '<span class="ct-down">&#10004; &#9660;</span>' if v < 0 else '<span class="ct-up">&#9650;</span>'


def _chg(v, base=None):
    pct = f" ({v / base:+.1%})" if base else ""
    cls = "ct-down" if v < -1 else ("ct-up" if v > 1 else "ct-flat")
    sign = "+" if v > 0 else ("-" if v < 0 else "")
    return f'<span class="{cls}">{sign}{inr(abs(v))}{pct}</span>'


def trend_tile(label, prev, cur, m0, m1):
    d = cur - prev
    return (f'<div class="ct-kpi {"good" if d < -1 else "accent" if d > 1 else ""}"><div class="ct-l">{esc(label)}</div><div class="ct-v">{inr(cur)}</div>'
            f'<div class="ct-s">{month_label(m0)} {inr(prev)} → {month_label(m1)} {inr(cur)}</div>'
            f'<div class="ct-delta">{_icon(d)} {_chg(d, prev)}</div></div>')


def trend_notes(C):
    m0, m1 = month_label(C["m0"]), month_label(C["m1"])
    D = C["drivers"]
    out = []
    t0, t1 = C["azure0"] + C["webber0"], C["azure1"] + C["webber1"]
    out.append(f"<b>Total cost {'rose' if t1 > t0 else 'fell'} by {inr(abs(t1 - t0))} ({(t1 / t0 - 1):+.1%})</b> from {m0} to {m1}: {inr(t0)} → {inr(t1)}.")
    for p, a, b in (("Webber", C["webber0"], C["webber1"]), ("Azure", C["azure0"], C["azure1"])):
        d = D[D["provider"] == p]
        if not len(d):
            continue
        top = d.reindex(d["effect"].abs().sort_values(ascending=False).index).head(3)
        parts = "; ".join(f"{r['driver'].lower()} {'+' if r['effect'] > 0 else '-'}{inr(abs(r['effect']))}" + (f" ({int(r['n'])})" if r["n"] > 1 else "") for _, r in top.iterrows())
        out.append(f"<b>{'WebberStop' if p == 'Webber' else 'Azure'} {'up' if b > a else 'down'} {inr(abs(b - a))}.</b> Largest drivers: {parts}.")
    return out


def _chg_table(df, n=12):
    if not len(df):
        return "<p class='ct-note'>No changes.</p>"
    rows = ['<table class="ct-plan"><thead><tr><th></th><th>Server / resource</th><th>What changed</th><th class="n">Before</th><th class="n">After</th><th class="n">Change</th></tr></thead><tbody>']
    for _, r in df.head(n).iterrows():
        rows.append(f'<tr><td>{_icon(r["change"])}</td><td><b>{esc(r["server"])}</b><br><span class="ct-note">{esc(r["line"])}</span></td>'
                    f'<td class="what"><b>{esc(r["driver"])}.</b> {esc(r["what"])}</td><td class="n">{inr(r["before"])}</td><td class="n">{inr(r["after"])}</td><td class="n">{_chg(r["change"])}</td></tr>')
    rows.append("</tbody></table>")
    return "".join(rows)


def _drivers_table(D, p):
    d = D[D["provider"] == p].copy()
    if not len(d):
        return ""
    d = d.reindex(d["effect"].abs().sort_values(ascending=False).index)
    rows = ['<table class="ct-plan"><thead><tr><th></th><th>Reason</th><th class="n">Servers / items</th><th class="n">Effect on monthly cost</th></tr></thead><tbody>']
    for _, r in d.iterrows():
        rows.append(f'<tr><td>{_icon(r["effect"])}</td><td><b>{esc(r["driver"])}</b></td><td class="n">{int(r["n"])}</td><td class="n">{_chg(r["effect"])}</td></tr>')
    rows.append("</tbody></table>")
    return "".join(rows)


def trend_report(C):
    buf = io.BytesIO()
    t0, t1 = C["azure0"] + C["webber0"], C["azure1"] + C["webber1"]
    m0_, m1_ = month_label(C["m0"]), month_label(C["m1"])
    S = pd.DataFrame([{"Provider": "Total", m0_: t0, m1_: t1, "Change": t1 - t0}, {"Provider": "Azure", m0_: C["azure0"], m1_: C["azure1"], "Change": C["azure1"] - C["azure0"]},
                      {"Provider": "WebberStop", m0_: C["webber0"], m1_: C["webber1"], "Change": C["webber1"] - C["webber0"]}])
    cols = {"server": "Server / resource", "line": "Line / type", "driver": "Reason", "before": "Before", "after": "After", "change": "Change", "what": "What changed"}
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        S.to_excel(xw, sheet_name="Summary", index=False)
        C["drivers"].rename(columns={"provider": "Provider", "driver": "Reason", "effect": "Effect on monthly cost", "n": "Servers / items"}).to_excel(xw, sheet_name="Reasons", index=False)
        C["webber"].rename(columns=cols).to_excel(xw, sheet_name="WebberStop changes", index=False)
        C["azure"].rename(columns=cols).to_excel(xw, sheet_name="Azure changes", index=False)
    return buf.getvalue()


def tab_trends(M):
    months = M["months"]
    if len(months) < 2:
        st.info("Trends need at least two billed months.")
        return
    c1, c2, _ = st.columns([1, 1, 3])
    m0 = c1.selectbox("From", months, index=len(months) - 2, format_func=month_label, key="tr_m0")
    m1 = c2.selectbox("To", months, index=len(months) - 1, format_func=month_label, key="tr_m1")
    if m0 >= m1:
        st.warning("Pick a 'From' month earlier than the 'To' month.")
        return
    key = ("cmp", id(M), m0, m1)
    if key not in st.session_state:
        for k in [k for k in st.session_state if isinstance(k, tuple) and k and k[0] == "cmp"]:
            del st.session_state[k]
        st.session_state[key] = month_compare(M, m0, m1)
    C = st.session_state[key]
    t0, t1 = C["azure0"] + C["webber0"], C["azure1"] + C["webber1"]
    c = st.columns(3)
    c[0].markdown(trend_tile("Total cost", t0, t1, m0, m1), unsafe_allow_html=True)
    c[1].markdown(trend_tile("WebberStop", C["webber0"], C["webber1"], m0, m1), unsafe_allow_html=True)
    c[2].markdown(trend_tile("Azure", C["azure0"], C["azure1"], m0, m1), unsafe_allow_html=True)
    st.caption("&#10004; &#9660; green = cost fell · &#9650; red = cost rose. Colour follows the effect on cost, so a planned capacity upgrade also shows red.")
    st.markdown("##### Summary for management")
    for f in trend_notes(C):
        st.markdown(f'<div class="ct-find">{f}</div>', unsafe_allow_html=True)
    D = C["drivers"].copy()
    D["label"] = D["provider"].map({"Webber": "WebberStop", "Azure": "Azure"}) + ": " + D["driver"]
    D = D.reindex(D["effect"].abs().sort_values(ascending=False).index).head(14).sort_values("effect", ascending=False)
    fig = go.Figure(go.Waterfall(x=list(D["label"]) + ["Net change"], y=list(D["effect"]) + [0], measure=["relative"] * len(D) + ["total"],
                                 text=[("+" if v > 0 else "-") + inr(abs(v)) for v in D["effect"]] + [""], textposition="outside",
                                 increasing=dict(marker=dict(color=PAL["maroon"])), decreasing=dict(marker=dict(color=PAL["good"])), totals=dict(marker=dict(color=PAL["slate_d"])),
                                 connector=dict(line=dict(color=PAL["line"]))))
    fig = style_fig(fig, 420, False, f"Why cost changed, {month_label(m0)} → {month_label(m1)} (red adds cost, green removes it)")
    fig.update_xaxes(tickangle=-30)
    fig.update_traces(cliponaxis=False)
    cs = np.cumsum(D["effect"].values)
    rupee_axis(fig, max(float(cs.max()), float(D["effect"].sum()), 1) * 1.25, min(float(cs.min()) * 1.25, 0))
    st.plotly_chart(fig, **STRETCH)
    w, a = st.tabs(["WebberStop", "Azure"])
    for tab, key_, nm in ((w, "webber", "Webber"), (a, "azure", "Azure")):
        with tab:
            st.markdown(f"**Reasons**")
            st.markdown(_drivers_table(C["drivers"], nm), unsafe_allow_html=True)
            st.markdown("**Largest changes, server by server**")
            R_ = C[key_]
            st.markdown(_chg_table(R_), unsafe_allow_html=True)
            if len(R_) > 12:
                with st.expander(f"All {len(R_)} changes"):
                    opts = sorted(R_["driver"].unique())
                    pick = st.multiselect("Reason", opts, default=opts, key="trf_" + key_)
                    q = R_[R_["driver"].isin(pick)].rename(columns={"server": "Server / resource", "line": "Line / type", "driver": "Reason", "before": "Before", "after": "After", "change": "Change", "what": "What changed"})
                    table(q, money_cols=("Before", "After", "Change"), names={c_: c_ for c_ in q.columns})
    T = M["monthly"]
    ch = T["total"].diff().iloc[1:]
    fig = go.Figure(go.Bar(x=[month_label(m) for m in T["month"].iloc[1:]], y=ch, marker_color=[PAL["good"] if v < 0 else PAL["maroon"] for v in ch],
                           text=[("+" if v > 0 else "-") + inr(abs(v)) for v in ch], textposition="outside"))
    fig = style_fig(fig, 300, False, "Month-on-month change in total cost (green = fell, red = rose)")
    fig.update_traces(cliponaxis=False)
    rupee_axis(fig, max(float(ch.max()), 1) * 1.3, min(float(ch.min()) * 1.3, -1))
    st.plotly_chart(fig, **STRETCH)
    st.download_button("Download this comparison for the management report (Excel)", trend_report(C), file_name=f"cost_trend_{m0}_to_{m1}.xlsx", key="dl_trend")


def ensure_server_config():
    """Allow large uploads (Zabbix hourly files are big). Takes effect the next time Streamlit starts."""
    path = os.path.join(".streamlit", "config.toml")
    try:
        old = open(path, encoding="utf-8").read() if os.path.exists(path) else ""
        if "maxUploadSize" not in old:
            os.makedirs(".streamlit", exist_ok=True)
            with open(path, "a", encoding="utf-8") as f:
                f.write(("\n" if old and not old.endswith("\n") else "") + "\n[server]\nmaxUploadSize = 4096\n")
    except OSError:
        pass


def data_sidebar(sb, base):
    """Saved data: status, save from the input folder, or replace with uploads. Every save replaces everything."""
    info = store_info()
    sb.markdown("### Saved data")
    if info["ready"]:
        k = info["kinds"]
        sb.success(f"Using saved data ({info['source']}, {info['saved_at']}): {k['azure']['files']} Azure, {k['webber']['files']} WebberStop and {k['zabbix']['files']} Zabbix files.")
    else:
        sb.info("Nothing saved yet: the app reads the input folder directly. Save once so the data stays available.")
    n = st.session_state.setdefault("up_n", 0)
    with sb.expander("Save or replace the data", expanded=not info["ready"]):
        st.caption("Each save **replaces everything** stored, so the app always holds one clean, current set.")
        if st.button("Save the files from the input folder", key="save_folder"):
            with st.spinner("Saving…"):
                res, err = store_from_folder(base)
            if err:
                st.error(err)
            else:
                st.cache_resource.clear()
                st.rerun()
        st.markdown("**or upload files**")
        ua = st.file_uploader("Azure cost files (.xlsx)", accept_multiple_files=True, type=["xlsx", "xls", "xlsm"], key=f"up_az_{n}")
        uw = st.file_uploader("WebberStop bills (.xls / .xlsx)", accept_multiple_files=True, type=["xls", "xlsx"], key=f"up_wb_{n}")
        uz = st.file_uploader("Zabbix audit exports (.csv)", accept_multiple_files=True, type=["csv"], key=f"up_zb_{n}")
        st.caption("Zabbix hourly files are large (hundreds of MB each); saving from the input folder is faster than uploading them.")
        if st.button("Replace saved data with these uploads", key="save_up", disabled=not (ua and uw and uz)):
            with st.spinner("Saving…"):
                res, err = store_from_uploads({"azure": ua, "webber": uw, "zabbix": uz})
            if err:
                st.error(err)
            else:
                st.session_state["up_n"] = n + 1
                st.cache_resource.clear()
                st.rerun()
    return info


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="₹", layout="wide")
    ensure_theme_file()
    ensure_server_config()
    st.markdown(CSS + CSS2, unsafe_allow_html=True)
    sb = st.sidebar
    sb.markdown("### Input folder")
    inbase = sb.text_input("Folder holding Azure Cost Logs, Webber Cost Logs and Zabbix Logs", value=st.session_state.get("base", default_base()), key="base_in")
    st.session_state["base"] = inbase
    sb.caption("Used when nothing is saved, and as the source for 'Save the files from the input folder'. Limits are fixed from the data (Evidence → Limits and assumptions).")
    data_sidebar(sb, inbase)
    base = active_base(inbase)
    if sb.button("Re-read all files"):
        st.cache_resource.clear()
    cfg = dict(DEFAULT_CFG)
    st.markdown(f'<div class="ct-title">{APP_TITLE}</div>', unsafe_allow_html=True)
    if not os.path.isdir(base):
        st.error(f"Folder not found: {base}")
        return
    sig = folder_signature(base)
    with st.spinner("Reading cost and Zabbix files (the first run takes a few minutes for the hourly files; later runs use the cache)…"):
        M = get_model(base, sig, tuple(sorted(cfg.items())), DEFAULT_YEAR)
    if M.get("error"):
        st.error(M["error"])
        st.info("Put 'Azure Cost Logs', 'Webber Cost Logs' and 'Zabbix Logs' inside the folder above (or point the sidebar at it).")
        return
    T = M["monthly"]
    zs = M["aud"].get("span")
    st.markdown(f'<div class="ct-sub">Cost {T.iloc[0]["label"]} – {T.iloc[-1]["label"]}'
                + (f' · Zabbix utilisation {zs[0]:%d %b} – {zs[1]:%d %b %Y}' if zs else " · no Zabbix utilisation data") + "</div>", unsafe_allow_html=True)
    t = st.tabs(["Overview", "Trends", "Action plan", "Evidence"])
    with t[0]:
        tab_overview(M)
    with t[1]:
        tab_trends(M)
    with t[2]:
        tab_plan(M)
    with t[3]:
        tab_evidence(M)


if __name__ == "__main__":
    main()
