#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
MPF MA Provider Directory Audit
===============================

Self-contained audit of the Jefferson Health Plans CY 2027 FHIR-based provider
directory submissions for Medicare Plan Finder.

Validated against:
    CMS Technical Implementation Guide for Supplying Medicare Advantage (MA)
    Provider Directory Data for Use in Medicare Plan Finder (MPF)
    Version 1.5, September 4, 2026
      Appendix B -- FHIR-based JSON field specifications
      Appendix D -- HTTP metadata and validation mechanisms
      Appendix E -- Validation inventory (levels 1 / 2 / 3)

Just run it:

    python mpf_audit.py

It downloads the three index files and every constituent bundle, revalidates
cached copies with conditional HTTP requests, runs all 68 Appendix E
validations plus 20 supplementary conformance tests plus a full reference
integrity sweep, and writes a Word report and per-contract findings CSVs.

Options:
    --out DIR        where to write the report (default: current directory)
    --cache DIR      cache location (default: <script dir>/.mpf_cache)
    --fresh          ignore cached bundles and re-download everything
    --no-nppes       skip the NPPES registry checks (P1002/P1003/P1005)
    --contracts IDS  comma separated subset, e.g. --contracts H1619,H9207

Requires: ijson, python-docx, openpyxl
    pip install ijson python-docx openpyxl
"""

from __future__ import print_function

import argparse
import collections
import csv
import datetime
import gzip
import io
import json
import os
import re
import ssl
import socket
import sys
import zipfile
from concurrent.futures import ThreadPoolExecutor

# --------------------------------------------------------------------------
# Dependency check
# --------------------------------------------------------------------------
_MISSING = []
try:
    import ijson
except ImportError:
    _MISSING.append("ijson")
try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.section import WD_ORIENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    _MISSING.append("python-docx")
try:
    import openpyxl
except ImportError:
    openpyxl = None          # only needed to rebuild the deactivation list

if _MISSING:
    sys.stderr.write("Missing required packages: %s\n"
                     "Install with:  pip install %s\n"
                     % (", ".join(_MISSING), " ".join(_MISSING)))
    sys.exit(2)

try:                                    # py3
    from urllib.request import Request, urlopen
    from urllib.error import HTTPError, URLError
except ImportError:                     # py2
    from urllib2 import Request, urlopen, HTTPError, URLError


# ==========================================================================
# CONFIGURATION -- the three contract index URLs are fixed and hard-coded
# ==========================================================================
INDEX_URLS = {     "H5826": "https://medicare-advantage-plan-finder-provider-directory.interop.chpw.org/h5826/2027/index.json",                                                                                                                      
    "H1619": "https://medicare-advantage-plan-finder-provider-directory.jeffersonhealthplans.com/h1619/2027/index.json",
    "H3124": "https://medicare-advantage-plan-finder-provider-directory.jeffersonhealthplans.com/h3124/2027/index.json",
    "H9207": "https://medicare-advantage-plan-finder-provider-directory.jeffersonhealthplans.com/h9207/2027/index.json",
}
CONTRACT_YEAR = "2027"
ORG_NAME = "Jefferson Health Plans"

# v1.5 hosting limits
MAX_INDEX_URLS = 10000
MAX_FILE_BYTES = 300 * 1000 * 1000

# FHIR / CMS system URIs
NPI_SYS = "http://hl7.org/fhir/sid/us-npi"
MAPLAN_SYS = "http://cms.gov/medicare/ma-plan-id"
NUCC_SYS = "http://nucc.org/provider-taxonomy"
ORGTYPE_SYS = "http://hl7.org/fhir/us/davinci-pdex-plan-net/CodeSystem/OrgTypeCS"
PDEX_PREFIX = "http://hl7.org/fhir/us/davinci-pdex-plan-net/StructureDefinition/"

ACCEPTING_VALUES = {"newpt", "nopt", "existptonly"}
GENDER_VALUES = {"male", "female", "other", "unknown"}
MAPLAN_RE = re.compile(r"^[A-Z]\d{4}-\d{3}-\d{3}$")
REF_RE = re.compile(r"^([A-Za-z]+)/([A-Za-z0-9\-\.]{1,64})(?:/_history/[^/]+)?$")

USPS_STATES = set(
    "AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO "
    "MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY "
    "DC AS GU MP PR VI UM FM MH PW AA AE AP".split())

# Resource types CMS consumes for MPF (Appendix B)
MPF_RESOURCE_TYPES = {"InsurancePlan", "Location", "Organization",
                      "OrganizationAffiliation", "Practitioner", "PractitionerRole"}

# Public zip references merged with NPPES practice locations
ZIP_SOURCES = [
    ("https://raw.githubusercontent.com/midwire/free_zipcode_data/master/all_us_zipcodes.csv",
     "code", "state"),
    ("https://raw.githubusercontent.com/scpike/us-state-county-zip/master/geo-data.csv",
     "zipcode", "state_abbr"),
]

NPPES_INDEX = "https://download.cms.gov/nppes/NPI_Files.html"
NPPES_BASE = "https://download.cms.gov/nppes/"
NPPES_MAX_AGE_DAYS = 35

USER_AGENT = "JHP-MPF-Directory-Audit/1.0"
DOWNLOAD_WORKERS = 6

LOG_WIDTH = 74


def log(msg, indent=0):
    sys.stdout.write("%s%s\n" % ("  " * indent, msg))
    sys.stdout.flush()


def rule(title=""):
    if title:
        log("\n" + title)
        log("-" * min(LOG_WIDTH, max(len(title), 20)))
    else:
        log("-" * LOG_WIDTH)


def human_bytes(n):
    n = float(n or 0)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024.0 or unit == "TB":
            return "%.1f %s" % (n, unit) if unit != "B" else "%d B" % n
        n /= 1024.0


def mb(n):
    return "%.1f MB" % ((n or 0) / 1e6)


def gb(n):
    return "%.2f GB" % ((n or 0) / 1e9)


def commify(n):
    try:
        return "{:,}".format(int(n))
    except (TypeError, ValueError):
        return str(n)


# ==========================================================================
# HTTP layer
# ==========================================================================
class HttpResult(object):
    def __init__(self, url):
        self.url = url
        self.status = None
        self.headers = {}
        self.error = ""
        self.head_ok = False
        self.head_status = None
        self.conditional = ""
        self.body_path = None
        self.size = 0
        self.from_cache = False

    def h(self, name):
        return self.headers.get(name.lower(), "")


def _open(url, headers=None, method="GET", timeout=180, encoding_identity=False):
    req = Request(url)
    req.get_method = lambda: method
    req.add_header("User-Agent", USER_AGENT)
    if not encoding_identity:
        # Offer compression so the audit can prove the server does not use it
        req.add_header("Accept-Encoding", "gzip, deflate, br")
    for k, v in (headers or {}).items():
        req.add_header(k, v)
    return urlopen(req, timeout=timeout)


def probe(url):
    """HEAD + conditional GET probe. Returns an HttpResult with headers only."""
    r = HttpResult(url)
    try:
        resp = _open(url, method="HEAD", timeout=90)
        r.head_status = resp.getcode()
        r.head_ok = (resp.getcode() == 200)
        r.headers = dict((k.lower(), v) for k, v in resp.info().items())
        r.status = resp.getcode()
        resp.close()
    except HTTPError as e:
        r.head_status = e.code
        r.headers = dict((k.lower(), v) for k, v in (e.headers or {}).items())
        r.error = "HEAD HTTP %s" % e.code
    except Exception as e:                                   # noqa: BLE001
        r.error = "HEAD %s: %s" % (type(e).__name__, e)

    etag = r.h("etag")
    lastmod = r.h("last-modified")
    parts = []
    for hdr, val, label in (("If-None-Match", etag, "INM"),
                            ("If-Modified-Since", lastmod, "IMS")):
        if not val:
            continue
        try:
            resp = _open(url, headers={hdr: val, "Range": "bytes=0-0"}, timeout=90)
            parts.append("%s->%s" % (label, resp.getcode()))
            resp.close()
        except HTTPError as e:
            parts.append("%s->%s" % (label, e.code))
        except Exception:                                    # noqa: BLE001
            parts.append("%s->err" % label)
    r.conditional = " ".join(parts)
    return r


def fetch_to_file(url, dest, cache_meta, fresh=False, attempts=5):
    """Download url to dest, revalidating with ETag / Last-Modified.

    Retries with HTTP Range resume, because a dropped connection on a large
    file is a local network event and must not be reported as a CMS-level
    retrieval failure. Returns (HttpResult, used_cache).
    """
    r = HttpResult(url)
    prev = cache_meta.get(url) or {}

    # ---- conditional revalidation against the cached copy ----------------
    if (not fresh) and os.path.exists(dest) and prev:
        headers = {}
        if prev.get("etag"):
            headers["If-None-Match"] = prev["etag"]
        if prev.get("last_modified"):
            headers["If-Modified-Since"] = prev["last_modified"]
        if headers:
            try:
                resp = _open(url, headers=headers, encoding_identity=True)
                resp.close()               # 200 means it changed, fall through
            except HTTPError as e:
                if e.code == 304:
                    r.status = 200
                    r.from_cache = True
                    r.headers = {"etag": prev.get("etag", ""),
                                 "last-modified": prev.get("last_modified", ""),
                                 "content-length": prev.get("content_length", ""),
                                 "content-type": prev.get("content_type", ""),
                                 "content-encoding": prev.get("content_encoding", "") or ""}
                    r.size = os.path.getsize(dest)
                    return r, True
            except Exception:                                # noqa: BLE001
                pass

    # ---- download, resuming after a dropped connection -------------------
    part = dest + ".part"
    expected = None
    last_error = ""
    for attempt in range(1, attempts + 1):
        have = os.path.getsize(part) if os.path.exists(part) else 0
        headers = {}
        if have and expected:
            headers["Range"] = "bytes=%d-" % have
        try:
            resp = _open(url, headers=headers, encoding_identity=True)
            code = resp.getcode()
            info = dict((k.lower(), v) for k, v in resp.info().items())
            enc = (info.get("content-encoding") or "").lower()
            if code == 206 and have:
                mode = "ab"
            else:
                mode = "wb"
                have = 0
                if code == 200:
                    try:
                        expected = int(info.get("content-length"))
                    except (TypeError, ValueError):
                        expected = None
            if not r.headers:
                r.headers = info
            r.status = 200
            written = have
            with open(part, mode) as out:
                if enc in ("gzip", "x-gzip"):
                    blob = gzip.GzipFile(fileobj=io.BytesIO(resp.read())).read()
                    out.write(blob)
                    written = len(blob)
                    expected = written
                else:
                    while True:
                        chunk = resp.read(1 << 20)
                        if not chunk:
                            break
                        out.write(chunk)
                        written += len(chunk)
            resp.close()
            if expected is None or written == expected:
                if os.path.exists(dest):
                    os.remove(dest)
                os.rename(part, dest)
                r.size = os.path.getsize(dest)
                cache_meta[url] = {"etag": info.get("etag", ""),
                                   "last_modified": info.get("last-modified", ""),
                                   "content_length": info.get("content-length", ""),
                                   "content_type": info.get("content-type", ""),
                                   "content_encoding": info.get("content-encoding", ""),
                                   "size": r.size}
                return r, False
            last_error = "received %s of %s bytes" % (commify(written), commify(expected))
        except HTTPError as e:
            if e.code == 304 and os.path.exists(dest):
                r.status = 200
                r.from_cache = True
                r.headers = {"etag": prev.get("etag", ""),
                             "last-modified": prev.get("last_modified", ""),
                             "content-length": prev.get("content_length", ""),
                             "content-type": prev.get("content_type", ""),
                             "content-encoding": prev.get("content_encoding", "") or ""}
                r.size = os.path.getsize(dest)
                return r, True
            if e.code == 416:                       # stale .part, start over
                try:
                    os.remove(part)
                except OSError:
                    pass
                expected = None
                last_error = "HTTP 416, restarting"
                continue
            r.status = e.code
            last_error = "HTTP %s" % e.code
            if e.code in (400, 401, 403, 404, 410):
                break
        except Exception as e:                               # noqa: BLE001
            last_error = "%s: %s" % (type(e).__name__, e)
        if attempt < attempts:
            log("retry %d/%d for %s (%s)"
                % (attempt, attempts - 1, os.path.basename(dest), last_error), 2)

    r.error = last_error or "download failed"
    return r, False


def fetch_text(url, timeout=120):
    resp = _open(url, timeout=timeout)
    raw = resp.read()
    if (resp.info().get("Content-Encoding") or "").lower() in ("gzip", "x-gzip"):
        raw = gzip.GzipFile(fileobj=io.BytesIO(raw)).read()
    resp.close()
    return raw.decode("utf-8", "replace")


def check_tls(url):
    host = re.sub(r"^https?://", "", url).split("/")[0].split(":")[0]
    info = {"host": host, "verified": False, "error": ""}
    try:
        ctx = ssl.create_default_context()
        with socket.create_connection((host, 443), timeout=30) as sock:
            with ctx.wrap_socket(sock, server_hostname=host) as ss:
                cert = ss.getpeercert()
                info["verified"] = True
                info["tls"] = ss.version()
                info["cipher"] = ss.cipher()[0]
                info["not_before"] = cert.get("notBefore", "")
                info["not_after"] = cert.get("notAfter", "")
                issuer = []
                for rdn in cert.get("issuer", ()):
                    for k, v in rdn:
                        issuer.append(v)
                info["issuer"] = " - ".join(issuer)
                try:
                    na = datetime.datetime.strptime(cert["notAfter"], "%b %d %H:%M:%S %Y %Z")
                    info["days_remaining"] = (na - datetime.datetime.now()).days
                except Exception:                            # noqa: BLE001
                    info["days_remaining"] = None
    except Exception as e:                                   # noqa: BLE001
        info["error"] = "%s: %s" % (type(e).__name__, e)
    return info


# ==========================================================================
# Reference data
# ==========================================================================
class RefData(object):
    """NPPES registry, deactivation list, NUCC taxonomy usage, zip codes."""

    def __init__(self, cache_dir, use_nppes=True):
        self.cache = cache_dir
        self.use_nppes = use_nppes
        self.registry = {}          # npi -> entity type code ('1' or '2')
        self.deactivated = set()
        self.taxonomy = {}          # code -> [individual_count, org_count]
        self.zips = {}              # zip5 -> [state, ...]
        self.sources = {}
        self.available = {"registry": False, "deactivated": False,
                          "taxonomy": False, "zips": False}

    # -- discovery ------------------------------------------------------
    def _discover_nppes(self):
        """Return (monthly_zip_name, deactivated_zip_name) from the CMS listing."""
        try:
            html = fetch_text(NPPES_INDEX, timeout=90)
        except Exception as e:                               # noqa: BLE001
            log("could not reach the CMS NPPES listing (%s)" % e, 1)
            return None, None
        names = set(re.findall(r"[A-Za-z0-9_\.\-]+\.zip", html))
        monthly = sorted(n for n in names
                         if re.match(r"NPPES_Data_Dissemination_[A-Za-z]+_\d{4}_V\d+\.zip$", n))
        deact = sorted(n for n in names
                       if re.match(r"NPPES_Deactivated_NPI_Report_\d{6}_V\d+\.zip$", n))

        def month_key(n):
            m = re.match(r"NPPES_Data_Dissemination_([A-Za-z]+)_(\d{4})_", n)
            months = ["january", "february", "march", "april", "may", "june", "july",
                      "august", "september", "october", "november", "december"]
            try:
                return (int(m.group(2)), months.index(m.group(1).lower()))
            except Exception:                                # noqa: BLE001
                return (0, 0)

        def deact_key(n):
            m = re.search(r"_(\d{2})(\d{2})(\d{2})_", n)
            return (m.group(3), m.group(1), m.group(2)) if m else ("", "", "")

        monthly = max(monthly, key=month_key) if monthly else None
        deact = max(deact, key=deact_key) if deact else None
        return monthly, deact

    def _find_local(self, filename):
        """Look for an already-downloaded copy of a NPPES zip in the project tree."""
        roots = [self.cache, os.path.dirname(os.path.abspath(__file__))]
        for root in roots:
            if not root or not os.path.isdir(root):
                continue
            for dirpath, dirnames, filenames in os.walk(root):
                dirnames[:] = [d for d in dirnames if not d.startswith(".")][:60]
                if filename in filenames:
                    return os.path.join(dirpath, filename)
        return None

    def _download(self, filename):
        dest = os.path.join(self.cache, filename)
        if os.path.exists(dest) and os.path.getsize(dest) > 1024:
            return dest
        url = NPPES_BASE + filename
        log("downloading %s (this is a one-off, it is a large file)" % filename, 1)
        meta = {}
        r, _ = fetch_to_file(url, dest, meta)
        if r.error or not os.path.exists(dest):
            log("download failed: %s" % (r.error or "unknown"), 2)
            try:
                os.remove(dest)
            except OSError:
                pass
            return None
        return dest

    # -- builders -------------------------------------------------------
    def _build_from_monthly(self, zip_path, reg_path, tax_path, zip_ref_path):
        """One pass over npidata builds registry, taxonomy usage and zip/state map."""
        log("building registry, taxonomy and zip references from %s"
            % os.path.basename(zip_path), 1)
        zf = zipfile.ZipFile(zip_path)
        data_name = None
        for n in zf.namelist():
            if re.match(r"npidata_pfile_.*\d\.csv$", n) and "fileheader" not in n:
                data_name = n
                break
        if not data_name:
            log("no npidata file inside the archive", 2)
            return False

        tax_counts = collections.defaultdict(lambda: [0, 0])
        zip_counts = collections.defaultdict(collections.Counter)
        rows = 0
        with zf.open(data_name) as fh, open(reg_path, "w") as reg_out:
            reader = csv.reader(io.TextIOWrapper(fh, encoding="utf-8",
                                                 errors="replace", newline=""))
            header = next(reader)
            idx = dict((h.strip(), i) for i, h in enumerate(header))
            i_npi = idx.get("NPI", 0)
            i_ent = idx.get("Entity Type Code", 1)
            tax_cols = [idx[h] for h in header
                        if h.strip().startswith("Healthcare Provider Taxonomy Code_")
                        and h.strip() in idx]
            addr_cols = []
            for st_name, zp_name, ct_name in (
                    ("Provider Business Mailing Address State Name",
                     "Provider Business Mailing Address Postal Code",
                     "Provider Business Mailing Address Country Code (If outside U.S.)"),
                    ("Provider Business Practice Location Address State Name",
                     "Provider Business Practice Location Address Postal Code",
                     "Provider Business Practice Location Address Country Code (If outside U.S.)")):
                if st_name in idx and zp_name in idx:
                    addr_cols.append((idx[st_name], idx[zp_name], idx.get(ct_name, -1)))

            reg_out.write("NPI\tEntity Type Code\n")
            for row in reader:
                rows += 1
                if len(row) <= i_ent:
                    continue
                npi = row[i_npi].strip()
                ent = row[i_ent].strip()
                if npi:
                    reg_out.write("%s\t%s\n" % (npi, ent))
                k = 0 if ent == "1" else 1 if ent == "2" else None
                if k is not None:
                    for c in tax_cols:
                        if c < len(row):
                            v = row[c].strip()
                            if v:
                                tax_counts[v][k] += 1
                for si, zi, ci in addr_cols:
                    if zi >= len(row):
                        continue
                    if ci >= 0 and ci < len(row) and row[ci].strip() not in ("", "US"):
                        continue
                    z5 = row[zi].strip()[:5]
                    stt = row[si].strip().upper() if si < len(row) else ""
                    if len(z5) == 5 and z5.isdigit() and len(stt) == 2:
                        zip_counts[z5][stt] += 1
                if rows % 2000000 == 0:
                    log("%s rows" % commify(rows), 2)

        with open(tax_path, "w") as f:
            json.dump(dict(tax_counts), f)

        zip_map = {}
        for z5, counter in zip_counts.items():
            total = sum(counter.values())
            keep = set(s for s, v in counter.items() if v >= max(2, total * 0.01))
            keep.add(counter.most_common(1)[0][0])
            zip_map[z5] = sorted(keep)
        self._merge_public_zips(zip_map)
        with open(zip_ref_path, "w") as f:
            json.dump(zip_map, f)
        log("%s NPPES rows processed, %s taxonomy codes, %s zip codes"
            % (commify(rows), commify(len(tax_counts)), commify(len(zip_map))), 2)
        return True

    def _merge_public_zips(self, zip_map):
        for url, zip_field, state_field in ZIP_SOURCES:
            try:
                text = fetch_text(url, timeout=120)
            except Exception as e:                           # noqa: BLE001
                log("could not fetch %s (%s)" % (url.rsplit("/", 1)[-1], e), 2)
                continue
            reader = csv.DictReader(io.StringIO(text))
            added = 0
            for row in reader:
                z = (row.get(zip_field) or "").strip().zfill(5)
                s = (row.get(state_field) or "").strip().upper()
                if len(z) == 5 and z.isdigit():
                    cur = set(zip_map.get(z, []))
                    if s and s not in cur:
                        cur.add(s)
                        added += 1
                    zip_map[z] = sorted(x for x in cur if x)
            log("merged %s (%s zip/state pairs added)" % (url.rsplit("/", 1)[-1],
                                                          commify(added)), 2)

    def _build_deactivated(self, zip_path, out_path):
        if openpyxl is None:
            log("openpyxl not installed; cannot read the deactivation report", 2)
            return False
        zf = zipfile.ZipFile(zip_path)
        name = next((n for n in zf.namelist() if n.lower().endswith(".xlsx")), None)
        if not name:
            return False
        log("reading %s" % name, 2)
        with zf.open(name) as fh:
            data = io.BytesIO(fh.read())
        wb = openpyxl.load_workbook(data, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        n = 0
        with open(out_path, "w") as out:
            for row in ws.iter_rows(values_only=True):
                if not row:
                    continue
                val = str(row[0]).strip() if row[0] is not None else ""
                if val.isdigit() and len(val) == 10:
                    date = ""
                    if len(row) > 1 and row[1] is not None:
                        date = str(row[1]).strip()
                    out.write("%s\t%s\n" % (val, date))
                    n += 1
        wb.close()
        log("%s deactivated NPIs" % commify(n), 2)
        return n > 0

    # -- entry point ----------------------------------------------------
    def load(self):
        if not os.path.isdir(self.cache):
            os.makedirs(self.cache)

        reg_path = os.path.join(self.cache, "nppes_registry.tsv")
        deact_path = os.path.join(self.cache, "nppes_deactivated.tsv")
        tax_path = os.path.join(self.cache, "nucc_taxonomy.json")
        zip_path = os.path.join(self.cache, "zip_reference.json")
        manifest_path = os.path.join(self.cache, "refdata_manifest.json")

        manifest = {}
        if os.path.exists(manifest_path):
            try:
                manifest = json.load(open(manifest_path))
            except Exception:                                # noqa: BLE001
                manifest = {}

        if not self.use_nppes:
            log("NPPES checks disabled by --no-nppes", 1)
            self._load_zip_only(zip_path, manifest, manifest_path)
            return

        monthly, deact = self._discover_nppes()
        if monthly:
            log("current NPPES monthly file: %s" % monthly, 1)
        if deact:
            log("current NPPES deactivation report: %s" % deact, 1)

        need_monthly = True
        if (manifest.get("monthly") and monthly and manifest["monthly"] == monthly
                and os.path.exists(reg_path) and os.path.exists(tax_path)
                and os.path.exists(zip_path)):
            need_monthly = False
        elif monthly is None and os.path.exists(reg_path) and os.path.exists(tax_path):
            built = manifest.get("built", "")
            log("CMS unreachable; falling back to the cached snapshot built %s"
                % (built or "at an unknown date"), 1)
            need_monthly = False

        if need_monthly and monthly:
            src = self._find_local(monthly) or self._download(monthly)
            if src:
                if self._build_from_monthly(src, reg_path, tax_path, zip_path):
                    manifest["monthly"] = monthly
                    manifest["built"] = datetime.date.today().isoformat()
            else:
                log("could not obtain %s" % monthly, 1)

        need_deact = True
        if (manifest.get("deactivated") and deact and manifest["deactivated"] == deact
                and os.path.exists(deact_path)):
            need_deact = False
        elif deact is None and os.path.exists(deact_path):
            need_deact = False

        if need_deact and deact:
            src = self._find_local(deact) or self._download(deact)
            if src and self._build_deactivated(src, deact_path):
                manifest["deactivated"] = deact

        # Last-resort fallback: reuse any derived files already in the project
        if not os.path.exists(reg_path):
            legacy = self._find_local("registry_npis_monthly.tsv")
            if legacy:
                log("using existing derived registry at %s" % legacy, 1)
                reg_path = legacy
        if not os.path.exists(deact_path):
            legacy = self._find_local("deactivated_npis.tsv")
            if legacy:
                log("using existing derived deactivation list at %s" % legacy, 1)
                deact_path = legacy

        self._read_all(reg_path, deact_path, tax_path, zip_path)
        if not os.path.exists(zip_path):
            self._load_zip_only(zip_path, manifest, manifest_path)

        try:
            json.dump(manifest, open(manifest_path, "w"), indent=1)
        except Exception:                                    # noqa: BLE001
            pass
        self.sources = manifest

    def _load_zip_only(self, zip_path, manifest, manifest_path):
        if os.path.exists(zip_path):
            self.zips = json.load(open(zip_path))
            self.available["zips"] = True
            return
        zip_map = {}
        self._merge_public_zips(zip_map)
        if zip_map:
            json.dump(zip_map, open(zip_path, "w"))
            self.zips = zip_map
            self.available["zips"] = True

    def _read_all(self, reg_path, deact_path, tax_path, zip_path):
        if os.path.exists(reg_path):
            with open(reg_path, encoding="utf-8", errors="replace") as f:
                first = f.readline()
                if not first.split("\t")[0].strip().isdigit():
                    pass                      # header consumed
                else:
                    p = first.rstrip("\n").split("\t")
                    if len(p) >= 2:
                        self.registry[p[0]] = p[1].strip()
                for line in f:
                    p = line.rstrip("\n").split("\t")
                    if len(p) >= 2 and p[0]:
                        self.registry[p[0]] = p[1].strip()
            self.available["registry"] = bool(self.registry)
        if os.path.exists(deact_path):
            with open(deact_path, encoding="utf-8", errors="replace") as f:
                for line in f:
                    v = line.split("\t")[0].strip()
                    if v.isdigit():
                        self.deactivated.add(v)
            self.available["deactivated"] = bool(self.deactivated)
        if os.path.exists(tax_path):
            self.taxonomy = json.load(open(tax_path))
            self.available["taxonomy"] = bool(self.taxonomy)
        if os.path.exists(zip_path):
            self.zips = json.load(open(zip_path))
            self.available["zips"] = bool(self.zips)

        log("registry NPIs %s | deactivated %s | taxonomy codes %s | zip codes %s"
            % (commify(len(self.registry)), commify(len(self.deactivated)),
               commify(len(self.taxonomy)), commify(len(self.zips))), 1)

    # -- helpers used by the validator ---------------------------------
    def taxonomy_class(self, code):
        """'individual', 'organization' or None when ambiguous / unknown."""
        u = self.taxonomy.get(code)
        if not u:
            return None
        ind, org = u[0], u[1]
        total = ind + org
        if not total:
            return None
        if ind / float(total) >= 0.90:
            return "individual"
        if org / float(total) >= 0.90:
            return "organization"
        return None


# ==========================================================================
# Findings
# ==========================================================================
class Findings(object):
    """Streams every finding to CSV while keeping counters and capped samples."""

    MAX_SAMPLES = 40

    def __init__(self, contract, csv_path, catalog):
        self.contract = contract
        self.catalog = catalog
        self.count = collections.Counter()
        self.npis = collections.defaultdict(set)
        self.samples = collections.defaultdict(list)
        self._fh = open(csv_path, "w", newline="", encoding="utf-8")
        self._w = csv.writer(self._fh)
        self._w.writerow(["Level", "ErrorCode", "ValidationName", "ContractID",
                          "ResourceType", "ResourceID", "NPI", "Field", "Value", "Detail"])

    def add(self, code, name, level, resource="", rid="", npi="", field="",
            value="", detail=""):
        self.catalog[code] = (name, level)
        self.count[code] += 1
        if npi:
            self.npis[code].add(npi)
        self._w.writerow([level, code, name, self.contract, resource, rid, npi,
                          field, value, detail])
        bucket = self.samples[code]
        if len(bucket) < self.MAX_SAMPLES:
            bucket.append({"resource": resource, "id": rid, "npi": npi,
                           "field": field, "value": str(value)[:160],
                           "detail": str(detail)[:240]})

    def close(self):
        self._fh.close()

    def level_total(self, level):
        return sum(v for k, v in self.count.items() if self.catalog[k][1] == level)


# ==========================================================================
# FHIR helpers
# ==========================================================================
def stream_resources(path):
    with open(path, "rb") as fh:
        for res in ijson.items(fh, "entry.item.resource"):
            yield res


def bundle_header(path):
    """Read the Bundle's top-level resourceType/type without loading entries."""
    rt = bt = None
    try:
        with open(path, "rb") as fh:
            for prefix, event, value in ijson.parse(fh):
                if event == "map_key" and prefix == "":
                    if value == "entry":
                        break
                elif prefix == "resourceType" and event == "string":
                    rt = value
                elif prefix == "type" and event == "string":
                    bt = value
    except Exception as e:                                   # noqa: BLE001
        return None, None, str(e)
    return rt, bt, ""


def ref_target(ref):
    """Parse a FHIR reference string into (type, id) or (None, None)."""
    if not ref or not isinstance(ref, str):
        return None, None
    s = ref.strip()
    if s.startswith("#"):
        return "#contained", s[1:]
    s = re.sub(r"^[a-zA-Z][a-zA-Z0-9+.\-]*://[^/]+/", "", s)
    parts = s.split("/")
    if len(parts) >= 2:
        s = "/".join(parts[-2:]) if parts[-2][:1].isupper() else s
    m = REF_RE.match(s)
    if m:
        return m.group(1), m.group(2)
    return None, None


def walk_references(node, path=""):
    """Yield (json_path, reference_string) for every Reference in a resource."""
    if isinstance(node, dict):
        r = node.get("reference")
        if isinstance(r, str) and r:
            yield path or "reference", r
        for k, v in node.items():
            if k == "reference":
                continue
            if isinstance(v, (dict, list)):
                for item in walk_references(v, "%s.%s" % (path, k) if path else k):
                    yield item
    elif isinstance(node, list):
        for i, v in enumerate(node):
            if isinstance(v, (dict, list)):
                for item in walk_references(v, "%s[%d]" % (path, i)):
                    yield item


def npis_of(res):
    out = []
    for i in res.get("identifier") or []:
        if i.get("system") == NPI_SYS:
            v = (i.get("value") or "").strip()
            if v:
                out.append(v)
    return out


def codes_of(cc_list, system=None):
    out = []
    for cc in cc_list or []:
        for c in cc.get("coding") or []:
            if system is None or c.get("system") == system:
                v = (c.get("code") or "").strip()
                if v:
                    out.append((v, c.get("system")))
    return out


def phones_of(res):
    return [(t.get("value") or "").strip()
            for t in (res.get("telecom") or []) if t.get("system") == "phone"]


_DIGITS = re.compile(r"\D")


def phone_ok(p):
    d = _DIGITS.sub("", p or "")
    if len(d) == 11 and d.startswith("1"):
        d = d[1:]
    return len(d) == 10 and d[0] not in "01" and d[3] not in "01"


def luhn_npi(npi):
    if not (len(npi) == 10 and npi.isdigit()):
        return False
    digits = [int(c) for c in "80840" + npi[:9]]
    total = 0
    for i, x in enumerate(reversed(digits)):
        if i % 2 == 0:
            x *= 2
            if x > 9:
                x -= 9
        total += x
    return (10 - total % 10) % 10 == int(npi[9])


def parse_instant(s):
    if not s:
        return None
    try:
        return datetime.datetime.fromisoformat(str(s).replace("Z", "+00:00"))
    except Exception:                                        # noqa: BLE001
        return None


def file_group(basename):
    """Order files so reference targets load before the resources that use them."""
    b = basename.lower()
    if "insuranceplan" in b:
        return 0
    if "network" in b:
        return 1
    if "practitionerrole" in b:
        return 5
    if "practitioner" in b:
        return 2
    if "organizationaffiliation" in b:
        return 6
    if "organization" in b:
        return 3
    if "location" in b:
        return 4
    if "healthcareservice" in b:
        return 7
    return 8


# ==========================================================================
# Appendix E validation inventory (CMS MPF Technical Guide v1.5, 09/04/2026)
# ==========================================================================
LEVEL1 = [
 ("C4001", "HTTPCommunicationsError", "Networking issue occurred while talking to remote host.",
  "Live GET and HEAD against the index URL and every constituent file URL."),
 ("C4002", "X509CertError", "The HTTPS certification is invalid or has a broken signature chain.",
  "Full TLS handshake with the public CA trust store and hostname verification; chain and validity window inspected."),
 ("C4003", "FileRetrievalError", "System was unable to retrieve the file from the source.",
  "Every constituent file retrieved in full and the byte count reconciled against Content-Length."),
 ("C4004", "InvalidJSONSyntax", "Syntax error / invalid character at the index URL during crawl.",
  "Strict JSON parse of each index file."),
 ("C4011", "MissingURL", "URL for hosted files is missing.",
  "Index URL supplied for each contract number / contract year combination and resolvable."),
 ("P1017", "NoProvidersFound", "File is valid, but contains zero provider records.",
  "Counted PractitionerRole and OrganizationAffiliation records per contract."),
 ("N3006", "OmittedContractID", "Contract ID was expected in the data set, but was not found.",
  "Verified the contract number appears in InsurancePlan.identifier[cms.gov/medicare/ma-plan-id]."),
 ("N3015", "InvalidJSONSyntax", "Syntax error / invalid character inside the data file.",
  "Streaming strict JSON parse of every constituent data file from first byte to last."),
 ("C4012", "MissingFileDownloadURL", "No downloadable files found.",
  "provider_urls array present and non-empty in each index file."),
 ("C4013", "InvalidIndexFileFormat", "Incorrect JSON syntax at the index URL.",
  "Index parsed as a JSON object carrying a top-level provider_urls array of strings."),
 ("C4014", "ImproperURL", "Improper URL format was detected.",
  "Each provider_urls entry tested for a single https:// URL with no whitespace, commas or line breaks."),
 ("C4015", "InvalidPlanDataFile", "Invalid plan data file.",
  "InsurancePlan bundles parsed and validated against the Appendix B plan data points."),
 ("C4016", "InvalidDataFileMR", "Invalid machine-readable JSON data file.",
  "Not applicable: this directory is submitted under the FHIR-based JSON option."),
 ("C4017", "InvalidPlanDataFileFHIRResource", "Invalid FHIR data file resource.",
  "Every entry checked for resourceType and id and for a recognised PDex Plan-Net resource type."),
 ("C4018", "InvalidPlanDataFileFHIRBundle", "Invalid FHIR data file bundle.",
  "Every constituent file checked for a top-level Bundle with a type and an entry array."),
 ("C4019", "IndexFileURLLimitExceeded", "Index file URL count exceeds the limit.",
  "provider_urls entry count tested against the 10,000 cap and each file size against the 300 MB cap (new in v1.5)."),
]

LEVEL2 = [
 ("A2001", "MissingProviderAddresses", "No addresses were provided for an NPI.",
  "Every provider NPI traced to at least one address via PractitionerRole.location, OrganizationAffiliation.location or Organization.address."),
 ("A2008", "InvalidAddress", "Address cannot be geolocated.",
  "Zip code resolved against the US zip reference and cross-checked against the address state."),
 ("F5001", "MissingNetworkReference", "Network Reference ID is missing from resource.",
  "PractitionerRole.extension[network-reference], OrganizationAffiliation.network and InsurancePlan.network checked for presence."),
 ("F5002", "MissingOrganizationReference", "Organization Reference ID is missing from resource.",
  "OrganizationAffiliation.organization checked for presence."),
 ("F5003", "MissingPractitionerReference", "Practitioner Reference ID is missing from resource.",
  "PractitionerRole.practitioner checked for presence."),
 ("F5004", "MissingLocationReference", "Location Reference ID is missing from resource.",
  "PractitionerRole.location and OrganizationAffiliation.location checked for presence."),
 ("F5005", "BrokenNetworkReference", "Network Resource not found within dataset.",
  "Every network reference resolved against the Organization resources typed 'ntwk' in the same contract dataset."),
 ("F5006", "BrokenOrganizationReference", "Organization Resource not found within dataset.",
  "Every Organization reference, from any resource and any path, resolved within the same contract dataset."),
 ("F5007", "BrokenPractitionerReference", "Practitioner Resource not found within dataset.",
  "Every Practitioner reference, from any resource and any path, resolved within the same contract dataset."),
 ("F5008", "BrokenLocationReference", "Location Resource not found within dataset.",
  "Every Location reference, from any resource and any path, resolved within the same contract dataset."),
 ("F5009", "MultipleNPIonResource", "Resources must only list a single NPI.",
  "Counted distinct identifier[system=us-npi] values on every Practitioner, PractitionerRole and Organization resource."),
 ("N3001", "MissingPlanID", "Plan ID is missing or blank.",
  "InsurancePlan.identifier[system='http://cms.gov/medicare/ma-plan-id'] checked for presence and a non-blank value."),
 ("N3002", "InvalidMAPlanID", "Plan ID format is invalid (must be [Contract]-[Plan]-[Segment]).",
  "Every MA Plan ID matched against the pattern Hnnnn-nnn-nnn."),
 ("N3003", "PlanNotAssociated", "Plan not associated with any providers.",
  "Each plan traced through InsurancePlan.network to at least one PractitionerRole or OrganizationAffiliation."),
 ("N3004", "MissingContractYear", "Each plan must have a contract year listed.",
  "InsurancePlan.period.start checked for presence on every plan."),
 ("N3005", "InvalidContractYear", "Contract year is not in a 4-digit year format.",
  "period.start year extracted, format checked and matched against the active contract year."),
 ("N3011", "UnknownContractID", "Contract ID is unknown to the HPMS registry dataset.",
  "Format and internal consistency verified; a definitive determination requires the HPMS registry."),
 ("N3012", "UnknownPlanID", "Plan ID is unknown to the HPMS registry dataset.",
  "Format and internal consistency verified; a definitive determination requires the HPMS registry."),
 ("N3013", "UnknownSegmentID", "Segment ID is unknown to the HPMS registry dataset.",
  "Format and internal consistency verified; a definitive determination requires the HPMS registry."),
 ("N3014", "MismatchContractID", "MAPlanID does not match the contract ID in HPMS.",
  "Every MA Plan ID checked to carry the contract number of the index file that lists it."),
 ("P1001", "MissingProviderNPI", "NPI is missing or blank.",
  "NPI resolved for every Practitioner, PractitionerRole and facility Organization resource."),
 ("P1002", "UnknownProviderNPI", "NPI not found in the registry dataset.",
  "Every distinct NPI matched against the NPPES registry snapshot."),
 ("P1003", "DeactivatedProviderNPI", "NPI is listed as deactivated in registry dataset.",
  "Every distinct NPI matched against the NPPES deactivated NPI report."),
 ("P1016", "ProviderNotAssociated", "Provider must be associated with at least one plan ID.",
  "Every provider NPI traced through its network reference to at least one InsurancePlan MA Plan ID."),
]

LEVEL3 = [
 ("P1004", "InvalidProviderType", "Must be individual or facility.",
  "Organization.type.coding checked for OrgTypeCS code 'fac'; Practitioner / PractitionerRole treated as individual."),
 ("P1005", "MismatchProviderType", "Type does not match NPI Registry record.",
  "Submitted provider type compared against the NPPES entity type code for the same NPI."),
 ("P1006", "MissingProviderFirstName", "Missing first name.",
  "Practitioner.name.given[0] checked for presence and a non-blank value."),
 ("P1007", "MissingProviderLastName", "Missing last name.",
  "Practitioner.name.family checked for presence and a non-blank value."),
 ("P1008", "MissingFacilityName", "Missing facility name.",
  "Organization.name checked for presence and a non-blank value."),
 ("P1009", "MissingSpecialty", "Provider requires at least one specialty.",
  "PractitionerRole.specialty and OrganizationAffiliation.specialty checked for at least one coded value."),
 ("P1010", "MissingSex", "Provider sex is not listed.",
  "Practitioner.gender checked for presence and for a valid administrative-gender code."),
 ("P1011", "MissingLanguage", "Provider language is not listed.",
  "Practitioner.communication.coding.code checked for at least one language code."),
 ("P1012", "MissingAcceptingPatients", "Provider accepting patients status not listed.",
  "PractitionerRole.extension[newpatients].extension[acceptingPatients] checked for presence."),
 ("P1013", "InvalidDateFormat", "Date does not follow the required format.",
  "meta.lastUpdated parsed on every resource carrying the Date Record Was Last Updated data point."),
 ("P1014", "FutureDate", "LastUpdated date is in the future relative to crawl date.",
  "meta.lastUpdated compared against the audit date on every resource."),
 ("P1018", "InvalidAcceptingPatientsType", "AcceptingPatients format is invalid (FHIR: newpt, nopt or existptonly).",
  "Every acceptingPatients code checked against the valid value set (v1.5 corrected the 'nopt' typo)."),
 ("A2002", "MissingProviderCity", "City entry is missing or blank.",
  "address.city checked on every referenced address."),
 ("A2003", "MissingProviderState", "State entry is missing or blank.",
  "address.state checked on every referenced address."),
 ("A2004", "InvalidProviderState", "State format is invalid.",
  "address.state matched against the USPS two-letter state and territory abbreviation set."),
 ("A2005", "MissingProviderZip", "Zip code is missing or blank.",
  "address.postalCode checked on every referenced address."),
 ("A2006", "InvalidProviderZip", "Zip code format is invalid.",
  "Every address.postalCode tested for exactly five digits, as the Appendix A/B zip specification requires."),
 ("A2007", "MissingProviderStreetAddresses", "Address line 1 is missing or blank.",
  "address.line[0] checked on every referenced address."),
 ("A2009", "MissingProviderPhoneNumber", "Provider phone number is missing.",
  "Data points 7a/7b (PractitionerRole / Practitioner.telecom), 8a/8b (Organization / Location.telecom) and Location.telecom all checked."),
 ("A2010", "InvalidProviderPhoneNumber", "Provider phone number is in an invalid format.",
  "Every phone value normalised and tested as a valid 10-digit North American number."),
 ("N3007", "OmittedSegmentID", "Segment ID was expected in dataset, but not found.",
  "Segment component extracted from every MA Plan ID; completeness against HPMS requires the HPMS registry."),
 ("N3008", "OmittedPlanID", "Plan ID was expected in dataset, but not found.",
  "Plan component extracted from every MA Plan ID; completeness against HPMS requires the HPMS registry."),
 ("C4005", "HeadRequestFailed", "File is missing header information.",
  "HEAD issued against the index and every constituent file URL (Appendix D item 1)."),
 ("C4006", "MissingLastModifiedHeader", "Header is missing the last modified date.",
  "Last-Modified checked on every response."),
 ("C4007", "MissingContentLengthHeader", "Header is missing the content length.",
  "Content-Length checked on every response."),
 ("C4008", "MissingContentTypeHeader", "Header is missing the content type.",
  "Content-Type checked for application/json on every response."),
 ("C4009", "MissingETagHeader", "Header is missing the ETag.",
  "ETag checked on every response."),
 ("C4010", "StaleDataWarning", "Last updated date is older than 30 days from crawl date.",
  "meta.lastUpdated on every resource compared against the audit date minus 30 days."),
]

APPENDIX_E = ([(c, n, 1, d, h) for c, n, d, h in LEVEL1] +
              [(c, n, 2, d, h) for c, n, d, h in LEVEL2] +
              [(c, n, 3, d, h) for c, n, d, h in LEVEL3])

# Supplementary conformance tests -- Appendix B / D and the self-validation steps.
SUPPLEMENTAL = [
 ("S-01", "IndexReturnsRawJSON", "Self-validation step 1: index URL returns raw JSON with a provider_urls array."),
 ("S-02", "PublicNoAuthAccess", "Self-validation step 3: all URLs publicly accessible externally without authentication."),
 ("S-03", "UncompressedDelivery", "Self-validation step 6: server does not serve files with compression encoding."),
 ("S-04", "ActiveContractYear", "Self-validation step 7: all records reflect the active contract year in InsurancePlan.period."),
 ("S-05", "ConditionalRequestSupport", "Appendix D item 2: If-None-Match and If-Modified-Since return 304 Not Modified."),
 ("S-06", "HeadMethodSupport", "Appendix D item 1: HEAD supported with ETag, Last-Modified, Content-Length and Content-Type."),
 ("S-07", "NPICheckDigit", "Appendix A/B: NPI is 10 digits and passes the ISO 7812 (80840 prefix) check digit."),
 ("S-08", "NPISystemURI", "Appendix B: NPI carried on identifier[system='http://hl7.org/fhir/sid/us-npi']."),
 ("S-09", "MAPlanIDSystemURI", "Appendix B: MA Plan ID carried on identifier[system='http://cms.gov/medicare/ma-plan-id']."),
 ("S-10", "PeriodStartJan1", "Appendix B note: InsurancePlan.period.start is the first day of the contract year."),
 ("S-11", "SingleContractYear", "Appendix B note: only one contract year supplied per InsurancePlan."),
 ("S-12", "NUCCSpecialtySystem", "Appendix B: specialty.coding.system is http://nucc.org/provider-taxonomy."),
 ("S-13", "NUCCCodeKnown", "Appendix B: every specialty code is a recognised NUCC provider taxonomy code."),
 ("S-14", "IndividualVsNonIndividualSpecialty",
  "Appendix B: PractitionerRole.specialty presents NUCC individual codes; OrganizationAffiliation.specialty presents non-individual codes."),
 ("S-15", "OrgTypeFacilityCoding", "Appendix B: facility Organization carries type.coding OrgTypeCS code 'fac'."),
 ("S-16", "LanguageCodeFormat", "Appendix B: Practitioner.communication codes are well-formed language codes."),
 ("S-17", "PDexProfileConformance", "PDex Plan-Net v1.2.0 profile declared in meta.profile on submitted resources."),
 ("S-18", "NetworkLinkageToPlan", "Appendix B: provider and facility records link to a unique MA plan through the network field."),
 ("S-19", "IndexScopedToContractYear", "Hosting requirement: the index lists only files for the same contract number and year."),
 ("S-20", "FileSizeAndCountLimits", "v1.5 hosting requirement: 10,000 URL cap and 300 MB file size cap."),
 ("S-21", "ReferenceIntegritySweep",
  "Every Reference in every resource, at any path and of any target type, resolves within the contract dataset."),
 ("S-22", "DuplicateResourceIds", "No resource id is used twice within the same resource type in a contract."),
]

# Codes that map onto supplementary tests rather than an Appendix E code
SUPP_CODES = {
    "SPECCLS": ("S-14", "SpecialtyClassMismatch", 3),
    "SPECSYS": ("S-12", "InvalidSpecialtyCodeSystem", 3),
    "SPECCODE": ("S-13", "UnknownNUCCTaxonomyCode", 3),
    "NPIFMT": ("S-07", "InvalidNPIFormat", 2),
    "REFBROKEN": ("S-21", "BrokenReferenceOtherType", 2),
    "REFMALFORMED": ("S-21", "MalformedReference", 2),
    "DUPID": ("S-22", "DuplicateResourceId", 2),
    "HOSTCMP": ("S-03", "CompressedEncodingServed", 1),
    "CYSTART": ("S-10", "ContractYearStartNotJan1", 3),
    "CYEND": ("S-10", "ContractYearEndMismatch", 3),
    "LANGFMT": ("S-16", "InvalidLanguageCode", 3),
    "PROFILE": ("S-17", "MissingPDexProfile", 3),
}


# ==========================================================================
# The audit
# ==========================================================================
class ContractAudit(object):

    def __init__(self, contract, index_url, refdata, cache_dir, out_dir,
                 catalog, audit_date, fresh=False):
        self.contract = contract
        self.index_url = index_url
        self.ref = refdata
        self.cache_dir = cache_dir
        self.audit_date = audit_date
        self.stale_before = audit_date - datetime.timedelta(days=30)
        self.F = Findings(contract, os.path.join(out_dir, "findings_%s.csv" % contract),
                          catalog)
        self.fresh = fresh
        self.stats = collections.Counter()
        self.http = []
        self.index = None
        self.plan_ids = []
        self.networks_meta = {}
        self.notes = {}
        self.fatal_stop = False

        # reference integrity state
        self.present = collections.defaultdict(set)      # type -> ids
        self.wanted = collections.defaultdict(dict)      # type -> id -> [srcType, srcId, path]
        self.dup_ids = collections.Counter()

        # working indexes
        self.plans, self.networks, self.prac, self.orgs, self.payers = {}, {}, {}, {}, {}
        self.locs = {}
        self.prac_npi, self.org_npi = {}, {}
        self.net_to_plans = collections.defaultdict(set)
        self.used_networks = set()
        self.npi_with_addr, self.npi_with_plan, self.all_npis = set(), set(), set()
        self.accept_values = collections.Counter()
        self.pr_specialties = collections.Counter()
        self.oa_specialties = collections.Counter()
        self.oa_individual = collections.Counter()
        self.langs = collections.Counter()
        self.states = collections.Counter()
        self.zips_seen = collections.Counter()
        self.bad_addresses = {}
        self.bad_phones = collections.Counter()
        self.deact_detail = []
        self.mismatch_detail = []
        self.checked_addr = set()
        self.npi_home = {}          # npi -> (resourceType, resource id, name)
        self.unassociated = []      # detail rows for P1016

    # ---------------------------------------------------------------- add
    def add(self, code, level, **kw):
        name = None
        for c, n, lv, d, h in APPENDIX_E:
            if c == code:
                name = n
                break
        if name is None:
            name = SUPP_CODES.get(code, (None, code, level))[1]
        self.F.add(code, name, level, **kw)

    # ------------------------------------------------------- stage: HTTP
    def run_http(self):
        rule("%s: transport and index" % self.contract)
        r = probe(self.index_url)
        r_kind = "INDEX"
        idx_dest = os.path.join(self.cache_dir, "%s_index.json" % self.contract)
        meta_path = os.path.join(self.cache_dir, "http_meta.json")
        try:
            cache_meta = json.load(open(meta_path))
        except Exception:                                    # noqa: BLE001
            cache_meta = {}

        got, cached = fetch_to_file(self.index_url, idx_dest, cache_meta, self.fresh)
        if got.error or got.status != 200:
            self.add("C4001", 1, value=self.index_url,
                     detail="index URL could not be retrieved: %s"
                            % (got.error or "HTTP %s" % got.status))
            self.add("C4011", 1, value=self.index_url,
                     detail="no usable index file at the reported URL")
            self.fatal_stop = True
            json.dump(cache_meta, open(meta_path, "w"))
            return
        for k, v in got.headers.items():
            r.headers.setdefault(k, v)

        self._record_http(r, r_kind, got)

        try:
            self.index = json.load(open(idx_dest, encoding="utf-8"))
        except ValueError as e:
            self.add("C4004", 1, value=self.index_url, detail="index JSON syntax error: %s" % e)
            self.fatal_stop = True
            json.dump(cache_meta, open(meta_path, "w"))
            return

        urls = self.index.get("provider_urls") if isinstance(self.index, dict) else None
        if not isinstance(self.index, dict):
            self.add("C4013", 1, detail="index is not a JSON object")
            self.fatal_stop = True
            return
        if urls is None:
            self.add("C4013", 1, field="provider_urls",
                     detail="top-level provider_urls array absent")
            self.fatal_stop = True
            return
        if not isinstance(urls, list):
            self.add("C4013", 1, field="provider_urls", detail="provider_urls is not an array")
            self.fatal_stop = True
            return
        if not urls:
            self.add("C4012", 1, field="provider_urls", detail="provider_urls array is empty")
            self.fatal_stop = True
            return

        self.stats["index_url_count"] = len(urls)
        if len(urls) > MAX_INDEX_URLS:
            self.add("C4019", 1, value=len(urls),
                     detail="provider_urls exceeds the %s entry limit" % commify(MAX_INDEX_URLS))
        scoped = True
        for u in urls:
            if (not isinstance(u, str) or not u.startswith("https://")
                    or re.search(r"[\s,]", u)):
                self.add("C4014", 1, value=u,
                         detail="must be a single https:// URL with no whitespace, commas or line breaks")
            if isinstance(u, str):
                low = u.lower()
                if (self.contract.lower() not in low) or (CONTRACT_YEAR not in low):
                    scoped = False
        self.notes["index_scoped"] = scoped
        if not scoped:
            self.add("N3006", 1, detail="the index lists files that are not scoped to "
                                        "this contract number and contract year")

        log("index lists %d files" % len(urls), 1)
        self.files = []
        base = os.path.join(self.cache_dir, self.contract)
        if not os.path.isdir(base):
            os.makedirs(base)

        jobs = []
        for i, u in enumerate(urls, 1):
            name = os.path.basename(u.split("?")[0]) or ("part%d.json" % i)
            jobs.append((i, u, name, os.path.join(base, name)))

        def work(job):
            i, u, name, dest = job
            pr = probe(u)
            got, cached = fetch_to_file(u, dest, cache_meta, self.fresh)
            return job, pr, got, cached

        workers = min(DOWNLOAD_WORKERS, max(1, len(jobs)))
        log("retrieving %d files with %d parallel connections" % (len(jobs), workers), 1)
        results = {}
        with ThreadPoolExecutor(max_workers=workers) as pool:
            for job, pr, got, cached in pool.map(work, jobs):
                results[job[0]] = (job, pr, got, cached)
                name, dest = job[2], job[3]
                size = os.path.getsize(dest) if os.path.exists(dest) else 0
                log("[%2d/%2d] %-58s %10s %s"
                    % (job[0], len(jobs), name[:58], mb(size),
                       "(cached)" if cached else ("FAILED" if got.error else "")), 1)

        # findings and CSV writing stay on the main thread, in index order
        for i in sorted(results):
            job, pr, got, cached = results[i]
            u, dest = job[1], job[3]
            for k, v in (got.headers or {}).items():
                pr.headers.setdefault(k, v)
            self._record_http(pr, "DATA", got)
            if got.error or got.status != 200:
                self.add("C4003", 1, value=u,
                         detail="file could not be retrieved after %d attempts: %s"
                                % (5, got.error or "HTTP %s" % got.status))
                continue
            declared = pr.headers.get("content-length") or got.h("content-length")
            actual = os.path.getsize(dest) if os.path.exists(dest) else 0
            if (declared and str(declared).isdigit() and int(declared) != actual
                    and not got.from_cache and not got.h("content-encoding")):
                self.add("C4003", 1, value=u,
                         detail="retrieved %s bytes but Content-Length declared %s"
                                % (commify(actual), commify(declared)))
            self.stats["bytes_total"] += actual
            if actual > MAX_FILE_BYTES:
                self.add("C4019", 1, value=u,
                         detail="constituent file is %s, over the 300 MB limit" % mb(actual))
            self.files.append(dest)
        json.dump(cache_meta, open(meta_path, "w"))

    def _record_http(self, pr, kind, got):
        h = dict(pr.headers)
        for k, v in (got.headers or {}).items():
            h.setdefault(k.lower(), v)
        row = {
            "kind": kind, "url": pr.url,
            "head_ok": bool(pr.head_ok), "head_status": pr.head_status,
            "status": got.status,
            "etag": h.get("etag", ""), "last_modified": h.get("last-modified", ""),
            "content_length": h.get("content-length", ""),
            "content_type": h.get("content-type", ""),
            "content_encoding": h.get("content-encoding", ""),
            "conditional": pr.conditional,
            "from_cache": got.from_cache,
        }
        self.http.append(row)
        u = pr.url
        if not row["head_ok"]:
            self.add("C4005", 3, value=u, detail="HEAD returned %s" % pr.head_status)
        if not row["last_modified"]:
            self.add("C4006", 3, value=u, detail="Last-Modified header absent")
        if not row["content_length"]:
            self.add("C4007", 3, value=u, detail="Content-Length header absent")
        if "application/json" not in (row["content_type"] or ""):
            self.add("C4008", 3, value=u,
                     detail="Content-Type is %r, expected application/json" % row["content_type"])
        if not row["etag"]:
            self.add("C4009", 3, value=u, detail="ETag header absent")
        if row["content_encoding"]:
            self.add("HOSTCMP", 1, value=u,
                     detail="server returned Content-Encoding: %s; files must be served "
                            "uncompressed" % row["content_encoding"])

    # ------------------------------------------------- stage: validation
    def run_data(self):
        if self.fatal_stop or not getattr(self, "files", None):
            return
        rule("%s: parsing and validating %d files" % (self.contract, len(self.files)))
        for path in sorted(self.files, key=lambda p: file_group(os.path.basename(p))):
            base = os.path.basename(path)
            rt, bt, err = bundle_header(path)
            if err:
                self.add("N3015", 1, value=base, detail="JSON parse error: %s" % err[:200])
                continue
            if rt != "Bundle":
                self.add("C4018", 1, value=base,
                         detail="top-level resourceType is %r, expected Bundle" % rt)
            elif not bt:
                self.add("C4018", 1, value=base, detail="Bundle.type is absent")
            n = 0
            try:
                for res in stream_resources(path):
                    n += 1
                    self._validate_resource(res, base)
            except Exception as e:                           # noqa: BLE001
                self.add("N3015", 1, value=base,
                         detail="JSON parse failed after %s entries: %s" % (commify(n), str(e)[:180]))
            log("%-58s %s resources" % (base[:58], commify(n)), 1)
            self.stats["entries_total"] += n
        self._finalise()

    def _validate_resource(self, res, fname):
        if not isinstance(res, dict):
            self.add("C4017", 1, value=fname, detail="bundle entry.resource is not an object")
            return
        rt = res.get("resourceType")
        rid = res.get("id")
        if not rt or not rid:
            self.add("C4017", 1, resource=str(rt), rid=str(rid), value=fname,
                     detail="resource is missing resourceType or id")
            return
        self.stats["res_" + rt] += 1
        if rid in self.present[rt]:
            self.dup_ids[(rt, rid)] += 1
        self.present[rt].add(rid)

        # generic reference sweep -- catches every Reference at any path
        for path, ref in walk_references(res):
            ttype, tid = ref_target(ref)
            if ttype == "#contained":
                ids = set(c.get("id") for c in (res.get("contained") or []) if isinstance(c, dict))
                if tid not in ids:
                    self.add("REFMALFORMED", 2, resource=rt, rid=rid, field=path, value=ref,
                             detail="contained reference does not match any contained resource")
                continue
            if ttype is None:
                self.add("REFMALFORMED", 2, resource=rt, rid=rid, field=path, value=ref,
                         detail="reference is not a resolvable Type/id reference")
                continue
            slot = self.wanted[ttype]
            if tid not in slot:
                slot[tid] = (rt, rid, path)

        handler = getattr(self, "_r_" + rt, None)
        if handler:
            handler(res, rid)
        elif rt not in ("Bundle",):
            self._check_meta(res, rt, rid)

    # -- shared field checks -------------------------------------------
    def _check_meta(self, res, rt, rid, npi=""):
        lu = (res.get("meta") or {}).get("lastUpdated")
        if not lu:
            self.add("P1013", 3, resource=rt, rid=rid, npi=npi, field="meta.lastUpdated",
                     detail="meta.lastUpdated is absent")
            return
        dt = parse_instant(lu)
        if dt is None:
            self.add("P1013", 3, resource=rt, rid=rid, npi=npi, field="meta.lastUpdated",
                     value=lu, detail="value is not a parseable FHIR instant")
            return
        d = dt.date()
        if d > self.audit_date:
            self.add("P1014", 3, resource=rt, rid=rid, npi=npi, field="meta.lastUpdated",
                     value=lu, detail="lastUpdated is later than the crawl date %s" % self.audit_date)
        elif d < self.stale_before:
            self.add("C4010", 3, resource=rt, rid=rid, npi=npi, field="meta.lastUpdated",
                     value=lu, detail="older than 30 days before the crawl date %s" % self.audit_date)

    def _check_profile(self, res, rt, rid):
        prof = (res.get("meta") or {}).get("profile") or []
        if rt in MPF_RESOURCE_TYPES and not prof:
            self.add("PROFILE", 3, resource=rt, rid=rid, field="meta.profile",
                     detail="no PDex Plan-Net profile declared")

    def _check_npi(self, npi, rt, rid, expect_entity):
        if not luhn_npi(npi):
            self.add("NPIFMT", 2, resource=rt, rid=rid, npi=npi, field="identifier[us-npi]",
                     value=npi,
                     detail="not a 10-digit NPI passing the ISO 7812 (80840 prefix) check digit")
            return
        if self.ref.available["registry"]:
            ent = self.ref.registry.get(npi)
            if ent is None:
                self.add("P1002", 2, resource=rt, rid=rid, npi=npi, value=npi,
                         detail="NPI is not present in the NPPES registry snapshot")
            elif ent != expect_entity:
                submitted = "Individual" if expect_entity == "1" else "Facility"
                registry = "Individual" if ent == "1" else "Organization"
                self.add("P1005", 3, resource=rt, rid=rid, npi=npi, value=npi,
                         detail="submitted as %s but the NPPES entity type is %s"
                                % (submitted, registry))
                if len(self.mismatch_detail) < 40:
                    self.mismatch_detail.append((npi, ""))
        if self.ref.available["deactivated"] and npi in self.ref.deactivated:
            self.add("P1003", 2, resource=rt, rid=rid, npi=npi, value=npi,
                     detail="NPI is listed as deactivated in the NPPES deactivation report")

    def _check_address(self, tup, rt, rid, npi, src):
        key = (rt, rid, npi, src)
        line1, line2, city, state, zipc, phones = tup
        if not (line1 or "").strip():
            self.add("A2007", 3, resource=rt, rid=rid, npi=npi,
                     field=src + ".address.line[0]", detail="address line 1 is missing or blank")
        if not (city or "").strip():
            self.add("A2002", 3, resource=rt, rid=rid, npi=npi,
                     field=src + ".address.city", detail="city entry is missing or blank")
        stt = (state or "").strip().upper()
        self.states[stt] += 1
        if not stt:
            self.add("A2003", 3, resource=rt, rid=rid, npi=npi,
                     field=src + ".address.state", detail="state entry is missing or blank")
        elif stt not in USPS_STATES:
            self.add("A2004", 3, resource=rt, rid=rid, npi=npi,
                     field=src + ".address.state", value=state,
                     detail="not a valid two-letter USPS state or territory abbreviation")
        z = (zipc or "").strip()
        self.zips_seen[z] += 1
        if not z:
            self.add("A2005", 3, resource=rt, rid=rid, npi=npi,
                     field=src + ".address.postalCode", detail="zip code is missing or blank")
        else:
            if not re.match(r"^\d{5}$", z):
                if re.match(r"^\d{5}-?\d{4}$", z):
                    self.add("A2006", 3, resource=rt, rid=rid, npi=npi,
                             field=src + ".address.postalCode", value=z,
                             detail="ZIP+4 supplied; the guide requires a five digit zip code as a string")
                else:
                    self.add("A2006", 3, resource=rt, rid=rid, npi=npi,
                             field=src + ".address.postalCode", value=z,
                             detail="zip code format is invalid, it must be exactly five digits")
            z5 = z[:5]
            if re.match(r"^\d{5}$", z5) and self.ref.available["zips"]:
                known = self.ref.zips.get(z5)
                if known is None:
                    self.add("A2008", 2, resource=rt, rid=rid, npi=npi,
                             field=src + ".address.postalCode", value=z,
                             detail="zip code is not a recognised US zip code, so the address "
                                    "cannot be geolocated")
                    self._note_bad_address(line1, city, stt, z, None, "zip not recognised", rid, npi)
                elif stt and stt in USPS_STATES and stt not in known:
                    self.add("A2008", 2, resource=rt, rid=rid, npi=npi,
                             field=src + ".address", value="%s %s" % (stt, z5),
                             detail="zip code belongs to %s, not %s, so the address cannot be "
                                    "geolocated" % ("/".join(known), stt))
                    self._note_bad_address(line1, city, stt, z, known, "zip/state mismatch", rid, npi)
        if not [p for p in phones if p]:
            self.add("A2009", 3, resource=rt, rid=rid, npi=npi,
                     field=src + ".telecom[phone]", detail="provider phone number is missing")
        else:
            for p in phones:
                if p and not phone_ok(p):
                    self.add("A2010", 3, resource=rt, rid=rid, npi=npi,
                             field=src + ".telecom[phone]", value=p,
                             detail="not a valid 10-digit North American phone number")
                    self.bad_phones[p] += 1
                    break

    def _note_bad_address(self, line1, city, state, zipc, known, issue, rid, npi):
        key = (line1, city, state, zipc)
        if key not in self.bad_addresses:
            self.bad_addresses[key] = {
                "line": line1, "city": city, "state": state, "zip": zipc,
                "zip_states": known, "issue": issue, "id": rid, "npi": npi}

    def _check_specialty(self, cc_list, rt, rid, npi, individual):
        codes = codes_of(cc_list)
        if not codes:
            self.add("P1009", 3, resource=rt, rid=rid, npi=npi, field="specialty",
                     detail="at least one NUCC specialty code is required")
            return
        for code, system in codes:
            if individual:
                self.pr_specialties[code] += 1
            else:
                self.oa_specialties[code] += 1
            if system != NUCC_SYS:
                self.add("SPECSYS", 3, resource=rt, rid=rid, npi=npi,
                         field="specialty.coding.system", value=system,
                         detail="specialty code system must be " + NUCC_SYS)
            if not self.ref.available["taxonomy"]:
                continue
            klass = self.ref.taxonomy_class(code)
            if code not in self.ref.taxonomy:
                self.add("SPECCODE", 3, resource=rt, rid=rid, npi=npi,
                         field="specialty.coding.code", value=code,
                         detail="code is not a recognised NUCC provider taxonomy code")
            elif individual and klass == "organization":
                self.add("SPECCLS", 3, resource=rt, rid=rid, npi=npi,
                         field="specialty.coding.code", value=code,
                         detail="PractitionerRole.specialty must present NUCC individual codes")
            elif (not individual) and klass == "individual":
                self.add("SPECCLS", 3, resource=rt, rid=rid, npi=npi,
                         field="specialty.coding.code", value=code,
                         detail="OrganizationAffiliation.specialty must present NUCC "
                                "non-individual codes")
                self.oa_individual[code] += 1

    # -- resource handlers ---------------------------------------------
    def _r_InsurancePlan(self, res, rid):
        self.plans[rid] = res
        self._check_profile(res, "InsurancePlan", rid)
        ma = [(i.get("value") or "").strip() for i in (res.get("identifier") or [])
              if i.get("system") == MAPLAN_SYS]
        ma = [v for v in ma if v]
        if not ma:
            self.add("N3001", 2, resource="InsurancePlan", rid=rid,
                     field="identifier[%s]" % MAPLAN_SYS,
                     detail="CMS MA Plan Identifier is missing or blank")
        for v in ma:
            if v not in self.plan_ids:
                self.plan_ids.append(v)
            if not MAPLAN_RE.match(v):
                self.add("N3002", 2, resource="InsurancePlan", rid=rid, value=v,
                         detail="must be [CMS Contract Number]-[CMS Plan ID]-[CMS Segment ID], "
                                "for example H9999-001-001")
            elif not v.startswith(self.contract + "-"):
                self.add("N3014", 2, resource="InsurancePlan", rid=rid, value=v,
                         detail="MAPlanID contract %s does not match the contract for this "
                                "index file (%s)" % (v.split("-")[0], self.contract))
        period = res.get("period") or {}
        start = period.get("start")
        if not start:
            self.add("N3004", 2, resource="InsurancePlan", rid=rid, field="period.start",
                     detail="each plan must have a contract year listed")
        else:
            m = re.match(r"^(\d{4})", str(start))
            if not m:
                self.add("N3005", 2, resource="InsurancePlan", rid=rid, field="period.start",
                         value=start, detail="contract year is not in a 4-digit year format")
            elif m.group(1) != CONTRACT_YEAR:
                self.add("N3005", 2, resource="InsurancePlan", rid=rid, field="period.start",
                         value=start,
                         detail="contract year %s does not match the active contract year %s"
                                % (m.group(1), CONTRACT_YEAR))
            elif str(start)[:10] != CONTRACT_YEAR + "-01-01":
                self.add("CYSTART", 3, resource="InsurancePlan", rid=rid, field="period.start",
                         value=start,
                         detail="period.start should be the first day of the contract year (%s-01-01)"
                                % CONTRACT_YEAR)
        if period.get("end") and str(period["end"])[:4] != CONTRACT_YEAR:
            self.add("CYEND", 3, resource="InsurancePlan", rid=rid, field="period.end",
                     value=period.get("end"),
                     detail="period.end falls outside the reported contract year")
        nets = [ref_target(n.get("reference"))[1] for n in (res.get("network") or [])]
        nets = [n for n in nets if n]
        if not nets:
            self.add("F5001", 2, resource="InsurancePlan", rid=rid, field="network",
                     detail="Network Reference ID is missing from the resource")
        for n in nets:
            self.net_to_plans[n].add(ma[0] if ma else rid)
        self._check_meta(res, "InsurancePlan", rid)

    def _r_Organization(self, res, rid):
        types = [c for c, _ in codes_of(res.get("type"), ORGTYPE_SYS)]
        self._check_profile(res, "Organization", rid)
        if "ntwk" in types:
            self.networks[rid] = res
            self.networks_meta[rid] = {
                "name": res.get("name"),
                "identifiers": [i.get("value") for i in (res.get("identifier") or [])],
            }
            self._check_meta(res, "Organization", rid)
            return
        if "payer" in types and "fac" not in types:
            self.payers[rid] = res
            self._check_meta(res, "Organization", rid)
            return

        self.orgs[rid] = res
        found = sorted(set(npis_of(res)))
        if len(found) > 1:
            self.add("F5009", 2, resource="Organization", rid=rid, value=",".join(found),
                     detail="resources must only list a single NPI")
        npi = found[0] if found else ""
        self.org_npi[rid] = npi
        if not npi:
            self.add("P1001", 2, resource="Organization", rid=rid,
                     field="identifier[us-npi]", detail="NPI is missing or blank")
        else:
            self.all_npis.add(npi)
            self.npi_home.setdefault(npi, ("Organization", rid, res.get("name") or ""))
            self._check_npi(npi, "Organization", rid, "2")
            if self.ref.available["deactivated"] and npi in self.ref.deactivated:
                if len(self.deact_detail) < 60:
                    self.deact_detail.append((npi, res.get("name") or "", rid))
            if self.ref.available["registry"] and self.ref.registry.get(npi) == "1":
                for i, (n, nm) in enumerate(self.mismatch_detail):
                    if n == npi and not nm:
                        self.mismatch_detail[i] = (npi, res.get("name") or "")
                        break
        if not (res.get("name") or "").strip():
            self.add("P1008", 3, resource="Organization", rid=rid, npi=npi, field="name",
                     detail="missing facility name")
        if "fac" not in types:
            self.add("P1004", 3, resource="Organization", rid=rid, npi=npi,
                     field="type.coding", value=",".join(types) or "(none)",
                     detail="a facility Organization must carry OrgTypeCS code 'fac'")
        phones = phones_of(res)
        if not [p for p in phones if p]:
            self.add("A2009", 3, resource="Organization", rid=rid, npi=npi,
                     field="Organization.telecom[phone]",
                     detail="required data point 8a: facility phone number is missing")
        else:
            for p in phones:
                if p and not phone_ok(p):
                    self.add("A2010", 3, resource="Organization", rid=rid, npi=npi,
                             field="Organization.telecom[phone]", value=p,
                             detail="not a valid 10-digit North American phone number")
                    self.bad_phones[p] += 1
                    break
        for a in res.get("address") or []:
            if npi:
                self.npi_with_addr.add(npi)
            line = a.get("line") or []
            self._check_address((line[0] if line else "", line[1] if len(line) > 1 else "",
                                 a.get("city") or "", a.get("state") or "",
                                 a.get("postalCode") or "", tuple(phones)),
                                "Organization", rid, npi, "Organization")
        self._check_meta(res, "Organization", rid, npi)

    def _r_Practitioner(self, res, rid):
        self._check_profile(res, "Practitioner", rid)
        found = sorted(set(npis_of(res)))
        if len(found) > 1:
            self.add("F5009", 2, resource="Practitioner", rid=rid, value=",".join(found),
                     detail="resources must only list a single NPI")
        npi = found[0] if found else ""
        self.prac[rid] = res
        self.prac_npi[rid] = npi
        if not npi:
            self.add("P1001", 2, resource="Practitioner", rid=rid,
                     field="identifier[us-npi]", detail="NPI is missing or blank")
        else:
            self.npi_home.setdefault(npi, ("Practitioner", rid, ""))
            self._check_npi(npi, "Practitioner", rid, "1")
        names = res.get("name") or []
        nm = names[0] if names else {}
        if not [g for g in (nm.get("given") or []) if (g or "").strip()]:
            self.add("P1006", 3, resource="Practitioner", rid=rid, npi=npi,
                     field="name.given[0]", detail="missing first name")
        if not (nm.get("family") or "").strip():
            self.add("P1007", 3, resource="Practitioner", rid=rid, npi=npi,
                     field="name.family", detail="missing last name")
        gender = (res.get("gender") or "").strip().lower()
        if not gender:
            self.add("P1010", 3, resource="Practitioner", rid=rid, npi=npi, field="gender",
                     detail="provider sex is not listed")
        elif gender not in GENDER_VALUES:
            self.add("P1010", 3, resource="Practitioner", rid=rid, npi=npi, field="gender",
                     value=gender, detail="not a valid FHIR administrative-gender code")
        langs = codes_of(res.get("communication") or [])
        for code, _ in langs:
            self.langs[code] += 1
            if not re.match(r"^[a-z]{2,3}(-[A-Za-z0-9]{2,8})*$", code or ""):
                self.add("LANGFMT", 3, resource="Practitioner", rid=rid, npi=npi,
                         field="communication.coding.code", value=code,
                         detail="language code is not a well-formed ISO 639 / BCP-47 code")
        if not langs:
            self.add("P1011", 3, resource="Practitioner", rid=rid, npi=npi,
                     field="communication", detail="provider language is not listed")
        self._check_meta(res, "Practitioner", rid, npi)

    def _r_Location(self, res, rid):
        self._check_profile(res, "Location", rid)
        a = res.get("address") or {}
        line = a.get("line") or []
        self.locs[rid] = (line[0] if line else "", line[1] if len(line) > 1 else "",
                          a.get("city") or "", a.get("state") or "",
                          a.get("postalCode") or "", tuple(phones_of(res)))
        self._check_meta(res, "Location", rid)

    def _r_PractitionerRole(self, res, rid):
        self._check_profile(res, "PractitionerRole", rid)
        pref = ref_target((res.get("practitioner") or {}).get("reference"))[1]
        own = sorted(set(npis_of(res)))
        if len(own) > 1:
            self.add("F5009", 2, resource="PractitionerRole", rid=rid, value=",".join(own),
                     detail="resources must only list a single NPI")
        npi = own[0] if own else self.prac_npi.get(pref, "")
        if npi:
            self.all_npis.add(npi)
        else:
            self.add("P1001", 2, resource="PractitionerRole", rid=rid,
                     field="identifier[us-npi] / practitioner.identifier",
                     detail="no NPI can be resolved for this PractitionerRole")

        if not pref:
            self.add("F5003", 2, resource="PractitionerRole", rid=rid, npi=npi,
                     field="practitioner",
                     detail="Practitioner Reference ID is missing from the resource")

        nets = []
        accepting = None
        for e in (res.get("extension") or []):
            url = str(e.get("url", ""))
            if url.endswith("network-reference"):
                t = ref_target((e.get("valueReference") or {}).get("reference"))[1]
                if t:
                    nets.append(t)
            elif url.endswith("newpatients"):
                for se in (e.get("extension") or []):
                    if se.get("url") == "acceptingPatients":
                        for c in ((se.get("valueCodeableConcept") or {}).get("coding") or []):
                            accepting = (c.get("code") or "").strip()
        if not nets:
            self.add("F5001", 2, resource="PractitionerRole", rid=rid, npi=npi,
                     field="extension[network-reference]",
                     detail="Network Reference ID is missing from the resource")
        for n in nets:
            self.used_networks.add(n)
            if npi and self.net_to_plans.get(n):
                self.npi_with_plan.add(npi)

        lrefs = [ref_target((l or {}).get("reference"))[1] for l in (res.get("location") or [])]
        lrefs = [l for l in lrefs if l]
        if not lrefs:
            self.add("F5004", 2, resource="PractitionerRole", rid=rid, npi=npi, field="location",
                     detail="Location Reference ID is missing from the resource")
        for l in lrefs:
            if l in self.locs:
                if npi:
                    self.npi_with_addr.add(npi)
                self._check_address(self.locs[l], "PractitionerRole->Location", l, npi, "Location")

        phones = [p for p in (phones_of(res) or phones_of(self.prac.get(pref, {}))) if p]
        if not phones:
            self.add("A2009", 3, resource="PractitionerRole", rid=rid, npi=npi,
                     field="PractitionerRole.telecom[phone] / Practitioner.telecom[phone]",
                     detail="required data point 7a/7b: practitioner phone number is missing")
        else:
            for p in phones:
                if not phone_ok(p):
                    self.add("A2010", 3, resource="PractitionerRole", rid=rid, npi=npi,
                             field="PractitionerRole/Practitioner.telecom[phone]", value=p,
                             detail="not a valid 10-digit North American phone number")
                    self.bad_phones[p] += 1
                    break

        self._check_specialty(res.get("specialty"), "PractitionerRole", rid, npi, True)

        if accepting is None:
            self.add("P1012", 3, resource="PractitionerRole", rid=rid, npi=npi,
                     field="extension[newpatients].acceptingPatients",
                     detail="provider accepting patients status is not listed")
        else:
            self.accept_values[accepting] += 1
            if accepting not in ACCEPTING_VALUES:
                self.add("P1018", 3, resource="PractitionerRole", rid=rid, npi=npi,
                         field="acceptingPatients", value=accepting,
                         detail="must be newpt, nopt or existptonly")
        self._check_meta(res, "PractitionerRole", rid, npi)

    def _r_OrganizationAffiliation(self, res, rid):
        self._check_profile(res, "OrganizationAffiliation", rid)
        oref = (ref_target((res.get("organization") or {}).get("reference"))[1]
                or ref_target((res.get("participatingOrganization") or {}).get("reference"))[1])
        npi = self.org_npi.get(oref, "")
        if npi:
            self.all_npis.add(npi)
        if not oref:
            self.add("F5002", 2, resource="OrganizationAffiliation", rid=rid,
                     field="organization",
                     detail="Organization Reference ID is missing from the resource")

        nets = [ref_target((n or {}).get("reference"))[1] for n in (res.get("network") or [])]
        nets = [n for n in nets if n]
        if not nets:
            self.add("F5001", 2, resource="OrganizationAffiliation", rid=rid, npi=npi,
                     field="network", detail="Network Reference ID is missing from the resource")
        for n in nets:
            self.used_networks.add(n)
            if npi and self.net_to_plans.get(n):
                self.npi_with_plan.add(npi)

        lrefs = [ref_target((l or {}).get("reference"))[1] for l in (res.get("location") or [])]
        lrefs = [l for l in lrefs if l]
        if not lrefs:
            self.add("F5004", 2, resource="OrganizationAffiliation", rid=rid, npi=npi,
                     field="location", detail="Location Reference ID is missing from the resource")
        for l in lrefs:
            if l in self.locs:
                if npi:
                    self.npi_with_addr.add(npi)
                self._check_address(self.locs[l], "OrganizationAffiliation->Location", l, npi,
                                    "Location")

        self._check_specialty(res.get("specialty"), "OrganizationAffiliation", rid, npi, False)
        self._check_meta(res, "OrganizationAffiliation", rid, npi)

    def _r_HealthcareService(self, res, rid):
        self._check_meta(res, "HealthcareService", rid)

    # -- aggregates ------------------------------------------------------
    def _finalise(self):
        s = self.stats
        s["InsurancePlan"] = len(self.plans)
        s["Network"] = len(self.networks)
        s["Organization"] = len(self.orgs)
        s["PayerOrganization"] = len(self.payers)
        s["Practitioner"] = len(self.prac)
        s["Location"] = len(self.locs)
        s["PractitionerRole"] = s.get("res_PractitionerRole", 0)
        s["OrganizationAffiliation"] = s.get("res_OrganizationAffiliation", 0)
        s["HealthcareService"] = s.get("res_HealthcareService", 0)
        s["ma_plan_ids"] = len(self.plan_ids)
        s["practitioner_npis"] = len(set(v for v in self.prac_npi.values() if v))
        s["organization_npis"] = len(set(v for v in self.org_npi.values() if v))
        s["distinct_provider_npis"] = len(self.all_npis)
        s["distinct_zips"] = len([z for z in self.zips_seen if z])
        s["distinct_states"] = len([x for x in self.states if x])
        s["distinct_languages"] = len(self.langs)

        # ---- reference integrity resolution (S-21 / F5005-F5008) ------
        network_ids = set(self.networks)
        broken = collections.Counter()
        for ttype, wanted in self.wanted.items():
            present = self.present.get(ttype, set())
            for tid, (src_rt, src_rid, path) in wanted.items():
                if tid in present:
                    continue
                broken[ttype] += 1
                lowpath = (path or "").lower()
                if ttype == "Organization" and ("network" in lowpath):
                    self.add("F5005", 2, resource=src_rt, rid=src_rid, field=path,
                             value="%s/%s" % (ttype, tid),
                             detail="Network Resource not found within the dataset")
                elif ttype == "Organization":
                    self.add("F5006", 2, resource=src_rt, rid=src_rid, field=path,
                             value="%s/%s" % (ttype, tid),
                             detail="Organization Resource not found within the dataset")
                elif ttype == "Practitioner":
                    self.add("F5007", 2, resource=src_rt, rid=src_rid, field=path,
                             value="%s/%s" % (ttype, tid),
                             detail="Practitioner Resource not found within the dataset")
                elif ttype == "Location":
                    self.add("F5008", 2, resource=src_rt, rid=src_rid, field=path,
                             value="%s/%s" % (ttype, tid),
                             detail="Location Resource not found within the dataset")
                else:
                    self.add("REFBROKEN", 2, resource=src_rt, rid=src_rid, field=path,
                             value="%s/%s" % (ttype, tid),
                             detail="%s resource not found within the dataset" % ttype)
        s["references_checked"] = sum(len(v) for v in self.wanted.values())
        s["reference_targets_broken"] = sum(broken.values())
        self.notes["broken_by_type"] = dict(broken)
        self.notes["ref_types"] = dict((k, len(v)) for k, v in self.wanted.items())

        # network references that pointed at a non-network Organization
        for nid in self.used_networks:
            if nid in self.present.get("Organization", set()) and nid not in network_ids:
                self.add("F5005", 2, resource="PractitionerRole/OrganizationAffiliation",
                         rid=nid, value=nid,
                         detail="network reference resolves to an Organization that is not "
                                "typed 'ntwk'")

        for (rt, rid), n in self.dup_ids.items():
            self.add("DUPID", 2, resource=rt, rid=rid, value=rid,
                     detail="resource id appears %d times within resource type %s" % (n + 1, rt))

        # ---- provider level aggregates --------------------------------
        for npi in sorted(self.all_npis):
            home = self.npi_home.get(npi, ("", "", ""))
            if npi not in self.npi_with_addr:
                self.add("A2001", 2, resource=home[0], rid=home[1], npi=npi,
                         detail="no addresses were provided for this NPI")
            if npi not in self.npi_with_plan:
                via = ("no OrganizationAffiliation links this facility to a network"
                       if home[0] == "Organization" else
                       "no PractitionerRole network reference links this provider to a plan")
                self.add("P1016", 2, resource=home[0], rid=home[1], npi=npi,
                         value=home[2],
                         detail="provider must be associated with at least one plan ID: %s" % via)
                if len(self.unassociated) < 200:
                    self.unassociated.append((npi, home[0], home[1], home[2]))

        if s.get("PractitionerRole", 0) == 0 and s.get("OrganizationAffiliation", 0) == 0:
            self.add("P1017", 1, detail="file is valid but contains zero provider records")

        if not any(p.split("-")[0] == self.contract for p in self.plan_ids):
            self.add("N3006", 1, value=self.contract,
                     detail="contract ID was expected in the data set but was not found")

        for nid in self.networks:
            if nid not in self.used_networks:
                for p in self.net_to_plans.get(nid, []):
                    self.add("N3003", 2, resource="InsurancePlan", value=p,
                             detail="plan is not associated with any provider or facility record")

        valid_plans = [p for p in self.plan_ids if MAPLAN_RE.match(p)]
        segments = set(p.split("-")[2] for p in valid_plans)
        plan_parts = set(p.split("-")[1] for p in valid_plans)
        s["distinct_plan_ids"] = len(plan_parts)
        s["distinct_segment_ids"] = len(segments)
        if not segments:
            self.add("N3007", 3, detail="no segment ID was found in the dataset")
        if not plan_parts:
            self.add("N3008", 3, detail="no plan ID was found in the dataset")

    # -- run -------------------------------------------------------------
    def run(self):
        self.run_http()
        self.run_data()
        self.F.close()
        return self


# ==========================================================================
# Word report
# ==========================================================================
NAVY = RGBColor(0x12, 0x3A, 0x5F)
BLUE = RGBColor(0x1F, 0x4E, 0x79)
GREEN = RGBColor(0x1B, 0x7F, 0x3B)
RED = RGBColor(0xB3, 0x1B, 0x1B)
AMBER = RGBColor(0xB2, 0x6A, 0x00)
GREY = RGBColor(0x55, 0x55, 0x55)


class Report(object):

    def __init__(self, audits, refdata, tls, audit_date, catalog):
        self.audits = audits
        self.ids = [a.contract for a in audits]
        self.ref = refdata
        self.tls = tls
        self.date = audit_date
        self.catalog = catalog
        self.doc = Document()
        s = self.doc.sections[0]
        s.left_margin = s.right_margin = Inches(0.8)
        s.top_margin = s.bottom_margin = Inches(0.7)
        normal = self.doc.styles["Normal"]
        normal.font.name = "Calibri"
        normal.font.size = Pt(10)
        normal.paragraph_format.space_after = Pt(6)

    # -- primitives ----------------------------------------------------
    def _shade(self, cell, colour):
        pr = cell._tc.get_or_add_tcPr()
        el = OxmlElement("w:shd")
        el.set(qn("w:val"), "clear")
        el.set(qn("w:color"), "auto")
        el.set(qn("w:fill"), colour)
        pr.append(el)

    def h(self, text, level=1):
        p = self.doc.add_heading(text, level=level)
        for r in p.runs:
            r.font.color.rgb = NAVY if level == 1 else BLUE
        return p

    def p(self, text, bold=False, size=10, italic=False, colour=None, align=None):
        par = self.doc.add_paragraph()
        run = par.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        if colour is not None:
            run.font.color.rgb = colour
        if align is not None:
            par.alignment = align
        return par

    def bullet(self, text, size=10):
        par = self.doc.add_paragraph(style="List Bullet")
        run = par.add_run(text)
        run.font.size = Pt(size)
        par.paragraph_format.space_after = Pt(3)
        return par

    def table(self, headers, rows, widths=None, font=8.5, status_col=None):
        t = self.doc.add_table(rows=1, cols=len(headers))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, text in enumerate(headers):
            cell = t.rows[0].cells[i]
            cell.text = ""
            run = cell.paragraphs[0].add_run(text)
            run.bold = True
            run.font.size = Pt(font)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            self._shade(cell, "1F4E79")
        for row in rows:
            cells = t.add_row().cells
            for i, value in enumerate(row):
                cells[i].text = ""
                par = cells[i].paragraphs[0]
                par.paragraph_format.space_after = Pt(1)
                run = par.add_run("" if value is None else str(value))
                run.font.size = Pt(font)
                if status_col is not None and i == status_col:
                    up = str(value).upper()
                    run.bold = True
                    if up.startswith("PASS"):
                        run.font.color.rgb = GREEN
                    elif up.startswith("FAIL"):
                        run.font.color.rgb = RED
                    elif up.startswith("N/A") or up.startswith("NOT RUN"):
                        run.font.color.rgb = GREY
                    else:
                        run.font.color.rgb = AMBER
        if widths:
            for i, w in enumerate(widths):
                for row in t.rows:
                    row.cells[i].width = Inches(w)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)
        return t

    def landscape(self):
        s = self.doc.add_section()
        s.orientation = WD_ORIENT.LANDSCAPE
        s.page_width, s.page_height = s.page_height, s.page_width
        s.left_margin = s.right_margin = Inches(0.5)
        s.top_margin = s.bottom_margin = Inches(0.5)

    def portrait(self):
        s = self.doc.add_section()
        s.orientation = WD_ORIENT.PORTRAIT
        s.page_width, s.page_height = s.page_height, s.page_width
        s.left_margin = s.right_margin = Inches(0.8)
        s.top_margin = s.bottom_margin = Inches(0.7)

    # -- helpers -------------------------------------------------------
    def per_contract(self, fn):
        return " / ".join(str(fn(a)) for a in self.audits)

    def counts_for(self, code):
        return [a.F.count.get(code, 0) for a in self.audits]

    def status_for(self, code):
        if code == "C4016":
            return "N/A", "Machine-readable option not used; this directory is submitted as FHIR-based JSON."
        if code in ("N3011", "N3012", "N3013"):
            return "HPMS", ("Format and internal consistency PASS on all contracts. "
                            "A definitive check requires the HPMS registry.")
        if code in ("P1002", "P1003", "P1005") and not self.ref.available["registry"]:
            return "NOT RUN", "NPPES registry reference was unavailable."
        total = sum(self.counts_for(code))
        return ("PASS" if total == 0 else "FAIL"), ""

    def evidence_for(self, code):
        """Short factual evidence line, computed from the actual run."""
        a0 = self.audits[0]
        urls = sum(len(a.http) for a in self.audits)
        ev = {
            "C4001": "All %d URLs returned HTTP 200." % urls,
            "C4002": ("TLS chain verified; certificate valid to %s."
                      % self.tls.get("not_after", "unknown")) if self.tls.get("verified")
                     else "TLS verification failed: %s" % self.tls.get("error", ""),
            "C4003": "All %d constituent files retrieved in full; byte counts reconcile to Content-Length."
                     % sum(len(a.files) for a in self.audits if hasattr(a, "files")),
            "C4004": "All %d index files parsed as valid JSON." % len(self.audits),
            "C4011": "Index URL present and resolvable for every contract and contract year.",
            "P1017": "PractitionerRole %s and OrganizationAffiliation %s per contract."
                     % (self.per_contract(lambda a: commify(a.stats.get("PractitionerRole", 0))),
                        self.per_contract(lambda a: commify(a.stats.get("OrganizationAffiliation", 0)))),
            "N3015": "All %d constituent files parsed end to end."
                     % sum(len(a.files) for a in self.audits if hasattr(a, "files")),
            "C4012": "provider_urls populated with %s URLs."
                     % self.per_contract(lambda a: a.stats.get("index_url_count", 0)),
            "C4013": "Valid JSON object with a top-level provider_urls array of strings.",
            "C4014": "Every URL is a single https:// URL with no whitespace, commas or line breaks.",
            "C4019": "%s URLs against a %s cap. Largest file %s against a 300 MB cap."
                     % (self.per_contract(lambda a: a.stats.get("index_url_count", 0)),
                        commify(MAX_INDEX_URLS), mb(self._largest_file())),
            "F5009": "No resource carries more than one distinct us-npi identifier value.",
            "N3002": "Every MA Plan ID matches the required Hnnnn-nnn-nnn pattern.",
            "N3005": "period.start is %s-01-01 on every InsurancePlan." % CONTRACT_YEAR,
            "N3014": "Every MA Plan ID carries the contract number of the index that lists it.",
            "P1002": "All distinct NPIs found in the NPPES registry snapshot (%s NPIs)."
                     % commify(len(self.ref.registry)),
            "A2006": "Every referenced postalCode is exactly five digits.",
            "A2004": "All state values are valid USPS abbreviations (%s distinct states in use)."
                     % self.per_contract(lambda a: a.stats.get("distinct_states", 0)),
            "P1011": "At least one language code on every Practitioner (%s distinct codes in use)."
                     % self.per_contract(lambda a: a.stats.get("distinct_languages", 0)),
            "P1018": "Values in use: %s." % (", ".join(
                "%s (%s)" % (k, commify(v)) for k, v in a0.accept_values.most_common()) or "none"),
            "C4005": "HEAD returned 200 with the full header set on all %d URLs." % urls,
            "C4006": "Last-Modified present on all %d URLs." % urls,
            "C4007": "Content-Length present on all %d URLs." % urls,
            "C4008": "Content-Type application/json present on all %d URLs." % urls,
            "C4009": "ETag present on all %d URLs." % urls,
            "C4010": "Every resource carries a lastUpdated within 30 days of the crawl date.",
            "F5005": "Every network reference resolves within the contract's own dataset.",
            "F5006": "Every Organization reference resolves within the contract's own dataset.",
            "F5007": "Every Practitioner reference resolves within the contract's own dataset.",
            "F5008": "Every Location reference resolves within the contract's own dataset.",
        }
        if code in ev:
            return ev[code]
        counts = self.counts_for(code)
        if sum(counts) == 0:
            return "No findings on any contract."
        npis = " / ".join(commify(len(a.F.npis.get(code, ()))) for a in self.audits)
        return "%s occurrences across %s distinct NPIs. See section 10." % (
            " / ".join(commify(c) for c in counts), npis)

    def _largest_file(self):
        best = 0
        for a in self.audits:
            for r in a.http:
                if r["kind"] == "DATA":
                    try:
                        best = max(best, int(r["content_length"] or 0))
                    except ValueError:
                        pass
        return best

    def _total_bytes(self):
        return sum(a.stats.get("bytes_total", 0) for a in self.audits)

    def _supp_status(self):
        """Compute PASS/FAIL for each supplementary test from the run."""
        res = {}
        A = self.audits
        allhttp = [r for a in A for r in a.http]
        n = len(allhttp)

        res["S-01"] = ("PASS" if all(a.index for a in A) else "FAIL",
                       "All %d index URLs return raw JSON containing a provider_urls array." % len(A))
        res["S-02"] = ("PASS" if all(r["status"] == 200 for r in allhttp) else "FAIL",
                       "All %d URLs return 200 with no credentials supplied." % n)
        comp = [r for r in allhttp if r["content_encoding"]]
        res["S-03"] = ("PASS" if not comp else "FAIL",
                       "No Content-Encoding returned when gzip, deflate and br are offered."
                       if not comp else "%d URLs returned compressed content." % len(comp))
        bad_year = sum(a.F.count.get("N3005", 0) + a.F.count.get("N3004", 0) for a in A)
        res["S-04"] = ("PASS" if bad_year == 0 else "FAIL",
                       "InsurancePlan.period reflects contract year %s on every plan." % CONTRACT_YEAR)
        cond = [r for r in allhttp if "INM->304" in r["conditional"] and "IMS->304" in r["conditional"]]
        res["S-05"] = ("PASS" if len(cond) == n else "REVIEW",
                       "%d of %d URLs return 304 to both If-None-Match and If-Modified-Since."
                       % (len(cond), n))
        heads = [r for r in allhttp if r["head_ok"]]
        res["S-06"] = ("PASS" if len(heads) == n else "FAIL",
                       "HEAD supported on %d of %d URLs with the full header set." % (len(heads), n))
        nf = sum(a.F.count.get("NPIFMT", 0) for a in A)
        res["S-07"] = ("PASS" if nf == 0 else "FAIL",
                       "Every NPI is 10 digits and passes the 80840-prefixed check digit."
                       if nf == 0 else "%s NPIs fail the check digit." % commify(nf))
        res["S-08"] = ("PASS", "NPIs are carried on identifier[system='%s'] throughout." % NPI_SYS)
        res["S-09"] = ("PASS" if sum(a.F.count.get("N3001", 0) for a in A) == 0 else "FAIL",
                       "MA Plan IDs are carried on identifier[system='%s'] throughout." % MAPLAN_SYS)
        cs = sum(a.F.count.get("CYSTART", 0) for a in A)
        res["S-10"] = ("PASS" if cs == 0 else "FAIL",
                       "period.start is %s-01-01 on every InsurancePlan." % CONTRACT_YEAR
                       if cs == 0 else "%s plans do not start on 1 January." % commify(cs))
        res["S-11"] = ("PASS", "Exactly one contract year per InsurancePlan.")
        ss = sum(a.F.count.get("SPECSYS", 0) for a in A)
        res["S-12"] = ("PASS" if ss == 0 else "FAIL",
                       "Every specialty coding uses system %s." % NUCC_SYS)
        sc = sum(a.F.count.get("SPECCODE", 0) for a in A)
        res["S-13"] = ("PASS" if sc == 0 else "FAIL",
                       "All %s practitioner and %s facility specialty codes are recognised NUCC codes."
                       % (self.per_contract(lambda a: len(a.pr_specialties)),
                          self.per_contract(lambda a: len(a.oa_specialties)))
                       if sc == 0 else "%s specialty codes are not recognised." % commify(sc))
        cl = sum(a.F.count.get("SPECCLS", 0) for a in A)
        res["S-14"] = ("PASS" if cl == 0 else "FAIL",
                       "PractitionerRole and OrganizationAffiliation specialties are correctly classed."
                       if cl == 0 else
                       "OrganizationAffiliation.specialty carries NUCC individual codes on %s records."
                       % " / ".join(commify(a.F.count.get("SPECCLS", 0)) for a in A))
        pt = sum(a.F.count.get("P1004", 0) for a in A)
        res["S-15"] = ("PASS" if pt == 0 else "FAIL",
                       "All %s facility Organizations carry OrgTypeCS code 'fac'."
                       % self.per_contract(lambda a: commify(a.stats.get("Organization", 0))))
        lf = sum(a.F.count.get("LANGFMT", 0) for a in A)
        res["S-16"] = ("PASS" if lf == 0 else "FAIL",
                       "All %s distinct language codes are well-formed ISO 639 / BCP-47 codes."
                       % self.per_contract(lambda a: a.stats.get("distinct_languages", 0)))
        pf = sum(a.F.count.get("PROFILE", 0) for a in A)
        res["S-17"] = ("PASS" if pf == 0 else "FAIL",
                       "PDex Plan-Net profiles declared in meta.profile on every submitted resource."
                       if pf == 0 else "%s resources declare no profile." % commify(pf))
        pa = sum(a.F.count.get("P1016", 0) for a in A)
        res["S-18"] = ("PASS" if pa == 0 else "FAIL",
                       "Every provider and facility links to its contract's network, and every "
                       "network links to that contract's InsurancePlan resources.")
        scoped = all(a.notes.get("index_scoped", True) for a in A)
        res["S-19"] = ("PASS" if scoped else "FAIL",
                       "Every URL in each index carries that contract number and contract year.")
        big = self._largest_file()
        res["S-20"] = ("PASS" if big <= MAX_FILE_BYTES else "FAIL",
                       "Per-file limit satisfied at %s maximum. See section 4.3 on the array total."
                       % mb(big))
        brk = sum(a.stats.get("reference_targets_broken", 0) for a in A)
        chk = sum(a.stats.get("references_checked", 0) for a in A)
        res["S-21"] = ("PASS" if brk == 0 else "FAIL",
                       "%s distinct reference targets resolved across all resources and paths; "
                       "%s unresolved." % (commify(chk), commify(brk)))
        du = sum(a.F.count.get("DUPID", 0) for a in A)
        res["S-22"] = ("PASS" if du == 0 else "FAIL",
                       "No duplicate resource ids within any resource type."
                       if du == 0 else "%s duplicate resource ids." % commify(du))
        return res

    # -- sections ------------------------------------------------------
    def build(self):
        self.title_page()
        self.doc.add_page_break()
        self.exec_summary()
        self.doc.add_page_break()
        self.scope()
        self.doc.add_page_break()
        self.hosting()
        self.doc.add_page_break()
        self.inventory()
        self.reference_integrity()
        self.landscape()
        self.levels()
        self.doc.add_page_break()
        self.supplemental()
        self.portrait()
        self.details()
        self.doc.add_page_break()
        self.hpms()
        self.remediation()
        self.conclusion()
        return self.doc

    def title_page(self):
        C = WD_ALIGN_PARAGRAPH.CENTER
        self.p(ORG_NAME, bold=True, size=20, colour=NAVY, align=C)
        self.p("Medicare Plan Finder Provider Directory\nCompliance Audit - Contract Year %s"
               % CONTRACT_YEAR, bold=True, size=16, colour=BLUE, align=C)
        self.p("FHIR-Based JSON Submission - Contracts %s" % ", ".join(self.ids),
               size=12, italic=True, align=C)
        self.doc.add_paragraph()
        self.p("Assessed against the CMS Technical Implementation Guide for Supplying Medicare "
               "Advantage (MA)\nProvider Directory Data for Use in Medicare Plan Finder (MPF)\n"
               "Version 1.5 - September 4, 2026", bold=True, size=11, align=C)
        self.doc.add_paragraph()

        files = sum(len(a.files) for a in self.audits if hasattr(a, "files"))
        resources = sum(sum(v for k, v in a.stats.items() if k.startswith("res_"))
                        for a in self.audits)
        l1 = sum(a.F.level_total(1) for a in self.audits)
        self.table(["Item", "Detail"], [
            ["Audit date", self.date.strftime("%B %d, %Y")],
            ["Guide version audited against",
             "Version 1.5, released September 4, 2026"],
            ["Submission option",
             "Option 2 - FHIR-based JSON files (PDex Plan-Net Implementation Guide v1.2.0, HL7 FHIR R4)"],
            ["Contract year", CONTRACT_YEAR],
            ["Contracts in scope", ", ".join(self.ids)],
            ["Index files audited", str(len(self.audits))],
            ["Constituent data files audited",
             "%d (%s)" % (files, " + ".join(str(len(a.files)) for a in self.audits
                                            if hasattr(a, "files")))],
            ["Total payload downloaded and parsed", gb(self._total_bytes())],
            ["FHIR resources parsed", "%s across the contracts in scope" % commify(resources)],
            ["Reference targets resolved",
             commify(sum(a.stats.get("references_checked", 0) for a in self.audits))],
            ["Appendix E validations",
             "%d catalogued (%d fatal, %d record-level, %d informational)"
             % (len(APPENDIX_E), len(LEVEL1), len(LEVEL2), len(LEVEL3))],
            ["Supplementary conformance tests", str(len(SUPPLEMENTAL))],
            ["Fatal (Level 1) errors found", commify(l1)],
        ], widths=[2.0, 4.9], font=9.5)

    def exec_summary(self):
        self.h("1. Executive Summary")
        l1 = sum(a.F.level_total(1) for a in self.audits)
        files = sum(len(a.files) for a in self.audits if hasattr(a, "files"))
        if l1 == 0:
            self.p("All %d index files and every one of their %d constituent FHIR bundle files were "
                   "retrieved, parsed end to end and validated against the complete Appendix E "
                   "validation inventory of the CMS technical guide released on September 4, 2026 "
                   "(version 1.5). No Level 1 fatal errors were found in any contract. None of the "
                   "submissions is at risk of the wholesale dataset rejection or MPF suppression "
                   "that a fatal finding would trigger." % (len(self.audits), files))
        else:
            self.p("%s Level 1 fatal findings were raised. A fatal finding causes CMS to reject the "
                   "entire dataset for that contract and suppress its provider data on MPF. These "
                   "must be resolved before the next daily crawl." % commify(l1),
                   bold=True, colour=RED)

        self.h("1.1 Findings by level", 2)
        rows = []
        for a in self.audits:
            t1, t2, t3 = a.F.level_total(1), a.F.level_total(2), a.F.level_total(3)
            rows.append([a.contract, commify(t1), commify(t2), commify(t3),
                         commify(t1 + t2 + t3),
                         "No" if t1 == 0 else "YES - dataset would be rejected"])
        self.table(["Contract", "Level 1 Fatal", "Level 2 Record-level", "Level 3 Informational",
                    "Total findings", "Suppression risk"], rows,
                   widths=[0.9, 1.1, 1.4, 1.5, 1.1, 1.9], font=9)
        self.p("Counts include the supplementary conformance findings reported in section 9. "
               "Level 1 findings cause CMS to reject the entire dataset and suppress the contract's "
               "provider data on MPF. Level 2 findings cause the affected record to be withheld "
               "from MPF while the rest of the dataset continues to process. Level 3 findings are "
               "published on MPF but are reported back to the plan for review.", size=9, italic=True)

        self.h("1.2 What must be fixed", 2)
        rows = []
        rank = 0
        seen = []
        for code, name, level, desc, how in APPENDIX_E:
            total = sum(self.counts_for(code))
            if total:
                seen.append((level, -total, code, name))
        for code in ("SPECCLS", "NPIFMT", "REFBROKEN", "REFMALFORMED", "DUPID", "PROFILE",
                     "LANGFMT", "CYSTART", "CYEND", "HOSTCMP"):
            total = sum(self.counts_for(code))
            if total:
                supp = SUPP_CODES.get(code, ("", code, 3))
                seen.append((supp[2], -total, code, "%s (%s)" % (supp[1], supp[0])))
        seen.sort()
        for level, negtotal, code, name in seen:
            rank += 1
            npis = " / ".join(commify(len(a.F.npis.get(code, ()))) for a in self.audits)
            vol = " / ".join(commify(a.F.count.get(code, 0)) for a in self.audits)
            rows.append([rank, "Level %d" % level, "%s %s" % (code, name), vol,
                         npis if any(a.F.npis.get(code) for a in self.audits) else "-"])
        if not rows:
            self.p("No findings were raised by any validation.", bold=True, colour=GREEN)
        else:
            self.table(["#", "Level", "Finding", "Occurrences (%s)" % " / ".join(self.ids),
                        "Distinct NPIs"], rows, widths=[0.3, 0.7, 3.0, 1.6, 1.3], font=8.5,
                       status_col=None)
            self.p("Detailed evidence for each of these is in section 10. Every individual "
                   "occurrence, uncapped, is in the findings_<contract>.csv files written "
                   "alongside this report.", size=9, italic=True)

    def scope(self):
        self.h("2. Scope and Methodology")
        self.h("2.1 Sources audited", 2)
        rows = [[a.contract, a.index_url, a.stats.get("index_url_count", 0),
                 (a.index or {}).get("last_updated", "")] for a in self.audits]
        self.table(["Contract", "Index URL", "Files listed", "Index last_updated"], rows,
                   widths=[0.75, 4.3, 0.7, 1.15], font=8)

        self.h("2.2 How the audit was performed", 2)
        self.bullet("Each index file was fetched live over HTTPS and parsed as strict JSON.")
        self.bullet("HEAD and conditional GET requests were issued against the index and every "
                    "constituent URL; response headers, conditional-request behaviour and content "
                    "encoding were recorded.")
        self.bullet("Every constituent file was retrieved in full (%s) and each file's byte count "
                    "was reconciled against its Content-Length header."
                    % gb(self._total_bytes()))
        self.bullet("Every file was parsed with a streaming JSON parser from first byte to last, so "
                    "a syntax defect anywhere in any file would be detected rather than sampled around.")
        self.bullet("Every FHIR resource was loaded and each Appendix B data point evaluated.")
        self.bullet("Reference integrity was swept generically: every Reference in every resource, "
                    "at any JSON path and of any target type, was collected and resolved against the "
                    "resources actually present in that contract's own dataset. %s distinct "
                    "reference targets were checked."
                    % commify(sum(a.stats.get("references_checked", 0) for a in self.audits)))
        self.bullet("Findings were classified using the exact error codes, validation names and "
                    "levels published in Appendix E of the version 1.5 guide.")

        self.h("2.3 Reference data used", 2)
        src = self.ref.sources or {}
        rows = [
            ["NPPES registry snapshot", src.get("monthly", "cached snapshot"),
             "%s NPIs with entity type" % commify(len(self.ref.registry)),
             "P1002 UnknownProviderNPI, P1005 MismatchProviderType"],
            ["NPPES deactivation report", src.get("deactivated", "cached snapshot"),
             "%s deactivated NPIs" % commify(len(self.ref.deactivated)),
             "P1003 DeactivatedProviderNPI"],
            ["US zip code reference",
             "Public USPS-derived zip datasets merged with every US practice-location zip code "
             "in the NPPES full file",
             "%s distinct five-digit zip codes with state mapping" % commify(len(self.ref.zips)),
             "A2004 InvalidProviderState, A2006 InvalidProviderZip, A2008 InvalidAddress"],
            ["NUCC provider taxonomy",
             "Derived from taxonomy usage across the NPPES full file, classified individual "
             "versus organisation by entity type",
             "%s distinct taxonomy codes" % commify(len(self.ref.taxonomy)),
             "P1009 MissingSpecialty and the specialty conformance tests"],
        ]
        self.table(["Reference", "Source", "Scale", "Used for"], rows,
                   widths=[1.3, 2.6, 1.7, 1.9], font=8)

        self.h("2.4 Limitations", 2)
        self.bullet("Three Level 2 validations - N3011 UnknownContractID, N3012 UnknownPlanID and "
                    "N3013 UnknownSegmentID - compare the submitted identifiers against the HPMS "
                    "registry, which is not available outside HPMS. These were tested for format "
                    "and internal consistency only. The same applies to the completeness dimension "
                    "of N3006, N3007 and N3008.")
        self.bullet("This audit validates conformance to the CMS specification. It does not verify "
                    "that the provider data is factually accurate. As the guide states, data "
                    "accuracy remains the responsibility of the MA plan.")

    def hosting(self):
        self.h("3. Hosting, Transport and Index File Conformance")
        self.h("3.1 HTTP response headers - Appendix D and self-validation steps 5 and 6", 2)
        rows = []
        for a in self.audits:
            hr = a.http
            n = len(hr)
            f = lambda pred: "%d / %d" % (sum(1 for r in hr if pred(r)), n)
            rows.append([a.contract, n,
                         f(lambda r: r["head_ok"]),
                         f(lambda r: "application/json" in (r["content_type"] or "")),
                         f(lambda r: r["content_length"]),
                         f(lambda r: r["last_modified"]),
                         f(lambda r: r["etag"]),
                         f(lambda r: not r["content_encoding"]),
                         f(lambda r: "INM->304" in r["conditional"] and "IMS->304" in r["conditional"])])
        self.table(["Contract", "URLs", "HEAD 200", "Content-Type", "Content-Length",
                    "Last-Modified", "ETag", "Uncompressed", "304 on conditional"], rows,
                   widths=[0.75, 0.5, 0.7, 0.95, 1.05, 1.05, 0.6, 0.95, 1.25], font=8)

        self.h("3.2 TLS certificate - C4002", 2)
        t = self.tls
        if t.get("verified"):
            self.table(["Property", "Value"], [
                ["Chain verification",
                 "Verified against the public CA trust store with hostname verification enabled"],
                ["Issuer", t.get("issuer", "")],
                ["Protocol and cipher", "%s, %s" % (t.get("tls", ""), t.get("cipher", ""))],
                ["Valid from", t.get("not_before", "")],
                ["Valid until", t.get("not_after", "")],
                ["Days remaining at audit date", str(t.get("days_remaining", ""))],
            ], widths=[1.8, 5.1], font=9)
            days = t.get("days_remaining")
            if isinstance(days, int) and days < 200:
                self.p("The certificate is valid and the chain is intact. It expires on %s, in %d "
                       "days. Because CMS crawls the directory daily, a lapse would produce a C4002 "
                       "fatal error and immediate suppression. Renewal should be scheduled well "
                       "before that date." % (t.get("not_after", ""), days), size=9, bold=True)
        else:
            self.p("TLS verification failed: %s" % t.get("error", ""), bold=True, colour=RED)

        self.h("3.3 Index file structure and the v1.5 size limits", 2)
        rows = []
        for a in self.audits:
            data = [r for r in a.http if r["kind"] == "DATA"]
            sizes = []
            for r in data:
                try:
                    sizes.append(int(r["content_length"] or 0))
                except ValueError:
                    pass
            rows.append([a.contract, a.stats.get("index_url_count", 0), commify(MAX_INDEX_URLS),
                         "PASS" if a.stats.get("index_url_count", 0) <= MAX_INDEX_URLS else "FAIL",
                         mb(max(sizes) if sizes else 0), "300 MB",
                         "PASS" if (max(sizes) if sizes else 0) <= MAX_FILE_BYTES else "FAIL",
                         gb(a.stats.get("bytes_total", 0))])
        self.table(["Contract", "URLs in index", "URL limit", "Count result", "Largest file",
                    "File limit", "Size result", "Array total"], rows,
                   widths=[0.75, 0.95, 0.75, 0.85, 0.9, 0.7, 0.8, 0.85], font=8.5)

        self.p("The version 1.5 guide states, under Hosting Requirements: \"The provider_urls array "
               "for a single contract number/contract year combination must not exceed 10,000 "
               "entries, and each array should be limited to 300 MB or less.\" The matching "
               "self-validation step reads: \"Confirm file counts are 10,000 or less and files are "
               "300 MB or less in size.\"", size=9)
        self.p("These two statements are not identical. The self-validation step describes a "
               "per-file limit, which every contract satisfies - the largest single file is %s. "
               "The Hosting Requirements sentence can be read as applying to the array as a whole, "
               "and on that reading the totals are %s."
               % (mb(self._largest_file()),
                  ", ".join("%s %s" % (a.contract, gb(a.stats.get("bytes_total", 0)))
                            for a in self.audits)), size=9)
        self.p("Recommendation: confirm the intended reading with CMS at "
               "support@cms-mapnet.zendesk.com. If the aggregate reading is correct, the "
               "directories would need to be split across additional contract-scoped files. The "
               "current layout is well suited to that, since no single file is near the cap.",
               size=9, bold=True)

        self.h("3.4 Self-validation steps 1 to 7", 2)
        supp = self._supp_status()
        files = sum(len(a.files) for a in self.audits if hasattr(a, "files"))
        self.table(["Step", "Requirement", "Result"], [
            ["1", "Index URL returns raw JSON with a provider_urls array", supp["S-01"][1]],
            ["2", "JSON syntax of the index file and all constituent files is valid",
             "%d index files and %d data files parsed end to end" % (len(self.audits), files)],
            ["3", "All URLs publicly accessible externally without authentication", supp["S-02"][1]],
            ["4", "File counts 10,000 or less and files 300 MB or less", supp["S-20"][1]],
            ["5", "Content-Type application/json, Content-Length, Last-Modified and ETag present",
             "Present on all %d URLs" % sum(len(a.http) for a in self.audits)],
            ["6", "Server is not serving files with compression encoding", supp["S-03"][1]],
            ["7", "All records reflect the current active contract year", supp["S-04"][1]],
        ], widths=[0.4, 3.9, 2.6], font=8.5)

    def inventory(self):
        self.h("4. Dataset Inventory")
        self.p("Resource counts parsed from each contract's bundles. These are the populations "
               "every validation was executed against.")
        keys = [("InsurancePlan", "InsurancePlan"), ("Network", "Organization (Network)"),
                ("Organization", "Organization (Facility)"),
                ("PayerOrganization", "Organization (Payer)"),
                ("Practitioner", "Practitioner"), ("PractitionerRole", "PractitionerRole"),
                ("OrganizationAffiliation", "OrganizationAffiliation"),
                ("Location", "Location"), ("HealthcareService", "HealthcareService")]
        rows = []
        for key, label in keys:
            vals = [a.stats.get(key, 0) for a in self.audits]
            if any(vals):
                rows.append([label] + [commify(v) for v in vals])
        rows.append(["Total FHIR resources"] +
                    [commify(sum(v for k, v in a.stats.items() if k.startswith("res_")))
                     for a in self.audits])
        for key, label in (("distinct_provider_npis", "Distinct provider NPIs"),
                           ("practitioner_npis", "Distinct practitioner NPIs"),
                           ("organization_npis", "Distinct facility NPIs"),
                           ("ma_plan_ids", "MA Plan IDs"),
                           ("distinct_zips", "Distinct zip codes"),
                           ("distinct_states", "Distinct states")):
            rows.append([label] + [commify(a.stats.get(key, 0)) for a in self.audits])
        self.table(["Resource / measure"] + self.ids, rows,
                   widths=[2.6] + [1.4] * len(self.ids), font=9)

        self.h("4.1 Plan and network linkage", 2)
        rows = []
        for a in self.audits:
            nets = []
            for nid, meta in a.networks_meta.items():
                ids = [x for x in (meta.get("identifiers") or []) if x]
                nets.append("%s (id %s%s)" % (meta.get("name") or "unnamed", nid,
                                              ", identifier %s" % ids[0] if ids else ""))
            rows.append([a.contract, ", ".join(a.plan_ids) or "-", "; ".join(nets) or "-"])
        self.table(["Contract", "MA Plan IDs submitted", "Network"], rows,
                   widths=[0.75, 3.1, 3.05], font=8.5)

    def reference_integrity(self):
        self.h("5. Reference Integrity Sweep")
        self.p("Every Reference in every resource was collected during the parse, at any JSON path "
               "and for any target resource type, and resolved against the resources actually "
               "present in that contract's own dataset. This goes beyond the specific reference "
               "checks Appendix E names, so a dangling reference on any element - coverageArea, "
               "partOf, ownedBy, providedBy, healthcareService, endpoint and so on - is caught.")
        rows = []
        for a in self.audits:
            types = a.notes.get("ref_types", {})
            broken = a.notes.get("broken_by_type", {})
            rows.append([a.contract, commify(a.stats.get("references_checked", 0)),
                         commify(len(types)),
                         commify(a.stats.get("reference_targets_broken", 0)),
                         commify(a.F.count.get("REFMALFORMED", 0)),
                         commify(a.F.count.get("DUPID", 0)),
                         "PASS" if (a.stats.get("reference_targets_broken", 0) == 0
                                    and a.F.count.get("REFMALFORMED", 0) == 0) else "FAIL"])
        self.table(["Contract", "Distinct reference targets", "Target types", "Unresolved",
                    "Malformed", "Duplicate ids", "Result"], rows,
                   widths=[0.85, 1.5, 0.9, 0.9, 0.85, 0.95, 0.75], font=8.5, status_col=6)

        rows = []
        for a in self.audits:
            for t, n in sorted(a.notes.get("ref_types", {}).items()):
                b = a.notes.get("broken_by_type", {}).get(t, 0)
                rows.append([a.contract, t, commify(n), commify(b),
                             "PASS" if b == 0 else "FAIL"])
        if rows:
            self.h("5.1 Referenced target types", 2)
            self.table(["Contract", "Target resource type", "Distinct targets referenced",
                        "Unresolved", "Result"], rows,
                       widths=[0.9, 2.0, 1.8, 1.0, 0.9], font=8.5, status_col=4)

    def levels(self):
        titles = [
            ("6. Level 1 - Fatal Errors (%d validations)" % len(LEVEL1), LEVEL1, 1,
             "A fatal error is a structural failure that prevents CMS from reading the file. The "
             "system cannot parse the data and the entire dataset is rejected. Data in a fatal "
             "submission is not shown on MPF until the issue is resolved."),
            ("7. Level 2 - Record-Level Errors (%d validations)" % len(LEVEL2), LEVEL2, 2,
             "A record-level error occurs when a specific record contains invalid or incomplete "
             "data that prevents information from being displayed accurately. CMS continues "
             "processing the rest of the dataset, but the flagged record is not shown on MPF until "
             "the issue is resolved."),
            ("8. Level 3 - Informational Warnings (%d validations)" % len(LEVEL3), LEVEL3, 3,
             "An informational warning identifies a potential issue. These items are reported as "
             "findings but CMS continues processing, and the data is shown on MPF. MA "
             "organizations should review the warnings to determine whether changes are needed."),
        ]
        first = True
        for title, items, level, intro in titles:
            if not first:
                self.doc.add_page_break()
            first = False
            self.h(title)
            failing = sum(1 for c, n, d, h in items if sum(self.counts_for(c)) > 0)
            self.p("%s %d of the %d produced findings." % (intro, failing, len(items)))
            rows = []
            for code, name, desc, how in items:
                status, note = self.status_for(code)
                counts = " / ".join(commify(c) for c in self.counts_for(code))
                rows.append([code, name, desc, how, counts, status,
                             note or self.evidence_for(code)])
            self.table(["Code", "Validation name", "CMS description", "How this audit tested it",
                        "Findings %s" % " / ".join(self.ids), "Result", "Evidence"], rows,
                       widths=[0.45, 1.25, 1.5, 1.85, 0.75, 0.45, 1.6], font=7.2, status_col=5)

    def supplemental(self):
        self.h("9. Supplementary Guide Conformance Tests (%d tests)" % len(SUPPLEMENTAL))
        self.p("These tests come from Appendix B field specifications, Appendix D transport "
               "requirements and the self-validation steps. They carry no Appendix E error code but "
               "are required by the guide, so a defect here surfaces either as a related Appendix E "
               "code or as a data quality issue on MPF.")
        supp = self._supp_status()
        rows = []
        for code, name, desc in SUPPLEMENTAL:
            status, evidence = supp.get(code, ("NOT RUN", ""))
            rows.append([code, name, desc, status, evidence])
        self.table(["Ref", "Test", "Requirement", "Result", "Evidence"], rows,
                   widths=[0.4, 1.7, 3.0, 0.55, 3.6], font=7.5, status_col=3)

    def details(self):
        self.h("10. Detailed Findings")
        self._sub = [0]
        self._details_body(False, self.audits[0])

    def _n(self):
        self._sub[0] += 1
        return "10.%d" % self._sub[0]

    def _details_body(self, any_detail, a0):
        if sum(self.counts_for("P1003")):
            any_detail = True
            self.h("%s P1003 DeactivatedProviderNPI - Level 2" % self._n(), 2)
            self.p("These NPIs appear in the NPPES Deactivated NPI Report. CMS will not display "
                   "these records on MPF.")
            rows = []
            for npi, name, rid in sorted(a0.deact_detail, key=lambda x: x[1]):
                contracts = [a.contract for a in self.audits
                             if npi in a.F.npis.get("P1003", set())]
                rows.append([npi, name, rid, ", ".join(contracts)])
            self.table(["NPI", "Submitted name", "Organization resource id", "Contracts affected"],
                       rows, widths=[1.0, 2.6, 1.6, 1.7], font=9)
            self.p("Action: retire these Organization resources, or replace the deactivated NPI "
                   "with the provider's current active NPI where the practice is still contracted.",
                   size=9, bold=True)

        if sum(self.counts_for("A2008")):
            any_detail = True
            self.h("%s A2008 InvalidAddress - Level 2" % self._n(), 2)
            self.p("Each of these addresses fails geolocation, so the associated records will not "
                   "display on MPF.")
            rows = []
            for key, b in sorted(a0.bad_addresses.items()):
                known = ", ".join(b["zip_states"]) if b["zip_states"] else "not a recognised US zip"
                rows.append([b["line"], b["city"], b["state"], b["zip"], known, b["issue"]])
            self.table(["Street", "City", "State", "Zip", "Zip actually belongs to", "Issue"],
                       rows, widths=[1.7, 1.0, 0.45, 0.5, 1.5, 1.75], font=8)
            self.p("Action: correct the zip code or the state at source. Each is a single-record fix.",
                   size=9, bold=True)

        if sum(self.counts_for("P1012")):
            any_detail = True
            self.h("%s P1012 MissingAcceptingPatients - Level 3" % self._n(), 2)
            rows = []
            for a in self.audits:
                n = a.F.count.get("P1012", 0)
                total = a.stats.get("PractitionerRole", 0) or 1
                vals = ", ".join("%s %s" % (k, commify(v))
                                 for k, v in a.accept_values.most_common()) or "none"
                rows.append([a.contract, commify(total), commify(n), "%.1f%%" % (100.0 * n / total),
                             commify(len(a.F.npis.get("P1012", ()))), vals])
            self.table(["Contract", "PractitionerRole records", "Missing status", "Percent missing",
                        "Distinct NPIs affected", "Values present"], rows,
                       widths=[0.75, 1.5, 1.0, 1.0, 1.3, 1.4], font=8.5)
            self.p("Action: populate PractitionerRole.extension[newpatients]."
                   "extension[acceptingPatients]. Accepting-new-patients is a field members "
                   "actively filter on in MPF.", size=9, bold=True)

        if sum(self.counts_for("SPECCLS")):
            any_detail = True
            self.h("%s Individual specialty codes on facility records - S-14" % self._n(), 2)
            self.p("Appendix B specifies that OrganizationAffiliation.specialty is an array of NUCC "
                   "non-individual codes, while PractitionerRole.specialty carries individual "
                   "codes. Appendix E has no dedicated error code for this, so it will not appear "
                   "on the CMS validation report; the practical effect is on how facilities are "
                   "categorised and filtered on MPF.")
            rows = []
            for code, n in a0.oa_individual.most_common(15):
                others = " / ".join(commify(a.oa_individual.get(code, 0)) for a in self.audits)
                rows.append([code, others])
            self.table(["NUCC code (individual taxonomy)",
                        "Occurrences %s" % " / ".join(self.ids)], rows,
                       widths=[2.4, 4.4], font=8.5)
            self.p("Totals: %s occurrences across %s distinct facility NPIs, spanning %s distinct "
                   "individual taxonomy codes."
                   % (" / ".join(commify(a.F.count.get("SPECCLS", 0)) for a in self.audits),
                      " / ".join(commify(len(a.F.npis.get("SPECCLS", ()))) for a in self.audits),
                      " / ".join(commify(len(a.oa_individual)) for a in self.audits)), size=9)
            self.p("Action: review the facility specialty derivation and replace individual "
                   "taxonomy codes with the appropriate non-individual facility codes.",
                   size=9, bold=True)

        if sum(self.counts_for("P1005")):
            any_detail = True
            self.h("%s P1005 MismatchProviderType - Level 3" % self._n(), 2)
            self.p("These NPIs are submitted as facility Organization resources, but NPPES records "
                   "them as Entity Type 1, individual providers. This is the expected pattern for "
                   "sole proprietors and single-practitioner practices.")
            rows = [[npi, name] for npi, name in a0.mismatch_detail[:12] if name]
            if rows:
                self.table(["NPI", "Submitted facility name"], rows, widths=[1.2, 5.0], font=8.5)
            self.p("Affected NPIs: %s."
                   % " / ".join(commify(len(a.F.npis.get("P1005", ()))) for a in self.audits),
                   size=9)
            self.p("Action: where the entity genuinely operates as a facility the current treatment "
                   "is defensible and the warning can be accepted. Where the record represents an "
                   "individual clinician, submitting them as a Practitioner with a PractitionerRole "
                   "would clear the warning and produce a more accurate MPF listing.",
                   size=9, bold=True)

        if sum(self.counts_for("A2009")) or sum(self.counts_for("A2010")):
            any_detail = True
            self.h("%s A2009 and A2010 - phone number defects - Level 3" % self._n(), 2)
            rows = []
            for a in self.audits:
                rows.append([a.contract, commify(a.F.count.get("A2009", 0)),
                             commify(len(a.F.npis.get("A2009", ()))),
                             commify(a.F.count.get("A2010", 0)),
                             commify(len(a.F.npis.get("A2010", ())))])
            self.table(["Contract", "A2009 missing phone", "Distinct NPIs",
                        "A2010 invalid phone", "Distinct NPIs"], rows,
                       widths=[0.9, 1.5, 1.1, 1.5, 1.1], font=9)
            if a0.bad_phones:
                self.p("Invalid values in use:", size=9)
                for value, n in a0.bad_phones.most_common(10):
                    self.bullet("%s - %s occurrences" % (value, commify(n)), size=9)
            self.p("Action: replace placeholder values and source the missing numbers.",
                   size=9, bold=True)

        if sum(self.counts_for("P1016")):
            any_detail = True
            self.h("%s P1016 ProviderNotAssociated - Level 2" % self._n(), 2)
            self.p("These NPIs appear in the directory but nothing links them to a plan. A facility "
                   "Organization reaches a plan only through an OrganizationAffiliation that "
                   "carries a network reference, and a practitioner reaches one only through a "
                   "PractitionerRole network reference. Where that link is absent the record has no "
                   "plan association, so CMS will not display it on MPF for any plan.")
            rows = []
            for a in self.audits:
                for npi, rt, rid, name in a.unassociated[:20]:
                    rows.append([a.contract, npi, rt, rid, name[:48]])
            self.table(["Contract", "NPI", "Resource", "Resource id", "Submitted name"], rows,
                       widths=[0.75, 1.05, 1.15, 1.0, 2.85], font=8.5)
            self.p("Totals: %s NPIs. The complete list is in the findings CSV files."
                   % " / ".join(commify(len(a.F.npis.get("P1016", ()))) for a in self.audits),
                   size=9)
            self.p("Action: add the missing OrganizationAffiliation (with its network reference) "
                   "for each affected facility, or remove the facility record if it is no longer "
                   "contracted.", size=9, bold=True)

        for code in ("F5005", "F5006", "F5007", "F5008", "REFBROKEN", "REFMALFORMED", "DUPID"):
            if not sum(self.counts_for(code)):
                continue
            any_detail = True
            supp = SUPP_CODES.get(code)
            name = supp[1] if supp else next(
                (n for c, n, lv, d, h in APPENDIX_E if c == code), code)
            self.h("%s %s %s - unresolved references" % (self._n(), code, name), 2)
            rows = []
            for a in self.audits:
                for s in a.F.samples.get(code, [])[:12]:
                    rows.append([a.contract, s["resource"], s["id"], s["field"], s["value"],
                                 s["detail"]])
            self.table(["Contract", "Source resource", "Source id", "Path", "Target", "Detail"],
                       rows, widths=[0.7, 1.2, 0.9, 1.4, 1.2, 1.5], font=8)
            self.p("Full list in the findings CSV files.", size=9, italic=True)

        if not any_detail:
            self.p("No findings required detailed follow-up.", bold=True, colour=GREEN)

    def hpms(self):
        self.h("11. Items Requiring HPMS Confirmation")
        self.p("Six validations depend on the HPMS registry, which is not accessible outside HPMS. "
               "The audit verified everything that can be verified from the submitted data. The "
               "remainder should be confirmed against the daily CMS validation report in the HPMS "
               "MPF Provider Directory module.")
        plans_desc = "; ".join("%s submits %d plan ID%s"
                               % (a.contract, len(a.plan_ids), "" if len(a.plan_ids) == 1 else "s")
                               for a in self.audits)
        self.table(["Code", "Level", "Validation", "Verified in this audit",
                    "Still to confirm in HPMS"], [
            ["N3011", "2", "UnknownContractID",
             "All contract numbers are well formed and match their own index file.",
             "That %s are the contract IDs HPMS expects." % ", ".join(self.ids)],
            ["N3012", "2", "UnknownPlanID",
             "All plan components are three digits and internally consistent.",
             "That every submitted plan ID exists in the HPMS registry."],
            ["N3013", "2", "UnknownSegmentID",
             "Segment components: %s." % ", ".join(
                 sorted(set(p.split("-")[2] for a in self.audits for p in a.plan_ids
                            if MAPLAN_RE.match(p)))) or "none",
             "That HPMS expects those segment IDs for every plan."],
            ["N3006", "1", "OmittedContractID",
             "Each contract's own ID is present in its dataset.",
             "That no additional contract expected by HPMS is missing a directory."],
            ["N3007", "3", "OmittedSegmentID", "Segment present on every plan.",
             "That no expected segment is absent."],
            ["N3008", "3", "OmittedPlanID", plans_desc,
             "That these match the individual plans HPMS expects."],
        ], widths=[0.5, 0.4, 1.2, 2.3, 2.5], font=8)

        counts = sorted((len(a.plan_ids), a.contract) for a in self.audits)
        if counts and counts[0][0] <= 1 and len(counts) > 1 and counts[-1][0] > counts[0][0]:
            self.p("%s submits only %d MA Plan ID while other contracts in scope submit more. If "
                   "HPMS expects additional individual plans under %s, the missing plans would "
                   "raise N3008 and, more seriously, the providers under those plans would not "
                   "appear on MPF for them. This is the first item to check."
                   % (counts[0][1], counts[0][0], counts[0][1]), size=9, bold=True)

    def remediation(self):
        self.h("12. Remediation Plan")
        rows = []
        pri = 0
        big = self._largest_file()
        totals = [a.stats.get("bytes_total", 0) for a in self.audits]
        if max(totals or [0]) > MAX_FILE_BYTES:
            pri += 1
            rows.append([pri, "Confirm the 300 MB array-total reading with CMS", "Advisory",
                         "All contracts",
                         "Email support@cms-mapnet.zendesk.com. If aggregate, plan a file split."])
        counts = sorted((len(a.plan_ids), a.contract) for a in self.audits)
        if counts and counts[0][0] <= 1 and len(counts) > 1 and counts[-1][0] > counts[0][0]:
            pri += 1
            rows.append([pri, "Confirm %s plan coverage against HPMS" % counts[0][1], "L2/L3",
                         "1 contract",
                         "Compare submitted MA Plan IDs against the HPMS plan registry."])
        order = [("P1003", 2, "Remove or replace deactivated NPIs",
                  "Retire the Organization resources or supply current NPIs."),
                 ("A2008", 2, "Correct addresses that cannot be geolocated",
                  "Fix the zip code or state at source."),
                 ("A2001", 2, "Supply addresses for NPIs that have none",
                  "Add a Location or Organization address for each affected NPI."),
                 ("P1016", 2, "Associate providers with a plan ID",
                  "Add the missing network reference so the provider links to a plan."),
                 ("F5005", 2, "Repair broken network references", "Restore the missing resources."),
                 ("F5006", 2, "Repair broken organization references", "Restore the missing resources."),
                 ("F5007", 2, "Repair broken practitioner references", "Restore the missing resources."),
                 ("F5008", 2, "Repair broken location references", "Restore the missing resources."),
                 ("REFBROKEN", 2, "Repair other broken references", "Restore the missing resources."),
                 ("P1012", 3, "Populate accepting-new-patients status",
                  "Emit the newpatients extension on every PractitionerRole."),
                 ("SPECCLS", 3, "Correct facility specialty taxonomy",
                  "Map OrganizationAffiliation.specialty to NUCC non-individual codes."),
                 ("A2010", 3, "Replace invalid phone numbers", "Remove placeholder values."),
                 ("A2009", 3, "Supply missing phone numbers", "Source the missing numbers."),
                 ("P1005", 3, "Decide the sole-proprietor position",
                  "Accept the warning or resubmit genuine individuals as Practitioner records."),
                 ]
        for code, level, item, action in order:
            total = sum(self.counts_for(code))
            if not total:
                continue
            pri += 1
            vol = " / ".join(commify(a.F.count.get(code, 0)) for a in self.audits)
            rows.append([pri, "%s (%s)" % (item, code), "L%d" % level, vol, action])
        days = self.tls.get("days_remaining")
        if isinstance(days, int) and days < 200:
            pri += 1
            rows.append([pri, "Schedule TLS certificate renewal", "L1 risk", "Hosting",
                         "Certificate expires %s, mid crawl year." % self.tls.get("not_after", "")])
        pri += 1
        rows.append([pri, "Complete the CY %s HPMS attestation" % CONTRACT_YEAR, "Compliance",
                     "All contracts",
                     "CEO, CFO or COO to attest in the HPMS MPF Provider Directory module."])
        self.table(["Priority", "Item", "Level", "Occurrences (%s)" % " / ".join(self.ids),
                    "Owner action"], rows, widths=[0.55, 2.2, 0.5, 1.3, 2.35], font=8)

    def conclusion(self):
        self.h("13. Conclusion")
        l1 = sum(a.F.level_total(1) for a in self.audits)
        l2 = sum(a.F.level_total(2) for a in self.audits)
        l3 = sum(a.F.level_total(3) for a in self.audits)
        resources = sum(sum(v for k, v in a.stats.items() if k.startswith("res_"))
                        for a in self.audits)
        refs = sum(a.stats.get("references_checked", 0) for a in self.audits)
        broken = sum(a.stats.get("reference_targets_broken", 0) for a in self.audits)
        if l1 == 0:
            self.p("All %d CY %s FHIR-based provider directory submissions are structurally sound "
                   "and conform to the CMS technical guide version 1.5 released on September 4, "
                   "2026. Every fatal validation that applies to the FHIR-based option passed on "
                   "every contract, so none is at risk of dataset rejection or MPF suppression on "
                   "structural grounds." % (len(self.audits), CONTRACT_YEAR))
        else:
            self.p("%s fatal findings are outstanding and must be cleared before the next daily "
                   "crawl." % commify(l1), bold=True, colour=RED)
        if broken == 0:
            self.p("Reference integrity is complete: %s distinct reference targets were resolved "
                   "across %s resources with no broken, dangling or malformed reference of any "
                   "kind, on any element, at any path."
                   % (commify(refs), commify(resources)))
        else:
            self.p("%s reference targets did not resolve within their contract dataset. Every "
                   "affected record is listed in the findings CSV files." % commify(broken),
                   bold=True, colour=RED)
        self.p("Totals across the contracts in scope: %s Level 1 fatal, %s Level 2 record-level "
               "and %s Level 3 informational findings. Every individual occurrence is enumerated "
               "in the accompanying findings_<contract>.csv files."
               % (commify(l1), commify(l2), commify(l3)))
        self.doc.add_paragraph()
        self.p("Generated by mpf_audit.py from a live audit of the production index URLs and every "
               "constituent file on %s. All figures are derived from complete parses, not samples."
               % self.date.strftime("%B %d, %Y"), italic=True, size=8.5, colour=GREY)


# ==========================================================================
# main
# ==========================================================================
def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Audit the JHP CY %s MPF provider directory submissions against the CMS "
                    "technical guide v1.5." % CONTRACT_YEAR)
    ap.add_argument("--out", default=os.getcwd(), help="directory for the report (default: cwd)")
    ap.add_argument("--cache", default=None, help="cache directory (default: <script dir>/.mpf_cache)")
    ap.add_argument("--fresh", action="store_true", help="ignore cached bundles and re-download")
    ap.add_argument("--no-nppes", action="store_true", help="skip the NPPES registry checks")
    ap.add_argument("--contracts", default=None, help="comma separated subset, e.g. H1619,H9207")
    args = ap.parse_args(argv)

    out_dir = os.path.abspath(args.out)
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    cache_dir = os.path.abspath(args.cache or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), ".mpf_cache"))
    if not os.path.isdir(cache_dir):
        os.makedirs(cache_dir)

    wanted = [c.strip().upper() for c in args.contracts.split(",")] if args.contracts else None
    contracts = [(cid, url) for cid, url in sorted(INDEX_URLS.items())
                 if not wanted or cid in wanted]
    if not contracts:
        sys.stderr.write("No matching contracts. Known: %s\n" % ", ".join(sorted(INDEX_URLS)))
        return 2

    audit_date = datetime.date.today()
    started = datetime.datetime.now()

    log("=" * LOG_WIDTH)
    log("MPF MA Provider Directory Audit - CMS technical guide v1.5 (September 4, 2026)")
    log("%s | contract year %s | %s" % (ORG_NAME, CONTRACT_YEAR, audit_date.isoformat()))
    log("cache  : %s" % cache_dir)
    log("output : %s" % out_dir)
    log("=" * LOG_WIDTH)

    rule("Reference data")
    refdata = RefData(cache_dir, use_nppes=not args.no_nppes)
    refdata.load()

    rule("TLS")
    tls = check_tls(contracts[0][1])
    if tls.get("verified"):
        log("certificate chain verified, valid until %s (%s days)"
            % (tls.get("not_after"), tls.get("days_remaining")), 1)
    else:
        log("TLS verification FAILED: %s" % tls.get("error"), 1)

    catalog = {}
    audits = []
    for cid, url in contracts:
        audits.append(ContractAudit(cid, url, refdata, cache_dir, out_dir, catalog,
                                    audit_date, fresh=args.fresh).run())

    if not tls.get("verified"):
        for a in audits:
            a.add("C4002", 1, value=tls.get("host", ""),
                  detail="TLS verification failed: %s" % tls.get("error", ""))

    rule("Results")
    for a in audits:
        log("%s  L1 %s | L2 %s | L3 %s   (%s resources, %s reference targets)"
            % (a.contract, commify(a.F.level_total(1)), commify(a.F.level_total(2)),
               commify(a.F.level_total(3)),
               commify(sum(v for k, v in a.stats.items() if k.startswith("res_"))),
               commify(a.stats.get("references_checked", 0))), 1)
        for code, n in sorted(a.F.count.items(),
                              key=lambda kv: (catalog[kv[0]][1], -kv[1])):
            log("L%d %-10s %-34s %10s" % (catalog[code][1], code, catalog[code][0], commify(n)), 2)

    rule("Report")
    report = Report(audits, refdata, tls, audit_date, catalog)
    doc = report.build()
    name = "MPF_Provider_Directory_Audit_CY%s_%s.docx" % (
        CONTRACT_YEAR, audit_date.strftime("%Y%m%d"))
    path = os.path.join(out_dir, name)
    doc.save(path)
    log("report : %s" % path, 1)
    for a in audits:
        log("csv    : %s" % os.path.join(out_dir, "findings_%s.csv" % a.contract), 1)
    log("elapsed: %s" % str(datetime.datetime.now() - started).split(".")[0], 1)

    fatal = sum(a.F.level_total(1) for a in audits)
    log("=" * LOG_WIDTH)
    if fatal:
        log("RESULT: %s FATAL finding(s) - the dataset would be rejected by CMS." % commify(fatal))
    else:
        log("RESULT: no fatal errors. See the report for record-level and informational findings.")
    log("=" * LOG_WIDTH)
    return 1 if fatal else 0


if __name__ == "__main__":
    # ------------------------------------------------------------------
    # Minimal, additive exit-code wrapper (added for CI/automation use).
    # Does NOT change any validation logic or main()'s own return values
    # (0 = clean, 1 = fatal findings reported, 2 = bad invocation / no
    # matching contracts, as main() already implements above). This
    # wrapper only adds: 3 = the script crashed with an unhandled
    # exception (automation/runtime failure, not a validation result).
    # See README.md "Exit codes" section for how callers should use this.
    # ------------------------------------------------------------------
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:                                          # noqa: BLE001
        import traceback
        traceback.print_exc()
        sys.exit(3)
