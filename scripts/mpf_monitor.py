#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
mpf_monitor.py -- change detection ONLY.

This script contains NO MPF validation logic. It exists purely to answer one
question per contract: "has the CMS-hosted index.json for this contract
changed since the last successful mpf_audit.py run?" All actual validation
(the CMS Appendix E rules, FHIR parsing, NPPES checks, report/CSV
generation) lives in mpf_audit.py, which remains the single source of truth.

What it does:
  1. Fetches each contract's index.json (a small JSON document -- see
     mpf_audit.py's INDEX_URLS / provider_urls handling; it is not the large
     provider data bundles) and reads its "last_updated" field.
  2. Compares that value against the value stored in state/mpf_state.json
     for the same contract.
  3. Reports, per contract: UNCHANGED (skip), CHANGED (run), or
     FETCH_ERROR (automation failure -- could not reach/parse the index).
  4. A contract absent from state/mpf_state.json is treated as CHANGED, so
     that the very first run performs an initial validation and populates
     state ("Option B" bootstrap behaviour -- see README.md).

Usage:
    python scripts/mpf_monitor.py --state state/mpf_state.json \
        --contracts H1619,H3124,H5826,H9207 [--json]

Exit code is always 0 unless invoked incorrectly; per-contract fetch errors
are reported in the output, not via process exit code, since one contract's
network hiccup must not block the others (contract isolation, spec item 3).
"""

from __future__ import print_function

import argparse
import json
import os
import sys

try:
    from urllib.request import Request, urlopen
    from urllib.error import HTTPError, URLError
except ImportError:                                         # py2 fallback
    from urllib2 import Request, urlopen, HTTPError, URLError

# Kept in sync with mpf_audit.py's INDEX_URLS. Not imported directly so this
# script has zero dependency on mpf_audit.py's (large) import chain
# (ijson / python-docx / openpyxl) -- it only needs to be fast and light.
INDEX_URLS = {
    "H5826": "https://medicare-advantage-plan-finder-provider-directory.interop.chpw.org/h5826/2027/index.json",
    "H1619": "https://medicare-advantage-plan-finder-provider-directory.jeffersonhealthplans.com/h1619/2027/index.json",
    "H3124": "https://medicare-advantage-plan-finder-provider-directory.jeffersonhealthplans.com/h3124/2027/index.json",
    "H9207": "https://medicare-advantage-plan-finder-provider-directory.jeffersonhealthplans.com/h9207/2027/index.json",
}

USER_AGENT = "JHP-MPF-Change-Monitor/1.0"


def fetch_index(url, timeout=60):
    """Fetch and parse an index.json. Returns (dict_or_None, error_str)."""
    req = Request(url)
    req.add_header("User-Agent", USER_AGENT)
    try:
        resp = urlopen(req, timeout=timeout)
        raw = resp.read()
        resp.close()
    except HTTPError as e:
        return None, "HTTP %s" % e.code
    except URLError as e:
        return None, "URLError: %s" % e.reason
    except Exception as e:                                   # noqa: BLE001
        return None, "%s: %s" % (type(e).__name__, e)
    try:
        return json.loads(raw.decode("utf-8", "replace")), ""
    except ValueError as e:
        return None, "invalid JSON: %s" % e


def load_state(path):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:                                          # noqa: BLE001
        return {}


def check_contracts(contracts, state):
    """Return a list of dicts describing the status of each contract.

    status is one of: CHANGED, UNCHANGED, FETCH_ERROR
    """
    results = []
    for cid in contracts:
        url = INDEX_URLS.get(cid)
        if not url:
            results.append({"contract": cid, "status": "FETCH_ERROR",
                            "error": "unknown contract id", "current": None,
                            "previous": state.get(cid)})
            continue
        idx, err = fetch_index(url)
        previous = state.get(cid)
        if err or idx is None:
            results.append({"contract": cid, "status": "FETCH_ERROR",
                            "error": err, "current": None, "previous": previous})
            continue
        current = idx.get("last_updated") if isinstance(idx, dict) else None
        if current is None:
            results.append({"contract": cid, "status": "FETCH_ERROR",
                            "error": "index.json has no 'last_updated' field",
                            "current": None, "previous": previous})
            continue
        if cid not in state:
            status = "CHANGED"          # bootstrap: never validated before
        elif previous == current:
            status = "UNCHANGED"
        else:
            status = "CHANGED"
        results.append({"contract": cid, "status": status,
                        "current": current, "previous": previous, "error": ""})
    return results


def main(argv=None):
    ap = argparse.ArgumentParser(description="MPF change-detection only (no validation logic).")
    ap.add_argument("--state", default="state/mpf_state.json", help="path to state JSON")
    ap.add_argument("--contracts", default="H1619,H3124,H5826,H9207",
                     help="comma separated contract IDs to check")
    ap.add_argument("--json", action="store_true", help="emit machine-readable JSON to stdout")
    args = ap.parse_args(argv)

    contracts = [c.strip().upper() for c in args.contracts.split(",") if c.strip()]
    state = load_state(args.state)
    results = check_contracts(contracts, state)

    if args.json:
        print(json.dumps(results, indent=2))
    else:
        for r in results:
            print("%-8s %-14s prev=%s current=%s %s"
                  % (r["contract"], r["status"], r.get("previous"), r.get("current"),
                     ("(%s)" % r["error"]) if r.get("error") else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
