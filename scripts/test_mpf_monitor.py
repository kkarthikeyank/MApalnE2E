#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
test_mpf_monitor.py -- quick, dependency-free unit test for the pure
change-detection comparison logic in mpf_monitor.py (spec item 17).

This does NOT hit the network. It calls check_contracts() directly with a
monkeypatched fetch_index() so we can prove the state-comparison logic
itself (UNCHANGED / CHANGED / bootstrap-CHANGED / FETCH_ERROR) is correct,
without needing pytest or live CMS servers.

Run with:  python scripts/test_mpf_monitor.py
Exits 0 and prints "ALL TESTS PASSED" on success, exits 1 and prints the
first failure otherwise.
"""

from __future__ import print_function

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mpf_monitor as M


def _fake_fetch(fixture):
    """Build a fetch_index replacement backed by a fixed {contract: (idx, err)} map."""
    def _fetch(url, timeout=60):
        for cid, u in M.INDEX_URLS.items():
            if u == url:
                return fixture.get(cid, (None, "no fixture for %s" % cid))
        return None, "unknown url in test fixture"
    return _fetch


def run():
    failures = []

    def check(label, cond):
        if not cond:
            failures.append(label)

    orig_fetch = M.fetch_index

    # ---- Case 1: no-change case -----------------------------------------
    # index.json's last_updated is identical to what's stored -> UNCHANGED
    state = {"H1619": "2026-09-01T00:00:00Z"}
    fixture = {"H1619": ({"last_updated": "2026-09-01T00:00:00Z", "provider_urls": ["x"]}, "")}
    M.fetch_index = _fake_fetch(fixture)
    try:
        results = M.check_contracts(["H1619"], state)
    finally:
        M.fetch_index = orig_fetch
    r = results[0]
    check("no-change case yields UNCHANGED", r["status"] == "UNCHANGED")
    check("no-change case preserves previous", r["previous"] == "2026-09-01T00:00:00Z")
    check("no-change case reports current", r["current"] == "2026-09-01T00:00:00Z")

    # ---- Case 2: changed case --------------------------------------------
    # index.json's last_updated differs from stored -> CHANGED
    state = {"H1619": "2026-09-01T00:00:00Z"}
    fixture = {"H1619": ({"last_updated": "2026-09-15T12:30:00Z", "provider_urls": ["x"]}, "")}
    M.fetch_index = _fake_fetch(fixture)
    try:
        results = M.check_contracts(["H1619"], state)
    finally:
        M.fetch_index = orig_fetch
    r = results[0]
    check("changed case yields CHANGED", r["status"] == "CHANGED")
    check("changed case reports new current", r["current"] == "2026-09-15T12:30:00Z")

    # ---- Case 3: bootstrap case -------------------------------------------
    # contract absent from state entirely -> treated as CHANGED (first run)
    state = {}
    fixture = {"H3124": ({"last_updated": "2026-09-01T00:00:00Z", "provider_urls": ["x"]}, "")}
    M.fetch_index = _fake_fetch(fixture)
    try:
        results = M.check_contracts(["H3124"], state)
    finally:
        M.fetch_index = orig_fetch
    r = results[0]
    check("bootstrap case yields CHANGED", r["status"] == "CHANGED")
    check("bootstrap case previous is None", r["previous"] is None)

    # ---- Case 4: fetch error -----------------------------------------------
    state = {"H5826": "2026-09-01T00:00:00Z"}
    fixture = {"H5826": (None, "HTTP 503")}
    M.fetch_index = _fake_fetch(fixture)
    try:
        results = M.check_contracts(["H5826"], state)
    finally:
        M.fetch_index = orig_fetch
    r = results[0]
    check("fetch error yields FETCH_ERROR", r["status"] == "FETCH_ERROR")
    check("fetch error does not clobber stored state", r["previous"] == "2026-09-01T00:00:00Z")

    # ---- Case 5: contract isolation ----------------------------------------
    # one contract errors, another is unchanged -- must not affect each other
    state = {"H1619": "A", "H3124": "B"}
    fixture = {
        "H1619": (None, "timeout"),
        "H3124": ({"last_updated": "B", "provider_urls": ["x"]}, ""),
    }
    M.fetch_index = _fake_fetch(fixture)
    try:
        results = M.check_contracts(["H1619", "H3124"], state)
    finally:
        M.fetch_index = orig_fetch
    by_id = dict((r["contract"], r) for r in results)
    check("isolation: H1619 errors independently", by_id["H1619"]["status"] == "FETCH_ERROR")
    check("isolation: H3124 still evaluated correctly", by_id["H3124"]["status"] == "UNCHANGED")

    if failures:
        print("FAILED:")
        for f in failures:
            print("  - %s" % f)
        return 1
    print("ALL TESTS PASSED (%d assertions)" % 12)
    return 0


if __name__ == "__main__":
    sys.exit(run())
