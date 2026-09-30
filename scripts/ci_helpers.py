#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ci_helpers.py -- small CLI utilities used only by the GitHub Actions
workflow to avoid embedding multi-line Python heredocs inside YAML (which
is fragile to indent correctly). No validation logic lives here either;
these are plumbing helpers around mpf_monitor.py's output and
state/mpf_state.json.

Subcommands:
  candidates <selection>
      Print "to_run=<csv>" for a manual-mode selection ("ALL" or a single
      contract id).

  monitor-outputs <monitor_result.json>
      Read mpf_monitor.py's --json output and print "to_run=<csv>" and
      "errored=<csv>" (contracts whose CHANGED / FETCH_ERROR status apply).

  previous-state <state.json> <contract>
      Print the stored last_updated value for a contract (empty if none).
"""

from __future__ import print_function

import json
import sys

ALL_CONTRACTS = ["H1619", "H3124", "H5826", "H9207"]


def cmd_candidates(argv):
    sel = argv[0] if argv else "ALL"
    sel = sel.strip().upper()
    if sel in ("ALL", ""):
        print("to_run=" + ",".join(ALL_CONTRACTS))
    else:
        print("to_run=" + sel)
    return 0


def cmd_monitor_outputs(argv):
    path = argv[0]
    with open(path, "r", encoding="utf-8") as f:
        results = json.load(f)
    changed = [r["contract"] for r in results if r.get("status") == "CHANGED"]
    errored = [r["contract"] for r in results if r.get("status") == "FETCH_ERROR"]
    print("to_run=" + ",".join(changed))
    print("errored=" + ",".join(errored))
    return 0


def cmd_previous_state(argv):
    state_path, contract = argv[0], argv[1]
    try:
        with open(state_path, "r", encoding="utf-8") as f:
            state = json.load(f)
    except Exception:                                          # noqa: BLE001
        state = {}
    print(state.get(contract, "") or "")
    return 0


def cmd_update_state(argv):
    """update-state <state.json> <"H1619=2026...;H3124=2026...;">"""
    state_path, updates = argv[0], (argv[1] if len(argv) > 1 else "")
    try:
        with open(state_path, "r", encoding="utf-8") as f:
            state = json.load(f)
    except Exception:                                          # noqa: BLE001
        state = {}
    for pair in updates.split(";"):
        pair = pair.strip()
        if not pair or "=" not in pair:
            continue
        cid, lu = pair.split("=", 1)
        state[cid] = lu
    with open(state_path, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, sort_keys=True)
    return 0


COMMANDS = {
    "candidates": cmd_candidates,
    "monitor-outputs": cmd_monitor_outputs,
    "previous-state": cmd_previous_state,
    "update-state": cmd_update_state,
}


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    if not argv or argv[0] not in COMMANDS:
        sys.stderr.write("usage: ci_helpers.py <%s> ...\n" % "|".join(COMMANDS))
        return 2
    return COMMANDS[argv[0]](argv[1:]) or 0


if __name__ == "__main__":
    sys.exit(main())
