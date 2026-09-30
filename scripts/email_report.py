#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
email_report.py -- email formatting/sending ONLY. No validation logic.

Sends one email per contract run, using smtplib and GitHub Actions secrets:
    EMAIL_USERNAME, EMAIL_PASSWORD, EMAIL_TO,
    EMAIL_SMTP_HOST (default: smtp.office365.com), EMAIL_SMTP_PORT (default: 587)

STARTTLS is used by default (matches the office365 default). Secrets are
read only from the environment and are never printed/logged.

Four message "kinds" (spec sections 13-16):
    updated-success      automatic run, contract changed, validated clean (exit 0)
    updated-findings      automatic run, contract changed, findings reported (exit 1)
    automation-crash      automation itself failed (download/crash/report-gen)
    manual-run             manual workflow_dispatch run (any script outcome)

The findings summary is derived from the real findings_<CONTRACT>.csv that
mpf_audit.py wrote (columns: Level,ErrorCode,ValidationName,ContractID,
ResourceType,ResourceID,NPI,Field,Value,Detail -- see Findings.__init__ in
mpf_audit.py). Nothing here is hardcoded PASS/FAIL text.

Usage:
    python scripts/email_report.py \
        --kind updated-findings \
        --contract H1619 --plan-year 2027 --run-mode automatic \
        --previous-last-updated "2026-09-01T00:00:00Z" \
        --new-last-updated "2026-09-15T12:30:00Z" \
        --exit-code 1 \
        --report-path reports/MPF_Provider_Directory_Audit_CY2027_20260930.docx \
        --csv-path reports/findings_H1619.csv \
        [--error-detail "..."] [--dry-run]
"""

from __future__ import print_function

import argparse
import collections
import csv
import os
import smtplib
import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def summarize_findings_csv(csv_path):
    """Return (counts_by_level dict, total_rows, top_error_codes list) or None."""
    if not csv_path or not os.path.exists(csv_path):
        return None
    by_level = collections.Counter()
    by_code = collections.Counter()
    total = 0
    try:
        with open(csv_path, "r", encoding="utf-8", errors="replace", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                total += 1
                by_level[row.get("Level", "?")] += 1
                by_code[row.get("ErrorCode", "?")] += 1
    except Exception as e:                                    # noqa: BLE001
        return {"error": "could not read findings CSV: %s" % e}
    top = by_code.most_common(8)
    return {"by_level": dict(by_level), "total": total, "top_codes": top}


SUBJECTS = {
    "updated-success": "[MPF] {contract} CY{plan_year} -- updated, validated clean",
    "updated-findings": "[MPF] {contract} CY{plan_year} -- updated, {n} finding(s) reported",
    "automation-crash": "[MPF] {contract} CY{plan_year} -- AUTOMATION FAILURE",
    "manual-run": "[MPF] {contract} CY{plan_year} -- manual run ({outcome})",
}


def build_message(args, summary):
    kind = args.kind
    lines = []
    lines.append("MPF Provider Directory Validation")
    lines.append("Contract:       %s" % args.contract)
    lines.append("Plan year:      CY%s" % args.plan_year)
    lines.append("Run mode:       %s" % args.run_mode)
    lines.append("Previous last_updated: %s" % (args.previous_last_updated or "(none -- first run)"))
    lines.append("New last_updated:      %s" % (args.new_last_updated or "(unknown)"))
    lines.append("mpf_audit.py exit code: %s" % args.exit_code)
    lines.append("")

    if kind == "automation-crash":
        lines.append("AUTOMATION FAILURE -- this is NOT a data validation result.")
        lines.append("The validation script/download/report-generation step itself did not")
        lines.append("complete, so no findings report was produced for this contract.")
        lines.append("Stored state for this contract was NOT advanced; the next scheduled")
        lines.append("run will retry it automatically.")
        lines.append("")
        if args.error_detail:
            lines.append("Details:")
            lines.append(args.error_detail)
        outcome = "automation failure"
    else:
        if summary and "error" not in summary:
            lines.append("Findings summary (from %s):" % os.path.basename(args.csv_path or ""))
            lines.append("  Total finding rows: %s" % summary["total"])
            for level in sorted(summary["by_level"], key=lambda k: (k == "?", k)):
                lines.append("  Level %s: %s" % (level, summary["by_level"][level]))
            if summary["top_codes"]:
                lines.append("  Top error codes:")
                for code, n in summary["top_codes"]:
                    lines.append("    %-10s %s" % (code, n))
            outcome = "findings reported" if summary["total"] else "no findings"
        elif summary and "error" in summary:
            lines.append("(Could not summarize findings CSV: %s)" % summary["error"])
            outcome = "unknown"
        else:
            lines.append("No findings CSV was available to summarize.")
            outcome = "clean" if args.exit_code == "0" else "unknown"
        lines.append("")
        lines.append("Report file: %s" % (os.path.basename(args.report_path) if args.report_path else "(not generated)"))

    subject_tmpl = SUBJECTS.get(kind, "[MPF] {contract} CY{plan_year} -- run complete")
    n = summary["total"] if (summary and "error" not in summary) else "?"
    subject = subject_tmpl.format(contract=args.contract, plan_year=args.plan_year,
                                   n=n, outcome=outcome)
    return subject, "\n".join(lines)


def send_email(subject, body, attachments):
    host = os.environ.get("EMAIL_SMTP_HOST") or "smtp.office365.com"
    port = int(os.environ.get("EMAIL_SMTP_PORT") or "587")
    username = os.environ.get("EMAIL_USERNAME")
    password = os.environ.get("EMAIL_PASSWORD")
    to_addr = os.environ.get("EMAIL_TO")

    missing = [n for n, v in (("EMAIL_USERNAME", username), ("EMAIL_PASSWORD", password),
                              ("EMAIL_TO", to_addr)) if not v]
    if missing:
        sys.stderr.write("email_report: missing required secret(s): %s -- not sending.\n"
                         % ", ".join(missing))
        return 1

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = username
    msg["To"] = to_addr
    msg.attach(MIMEText(body, "plain"))

    for path in attachments:
        if not path or not os.path.exists(path):
            continue
        try:
            with open(path, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(path))
            part["Content-Disposition"] = 'attachment; filename="%s"' % os.path.basename(path)
            msg.attach(part)
        except Exception as e:                                # noqa: BLE001
            sys.stderr.write("email_report: could not attach %s: %s\n" % (path, e))

    try:
        with smtplib.SMTP(host, port, timeout=60) as server:
            server.starttls()
            server.login(username, password)
            server.sendmail(username, [a.strip() for a in to_addr.split(",") if a.strip()],
                            msg.as_string())
    except Exception as e:                                    # noqa: BLE001
        # Never print secrets; smtplib exceptions do not include the password.
        sys.stderr.write("email_report: send failed: %s\n" % e)
        return 1
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Send an MPF validation result email (formatting/sending only).")
    ap.add_argument("--kind", required=True,
                     choices=["updated-success", "updated-findings", "automation-crash", "manual-run"])
    ap.add_argument("--contract", required=True)
    ap.add_argument("--plan-year", default="2027")
    ap.add_argument("--run-mode", default="automatic", choices=["automatic", "manual"])
    ap.add_argument("--previous-last-updated", default="")
    ap.add_argument("--new-last-updated", default="")
    ap.add_argument("--exit-code", default="")
    ap.add_argument("--report-path", default="")
    ap.add_argument("--csv-path", default="")
    ap.add_argument("--error-detail", default="")
    ap.add_argument("--dry-run", action="store_true", help="print the message instead of sending")
    args = ap.parse_args(argv)

    summary = summarize_findings_csv(args.csv_path) if args.kind != "automation-crash" else None
    subject, body = build_message(args, summary)

    if args.dry_run:
        print("SUBJECT: %s" % subject)
        print(body)
        return 0

    attachments = [p for p in (args.report_path, args.csv_path) if p]
    return send_email(subject, body, attachments)


if __name__ == "__main__":
    sys.exit(main())
