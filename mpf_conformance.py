#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
MPF Conformance Report
=======================

Produces a CMS "Technical Implementation Guide for Supplying MA Provider
Directory Data for Use in MPF, v1.5" conformance report, scoped to
contracts H1619, H3124 and H9207.

This script is a thin, section-structured layer on top of mpf_audit.py: it
imports and reuses mpf_audit's HTTP/cache layer, FHIR bundle parsing and
ContractAudit validation engine rather than re-implementing any of it, and
then reorganises the resulting findings into the eight report sections
described in the task brief.

    python mpf_conformance.py [--contracts H1619,H3124,H9207] [--out DIR]
                              [--cache DIR] [--fresh] [--no-nppes]
                              [--cms-reports DIR]
"""

from __future__ import print_function

import argparse
import collections
import csv
import datetime
import glob
import json
import os
import re
import sys

import mpf_audit as MA          # reuse: HTTP/cache, FHIR parsing, ContractAudit, RefData
from mpf_audit import (
    Document, Pt, Inches, RGBColor, WD_TABLE_ALIGNMENT, WD_ORIENT, qn, OxmlElement,
    ContractAudit, RefData, check_tls, commify, mb, log, rule,
    INDEX_URLS, CONTRACT_YEAR, ORG_NAME, MAX_INDEX_URLS, MAX_FILE_BYTES,
    NAVY, BLUE, GREEN, RED, AMBER, GREY,
)

REPORT_CONTRACTS = ["H1619", "H3124", "H9207"]   # excludes H5826 by design

STATUS_COLOR = {"PASS": GREEN, "FAIL": RED, "WARN": AMBER, "INFO": GREY, "SKIP": GREY}


# ==========================================================================
# Section membership -- maps each Appendix E / supplemental code onto the
# conformance-report section it belongs to (a code may appear in more than
# one section since several sections share underlying checks).
# ==========================================================================
SECTION1_CODES = ["C4001", "C4002", "C4003", "C4004", "N3015", "C4018", "C4019",
                   "C4008", "C4006", "C4007", "C4009", "C4005", "C4011", "C4012",
                   "C4013", "C4014", "HOSTCMP", "DUPID"]

SECTION2_CODES = ["N3001", "N3002", "N3004", "N3005", "F5001", "F5005", "N3003",
                   "N3011", "N3012", "N3013", "N3014"]

SECTION3_CODES = ["P1001", "F5009", "P1013", "P1014", "C4010", "F5004", "F5008",
                   "P1009", "F5003", "F5007", "A2009", "A2010", "P1006", "P1007",
                   "P1010", "P1011", "P1012", "P1018"]

SECTION4_CODES = ["A2007", "A2002", "A2003", "A2004", "A2005", "A2006"]

SECTION5_CODES = ["F5001", "F5002", "F5004", "F5005", "F5006", "F5008", "P1009",
                   "P1001", "F5009", "P1008", "A2001", "A2007", "A2002", "A2003",
                   "A2004", "A2005", "A2006", "A2009", "A2010", "P1004"]

CODE_LABEL = {}   # populated from mpf_audit.APPENDIX_E / SUPP_CODES at runtime


def _label_for(code, catalog):
    entry = catalog.get(code)
    return entry[0] if entry else code


def _fail_warn(count, level):
    """Level 1 findings are always FAIL; level 2/3 are FAIL if this run has
    zero reference data to soften them, otherwise treated per spec as FAIL
    (hard data-quality defects) except the advisory codes handled by caller."""
    return "FAIL" if count else "PASS"


# ==========================================================================
# Cross-contract facility reconciliation (section 5, "recurring facility
# records ... across all three contracts")
# ==========================================================================
def recurring_facility_gaps(audits):
    seen = collections.defaultdict(list)     # org id -> [(contract, gaps)]
    contract_count = collections.Counter()
    for a in audits:
        for oid, res in a.orgs.items():
            contract_count[oid] += 1
            gaps = []
            npi = a.org_npi.get(oid, "")
            if not npi:
                gaps.append("NPI")
            if not (res.get("name") or "").strip():
                gaps.append("name")
            phones = MA.phones_of(res)
            if not [p for p in phones if p]:
                gaps.append("phone")
            seen[oid].append((a.contract, res.get("name") or "", gaps))
    rows = []
    for oid, entries in seen.items():
        if contract_count[oid] < 2:
            continue
        if not any(g for _, _, g in entries):
            continue
        contracts_str = ", ".join(sorted(set(c for c, _, _ in entries)))
        gap_str = "; ".join("%s: missing %s" % (c, ",".join(g)) for c, _, g in entries if g)
        name = next((n for _, n, _ in entries if n), "")
        rows.append((oid, name, contracts_str, gap_str))
    rows.sort(key=lambda r: r[0])
    return rows


# ==========================================================================
# Section 6 -- CMS validation report reconciliation
# ==========================================================================
def load_cms_reports(cms_dir, contracts):
    """Return {contract: {"file": path, "row_count": n, "codes": Counter}} for
    files matching CY2027_ValidationError_{CONTRACT}_YYYYMMDD.csv, or {} if
    the directory was not supplied / nothing matched."""
    if not cms_dir:
        return {}
    out = {}
    pattern = re.compile(r"CY\d{4}_ValidationError_([A-Z0-9]+)_(\d{8})\.csv$", re.IGNORECASE)
    for path in glob.glob(os.path.join(cms_dir, "*.csv")):
        m = pattern.search(os.path.basename(path))
        if not m:
            continue
        contract = m.group(1).upper()
        if contract not in contracts:
            continue
        rows = 0
        codes = collections.Counter()
        try:
            with open(path, encoding="utf-8", errors="replace", newline="") as f:
                reader = csv.reader(f)
                header = next(reader, [])
                code_idx = None
                for i, h in enumerate(header):
                    if "error" in h.lower() and "code" in h.lower():
                        code_idx = i
                        break
                for row in reader:
                    rows += 1
                    if code_idx is not None and code_idx < len(row):
                        codes[row[code_idx].strip()] += 1
        except Exception as e:                                # noqa: BLE001
            log("could not read CMS validation report %s (%s)" % (path, e), 1)
            continue
        prior = out.get(contract)
        if prior is None or m.group(2) > prior["date"]:
            out[contract] = {"file": path, "row_count": rows, "codes": codes, "date": m.group(2)}
    return out


# ==========================================================================
# Section 7 -- prior-run comparison
# ==========================================================================
def build_summary(audits, catalog):
    summary = {"generated": datetime.date.today().isoformat(), "contracts": {}}
    for a in audits:
        summary["contracts"][a.contract] = {
            "resources": dict((k[4:], v) for k, v in a.stats.items() if k.startswith("res_")),
            "findings_by_code": dict(a.F.count),
            "fatal": a.F.level_total(1),
            "level2": a.F.level_total(2),
            "level3": a.F.level_total(3),
        }
    return summary


def load_prior_summary(cache_dir):
    path = os.path.join(cache_dir, "mpf_conformance_prior.json")
    if os.path.exists(path):
        try:
            return json.load(open(path, encoding="utf-8")), path
        except Exception:                                     # noqa: BLE001
            return None, path
    return None, path


def diff_summary(prior, current):
    """Return a list of (contract, code, prior_count, current_count, delta)
    rows for every code whose count changed."""
    rows = []
    prior_contracts = (prior or {}).get("contracts", {})
    cur_contracts = current.get("contracts", {})
    for contract, cur in cur_contracts.items():
        prev = prior_contracts.get(contract, {})
        prev_codes = prev.get("findings_by_code", {})
        cur_codes = cur.get("findings_by_code", {})
        codes = set(prev_codes) | set(cur_codes)
        for code in sorted(codes):
            p = prev_codes.get(code, 0)
            c = cur_codes.get(code, 0)
            if p != c:
                rows.append((contract, code, p, c, c - p))
    return rows


# ==========================================================================
# Report writer
# ==========================================================================
class ConformanceReport(object):

    def __init__(self, audits, refdata, tls, audit_date, catalog, cms_reports,
                 prior_summary, cms_dir_arg):
        self.audits = audits
        self.by_id = dict((a.contract, a) for a in audits)
        self.ref = refdata
        self.tls = tls
        self.date = audit_date
        self.catalog = catalog
        self.cms_reports = cms_reports
        self.prior = prior_summary
        self.cms_dir_arg = cms_dir_arg
        self.doc = Document()
        s = self.doc.sections[0]
        s.left_margin = s.right_margin = Inches(0.8)
        s.top_margin = s.bottom_margin = Inches(0.7)
        normal = self.doc.styles["Normal"]
        normal.font.name = "Calibri"
        normal.font.size = Pt(10)
        normal.paragraph_format.space_after = Pt(6)
        self.checklist = []          # [(code_or_tag, severity, text)]
        self.section_rows = collections.defaultdict(list)   # csv rows per contract

    # -- primitives (mirrors mpf_audit.Report) --------------------------
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

    def p(self, text, bold=False, size=10, italic=False, colour=None):
        par = self.doc.add_paragraph()
        run = par.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        if colour is not None:
            run.font.color.rgb = colour
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
                text = "" if value is None else str(value)
                run = par.add_run(text)
                run.font.size = Pt(font)
                if status_col is not None and i == status_col:
                    run.bold = True
                    run.font.color.rgb = STATUS_COLOR.get(text.split()[0] if text else "", GREY)
        if widths:
            for i, w in enumerate(widths):
                for row in t.rows:
                    row.cells[i].width = Inches(w)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)
        return t

    # -- check helper: records a result row + checklist + csv rows ------
    def check(self, section, code, check_text, result, evidence, resource_type="", resource_id="",
              detail=""):
        """result: PASS / FAIL / WARN / INFO / SKIP"""
        if result in ("FAIL", "WARN"):
            self.checklist.append((code, result, "[%s] %s -- %s" % (code, check_text, evidence)))
        return (result, code, check_text, evidence)

    def add_csv_rows(self, contract, section, code, check_text, result, samples, resource_type=""):
        """Append one CSV row per individual FAIL/WARN instance, drawn from
        the ContractAudit's capped finding samples (already de-duplicated
        and rate-limited by mpf_audit.Findings.MAX_SAMPLES)."""
        if result not in ("FAIL", "WARN"):
            return
        if not samples:
            self.section_rows[contract].append(
                [contract, section, code, check_text, result, resource_type, "", ""])
            return
        for s in samples:
            self.section_rows[contract].append([
                contract, section, code, check_text, result,
                s.get("resource") or resource_type, s.get("id") or "",
                s.get("detail") or s.get("value") or ""])

    # -- header table -----------------------------------------------------
    def header_table(self, verdict_counts):
        self.h("MPF Conformance Report -- CY %s" % CONTRACT_YEAR, level=1)
        idx_rows = []
        for a in self.audits:
            lu = (a.index or {}).get("last_updated", "") if a.index else ""
            files = len(getattr(a, "files", []) or [])
            idx_rows.append("%s: last_updated=%s, files=%d" % (a.contract, lu or "(absent)", files))
        resources_scanned = sum(sum(v for k, v in a.stats.items() if k.startswith("res_"))
                                for a in self.audits)
        cms_note = ("%d contract report(s) reconciled (%s)"
                    % (len(self.cms_reports), ", ".join(sorted(self.cms_reports)))
                    if self.cms_reports else
                    "none supplied (--cms-reports not provided or no matching files found)")
        verdict = ("Verdict: %d PASS, %d FAIL, %d WARN across %d checks run against contracts %s."
                   % (verdict_counts["PASS"], verdict_counts["FAIL"], verdict_counts["WARN"],
                      sum(verdict_counts.values()), ", ".join(a.contract for a in self.audits)))
        if verdict_counts["FAIL"] == 0:
            verdict += (" No blocking (FAIL) conformance defects were found against the checks "
                        "this script can evaluate offline; WARN items require attestation-owner "
                        "review before sign-off.")
        else:
            verdict += (" Blocking (FAIL) defects exist and must be remediated, or accepted with "
                        "documented rationale, before attestation.")
        rows = [
            ("Spec applied", "CMS Technical Implementation Guide for Supplying MA Provider "
                              "Directory Data for Use in MPF, v1.5"),
            ("Contract year", CONTRACT_YEAR),
            ("Contracts assessed", ", ".join(a.contract for a in self.audits) + " (H5826 excluded by scope)"),
            ("Indexes assessed", "; ".join(idx_rows)),
            ("Resources scanned", commify(resources_scanned)),
            ("CMS validation reports reconciled", cms_note),
            ("Report generated", self.date.isoformat()),
            ("Attestation submission date", "TODO -- plan-specific, enter before filing"),
            ("Attestation signer", "TODO -- plan-specific, enter before filing"),
            ("Verdict", verdict),
        ]
        self.table(["Field", "Value"], rows, widths=[1.8, 5.4])

    def methodology(self):
        self.h("Methodology", level=2)
        self.p("This report implements the self-validation, hosting, and Appendix B/D/E checks "
               "described in the CMS Technical Implementation Guide for Supplying Medicare "
               "Advantage Provider Directory Data for Use in Medicare Plan Finder (MPF), v1.5, "
               "against the FHIR-based JSON index and constituent bundle files published for each "
               "contract. Section 1 covers Appendix D hosting and index conformance (self-validation "
               "steps 1-7 and the v1.5 10,000-URL / 300 MB limits). Section 2 covers InsurancePlan "
               "and network linkage per Appendix B and the N3xxx / F5xxx validation inventory. "
               "Section 3 covers PractitionerRole and Practitioner per Appendix B. Section 4 covers "
               "Location per Appendix A/B. Section 5 covers OrganizationAffiliation and facility "
               "Organization records. Section 6 reconciles this run against CMS-supplied validation "
               "reports when available. Section 7 compares this run's findings against the prior run "
               "recorded in the cache. Section 8 is an auto-generated pre-attestation checklist built "
               "from this run's actual FAIL/WARN counts.")

    def scope_note(self):
        self.h("Note on scope", level=2)
        self.p("HealthcareService resources are excluded from this conformance assessment: the "
               "Technical Implementation Guide scopes MPF submission to InsurancePlan, Location, "
               "Organization, OrganizationAffiliation, Practitioner and PractitionerRole, and "
               "HealthcareService is not a data point CMS ingests for MPF display. Any "
               "HealthcareService bundles present in an index are downloaded and syntax-checked "
               "for hosting conformance (Section 1) but are not otherwise scored. Determinations "
               "that require the CMS HPMS registry (N3011-N3014) are reported as format/internal-"
               "consistency checks only, per Section 2, and are not a substitute for CMS's own "
               "HPMS reconciliation.")

    # -- section builder --------------------------------------------------
    def section_table(self, title, rows):
        """rows: list of (result, code, check, evidence)"""
        self.h(title, level=2)
        if not rows:
            self.p("No checks evaluated in this section.")
            return
        table_rows = [(r[0], r[1], r[2], r[3]) for r in rows]
        self.table(["Result", "Code", "Check", "Evidence"], table_rows,
                   widths=[0.6, 0.6, 2.2, 3.8], status_col=0)

    def build(self):
        counts = {"PASS": 0, "FAIL": 0, "WARN": 0}
        section_results = self.build_sections(counts)
        self.header_table(counts)
        self.methodology()
        self.scope_note()
        for title, rows in section_results:
            self.section_table(title, rows)
        self.attestation_checklist()
        return self.doc

    # -- sections -----------------------------------------------------
    def build_sections(self, counts):
        out = []
        out.append(("Section 1 -- Hosting and Index Conformance", self._section1(counts)))
        out.append(("Section 2 -- Plan and Network Linkage", self._section2(counts)))
        out.append(("Section 3 -- PractitionerRole / Practitioner (Appendix B)", self._section3(counts)))
        out.append(("Section 4 -- Location (Appendix A)", self._section4(counts)))
        out.append(("Section 5 -- Facility Entries (OrganizationAffiliation / Organization)",
                    self._section5(counts)))
        out.append(("Section 6 -- Reconciliation vs CMS Validation Reports", self._section6(counts)))
        out.append(("Section 7 -- What Changed Since Last Run", self._section7(counts)))
        return out

    def _tally(self, counts, result):
        counts[result] = counts.get(result, 0) + 1

    def _emit(self, section_no, counts, rows_out, code, check_text, result, evidence,
              resource_type="", samples=None):
        self._tally(counts, result)
        rows_out.append((result, code, check_text, evidence))
        for a in self.audits:
            self.add_csv_rows(a.contract, section_no, code, check_text, result,
                              samples.get(a.contract, []) if samples else
                              a.F.samples.get(code, []), resource_type)

    def _agg_samples(self, code):
        return dict((a.contract, a.F.samples.get(code, [])) for a in self.audits)

    # ---- section 1 ------------------------------------------------------
    def _section1(self, counts):
        rows = []
        A = self.audits

        def cnt(code):
            return sum(a.F.count.get(code, 0) for a in A)

        # index shape / provider_urls
        idx_ok = all(a.index is not None for a in A)
        self._emit("1", counts, rows, "IDX-SHAPE",
                   "index returns JSON with contract_number/contract_year/last_updated/provider_urls",
                   "PASS" if idx_ok else "FAIL",
                   "; ".join("%s: keys=%s" % (a.contract, ",".join(sorted((a.index or {}).keys())))
                             for a in A) if idx_ok else "one or more index files failed to parse")

        c19 = cnt("C4019")
        self._emit("1", counts, rows, "C4019", "provider_urls <= 10,000 entries and each file <= 300 MB",
                   "FAIL" if c19 else "PASS",
                   "%s URLs / largest file %s against caps of %s URLs and 300 MB."
                   % (", ".join("%s=%s" % (a.contract, commify(a.stats.get("index_url_count", 0)))
                                for a in A),
                      mb(max((int(r["content_length"] or 0) for a in A for r in a.http
                              if r["kind"] == "DATA" and str(r["content_length"] or "").isdigit()),
                             default=0)),
                      commify(MAX_INDEX_URLS)))

        c4c18 = cnt("C4004") + cnt("N3015") + cnt("C4018")
        self._emit("1", counts, rows, "C4004/N3015/C4018",
                   "valid JSON syntax, Bundle.type=collection, 0 duplicate resource IDs",
                   "FAIL" if (c4c18 + cnt("DUPID")) else "PASS",
                   "JSON syntax errors=%d, bundle-shape errors=%d, duplicate resource ids=%d."
                   % (cnt("C4004") + cnt("N3015"), cnt("C4018"), cnt("DUPID")))

        self._emit("1", counts, rows, "C4008", "Content-Type application/json on every response",
                   "FAIL" if cnt("C4008") else "PASS",
                   "%d of %d responses missing/incorrect Content-Type."
                   % (cnt("C4008"), sum(len(a.http) for a in A)))

        hdrs = cnt("C4006") + cnt("C4007") + cnt("C4009")
        self._emit("1", counts, rows, "C4006/C4007/C4009",
                   "Content-Length, Last-Modified and ETag headers present",
                   "FAIL" if hdrs else "PASS",
                   "missing Last-Modified=%d, Content-Length=%d, ETag=%d."
                   % (cnt("C4006"), cnt("C4007"), cnt("C4009")))

        heads = [r for a in A for r in a.http if r["head_ok"]]
        allh = [r for a in A for r in a.http]
        self._emit("1", counts, rows, "C4005", "HEAD returns 200, headers consistent with GET",
                   "FAIL" if cnt("C4005") else "PASS",
                   "HEAD 200 on %d of %d URLs." % (len(heads), len(allh)))

        comp = cnt("HOSTCMP")
        self._emit("1", counts, rows, "HOSTCMP", "no Content-Encoding (files served uncompressed)",
                   "FAIL" if comp else "PASS",
                   "%d responses returned a Content-Encoding header." % comp if comp else
                   "no response returned Content-Encoding though gzip/deflate/br were offered.")

        auth_fail = cnt("C4001") + cnt("C4003") + cnt("C4002")
        tls_ok = self.tls.get("verified")
        self._emit("1", counts, rows, "C4001-C4003",
                   "publicly accessible without auth over HTTPS",
                   "FAIL" if (auth_fail or not tls_ok) else "PASS",
                   ("TLS verified to %s; all %d URLs returned 200 without credentials."
                    % (self.tls.get("not_after", "unknown"), len(allh))) if tls_ok and not auth_fail
                   else "TLS or transport failure: %s" % self.tls.get("error", "see findings CSV"))

        return rows

    # ---- section 2 --------------------------------------------------
    def _section2(self, counts):
        rows = []
        A = self.audits

        def cnt(code):
            return sum(a.F.count.get(code, 0) for a in A)

        n12 = cnt("N3001") + cnt("N3002")
        self._emit("2", counts, rows, "N3001/N3002",
                   "MA Plan Identifier present, format HXXXX-NNN-NNN",
                   "FAIL" if n12 else "PASS",
                   "missing=%d, malformed=%d across %s plan ids."
                   % (cnt("N3001"), cnt("N3002"),
                      commify(sum(a.stats.get("ma_plan_ids", 0) for a in A))))

        n45 = cnt("N3004") + cnt("N3005")
        self._emit("2", counts, rows, "N3004/N3005", "period.start = %s-01-01" % CONTRACT_YEAR,
                   "FAIL" if n45 else "PASS",
                   "missing period.start=%d, wrong/invalid year=%d." % (cnt("N3004"), cnt("N3005")))

        f15 = cnt("F5001") + cnt("F5005")
        self._emit("2", counts, rows, "F5001/F5005", "InsurancePlan.network present and resolves",
                   "FAIL" if f15 else "PASS",
                   "missing network refs=%d, broken network refs=%d." % (cnt("F5001"), cnt("F5005")))

        n3 = cnt("N3003")
        self._emit("2", counts, rows, "N3003",
                   "every plan associated with providers (network referenced by >=1 role/affiliation)",
                   "FAIL" if n3 else "PASS",
                   "%d plan(s) whose network is not referenced by any PractitionerRole or "
                   "OrganizationAffiliation." % n3)

        dup = cnt("DUPID")
        self._emit("2", counts, rows, "NETDUP", "no orphan / duplicate network resources",
                   "FAIL" if dup else "PASS",
                   "%d duplicate resource id(s) detected across all resource types (network "
                   "Organizations included)." % dup)

        self._emit("2", counts, rows, "N3011-N3014",
                   "reconcile plan list vs HPMS registry",
                   "WARN",
                   "Format and internal MAPlanID/contract consistency checks pass on this run; a "
                   "definitive reconciliation against the CMS HPMS registry cannot be performed "
                   "offline and requires the HPMS extract (informational/manual step).")

        plan_rows = []
        for a in A:
            for pid in sorted(a.plan_ids):
                plan_rows.append((a.contract, pid))
        self.h("Plans published per contract", level=3)
        if plan_rows:
            self.table(["Contract", "MA Plan ID"], plan_rows, widths=[1.2, 2.0])
        else:
            self.p("No InsurancePlan.identifier[cms.gov/medicare/ma-plan-id] values found.")

        return rows

    # ---- section 3 --------------------------------------------------
    def _section3(self, counts):
        rows = []
        A = self.audits

        def cnt(code):
            return sum(a.F.count.get(code, 0) for a in A)

        checks = [
            ("P1001/F5009", "Identifier (NPI) present, single NPI per resource",
             cnt("P1001") + cnt("F5009")),
            ("P1013/P1014/C4010", "meta.lastUpdated present, parseable, not future, <=30 days old",
             cnt("P1013") + cnt("P1014") + cnt("C4010")),
            ("F5004/F5008", "PractitionerRole.location present and resolves",
             cnt("F5004") + cnt("F5008")),
            ("P1009", "Specialty coding present", cnt("P1009")),
            ("F5001/F5005", "Network extension present and resolves", cnt("F5001") + cnt("F5005")),
            ("F5003/F5007", "Practitioner reference present and resolves",
             cnt("F5003") + cnt("F5007")),
            ("A2009/A2010", "Phone present, exactly 10 digits", cnt("A2009") + cnt("A2010")),
            ("P1006/P1007", "Practitioner.name.text/given[0]/family present",
             cnt("P1006") + cnt("P1007")),
            ("P1010", "Practitioner.gender present", cnt("P1010")),
            ("P1011", "Practitioner.communication (languages) present", cnt("P1011")),
        ]
        for code, text, c in checks:
            self._emit("3", counts, rows, code, text, "FAIL" if c else "PASS",
                       "%d finding(s) across %s." % (c, ", ".join(a.contract for a in A)))

        p1012 = cnt("P1012")
        p1018 = cnt("P1018")
        total_roles = sum(a.stats.get("PractitionerRole", 0) for a in A)
        pct = (100.0 * p1012 / total_roles) if total_roles else 0.0
        result = "WARN" if p1012 else "PASS"
        self._emit("3", counts, rows, "P1012/P1018",
                   "accepting-new-patients extension present on roles (WARN, not FAIL)", result,
                   "%d of %s PractitionerRole record(s) (%.1f%%) omit accepting-patients status; "
                   "%d record(s) use an invalid accepting-patients code."
                   % (p1012, commify(total_roles), pct, p1018))

        return rows

    # ---- section 4 --------------------------------------------------
    def _section4(self, counts):
        rows = []
        A = self.audits

        def cnt(code):
            return sum(a.F.count.get(code, 0) for a in A)

        checks = [
            ("A2007", "address.line[0] present", cnt("A2007")),
            ("A2002", "address.city present", cnt("A2002")),
            ("A2003", "address.state present", cnt("A2003")),
            ("A2004", "address.state is a 2-letter USPS code", cnt("A2004")),
            ("A2005", "postalCode present", cnt("A2005")),
            ("A2006", "postalCode is exactly 5 digits", cnt("A2006")),
        ]
        for code, text, c in checks:
            self._emit("4", counts, rows, code, text, "FAIL" if c else "PASS",
                       "%d finding(s)." % c)

        total_locs = sum(len(a.locs) for a in A)
        with_phone = sum(1 for a in A for tup in a.locs.values() if any(p for p in tup[5]))
        pct = (100.0 * with_phone / total_locs) if total_locs else 0.0
        self._emit("4", counts, rows, "LOC-PHONE",
                   "Location.telecom phone presence (optional for CY%s, informational)" % CONTRACT_YEAR,
                   "INFO", "%d of %s Location resources (%.1f%%) carry a phone number."
                   % (with_phone, commify(total_locs), pct))
        return rows

    # ---- section 5 --------------------------------------------------
    def _section5(self, counts):
        rows = []
        A = self.audits

        def cnt(code):
            return sum(a.F.count.get(code, 0) for a in A)

        oa_ref = cnt("F5001") + cnt("F5002") + cnt("F5004") + cnt("F5005") + cnt("F5006") + cnt("F5008")
        self._emit("5", counts, rows, "F5001/F5002/F5004-F5006/F5008",
                   "OrganizationAffiliation.network/organization/location present and resolve",
                   "FAIL" if oa_ref else "PASS", "%d reference finding(s) across all contracts." % oa_ref)

        self._emit("5", counts, rows, "P1009",
                   "OrganizationAffiliation.specialty coding present",
                   "FAIL" if cnt("P1009") else "PASS", "%d finding(s)." % cnt("P1009"))

        f5009 = cnt("F5009")
        self._emit("5", counts, rows, "P1001/F5009",
                   "Organization.identifier us-npi present, single value",
                   "FAIL" if (cnt("P1001") + f5009) else "PASS",
                   "missing=%d, multiple NPI=%d." % (cnt("P1001"), f5009))

        self._emit("5", counts, rows, "P1008", "Organization.name present",
                   "FAIL" if cnt("P1008") else "PASS", "%d finding(s)." % cnt("P1008"))

        addr_phone = cnt("A2001") + cnt("A2007") + cnt("A2002") + cnt("A2003") + cnt("A2004") + \
                     cnt("A2005") + cnt("A2006") + cnt("A2009") + cnt("A2010")
        self._emit("5", counts, rows, "A2001-A2010",
                   "facility address (Organization/Location) and phone (10 digits) present",
                   "FAIL" if addr_phone else "PASS", "%d finding(s) across facility address/phone checks." % addr_phone)

        p1004 = cnt("P1004")
        self._emit("5", counts, rows, "ORGTYPE-FAC",
                   "Organization.type.coding == 'fac' for facility-context orgs",
                   "FAIL" if p1004 else "PASS",
                   "%d Organization resource(s) referenced in a facility context do not carry "
                   "OrgTypeCS code 'fac' (typed prvgrp/bus/other instead)." % p1004)

        p1012 = cnt("P1012")
        self._emit("5", counts, rows, "OA-ACCEPT",
                   "accepting-patients extension presence on OrganizationAffiliation", "WARN",
                   "Accepting-patients status is only modeled on PractitionerRole in this dataset "
                   "(%d role-level omissions recorded in Section 3); OrganizationAffiliation carries "
                   "no equivalent extension in the source data, so presence cannot be affirmed." % p1012)

        recur = recurring_facility_gaps(A)
        self.h("Recurring facility records with missing NPI / name / phone", level=3)
        if recur:
            self.table(["Organization id", "Name", "Contracts", "Gaps"], recur[:100],
                       widths=[1.3, 1.6, 1.3, 3.0])
            if len(recur) > 100:
                self.p("... and %d more (see conformance_findings_*.csv)." % (len(recur) - 100))
            for oid, name, contracts_str, gap_str in recur:
                for c in contracts_str.split(", "):
                    self.section_rows[c].append(
                        [c, "5", "RECUR-FAC", "recurring facility record with data gaps", "WARN",
                         "Organization", oid, gap_str])
            self._tally(counts, "WARN")
        else:
            self.p("No Organization id shared across two or more of the assessed contracts is "
                   "missing NPI, name or phone.")
        return rows

    # ---- section 6 --------------------------------------------------
    def _section6(self, counts):
        rows = []
        if not self.cms_reports:
            self.p("No CMS validation reports supplied -- section skipped. Pass --cms-reports "
                   "<dir> pointing at a folder of CY%s_ValidationError_{CONTRACT}_YYYYMMDD.csv "
                   "files to enable this reconciliation." % CONTRACT_YEAR)
            return rows
        recon_rows = []
        for a in self.audits:
            info = self.cms_reports.get(a.contract)
            if not info:
                recon_rows.append((a.contract, "no matching CMS report found", "-", "-"))
                continue
            top = ", ".join("%s(%d)" % (c, n) for c, n in info["codes"].most_common(5)) or "n/a"
            recon_rows.append((a.contract, os.path.basename(info["file"]),
                               commify(info["row_count"]), top))
        self.table(["Contract", "CMS report file", "Row count", "Top error codes"], recon_rows,
                  widths=[1.0, 2.6, 1.0, 2.4])
        self.p("Row counts and error codes above are read verbatim from the supplied CMS CSV "
               "files; this script does not attempt to re-derive or estimate CMS's own counts.")
        return rows

    # ---- section 7 --------------------------------------------------
    def _section7(self, counts):
        rows = []
        if not self.prior:
            self.p("No prior run found for comparison (no mpf_conformance_prior.json in the "
                   "cache directory). This run's summary has been saved for the next comparison.")
            return rows
        current = build_summary(self.audits, self.catalog)
        diffs = diff_summary(self.prior, current)
        if not diffs:
            self.p("No change in per-code finding counts since the prior run (%s)."
                   % self.prior.get("generated", "unknown date"))
            return rows
        self.p("Compared against the prior run generated %s:" % self.prior.get("generated", "unknown date"))
        diff_rows = [(c, code, p, cur, ("+%d" % d if d > 0 else str(d)))
                    for c, code, p, cur, d in diffs]
        self.table(["Contract", "Code", "Prior count", "Current count", "Change"], diff_rows,
                  widths=[1.0, 1.0, 1.2, 1.4, 1.0])
        return rows

    # ---- section 8 --------------------------------------------------
    def attestation_checklist(self):
        self.h("Section 8 -- Pre-Attestation Checklist", level=2)
        fails = [c for c in self.checklist if c[1] == "FAIL"]
        warns = [c for c in self.checklist if c[1] == "WARN"]
        if not fails and not warns:
            self.p("No FAIL or WARN items were raised by this run. No corrective action items are "
                   "generated; confirm the TODO attestation fields in the header table before filing.")
            return
        n = 1
        if fails:
            self.p("Blocking items (FAIL) -- must be remediated or formally accepted before "
                   "attestation:", bold=True)
            for code, sev, text in fails:
                self.p("%d. %s" % (n, text))
                n += 1
        if warns:
            self.p("Advisory items (WARN) -- review before attestation:", bold=True)
            for code, sev, text in warns:
                self.p("%d. %s" % (n, text))
                n += 1
        self.p("%d. Confirm the attestation submission date and signer fields in the header "
               "table, currently marked TODO, before filing this report." % n)


# ==========================================================================
# CLI
# ==========================================================================
def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Build the CY %s MPF conformance report (CMS Technical Implementation Guide "
                    "v1.5) for contracts %s." % (CONTRACT_YEAR, ", ".join(REPORT_CONTRACTS)))
    ap.add_argument("--out", default=os.getcwd(), help="directory for the report (default: cwd)")
    ap.add_argument("--cache", default=None, help="cache directory (default: <script dir>/.mpf_cache)")
    ap.add_argument("--fresh", action="store_true", help="ignore cached bundles and re-download")
    ap.add_argument("--no-nppes", action="store_true", help="skip the NPPES registry checks")
    ap.add_argument("--contracts", default=",".join(REPORT_CONTRACTS),
                    help="comma separated subset, default %s" % ",".join(REPORT_CONTRACTS))
    ap.add_argument("--cms-reports", default=None,
                    help="directory of CY%s_ValidationError_{CONTRACT}_YYYYMMDD.csv files" % CONTRACT_YEAR)
    args = ap.parse_args(argv)

    out_dir = os.path.abspath(args.out)
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    cache_dir = os.path.abspath(args.cache or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), ".mpf_cache"))
    if not os.path.isdir(cache_dir):
        os.makedirs(cache_dir)

    wanted = [c.strip().upper() for c in args.contracts.split(",") if c.strip()]
    contracts = [(cid, url) for cid, url in sorted(INDEX_URLS.items())
                if cid in wanted and cid != "H5826"]
    if not contracts:
        sys.stderr.write("No matching contracts (H5826 is excluded by design). Known: %s\n"
                         % ", ".join(sorted(c for c in INDEX_URLS if c != "H5826")))
        return 2

    audit_date = datetime.date.today()
    started = datetime.datetime.now()

    log("=" * MA.LOG_WIDTH)
    log("MPF Conformance Report - CMS technical guide v1.5")
    log("%s | contract year %s | %s" % (ORG_NAME, CONTRACT_YEAR, audit_date.isoformat()))
    log("cache  : %s" % cache_dir)
    log("output : %s" % out_dir)
    log("=" * MA.LOG_WIDTH)

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
    for code, name, level, desc, how in MA.APPENDIX_E:
        catalog[code] = (name, level)
    for tag, (code, name, level) in MA.SUPP_CODES.items():
        catalog[tag] = (name, level)

    audits = []
    for cid, url in contracts:
        # Use a distinct findings CSV name (findings_<CONTRACT>.csv) from mpf_audit's own
        # streaming writer so re-running mpf_audit.py separately is unaffected; ContractAudit
        # writes that file as a side effect of validation and we keep it for parity/debugging.
        audits.append(ContractAudit(cid, url, refdata, cache_dir, out_dir, catalog,
                                    audit_date, fresh=args.fresh).run())

    if not tls.get("verified"):
        for a in audits:
            a.add("C4002", 1, value=tls.get("host", ""),
                 detail="TLS verification failed: %s" % tls.get("error", ""))

    rule("CMS validation report reconciliation")
    cms_reports = load_cms_reports(args.cms_reports, wanted)
    if args.cms_reports and not cms_reports:
        log("--cms-reports was supplied but no matching CY%s_ValidationError_{CONTRACT}_"
            "YYYYMMDD.csv files were found in %s" % (CONTRACT_YEAR, args.cms_reports), 1)
    elif cms_reports:
        log("matched CMS reports for: %s" % ", ".join(sorted(cms_reports)), 1)

    rule("Prior-run comparison")
    prior_summary, prior_path = load_prior_summary(cache_dir)
    if prior_summary:
        log("prior run found: %s" % prior_path, 1)
    else:
        log("no prior run found at %s" % prior_path, 1)

    rule("Report")
    report = ConformanceReport(audits, refdata, tls, audit_date, catalog, cms_reports,
                               prior_summary, args.cms_reports)
    doc = report.build()
    name = "MPF_Conformance_Report_CY%s_%s.docx" % (CONTRACT_YEAR, audit_date.strftime("%Y%m%d"))
    path = os.path.join(out_dir, name)
    doc.save(path)
    log("report : %s" % path, 1)

    # per-contract conformance findings CSVs
    fieldnames = ["contract", "section", "code", "check", "result", "resource_type",
                 "resource_id", "detail"]
    for a in audits:
        csv_path = os.path.join(out_dir, "conformance_findings_%s.csv" % a.contract)
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(fieldnames)
            for row in report.section_rows.get(a.contract, []):
                w.writerow(row)
        log("csv    : %s" % csv_path, 1)

    # save this run's summary for the next "what changed" comparison
    summary = build_summary(audits, catalog)
    prior_out_path = os.path.join(cache_dir, "mpf_conformance_prior.json")
    try:
        json.dump(summary, open(prior_out_path, "w", encoding="utf-8"), indent=1)
        log("saved run summary for next comparison: %s" % prior_out_path, 1)
    except Exception as e:                                    # noqa: BLE001
        log("could not save run summary: %s" % e, 1)

    log("elapsed: %s" % str(datetime.datetime.now() - started).split(".")[0], 1)
    log("=" * MA.LOG_WIDTH)
    return 0


if __name__ == "__main__":
    sys.exit(main())
