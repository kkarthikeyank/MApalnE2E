# MPF Provider Directory Validation -- Automation Layer

This repo wraps GitHub Actions automation around the **existing, unmodified**
CMS Medicare Advantage Plan Finder (MPF) provider-directory validation logic
in `mpf_audit.py` (and its companion `mpf_conformance.py`). `mpf_audit.py`
remains the single source of truth for every validation rule (CMS Appendix E
inventory, FHIR parsing, NPPES checks, Word report + CSV generation). Nothing
in `.github/` or `scripts/` duplicates or reimplements any of that logic --
they only decide *when* it runs, *which* contract runs, detect change,
handle the resulting report, and send email.

Contracts: `H1619`, `H3124`, `H5826`, `H9207` -- Plan Year 2027.

## Architecture

```
.github/workflows/mpf_end_to_end_validation.yml   orchestration only
scripts/mpf_monitor.py                             change detection only (reads index.json, compares to state)
scripts/email_report.py                             email formatting/sending only
scripts/test_mpf_monitor.py                         offline unit test for the comparison logic
scripts/ci_helpers.py                                small workflow-only plumbing helpers (no validation logic)
state/mpf_state.json                                 {"H1619": "<last_updated>", ...}
reports/                                              workflow's output dir for mpf_audit.py (gitignored contents)
mpf_audit.py                                          UNCHANGED validation logic (single source of truth)
mpf_conformance.py                                    companion script, untouched
```

### Two execution modes, one script

Both AUTOMATIC (`schedule: */15 * * * *`) and MANUAL (`workflow_dispatch`,
choice of `ALL` / `H1619` / `H3124` / `H5826` / `H9207`) call the exact same
command:

```
python mpf_audit.py --out reports --contracts <ID>
```

`ALL` in manual mode runs this once per contract (four separate invocations)
so each contract gets its own report, CSV and email, consistent with
contract isolation.

### Change detection (automatic mode only)

`scripts/mpf_monitor.py` fetches each contract's small `index.json` (the
same URLs as `mpf_audit.py`'s hardcoded `INDEX_URLS`; this is a lightweight
JSON document with a `provider_urls` array, not the multi-hundred-MB data
bundles themselves -- see `mpf_audit.py`'s own `fetch_to_file`/`ContractAudit`
handling), reads its `last_updated` field (confirmed against `mpf_audit.py`'s
own parsing: `(a.index or {}).get("last_updated", "")`), and compares it to
`state/mpf_state.json`. A contract only proceeds to `mpf_audit.py` when its
`last_updated` differs from the stored value (or has no stored value yet --
see Bootstrap below). Because only the changed contracts trigger the heavy
`mpf_audit.py` download, the 15-minute cron is safe even though a full
contract validation can be large.

Manual mode bypasses this comparison entirely for the selected contract(s).

### Exit codes -- approach taken

`mpf_audit.py`'s own `main()` already returns:
- `0` -- ran clean, no fatal (Level 1) findings
- `1` -- ran clean, fatal findings were reported
- `2` -- bad invocation (e.g. `--contracts` matched nothing)

We made one **minimal, additive** change at the very bottom of the file (the
`if __name__ == "__main__":` block only -- no validation logic touched):
wrapped `sys.exit(main())` in a `try/except` so that any *unhandled
exception* (a genuine crash) now exits `3` instead of an uncaught traceback
with a non-standard exit status. We deliberately did **not** attempt to
invent exit codes `2` (download failure) or `4` (report-generation failure)
by threading new state through `main()`'s internals, since that would mean
touching the validation flow itself.

Instead, the workflow's "Run mpf_audit.py" step infers automation failure
using the documented fallback from the spec:
- exit code `3` (crash) or `2` (bad invocation) => automation failure, **or**
- the expected `MPF_Provider_Directory_Audit_CY2027_*.docx` and
  `findings_<CONTRACT>.csv` were **not** actually created => automation
  failure, regardless of exit code.

Only when neither condition is true is the run classified as a completed
validation, and then `0` vs `1` distinguishes "clean" from "findings
reported" for the email subject/body -- never hardcoded, always derived from
the real exit code and CSV contents.

### Validation result vs. automation failure

These are never conflated:
- **Automation failure** (download/network error, script crash, report not
  produced): email kind `automation-crash`; state is **not** advanced for
  that contract, so the next scheduled run retries it.
- **Validation completed** (script ran to completion and produced its
  report + CSV, whether or not it found issues): email kind
  `updated-success` / `updated-findings` (or `manual-run` for
  workflow_dispatch); state **is** advanced (automatic mode only, after a
  successful run).

### State management

`state/mpf_state.json` holds `{"H1619": "<last_updated>", ...}`. It is
updated **only** after that contract's validation + report completed
successfully, and only for automatic-mode runs (manual runs intentionally
do not perturb the change-detection baseline). The workflow commits and
pushes the updated file using `GITHUB_TOKEN` via the default checkout
credentials. The workflow-level

```yaml
concurrency:
  group: mpf-validation
  cancel-in-progress: false
```

queues overlapping scheduled/manual runs instead of letting them race on
this commit or double-process a contract.

### Bootstrap / first run

`state/mpf_state.json` ships as `{}` in the initial commit. `mpf_monitor.py`
treats "no stored value for a contract" as `CHANGED`, so the **first**
automatic run performs an initial validation for all four contracts and
populates state (**Option B**). If you would rather the first scheduled run
do nothing until a real change is observed, switch to **Option A**: seed
`state/mpf_state.json` with each contract's current `last_updated` value
(fetch each contract's `index.json` once, e.g. via
`python scripts/mpf_monitor.py --json`, and copy the `current` values in)
before the first scheduled run.

### Report handling

Only the files `mpf_audit.py` actually produces are used -- no second report
format was invented:
- `MPF_Provider_Directory_Audit_CY2027_<YYYYMMDD>.docx`
- `findings_<CONTRACT>.csv` (columns: `Level,ErrorCode,ValidationName,
  ContractID,ResourceType,ResourceID,NPI,Field,Value,Detail`)

Both are uploaded as workflow artifacts (`actions/upload-artifact`) and
attached to the contract's notification email.

### Email

`scripts/email_report.py` uses `smtplib` and reads secrets from the
environment only:

| Secret | Required | Default if unset |
|---|---|---|
| `EMAIL_USERNAME` | yes | -- |
| `EMAIL_PASSWORD` | yes | -- |
| `EMAIL_TO` | yes (comma-separated OK) | -- |
| `SMTP_SERVER` | no | `smtp.office365.com` |
| `SMTP_PORT` | no | `587` (STARTTLS) |

Add these under **Settings -> Secrets and variables -> Actions -> New
repository secret**. Nothing secret is ever printed to logs -- only the
existence of a *missing* secret name is reported. The findings summary in
the email body (row counts per `Level`, top `ErrorCode`s) is computed live
from the real `findings_<CONTRACT>.csv`; PASS/FAIL wording is never
hardcoded, it is derived from the exit code and CSV contents each time.

### Job summary

Each run writes a Markdown job summary (`$GITHUB_STEP_SUMMARY`) with: a
contract status table (will run / skipped - no change), a per-contract
validation results table (exit code, report generated, result, email
status), and a final result line.

## First-run checklist

1. Push this repo to GitHub.
2. Add the five email secrets above.
3. Either let the first scheduled run bootstrap all four contracts
   (Option B, default), or seed `state/mpf_state.json` first (Option A).
4. Optionally trigger a manual run (`Actions -> MPF End-to-End Validation ->
   Run workflow`) to validate one contract immediately.

## Test scenarios (spec section 24)

These are documented rather than executed against live CMS servers, since
several require simulating network/crash failures. `scripts/test_mpf_monitor.py`
independently proves the core comparison logic (run: `python
scripts/test_mpf_monitor.py`).

| # | Scenario | Expected behavior | Implemented by |
|---|---|---|---|
| 1 | No contract has changed (scheduled run) | All 4 SKIPPED - NO CHANGE; no validation, no report, no email | `mpf_monitor.py` returns `UNCHANGED`; workflow's `to_run` list is empty |
| 2 | Exactly one contract changed | Only that contract runs `mpf_audit.py --contracts <ID>`, gets its own report/CSV/email; state advances for it only | `mpf_monitor.py` status `CHANGED` filtering + per-contract loop in the workflow |
| 3 | All 4 contracts changed simultaneously | Each runs independently, 4 separate reports/CSVs/emails, 4 independent state updates | Same per-contract loop; contracts never share report objects since each invocation is `--contracts <single ID>` |
| 4 | mpf_audit.py runs clean (no findings) | Email kind `updated-success` / `manual-run`, state advances | exit code `0` + docx/csv present -> `updated-success` branch |
| 5 | mpf_audit.py runs and reports findings | Email kind `updated-findings` with real findings summary from CSV, state advances | exit code `1` + docx/csv present -> `updated-findings` branch, `summarize_findings_csv()` |
| 6 | Download/network failure | `automation-crash` email; state NOT advanced; next run retries | inferred via missing docx/csv output regardless of exit code |
| 7 | Script/runtime crash (unhandled exception) | `automation-crash` email; state NOT advanced | `mpf_audit.py`'s added exit-code-3 wrapper + missing-output fallback check |
| 8 | Report-generation failure | `automation-crash` email; state NOT advanced | missing-output fallback check (docx not found even if exit code was 0/1) |
| 9 | Manual run of a single contract | Runs regardless of change detection, unconditional email kind `manual-run`, does not touch automatic state baseline (unless contract has no prior state) | `run_mode == manual` bypasses `mpf_monitor.py`; state update guarded to automatic mode only |
| 10 | Manual run of ALL | Runs once per contract unconditionally, 4 separate reports/CSVs/emails | manual-mode `candidates` set to all 4 IDs, same per-contract loop as scheduled runs |

## Notes

- The workflow runs entirely on GitHub-hosted `ubuntu-latest` runners --
  no dependency on any local machine.
- `mpf_audit.py`'s own CLI/report/CSV behavior, hardcoded `INDEX_URLS`, and
  all validation rules are untouched except the single additive exit-code-3
  wrapper described above.
