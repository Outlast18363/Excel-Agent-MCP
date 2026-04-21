# Multi-Agent Framework — Usage Guide

A three-role (Planner → Executor ⇄ Evaluator) loop that drives the Codex CLI
(`codex exec --json`) against an Excel workbook via the `excel-mcp` MCP server.
The Planner produces a plan, the Executor mutates the workbook, and the
Evaluator verifies the result and emits a `<verdict>` tag that steers the loop.

---

## 1. Prerequisites

- **Codex CLI** available on `PATH` (the framework invokes `codex exec --json`).
- **`excel-mcp` MCP server** installed at the path hard-coded in
  `config.py::EXCEL_MCP_ROOT` (currently
  `/Users/jz/Desktop/spreadsheet FINCH proj/excel_mcp`), with its own
  `.venv/bin/python` interpreter. If you move it, update `EXCEL_MCP_ROOT`.
- **Python 3.10+** (the code uses `X | None` union syntax).
- No extra Python dependencies — the package uses only the stdlib
  (`subprocess`, `json`, `pathlib`, `argparse`, `shutil`, `re`, `dataclasses`).

---

## 2. Quick start

Run the framework as a module from the directory that *contains* the
`multi_agent_framework/` package (i.e. the project root, not inside the
package itself):

```bash
cd /Users/jz/dev_space/codex_agent_test

python -m multi_agent_framework.runner \
    --task "Add a pivot summary sheet that totals revenue by region." \
    --workbook /absolute/path/to/book.xlsx
```

Optional: pin the run directory (otherwise one is auto-generated, see §4):

```bash
python -m multi_agent_framework.runner \
    --task "..." \
    --workbook /absolute/path/to/book.xlsx \
    --run-dir trace_logs/run_demo
```

Optional: supply `--task-id` to stamp the orchestrator-published workbook
copy inside `final_result/` with a meaningful signature (e.g. a Finch
dataset id). If omitted, the `--run-dir` basename is used.

```bash
python -m multi_agent_framework.runner \
    --task "..." \
    --workbook /absolute/path/to/book.xlsx \
    --run-dir trace_logs/run_demo \
    --task-id finch_0042
```

On exit, `runner.py` prints a JSON summary to stdout, for example:

```json
{
  "verdict": "success",
  "iterations": 2,
  "redo_count": 1,
  "reset_count": 0,
  "max_redo": 3,
  "max_reset": 1,
  "trace": "/abs/.../trace_logs/run_20260420_143000/trace.jsonl",
  "run_dir": "/abs/.../trace_logs/run_20260420_143000",
  "final_dir": "/abs/.../trace_logs/run_20260420_143000/final_result"
}
```

Exit code is `0` on `verdict=="success"`, `1` otherwise, and `2` if the
workbook path does not exist.

---

## 3. How the loop works

```
Planner ──► [snapshot.xlsx]
                │
                ▼
        ┌─────────────────────────────────────────────────────────────┐
        │ each iteration:                                             │
        │   1. Orchestrator wipes final_result/                       │
        │   2. Executor edits --workbook in place; writes any         │
        │      non-xlsx deliverable into final_result/                │
        │   3. Orchestrator copies --workbook to                      │
        │      final_result/{task_id}_final_result.xlsx               │
        │   4. Evaluator scans final_result/ (non-xlsx if present is  │
        │      the deliverable; else the workbook copy is)            │
        └─────────────────────────────────────────────────────────────┘
                │
                ▼
            verdict
              ├─ success ──► done
              ├─ redo    ──► next iteration (workbook kept)
              └─ reset   ──► restore snapshot + delete impl_report.md
                             ──► next iteration
```

- **Planner** runs exactly once. It only *reads* the workbook (inspection
  tools only) and writes `plan.md`.
- **Snapshot**: after the plan is written, the orchestrator copies the
  workbook to `snapshot.xlsx` so a `reset` verdict can roll changes back.
- **`final_result/` wipe**: the orchestrator empties this folder at the
  top of every iteration (initial, redo, and reset) so the Evaluator can
  never see a stale workbook copy or a stale non-xlsx report.
- **Executor** mutates the workbook in place and writes `impl_report.md`.
  If the task calls for a non-Excel deliverable (`.txt`, `.md`, `.pdf`,
  `.docx`, `.csv`, …), the Executor writes it *directly* into
  `final_result/`. The Executor is explicitly forbidden from copying the
  workbook itself — that is the orchestrator's job.
- **Workbook publication**: as soon as the Executor returns, the
  orchestrator unconditionally copies `--workbook` to
  `final_result/{task_id}_final_result.xlsx`. The `{task_id}_` prefix is
  a signature: any workbook copy in `final_result/` that does not match
  that name was not produced by the orchestrator.
- **Evaluator** scans `final_result/` and applies a priority rule: if any
  non-xlsx file is present, that file is the primary deliverable and the
  workbook copy is supporting context; otherwise the workbook copy is
  itself the deliverable. It writes `eval_report.md` and ends its final
  message with exactly one of:
  - `<verdict>success</verdict>` — done.
  - `<verdict>redo</verdict>` — re-run Executor with the eval report in hand.
  - `<verdict>reset</verdict>` — restore snapshot, delete the stale
    `impl_report.md`, run Executor fresh.
- **Loop caps** (see `orchestrator.py`):
  - `MAX_REDO = 3` — successive in-place fixes before escalating to reset.
  - `MAX_RESET = 1` — how many times the workbook may be rolled back.
  - `EvaluatorAgent.MAX_VERDICT_RETRY = 2` — Evaluator re-invocations when
    the verdict tag is missing from its final message (fresh CLI subprocess,
    fresh context window each time).

---

## 4. Where everything is written (vital paths)

All run artifacts live under a single **run directory**.

- **`--run-dir` explicit**: the path you pass.
- **`--run-dir` omitted**: `trace_logs/run_<YYYYMMDD_HHMMSS>`, resolved
  relative to the directory you launched `python -m` from.

Inside `<run_dir>/`:

| Artifact | Path | Produced by |
|---|---|---|
| **Event bus JSONL** (every event from every agent + orchestrator transitions/verdicts) | `<run_dir>/trace.jsonl` | `EventBus` (`event_bus.py`) |
| **Pre-mutation snapshot** of the workbook (used for `reset`) | `<run_dir>/snapshot.xlsx` | `Orchestrator` (copied right after Planner finishes) |
| **Final deliverable folder** (wiped at the top of each iteration) | `<run_dir>/final_result/` | `Orchestrator.__init__` |
| ├─ Published workbook copy (the signed deliverable) | `<run_dir>/final_result/{task_id}_final_result.xlsx` | Orchestrator — re-copied from `--workbook` after every Executor run |
| └─ Standalone report (optional, when the task asks for one) | `<run_dir>/final_result/<name>.{txt,md,pdf,docx,csv,…}` | Executor — written directly into `final_result/` |
| **Handover directory** (shared docs between agents) | `<run_dir>/handover/` | `Orchestrator.__init__` |
| ├─ Plan (Markdown) | `<run_dir>/handover/plan.md` | Planner |
| ├─ Implementation report (Markdown) | `<run_dir>/handover/impl_report.md` | Executor — deleted on `reset` |
| └─ Evaluation report (Markdown) | `<run_dir>/handover/eval_report.md` | Evaluator |

### Final output: `final_result/`

`<run_dir>/final_result/` is the **authoritative deliverable surface** of
a run. Its contract:

- **Per-iteration wipe.** The orchestrator empties this folder at the top
  of every iteration (initial, redo, and reset). No file ever carries
  over from a prior iteration, so the Evaluator only ever sees artifacts
  produced by the current Executor pass.
- **Workbook publication is orchestrator-only.** After every Executor
  run, the orchestrator copies `--workbook` to
  `final_result/{task_id}_final_result.xlsx`. `{task_id}_` is a
  signature: any `*.xlsx` inside `final_result/` that does not start
  with that prefix was **not** produced by the orchestrator and is
  almost certainly a bug (the Executor prompt explicitly forbids
  copying the workbook itself).
- **Non-xlsx deliverables are Executor-owned.** When the task calls for
  a `.txt`/`.md`/`.pdf`/`.docx`/`.csv`/… report, the Executor writes it
  *directly* into `final_result/`. Anything it leaves elsewhere is
  invisible to the Evaluator.
- **Evaluator priority rule.** If any non-xlsx file is present in
  `final_result/`, that file is the primary deliverable and the
  workbook copy is supporting context. Otherwise the workbook copy is
  the deliverable.

`--workbook` is still the single **edit** target and the snapshot /
rollback pivot — nothing in the `final_result/` flow mutates it.
`<run_dir>/snapshot.xlsx` remains the rollback source on `reset`.

> If you want to keep the original untouched, copy it yourself before
> invoking the runner and pass the copy as `--workbook` — the
> orchestrator will still publish its own signed copy into
> `final_result/`.

### Summary JSON

The end-of-run summary (shown in §2) is printed to **stdout only** — it is
not persisted by default. Redirect if you want a file:

```bash
python -m multi_agent_framework.runner ... > run_summary.json
```

---

## 5. Inspecting a run

- **Live tail of the event stream** (agent messages, tool calls, reasoning,
  orchestrator transitions, verdicts):

```bash
tail -f trace_logs/run_<ts>/trace.jsonl | jq .
```

- **Reading the handover chain** (plan → impl → eval) gives you a
  human-readable narrative of the run without decoding the JSONL.

- **Verdict history**: filter the trace for verdict events:

```bash
jq -c 'select(.type=="verdict")' trace_logs/run_<ts>/trace.jsonl
```

---

## 6. Configuration knobs

| Where | Name | Purpose |
|---|---|---|
| `config.py` | `EXCEL_MCP_ROOT` | Path to the `excel-mcp` checkout (its `.venv/bin/python` and `cwd` are derived from this). |
| `config.py` | `MODEL` | Codex model passed via `-m`. |
| `config.py` | `REASONING_EFFORT` | `model_reasoning_effort` override. |
| `config.py` | `APPROVAL_POLICY` | `approval_policy` override (default `never` — fully autonomous). |
| `config.py` | `SANDBOX_MODE` | `-s` flag; `workspace-write` lets agents write inside `run_dir`. |
| `config.py` | `ROLE_TOOLS` | Per-role MCP tool allowlist. Planner and Evaluator get read-only tools; Executor additionally gets `web_search` and `xlwing_skills`. |
| `orchestrator.py` | `MAX_REDO` | Successive in-place Executor retries. |
| `orchestrator.py` | `MAX_RESET` | Snapshot rollbacks allowed. |
| `agent.py` | `EvaluatorAgent.MAX_VERDICT_RETRY` | Retries when the verdict tag is missing. |

The `-C <run_dir>` flag in `build_codex_cmd` sets the Codex working
directory to the run directory, so any relative paths an agent writes land
under `<run_dir>/` rather than polluting the caller's cwd.

---

## 7. Programmatic use

If you'd rather skip the CLI wrapper:

```python
from pathlib import Path
from multi_agent_framework.orchestrator import Orchestrator

result = Orchestrator(
    task="Add a pivot summary sheet…",
    workbook=Path("/abs/path/to/book.xlsx"),
    run_dir=Path("trace_logs/run_demo"),
    task_id="finch_0042",
).run()

print(result.verdict, result.iterations, result.trace_path)
```

`RunResult` fields: `verdict`, `iterations`, `redo_count`, `reset_count`,
`trace_path`, `run_dir`, `final_dir`.

---

## 8. Troubleshooting

- **`workbook not found`** (exit code 2): the `--workbook` path is wrong or
  relative to an unexpected cwd. Use an absolute path.
- **Codex subprocess hangs**: the framework uses `stdin=DEVNULL`, so the CLI
  never waits on a TTY. If it still hangs, the MCP server is likely
  unreachable — confirm `EXCEL_MCP_ROOT/.venv/bin/python -m excel_mcp` runs
  standalone.
- **Verdict missing**: the Evaluator retries up to `MAX_VERDICT_RETRY + 1`
  times; after that the orchestrator defaults to `redo` and logs a
  `warning` event to the trace with the tail of the final message.
- **Looks like nothing changed after `reset`**: that's expected — the
  workbook is restored from `snapshot.xlsx` and `impl_report.md` is deleted
  so the next Executor doesn't read stale instructions.
