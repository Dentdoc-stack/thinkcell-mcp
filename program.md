# autoresearch — think-cell MCP self-improvement loop

This adapts [karpathy/autoresearch](https://github.com/karpathy/autoresearch) to
this MCP server. Instead of training a model to minimise `val_bpb`, an agent
edits the chart-building library to **maximise a deck-quality score** produced
by a fixed, read-only judge.

## Setup

1. **Run tag**: propose a tag from today's date (e.g. `jun2`). The branch
   `autoresearch/<tag>` must not already exist — this is a fresh run.
2. **Create the branch**: `git checkout -b autoresearch/<tag>` from `main`.
3. **Read the in-scope files** (the repo is small):
   - `README.md`, `CLAUDE.md` — repo context and the project's learnings.
   - `eval.py` — the judge. **Read it, never edit it.** It defines the metric.
   - `charts/*.py`, `autodeck.py`, `validator.py` — the files you edit.
4. **Verify the toolchain**: `py eval.py` runs and prints a `score:` line, and
   `py tests.py` prints `All checks passed.`
5. **Initialise `results.tsv`** with the header row only. The baseline is
   recorded after the first run.
6. **Confirm and go.**

## What you CAN and CANNOT do

**CAN edit** — the library under optimization:
`charts/base.py`, `charts/*.py`, `autodeck.py`, `validator.py`.

**CANNOT edit**:
- `eval.py` — the judge / ground-truth metric. Mirrors `prepare.py` in
  autoresearch. Touching it invalidates the run.
- `tests.py` — the existing regression suite. It is a **gate**: every kept
  change must keep `py tests.py` green.
- The public tool surface in `server.py` must stay backward compatible. You may
  *add* optional return keys (e.g. `warnings`), never remove or rename existing
  ones.

**CANNOT** add new third-party dependencies. Standard library only.

## The metric

`py eval.py` prints two machine-readable lines:

```
score: <achieved> / <max>
loc:   <editable-lines-of-code>
```

**Goal: maximise `score` (higher is better).** `loc` is the cost column
(analogue of peak VRAM): all else equal, fewer lines is better.

The score has a fixed core that anchors it (every chart type must accept its
own example and emit strict-JSON `.ppttc`) plus headroom dimensions that map to
real defects:

- **Non-finite robustness** — `NaN`/`Infinity` are not valid JSON and break
  `ppttc.exe`. Builders and `autodeck` must reject them (or never emit them);
  `validate_ppttc_data` must flag them.
- **Documented-gotcha guardrails** — surface a non-fatal `warnings` channel for
  the two silent failures recorded in `CLAUDE.md`: percentages passed as
  decimals (every `|value| < 1`), and a date-axis chart with a single category
  (renders empty).

## Simplicity criterion

All else equal, simpler is better. A small score gain that adds ugly complexity
is not worth it; a change that holds the score and *deletes* lines is a win.
Weigh the `loc` cost against the score delta before keeping a change.

## Logging results

Append to `results.tsv` (tab-separated). Header + 5 columns:

```
commit	score	loc	status	description
```

1. short git commit hash (7 chars) — `pending` before committing a kept change
2. score achieved (e.g. `89/126`) — `0/126` for crashes
3. editable LOC from the `loc:` line
4. status: `keep`, `discard`, or `crash`
5. short description (NO tabs)

Leave `results.tsv` untracked by git.

## The experiment loop

LOOP:

1. Note the current branch/commit.
2. Edit the in-scope library with one experimental idea.
3. `py tests.py` — if it fails, the change regressed; treat as discard and
   revert. The existing suite must stay green.
4. `py eval.py > eval.log 2>&1` then read `grep -E "^score:|^loc:" eval.log`.
5. If `score` improved (and tests pass), **keep**: commit and advance.
6. If `score` is equal or worse, **discard**: `git reset --hard` to the prior
   commit. (A pure simplification that holds the score and lowers `loc` may be
   kept.)
7. Record the row in `results.tsv`.

Crashes: if `eval.py` errors out, read the traceback, fix obvious mistakes
(typo, import), else revert and log `crash`.

When the headroom is exhausted (score at max with tests green), stop and report
the kept changes — unlike the open-ended training loop, this judge has a finite
maximum.
