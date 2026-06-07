# conference-scheduler

A small, config-driven toolkit for running a scientific conference's abstract
review and programme scheduling in Excel. Three stages, one YAML config:

```
 raw abstracts ──► generate_review.py ──► reviewer workbook (reviewers score)
                                                   │
                                                   ▼
 blank grid  ◄──── generate_schedule.py      scored abstracts
       │                                           │
       └──────────────► fill_schedule.py ◄─────────┘
                              │
                              ▼
                    filled programme grid
```

Everything conference-specific — dates, sessions, breaks, file names, SOC
members, **and the spreadsheet column names** — lives in a single
[`conference.yaml`](conference.yaml). The scripts never need editing: point them
at a different YAML and they work for any conference.

## Why it's generic

The fragile part of automating a review spreadsheet is column positions. This
toolkit locates every field (abstract ID, score, suggested format, submitter,
topic tag, …) **by header name** using configurable aliases, so a sheet whose
columns are renamed or reordered still works:

```yaml
column_aliases:
  avg_score: ["Avg Score", "Average Score", "Score"]
```

If a required column can't be found, you get a clear error listing the aliases
tried and the headers actually present — never a silent wrong-column read.

## Install

```bash
pip install -r requirements.txt    # openpyxl, PyYAML
```

## Usage

Put your input workbooks in `data/` and configure paths in `conference.yaml`.

```bash
# 1. Build the reviewer workbook from raw submissions
python generate_review.py --config conference.yaml

#    … reviewers open it and enter scores + suggested format …

# 2. Build the blank programme grid
python generate_schedule.py --config conference.yaml

# 3. Place the scored talks into the grid
python fill_schedule.py --config conference.yaml          # or --dry-run
```

`--config` defaults to `conference.yaml`, so you can omit it if that's your file.

## The pipeline

| Script | Reads | Writes |
|--------|-------|--------|
| `generate_review.py` | `review.raw_abstracts` (raw submissions) | reviewer workbook: Summary, Distribution, Abstracts, Scores Summary, one tab per SOC member |
| `generate_schedule.py` | the YAML only | blank `schedule_file`: Programme grid + hidden `_META` |
| `fill_schedule.py` | `input_abstracts` (scored) + the blank grid | filled Programme grid + Posters sheet |

`schedule_config.py` is the shared loader that parses the YAML into a `Conf`
object used by all three scripts.

## Key config knobs

- **`session_blocks`** — `[label, start_h, start_m, end_h, end_m]` per session; the
  grid and slot rows are derived from these and `slot_min`.
- **`fixed_entries`** — lock cells (registration, ceremonies, excursions):
  `(day, session, text [, num_slots [, start_slot]])`.
- **`sparkler_sessions`** — reserve part of a session for poster sparklers.
- **`format_map`** — maps each "suggested format" string to a talk type + duration.
- **`station_tag` / `station_threshold` / `station_sessions`** — cluster talks
  carrying a topic tag (e.g. `Station`) above a score threshold into dedicated
  session(s), filling preferred sessions first then spilling over.
- **`random_seed`** — fixes the layout; change it for a different draw, set `null`
  for a fresh random layout each run.

See the comments in [`conference.yaml`](conference.yaml) for the full set.

## Notes

- `data/` and `output/` are gitignored — conference workbooks contain author
  names/emails and should not be committed.
- Scripts never overwrite a non-empty cell, so anything you pre-fill (or fix via
  `fixed_entries`) is preserved across re-runs.
- `generate_review.py` also retains a legacy CLI mode (pass individual flags
  instead of `--config`); run `python generate_review.py -h` for details.

## License

MIT — see [LICENSE](LICENSE).
