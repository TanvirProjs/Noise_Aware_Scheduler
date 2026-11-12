# Reproduction and manuscript consistency check

**Check date:** 2026-11-08  
**Environment:** Python 3.12.14, NumPy 2.3.5, pandas 2.2.3, DuckDB 2.0.0, python-docx 1.2.0, Matplotlib 3.10.8.

The repository was unpacked into a clean staging directory and its scripts were rerun. The following checks passed:

1. `src/simulate.py` completed all five telemetry windows, three workload seeds, two arrival rates, and five policies. Its eight output CSV files (`summary`, `scenario_summary`, `run_manifest`, `workflow_metrics`, plus aggregate and per-record forecast and budget-feasibility outputs) were byte-for-byte identical to the recorded run outputs.
2. `src/validate_forecast.py` and `src/diagnose_budget_feasibility.py` completed successfully. Their aggregate and detailed CSV outputs matched the recorded files byte-for-byte.
3. `src/extract_calibration.py` regenerated all 10,020 rows from the pinned source Parquet file. The resulting compact CSV was byte-for-byte identical to `data/ibm_heron_calibration_snapshots.csv`.
4. `src/make_paper.py` rebuilt the manuscript. The generated Word document's `word/document.xml` was byte-for-byte identical to the checked-in paper, including its text and tables. Rendering that rebuilt DOCX produced a six-page A4 PDF whose extracted text was identical to the checked-in PDF.

## Primary pooled results reproduced

These values are generated from `results/summary.csv`; the manuscript's results text is generated from that file. Table cells separately summarize means and sample standard deviations across the 15 window/seed blocks per arrival rate.

| Policy | Rate | Workflows | Mean response (min) | P95 response (min) | Budget exceedance | Deadline misses |
|---|---:|---:|---:|---:|---:|---:|
| FIFO-current | 4/hour | 10,045 | 30.716361 | 40.498746 | 80.4878% | 0.0100% |
| EB-UCB | 4/hour | 10,045 | 31.345420 | 41.723907 | 80.7666% | 0.0498% |
| FIFO-current | 8/hour | 20,080 | 93.716092 | 372.794344 | 83.0827% | 15.0896% |
| EB-UCB | 8/hour | 20,080 | 157.067399 | 829.094428 | 79.4124% | 19.3127% |

“Budget exceedance” and “failure” are model-internal proxy outcomes under the independent-operation assumption. They are not measured hardware fidelity or service-level guarantees. No quantum jobs were submitted.

The checks establish consistency of the supplied code, data, generated outputs, and manuscript in the stated environment. They do not establish hardware validity, generalization to unseen devices, or statistical significance.
