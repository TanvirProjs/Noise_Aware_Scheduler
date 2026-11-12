"""Extract unique hardware calibration snapshots used by the scheduler paper."""
import duckdb
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "train-00000-of-00001.parquet"
OUT = ROOT / "data" / "ibm_heron_calibration_snapshots.csv"
con = duckdb.connect()
query = f"""
WITH canonical AS (
  SELECT DISTINCT backend, observed_time, property_family,
    CASE WHEN property_family = 'cz_error'
      THEN least(qubit_a, CAST(qubit_b AS BIGINT)) ELSE qubit_a END AS item_a,
    CASE WHEN property_family = 'cz_error'
      THEN greatest(qubit_a, CAST(qubit_b AS BIGINT)) ELSE NULL END AS item_b,
    value, calibration_age_seconds
  FROM read_parquet('{RAW.as_posix()}')
  WHERE backend IN ('ibm_fez','ibm_kingston','ibm_marrakesh')
    AND property_family IN ('cz_error','readout_error')
    AND value IS NOT NULL AND value > 0 AND value < 0.999
    AND observed_time IS NOT NULL
)
SELECT backend, observed_time, property_family,
       median(value) AS median_error,
       quantile_cont(value, 0.90) AS p90_error,
       count(*) AS n_components,
       quantile_cont(calibration_age_seconds, 0.50) AS median_age_seconds,
       quantile_cont(calibration_age_seconds, 0.90) AS p90_age_seconds,
       max(calibration_age_seconds) AS max_age_seconds
FROM canonical
GROUP BY backend, observed_time, property_family
ORDER BY backend, observed_time, property_family
"""
df = con.execute(query).fetchdf()
df.to_csv(OUT, index=False)
print(f"wrote {len(df)} backend-snapshot-property rows to {OUT}")
print(df.groupby(['backend','property_family']).agg(rows=('observed_time','count'),first=('observed_time','min'),last=('observed_time','max')).to_string())
