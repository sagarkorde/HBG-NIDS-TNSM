"""Split the dataset zips into small gzip CSV parts (~30-40 MB each) for transfer.

Run from the ICMLDE 2026 folder (Anaconda Prompt):
    python prepare_data.py

Output: data_parts\\<dataset>\\<day>_partNNN.csv.gz
CICIDS2017 is kept in full. CSE-CIC-IDS2018 is uniformly sampled (SAMPLE_2018 of flows, same seed),
because the full set is ~10 GB compressed.
"""
import os
import sys
import zipfile

import pandas as pd

SAMPLE_2018 = 0.05          # fraction of CSE-CIC-IDS2018 flows to keep (uniform random)
ROWS_PER_PART = 250_000
DROP = {"id", "Flow ID", "Src Port"}
JOBS = {"CICIDS2017_improved.zip": 1.0, "CSECICIDS2018_improved.zip": SAMPLE_2018}

here = os.path.dirname(os.path.abspath(__file__))
for zname, frac in JOBS.items():
    zpath = os.path.join(here, zname)
    if not os.path.exists(zpath):
        print("skip (not found):", zname)
        continue
    out = os.path.join(here, "data_parts", zname.replace(".zip", ""))
    os.makedirs(out, exist_ok=True)
    total_in = total_out = 0
    with zipfile.ZipFile(zpath) as z:
        for member in z.namelist():
            if not member.lower().endswith(".csv"):
                continue
            stem = os.path.splitext(os.path.basename(member))[0].replace(" ", "_")
            print(f"{zname}: {member}", flush=True)
            buf, part = [], 0
            with z.open(member) as fh:
                for i, chunk in enumerate(pd.read_csv(fh, chunksize=200_000, low_memory=False, encoding="latin1")):
                    chunk.columns = [str(c).strip() for c in chunk.columns]
                    chunk = chunk.drop(columns=[c for c in chunk.columns if c in DROP])
                    total_in += len(chunk)
                    if frac < 1.0:
                        chunk = chunk.sample(frac=frac, random_state=i)
                    buf.append(chunk)
                    if sum(len(b) for b in buf) >= ROWS_PER_PART:
                        df = pd.concat(buf)
                        df.to_csv(os.path.join(out, f"{stem}_part{part:03d}.csv.gz"), index=False,
                                  float_format="%.6g", compression="gzip")
                        total_out += len(df); part += 1; buf = []
            if buf:
                df = pd.concat(buf)
                df.to_csv(os.path.join(out, f"{stem}_part{part:03d}.csv.gz"), index=False,
                          float_format="%.6g", compression="gzip")
                total_out += len(df)
    print(f"DONE {zname}: read {total_in:,} flows, wrote {total_out:,}", flush=True)
print("All done. Tell Claude the parts are ready.")
