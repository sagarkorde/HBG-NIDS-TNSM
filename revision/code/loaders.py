"""Unified loaders: original CICIDS2017, improved CIC-IDS-2017/CSE-CIC-IDS-2018 (Engelen/Liu et al., CNS 2022),
and UNSW-NB15. Every loader returns a canonical flow DataFrame:

    ts (datetime64), src, dst, dport, proto, label (0/1), attack (str category),
    core numeric columns (CORE), plus all other numeric columns prefixed 'f_' for flow-level baselines.
"""
from __future__ import annotations

import re
from pathlib import Path

import numpy as np
import pandas as pd

CORE = ["duration", "fwd_pkts", "bwd_pkts", "fwd_bytes", "bwd_bytes",
        "flow_bytes_s", "flow_pkts_s", "iat_mean", "fwd_iat_mean", "bwd_iat_mean",
        "pkt_len_mean", "total_bytes", "total_pkts"]


def _key(c: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(c).lower())


# alias table: canonical -> list of normalised source names
CIC_ALIASES = {
    "src": ["sourceip", "srcip"],
    "dst": ["destinationip", "dstip"],
    "dport": ["destinationport", "dstport"],
    "proto": ["protocol"],
    "ts": ["timestamp"],
    "label": ["label"],
    "attempted": ["attemptedcategory"],
    "duration": ["flowduration"],
    "fwd_pkts": ["totalfwdpackets", "totalfwdpacket", "totfwdpkts"],
    "bwd_pkts": ["totalbackwardpackets", "totalbwdpackets", "totbwdpkts"],
    "fwd_bytes": ["totallengthoffwdpackets", "totallengthoffwdpacket", "totlenfwdpkts"],
    "bwd_bytes": ["totallengthofbwdpackets", "totallengthofbwdpacket", "totlenbwdpkts"],
    "flow_bytes_s": ["flowbytess"],
    "flow_pkts_s": ["flowpacketss", "flowpktss"],
    "iat_mean": ["flowiatmean"],
    "fwd_iat_mean": ["fwdiatmean"],
    "bwd_iat_mean": ["bwdiatmean"],
    "pkt_len_mean": ["packetlengthmean", "pktlenmean"],
}
CIC_DROP = {"flowid", "sourceip", "srcip", "destinationip", "dstip", "timestamp", "label", "id",
            "attemptedcategory", "sourceport", "srcport"}


def _fix_cic2017_original_ts(s: pd.Series) -> pd.Series:
    """Original CICIDS2017 CSVs use 'd/m/yyyy h:mm' with 12-hour clock and no AM/PM.
    Office hours were ~08:00-18:00, so hours 1..7 are afternoon (add 12h)."""
    ts = pd.to_datetime(s, errors="coerce", dayfirst=True, format="mixed")
    h = ts.dt.hour
    return ts + pd.to_timedelta(((h >= 1) & (h <= 7)).astype(int) * 12, unit="h")


def load_cic_csv(path: Path, original_2017: bool = False, usecols_numeric: bool = True) -> pd.DataFrame:
    raw = pd.read_csv(path, low_memory=False, encoding="latin1", skipinitialspace=True)
    raw.columns = [str(c).strip() for c in raw.columns]
    if "Label" not in raw.columns and "label" not in raw.columns:
        raise ValueError(f"no Label column in {path}")
    keymap = {_key(c): c for c in raw.columns}
    out = pd.DataFrame(index=raw.index)
    for canon, aliases in CIC_ALIASES.items():
        for a in aliases:
            if a in keymap:
                out[canon] = raw[keymap[a]]
                break
    # repeated header lines inside some CSE-CIC-IDS2018 files
    out = out[out["label"].astype(str).str.strip().str.lower() != "label"]
    raw = raw.loc[out.index]
    lab = out["label"].astype(str).str.strip()
    # CNS2022 recommendation: "Attempted" flows are benign
    attempted = lab.str.contains("attempted", case=False)
    if "attempted" in out:
        attempted |= pd.to_numeric(out["attempted"], errors="coerce").fillna(-1).ne(-1)
    lab = lab.where(~attempted, "BENIGN")
    out["attack"] = lab.str.upper().replace({"BENIGN": "BENIGN"})
    out["label"] = (out["attack"] != "BENIGN").astype(np.int8)
    if original_2017:
        out["ts"] = _fix_cic2017_original_ts(out["ts"])
    else:
        ts = pd.to_datetime(out["ts"], errors="coerce", format="ISO8601")
        if ts.isna().mean() > 0.5:  # non-ISO (e.g. original CSE-CIC-IDS2018 'dd/mm/yyyy hh:mm:ss')
            ts = pd.to_datetime(out["ts"], errors="coerce", format="mixed", dayfirst=True)
        out["ts"] = ts
    for c in CORE:
        if c in out:
            out[c] = pd.to_numeric(out[c], errors="coerce").astype("float64")
    out["total_bytes"] = out["fwd_bytes"].fillna(0) + out["bwd_bytes"].fillna(0)
    out["total_pkts"] = out["fwd_pkts"].fillna(0) + out["bwd_pkts"].fillna(0)
    for c in ["src", "dst"]:
        if c not in out:
            out[c] = "na"
        out[c] = out[c].astype(str)
    out["dport"] = pd.to_numeric(out.get("dport", -1), errors="coerce").fillna(-1).astype(np.int32)
    out["proto"] = pd.to_numeric(out.get("proto", -1), errors="coerce").fillna(-1).astype(np.int16)
    if usecols_numeric:
        for c in raw.columns:
            k = _key(c)
            if k in CIC_DROP or k in {"destinationport", "dstport", "protocol"}:
                continue
            v = pd.to_numeric(raw[c], errors="coerce")
            if v.notna().mean() > 0.9:
                out["f_" + k] = v.astype("float32")
    out = out.drop(columns=[c for c in ["attempted"] if c in out])
    out = out.replace([np.inf, -np.inf], np.nan)
    out = out.dropna(subset=["ts"])
    out["file"] = path.name
    out["dayfile"] = re.sub(r"(_part\d+)?\.csv(\.gz)?$", "", path.name)
    return out.reset_index(drop=True)


UNSW_COLS = ["srcip", "sport", "dstip", "dsport", "proto", "state", "dur", "sbytes", "dbytes", "sttl", "dttl",
             "sloss", "dloss", "service", "Sload", "Dload", "Spkts", "Dpkts", "swin", "dwin", "stcpb", "dtcpb",
             "smeansz", "dmeansz", "trans_depth", "res_bdy_len", "Sjit", "Djit", "Stime", "Ltime", "Sintpkt",
             "Dintpkt", "tcprtt", "synack", "ackdat", "is_sm_ips_ports", "ct_state_ttl", "ct_flw_http_mthd",
             "is_ftp_login", "ct_ftp_cmd", "ct_srv_src", "ct_srv_dst", "ct_dst_ltm", "ct_src_ ltm",
             "ct_src_dport_ltm", "ct_dst_sport_ltm", "ct_dst_src_ltm", "attack_cat", "Label"]


def load_unsw_csv(path: Path) -> pd.DataFrame:
    raw = pd.read_csv(path, header=None, names=UNSW_COLS, low_memory=False, encoding="latin1")
    out = pd.DataFrame(index=raw.index)
    out["ts"] = pd.to_datetime(pd.to_numeric(raw["Stime"], errors="coerce"), unit="s")
    out["src"] = raw["srcip"].astype(str)
    out["dst"] = raw["dstip"].astype(str)
    out["dport"] = pd.to_numeric(raw["dsport"], errors="coerce").fillna(-1).astype(np.int32)
    out["proto"] = raw["proto"].astype("category").cat.codes.astype(np.int16)
    out["label"] = pd.to_numeric(raw["Label"], errors="coerce").fillna(0).astype(np.int8)
    cat = raw["attack_cat"].fillna("").astype(str).str.strip()
    cat = cat.replace({"Backdoors": "Backdoor"})
    out["attack"] = np.where(out["label"] == 1, cat.str.upper(), "BENIGN")
    num = lambda c: pd.to_numeric(raw[c], errors="coerce").astype("float64")
    out["duration"] = num("dur") * 1e6  # to microseconds, like CICFlowMeter
    out["fwd_pkts"], out["bwd_pkts"] = num("Spkts"), num("Dpkts")
    out["fwd_bytes"], out["bwd_bytes"] = num("sbytes"), num("dbytes")
    out["total_bytes"] = out["fwd_bytes"] + out["bwd_bytes"]
    out["total_pkts"] = out["fwd_pkts"] + out["bwd_pkts"]
    d = num("dur").replace(0, np.nan)
    out["flow_bytes_s"] = out["total_bytes"] / d
    out["flow_pkts_s"] = out["total_pkts"] / d
    out["fwd_iat_mean"], out["bwd_iat_mean"] = num("Sintpkt"), num("Dintpkt")
    out["iat_mean"] = out["duration"] / out["total_pkts"].clip(lower=1)
    out["pkt_len_mean"] = out["total_bytes"] / out["total_pkts"].clip(lower=1)
    skip = {"srcip", "dstip", "sport", "dsport", "proto", "state", "service", "attack_cat", "Label", "Stime", "Ltime"}
    for c in UNSW_COLS:
        if c in skip:
            continue
        out["f_" + _key(c)] = pd.to_numeric(raw[c], errors="coerce").astype("float32")
    for c in ["state", "service"]:
        out["f_" + c] = raw[c].astype("category").cat.codes.astype("float32")
    out = out.replace([np.inf, -np.inf], np.nan).dropna(subset=["ts"])
    out["file"] = path.name
    out["dayfile"] = out["ts"].dt.strftime("%Y-%m-%d")
    return out.reset_index(drop=True)


def load_dir(path: Path, kind: str) -> pd.DataFrame:
    files = sorted(list(Path(path).rglob("*.csv")) + list(Path(path).rglob("*.csv.gz")))
    parts = []
    for f in files:
        if kind == "unsw":
            if not re.match(r"UNSW-NB15_\d\.csv", f.name):
                continue
            parts.append(load_unsw_csv(f))
        else:
            try:
                parts.append(load_cic_csv(f, original_2017=(kind == "cic2017_orig")))
            except ValueError:
                continue
        print(f"  loaded {f.name}: {len(parts[-1]):,} flows, {parts[-1]['label'].mean():.3f} attack", flush=True)
    df = pd.concat(parts, ignore_index=True)
    fcols = [c for c in df.columns if c.startswith("f_")]
    df[fcols] = df[fcols].astype("float32")
    return df.sort_values("ts", kind="stable").reset_index(drop=True)
