#!/usr/bin/env python3
"""
PSI Batch Tester — calls Google PageSpeed Insights API for a list of URLs.

Usage:
    python psi_batch.py --key YOUR_API_KEY --urls urls.txt --runs 5 --strategy both --output results.json

Arguments:
    --key        Google PSI API key (required)
    --urls       Path to a text file with one URL per line, OR a comma-separated list of URLs
    --runs       Number of times to test each URL per strategy (default: 5)
    --strategy   mobile, desktop, or both (default: mobile)
    --output     Output JSON file path (default: psi_results.json)
"""

import argparse
import json
import statistics
import sys
import time
import urllib.request
import urllib.parse
import urllib.error

PSI_API = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"

# ── metric extraction helpers ─────────────────────────────────────────────────

def safe_get(d, *keys, default=None):
    for k in keys:
        if not isinstance(d, dict):
            return default
        d = d.get(k, {})
    return d if d != {} else default


def extract_lab(data):
    audits = safe_get(data, "lighthouseResult", "audits") or {}
    perf = safe_get(data, "lighthouseResult", "categories", "performance", "score")
    lcp_ms = safe_get(audits, "largest-contentful-paint", "numericValue")
    tbt_ms = safe_get(audits, "total-blocking-time", "numericValue")
    cls    = safe_get(audits, "cumulative-layout-shift", "numericValue")
    fcp_ms = safe_get(audits, "first-contentful-paint", "numericValue")
    return {
        "lcp_ms":   round(lcp_ms, 1)  if lcp_ms  is not None else None,
        "tbt_ms":   round(tbt_ms, 1)  if tbt_ms  is not None else None,
        "cls":      round(cls, 4)     if cls      is not None else None,
        "fcp_ms":   round(fcp_ms, 1)  if fcp_ms  is not None else None,
        "perf_score": round(perf * 100) if perf is not None else None,
    }


def extract_field(data):
    """Extract CrUX field data. Returns None values if no field data available."""
    exp = data.get("loadingExperience", {})
    metrics = exp.get("metrics", {})
    overall = exp.get("overall_category", None)

    def crux_p75(key):
        m = metrics.get(key, {})
        return m.get("percentile", None)

    return {
        "lcp_p75_ms": crux_p75("LARGEST_CONTENTFUL_PAINT_MS"),
        "inp_p75_ms": crux_p75("INTERACTION_TO_NEXT_PAINT"),
        "cls_p75":    crux_p75("CUMULATIVE_LAYOUT_SHIFT_SCORE"),
        "fcp_p75_ms": crux_p75("FIRST_CONTENTFUL_PAINT_MS"),
        "overall":    overall,
    }


# ── API call with retry ───────────────────────────────────────────────────────

def call_psi(url, key, strategy, attempt=1, max_attempts=3):
    params = urllib.parse.urlencode({
        "url": url,
        "key": key,
        "strategy": strategy,
        "category": "performance",
    })
    req_url = f"{PSI_API}?{params}"
    try:
        with urllib.request.urlopen(req_url, timeout=60) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        if e.code == 429 and attempt < max_attempts:
            wait = 30 * attempt
            print(f"  Rate limited. Waiting {wait}s before retry {attempt+1}/{max_attempts}...", flush=True)
            time.sleep(wait)
            return call_psi(url, key, strategy, attempt + 1, max_attempts)
        raise RuntimeError(f"PSI API error {e.code}: {body[:200]}")
    except Exception as e:
        if attempt < max_attempts:
            wait = 10 * attempt
            print(f"  Request failed ({e}). Retrying in {wait}s...", flush=True)
            time.sleep(wait)
            return call_psi(url, key, strategy, attempt + 1, max_attempts)
        raise


# ── median helper (ignores None) ─────────────────────────────────────────────

def median_or_none(values):
    clean = [v for v in values if v is not None]
    return round(statistics.median(clean), 4) if clean else None


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="PSI Batch Tester")
    parser.add_argument("--key",      required=True, help="Google PSI API key")
    parser.add_argument("--urls",     required=True, help="File with one URL per line, or comma-separated URLs")
    parser.add_argument("--runs",     type=int, default=5, help="Runs per URL per strategy (default: 5)")
    parser.add_argument("--strategy", default="mobile", choices=["mobile", "desktop", "both"])
    parser.add_argument("--output",   default="psi_results.json")
    parser.add_argument("--delay",    type=float, default=2.0, help="Seconds between API calls (default: 2)")
    args = parser.parse_args()

    # Load URLs
    import os
    if os.path.isfile(args.urls):
        with open(args.urls) as f:
            urls = [line.strip() for line in f if line.strip() and not line.startswith("#")]
    else:
        urls = [u.strip() for u in args.urls.split(",") if u.strip()]

    if not urls:
        print("No URLs found. Check your --urls argument.", file=sys.stderr)
        sys.exit(1)

    strategies = ["mobile", "desktop"] if args.strategy == "both" else [args.strategy]
    total_calls = len(urls) * len(strategies) * args.runs
    print(f"\nPSI Batch Tester")
    print(f"URLs: {len(urls)} | Strategies: {', '.join(strategies)} | Runs per URL: {args.runs}")
    print(f"Total API calls: {total_calls}")
    print(f"Estimated time: {total_calls * args.delay / 60:.1f}–{total_calls * (args.delay + 5) / 60:.1f} minutes\n")

    results = []
    call_num = 0

    for url in urls:
        for strategy in strategies:
            print(f"Testing: {url} [{strategy}]")
            runs_lab  = []
            runs_field = []

            for run in range(1, args.runs + 1):
                call_num += 1
                print(f"  Run {run}/{args.runs} (call {call_num}/{total_calls})...", end=" ", flush=True)
                try:
                    data  = call_psi(url, args.key, strategy)
                    lab   = extract_lab(data)
                    field = extract_field(data)
                    runs_lab.append(lab)
                    runs_field.append(field)
                    print(f"LCP {lab.get('lcp_ms')}ms  TBT {lab.get('tbt_ms')}ms  CLS {lab.get('cls')}  Score {lab.get('perf_score')}")
                except Exception as e:
                    print(f"FAILED: {e}")
                    runs_lab.append({})
                    runs_field.append({})

                if run < args.runs:
                    time.sleep(args.delay)

            # Aggregate medians
            median_lab = {
                k: median_or_none([r.get(k) for r in runs_lab])
                for k in ["lcp_ms", "tbt_ms", "cls", "fcp_ms", "perf_score"]
            }
            # Field data doesn't change run-to-run (same 28-day CrUX window),
            # so just take the last successful response
            last_field = next((r for r in reversed(runs_field) if r), {})

            results.append({
                "url":        url,
                "strategy":   strategy,
                "runs":       args.runs,
                "lab_median": median_lab,
                "field":      last_field,
                "raw_lab":    runs_lab,
            })
            print()

    # Write JSON
    with open(args.output, "w") as f:
        json.dump({"results": results, "meta": {
            "runs": args.runs,
            "strategies": strategies,
            "url_count": len(urls),
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        }}, f, indent=2)

    print(f"\nDone! Results saved to: {args.output}")
    print(f"Run generate_xlsx.py to create the formatted spreadsheet.\n")


if __name__ == "__main__":
    main()
