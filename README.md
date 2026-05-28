# Zorlu Enerji – Plant-Level Monthly Electricity Generation

Historical monthly electricity output (MWh) for each Zorlu Enerji power plant in Turkey, sourced from the official Turkish electricity market transparency platform.

## What this data is

Monthly generation figures (MWh) per individual power plant for **Zorlu Enerji Elektrik Üretim A.Ş. (ZOREN)** and its subsidiaries, covering:

| Plant | Fuel | Capacity (MW) | Location |
|---|---|---|---|
| Kızıldere I / II / III | Geothermal | 15 / 80 / 165 | Denizli |
| Alaşehir | Geothermal | 45 | Manisa |
| İkizdere / Tercan / Mercan / Beyköy / Kuzgun / Çıldır / Ataköy | Hydro | 5–25 each | Various |
| Gökçedağ† | Wind | 135 | Osmaniye |
| Lüleburgaz‡ | Natural Gas | 49.5 | Kırklareli |

† Gökçedağ was sold to Rönesans Enerji in December 2025. Data in this dataset covers only the period while Zorlu-owned (through 2025-12-31); post-sale rows are excluded.  
‡ Lüleburgaz produced only 7.5 GWh in total (almost entirely in 2019) and has been effectively inactive since — consistent with Zorlu's stated exit from fossil fuels.  
**Not included:** Pakistan (Jhimpir, 56.4 MW wind) and Palestine (Dead Sea, 1.5 MW solar) are outside Turkey and not in EPIAS. Solar hybrid units at Alaşehir (3.75 MW) and Kızıldere (0.99 MW) are below EPIAS registration threshold and absent from this dataset.

## Where it comes from

**EPİAŞ Transparency Platform** (`seffaflik.epias.com.tr`)  
Endpoint: **rt-gen-bulk** – *Santral Bazlı Toplu Gerçek Zamanlı Üretim* (plant-level bulk real-time generation).

This is the EPIAS real-time metered generation feed — what the meter at each plant reads, reported hourly. Note that **UEVM** (Uzlaştırmaya Esas Veriş Miktarı, settlement-basis generation) is a related but distinct dataset used for financial settlement and may differ from real-time figures. This dataset uses rt-gen-bulk, not UEVM.

**Coverage:** 16 May 2019 – present (EPIAS began publishing plant-level data on this date).

## How to run

### 1. Prerequisites
```bash
pip install -r requirements.txt
```

### 2. Credentials
Create a `.env` file in the project folder (copy from `.env.example`):
```
EPTR_USERNAME=your_email@example.com
EPTR_PASSWORD=your_password
```

Register free at [seffaflik.epias.com.tr](https://seffaflik.epias.com.tr).

> **Important:** Registration on the EPİAŞ platform only works from a Turkish IP address.
> If you are outside Turkey, **enable a VPN with a Turkish server** before registering.
> Once the account is created, the VPN is no longer needed — the API works from any location.

### 3. Run
```bash
python fetch_zorlu.py
```

The script prints progress and saves a checkpoint CSV every 50 days so the run can be resumed if interrupted.  
**Expected runtime: ~20 minutes** (≈2,570 API calls at one per day, ~0.35 s/call).

Re-run the same command after an interruption — it resumes from the last saved day.

### 4. Rebuild dashboard (optional)
```bash
python build_dashboard.py
```

Regenerates the Dashboard sheet in the Excel file without re-fetching data.

## Output files

| File | Description |
|---|---|
| `zorlu_enerji_generation.xlsx` | Main output — analyst-ready Excel with 8 sheets |
| `zorlu_generation_raw.csv` | Raw hourly data at physical plant level (date / hour / pp_id / MWh) |

### Excel sheets

| Sheet | Contents |
|---|---|
| **Dashboard** | Visual summary: key metrics, annual production heatmap, capacity factors, seasonality |
| **Guide** | Table of contents explaining every sheet |
| **Summary** | All plants combined, one row per month |
| **By Plant** | One row per plant per month — main analysis sheet |
| **Wide Pivot** | Plants as columns, months as rows — ready for charting |
| **By Fuel Type** | Geothermal / Hydro / Wind / Gas monthly totals |
| **Raw Hourly (sample)** | First 50,000 hourly source readings from EPIAS |
| **Metadata** | Data source, date pulled, coverage, caveats |
| **Plant Reference** | Plant list with EPIAS ID, fuel type, capacity, location |

## Known limitations

- **Pre-2019 data not available** in EPIAS — use Zorlu annual reports for earlier years (annual totals only, fuel-type aggregates).
- **rt-gen-bulk vs UEVM:** This dataset uses real-time metered generation. Settlement figures (UEVM) used in financial contracts may differ slightly.
- **Mercan (Yukarı) and Mercan (Hacı):** Installed capacity for these sub-units is not available in public sources. Sum all three Mercan entries for total Mercan output.
- **2019 is a partial year** — data starts 16 May 2019. 2026 is partial depending on when data was pulled.
- **No native Excel charts** — the Dashboard contains formatted tables with conditional formatting; charts must be built manually in Excel from the data sheets.

---

## Changelog — issues found and fixed

Documented for transparency and reproducibility.

### Data accuracy
| # | Issue | Fix |
|---|---|---|
| 1 | README described endpoint as UEVM (settlement) but code uses rt-gen-bulk (real-time) — different datasets | README and Metadata tab corrected; difference explained |
| 2 | Gökçedağ (sold Dec 2025) had post-sale rows through May 2026 attributed to Zorlu | Rows after 2025-12-31 excluded in aggregation |
| 3 | Mercan sub-units had `capacity_mw: 0.0` — misleading for capacity factor calc | Set to `None`; Metadata note added |
| 4 | 2025 omitted from Capacity Factor table (hardcoded to 2020–2024) | Replaced with dynamic `FULL_YEARS` computed from data |

### Code correctness
| # | Issue | Fix |
|---|---|---|
| 5 | `fetch_day()` returned empty DataFrame for all failures — no distinction between API error, empty response, schema change, partial day | Returns `(df, status, detail)` tuple with 5 status codes |
| 6 | Empty-response days not retried; `failed_days.csv` had no reason column | Empty responses are retried; CSV now has `status` + `detail` |
| 7 | Partial days (< 80% of expected rows) were silently discarded | Partial days saved with `FETCH_PARTIAL` flag |
| 8 | `build_dashboard.py` printed "Charts added" when no charts exist in the workbook | Print statement corrected |
| 9 | Dead code: `_set_col_width`, `ID_TO_META`, `pull_ts`, unused imports | Removed |
| 10 | Dashboard values hardcoded (coverage dates, data point count, top plant) | All derived from data at build time |

### Documentation
| # | Issue | Fix |
|---|---|---|
| 11 | README listed non-existent output files (`zorlu_organizations.csv`, `zorlu_uevcbs.csv`) | Removed |
| 12 | Tab names in README didn't match workbook; `Guide` tab missing from list | Corrected |
| 13 | Solar plants listed in plant table but absent from dataset | Removed from table; absence noted with explanation |
| 14 | API call count stated as ~1,000; checkpoint as "every 200 rows" | Corrected to ~2,570 and "every 50 days" |

### Data quality flags (open — require further investigation)
- **Geothermal 2021 is 10% below the annual report** (expected direction: above). Likely a Kızıldere III plant ID issue for that year. Do not use 2021 geothermal in year-on-year comparisons without verification.
- **Hydro is 16–21% above the annual report** every year. Gross/net explains ~5%; the remaining gap may indicate Yukarı Mercan / Hacı Mercan include output not wholly attributed to Zorlu in EPIAS. Verify before using hydro in a financial model.

---

## What was tried and why it was ruled out

Finding plant-level monthly generation data for a Turkish IPP is not straightforward. Below is the full path taken, including dead ends.

### Sources investigated

| Source | Result |
|---|---|
| **EPİAŞ Transparency Platform** | ✅ Used — plant-level hourly data since May 2019 |
| Zorlu Enerji Annual Reports (PDF) | ✅ Used as cross-check — annual totals by fuel type only, no plant-level |
| TEIAS (grid operator) monthly reports | ❌ National aggregates only, no individual plants |
| TEIAS YTBS statistics portal | ❌ Portal refused connections during collection |
| ENTSO-E Transparency Platform | ❌ Turkey participates as observer only; does not submit unit-level generation data |
| EPIAS monthly market reports (public PDFs) | ❌ National totals by fuel type only, no company or plant breakdown |
| KAP / Borsa Istanbul disclosures | ❌ Financial disclosures only, no operational generation data |
| Global Energy Monitor | ❌ Installed capacity only, no generation output |

### Technical obstacles encountered

**1. EPİAŞ registration blocked outside Turkey**  
The registration page on `seffaflik.epias.com.tr` silently fails without a Turkish IP address. Solved by enabling a VPN with a Turkish server before signing up. After registration, the API works from any location.

**2. All API endpoints require authentication**  
Unlike many European transparency platforms, EPİAŞ provides no public data without a login token (TGT). Attempted direct API calls without a token — all returned HTTP 401. Registration was the only path forward.

**3. Three different ID systems in EPIAS**  
The platform uses at least three different plant identifier types that are not interchangeable:
- `organizationId` — the legal entity (e.g. Zorlu Doğal Elektrik)
- `uevcbId` — the injection/withdrawal sub-unit within an org
- physical plant ID (from `pp-list`) — what `rt-gen-bulk` actually requires

Initial attempts used UEVCB IDs in `rt-gen-bulk` and got empty results. Switched to `pp-list` to get physical plant IDs, which resolved the issue.

**4. `uevm` endpoint silently ignored the plant filter**  
Calling `uevm` with a `pp_id` parameter appeared to work (returned 744 rows) but actually returned national aggregate data — the filter was silently ignored. Only caught by noticing the totals matched all-Turkey generation, not a single 135 MW plant. Switched to `rt-gen-bulk`.

**5. Keyword search returned false positives**  
Searching `pp-list` for Zorlu-related plant names matched 23 entries, of which 8 were unrelated (e.g. TÜRKERLER ALAŞEHİR JES — a different company operating in the same geothermal field). Manual review against the annual report plant list was required.

**6. Annual report PDFs are partially image-scanned**  
Older TEIAS statistics books are fully scanned PDFs with no extractable text. Zorlu's own annual reports are text-based but the production data tables are deep in the appendix, requiring calculation of the correct PDF page offset.

**7. Kızıldere I shows near-zero production**  
The oldest geothermal unit (15 MW, commissioned ~1984) shows only 8.8 GWh total across 7 years. This appears correct — the unit runs intermittently and its steam may be attributed to the adjacent Kızıldere II plant in EPIAS reporting.
