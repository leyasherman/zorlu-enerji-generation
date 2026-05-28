# Zorlu Enerji – Plant-Level Monthly Electricity Generation

Historical monthly electricity output (MWh) for each Zorlu Enerji power plant in Turkey, sourced from the official Turkish electricity market transparency platform.

## What this data is

Monthly generation figures (MWh) per individual power plant for **Zorlu Enerji Elektrik Üretim A.Ş. (ZOREN)** and its subsidiaries, covering:

| Plant | Fuel | Capacity (MW) | Location |
|---|---|---|---|
| Kızıldere I / II / III | Geothermal | 15 / 80 / 165 | Denizli |
| Alaşehir | Geothermal | 45 | Manisa |
| İkizdere / Tercan / Mercan / Beyköy / Kuzgun / Çıldır / Ataköy | Hydro | 5–25 each | Various |
| Gökçedağ* | Wind | 135 | Osmaniye |
| Lüleburgaz | Natural Gas | 49.5 | Kırklareli |
| Alaşehir Solar / Kızıldere Solar | Solar | 3.75 / 0.99 | Manisa / Denizli |

*Gökçedağ was sold in December 2025; historical data included.

## Where it comes from

**EPİAŞ Transparency Platform** (`seffaflik.epias.com.tr`)  
Endpoint: UEVM – *Uzlaştırmaya Esas Veriş Miktarı* (settlement-basis metered generation).  
This is the official Turkish electricity market operator's data — the same numbers used for financial settlement of energy contracts.

**Coverage:** May 2019 – present (EPİAŞ began publishing plant-level data on 16 May 2019).

## How to run

### 1. Prerequisites
```bash
pip install -r requirements.txt
```

### 2. Credentials
Edit `.env` and fill in your EPİAŞ account details:
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

The script prints progress and saves checkpoints every 200 rows.  
**Expected runtime: 20–40 minutes** (≈1 000 API calls at 3 req/s).

If interrupted, re-run the same command — it will resume from where it stopped.

## Output files

| File | Description |
|---|---|
| `zorlu_enerji_generation.xlsx` | Main output — analyst-ready Excel |
| `zorlu_generation_raw.csv` | Raw data at UEVCB (sub-unit) level |
| `zorlu_organizations.csv` | Zorlu org IDs discovered in EPIAS |
| `zorlu_uevcbs.csv` | Plant sub-unit IDs and names |

### Excel tabs

| Tab | Contents |
|---|---|
| **Summary** | All plants combined, monthly totals |
| **By Plant (Long)** | One row per plant per month — main analysis tab |
| **Wide Pivot** | Plants as columns, months as rows |
| **By Fuel Type** | Geothermal / Hydro / Wind / Gas / Solar monthly |
| **Raw (UEVCB level)** | Raw API data before plant grouping |
| **Metadata** | Source, date pulled, caveats |
| **Plant Reference** | Plant list with capacity and fuel type |

## Known limitations

- **Pre-2019 data not available** in EPIAS — use Zorlu annual reports for earlier years (annual totals only).
- **Pakistan (Jhimpir, 56.4 MW) and Palestine (Dead Sea, 1.5 MW)** plants are outside Turkey and not in EPIAS.
- API returns net settled generation; may differ slightly (~8–10%) from gross output figures in Zorlu annual reports.
- The script auto-matches UEVCB names to plant names via keyword matching — review `zorlu_uevcbs.csv` to verify the mapping is correct.

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
Unlike many European transparency platforms, EPİAŞ provides no public data without a login token (TGT). Attempted direct API calls to endpoints like `/generation/data/powerplant-list` without a token — all returned HTTP 401. Registration was the only path forward.

**3. Three different ID systems in EPIAS**  
The platform uses at least three different plant identifier types that are not interchangeable:
- `organizationId` — the legal entity (e.g. Zorlu Doğal Elektrik)
- `uevcbId` — the injection/withdrawal sub-unit within an org
- physical plant ID (from `pp-list`) — what `rt-gen-bulk` actually requires

Initial attempts used UEVCB IDs in `rt-gen-bulk` and got empty results. Switched to `pp-list` to get physical plant IDs, which resolved the issue.

**4. `uevm` endpoint silently ignored the plant filter**  
Calling `uevm` with a `pp_id` parameter appeared to work (returned 744 rows) but actually returned national aggregate data — the filter was silently ignored. Only caught by noticing the "Gökçedağ" wind total matched all-Turkey wind, not a single 135 MW plant.

**5. Keyword search returned false positives**  
Searching `pp-list` for Zorlu-related plant names matched 23 entries, of which 8 were unrelated (e.g. TÜRKERLER ALAŞEHİR JES — a different company operating in the same geothermal field). Manual review against the annual report plant list was required.

**6. Annual report PDFs are image-scanned**  
The older TEIAS statistics books are fully scanned PDFs with no extractable text. Zorlu's own annual reports are text-based but the production data tables are in the appendix (document page 419 out of 451), which required calculating the correct PDF page offset (each visible page = 2 document pages).

**7. Kızıldere I shows near-zero production**  
The oldest geothermal unit (15 MW, commissioned ~1984) shows only 8.8 GWh total across 7 years. This appears correct — the unit runs intermittently and its steam output may be attributed to the adjacent Kızıldere II plant in EPIAS reporting.
