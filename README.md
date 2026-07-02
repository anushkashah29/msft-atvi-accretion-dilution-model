# Mini M&A Deal Model — Accretion / Dilution Analysis

**Category:** Investment Banking | **Difficulty:** Intermediate | **Duration:** 3–4 weeks

---

## Project Overview

This project models Microsoft's $68.7 billion all-cash acquisition of Activision Blizzard (deal announced January 2022, closed October 13, 2023) using the standard investment banking accretion/dilution framework. The goal is to determine whether the acquisition increased or decreased Microsoft's earnings per share (EPS) — the primary financial health check every board reviews before approving a deal.

> **Accretion/dilution analysis** measures whether pro forma EPS (post-deal) exceeds or falls short of standalone EPS (pre-deal). A deal is **accretive** if pro forma EPS > standalone EPS, and **dilutive** if the reverse is true.

---

## Deal at a Glance

| Item | Detail |
|---|---|
| Acquirer | Microsoft Corporation (MSFT) |
| Target | Activision Blizzard, Inc. (ATVI) |
| Deal Value | $68.7B equity value ($95.00/share) |
| Total Deal Cost | $70.5B (incl. fees + ATVI net debt assumed) |
| Structure | 100% all-cash — no new MSFT shares issued |
| Announcement | January 18, 2022 |
| Close | October 13, 2023 |
| Financing | $20B MSFT balance sheet cash + $50.5B new debt |

---

## Model Structure

The workbook (`MSFT_ATVI_Accretion_Dilution_Model.xlsx`) is organised into six sequential steps, each on its own tab:

### Step 1 — Sources & Uses of Funds
Maps exactly how the $70.5B deal cost is financed. Key finding: the deal is 100% cash — no new MSFT shares are issued, so the EPS denominator (7,434M diluted shares) is fixed throughout the model.

| Source | Amount |
|---|---|
| MSFT balance sheet cash | $20,000M |
| New long-term bonds (10–30 yr) | $35,000M |
| New medium-term notes (3–5 yr) | $10,000M |
| Commercial paper | $5,470M |
| **Total** | **$70,470M** |

### Step 2 — Standalone Income Statements
Establishes the pre-deal baseline for both companies at FY2023E.

| Metric | MSFT | ATVI |
|---|---|---|
| Revenue | $211.9B | $8.8B |
| Net Income | $76,471M | $3,572M |
| Diluted Shares | 7,434M | 724M |
| **Standalone EPS** | **$10.29** | **$4.93** |

MSFT's standalone EPS of **$10.29** is the baseline every pro forma calculation compares against.

### Step 3 — Purchase Price Allocation (PPA) & Amortisation
Under GAAP (ASC 805), the premium paid above book value must be allocated to identifiable intangible assets, which are then amortised through the income statement — creating a recurring EPS headwind post-close.

| Intangible Asset | Fair Value | Useful Life | Annual Amortisation |
|---|---|---|---|
| Developed game titles (CoD, WoW, Overwatch) | $8,500M | 6 yrs | $1,417M |
| Customer relationships (300M+ MAUs) | $3,200M | 10 yrs | $320M |
| In-process R&D | $1,400M | Expensed at close | — |
| Technology platform / engine | $1,100M | 5 yrs | $220M |
| Trade names (Blizzard, King, Activision) | $1,800M | Indefinite | — |
| **Goodwill** | **~$46,000M** | Not amortised | — |
| **Total annual amortisation (GAAP)** | | | **$1,957M** |

After-tax amortisation (at MSFT's 18.8% effective rate): **~$1,589M/yr** — an EPS drag of approximately **-$0.21/share** annually.

### Step 4 — Financing Cost
$50.5B of new debt at a blended rate of ~4.6%, plus $900M of foregone interest income on $20B cash deployed.

| Item | Amount |
|---|---|
| Gross interest on new debt | $2,322M |
| Foregone interest on cash deployed | $900M |
| **Total gross financing cost** | **$3,222M** |
| Tax shield (18.8%) | ($606M) |
| **After-tax financing cost** | **$2,616M** |
| **EPS drag** | **~-$0.35/share** |

### Step 5 — Synergy Assumptions & Realisation Schedule
Synergies are modelled conservatively: cost synergies taken at full credit; revenue synergies discounted 50% to reflect execution risk.

| Synergy Type | Run-Rate | Confidence |
|---|---|---|
| Corporate function overlap | $450M | High |
| Azure cloud migration (ATVI off AWS/GCP) | $380M | High |
| Procurement savings | $180M | Medium |
| Real estate consolidation | $90M | Medium |
| **Total cost synergies** | **$1,100M** | |
| Game Pass subscriber uplift | $800M | Medium |
| Mobile gaming (King + MSFT distribution) | $600M | Low |
| Cross-sell / upsell | $350M | Low |
| International expansion | $250M | Low |
| **Total revenue synergies (post-50% haircut)** | **$1,000M** | |
| **Total run-rate synergies (base case)** | **$2,100M** | |

Synergies ramp over three years: **40% Year 1 → 70% Year 2 → 100% Year 3**.

### Step 6 — Pro Forma EPS & Verdict

| Line Item | Year 1 | Year 2 | Year 3 | Run-Rate |
|---|---|---|---|---|
| MSFT standalone net income | $76,471M | $76,471M | $76,471M | $76,471M |
| (+) ATVI acquired net income | $3,572M | $3,572M | $3,572M | $3,572M |
| (−) After-tax financing cost | ($2,616M) | ($2,616M) | ($2,616M) | ($2,616M) |
| (−) After-tax PPA amortisation | ($1,589M) | ($1,589M) | ($1,589M) | ($1,589M) |
| (+) After-tax synergies | $682M | $1,194M | $1,705M | $1,705M |
| **Pro Forma Net Income** | **~$76,520M** | **~$77,032M** | **~$77,543M** | **~$77,543M** |
| Diluted shares (unchanged) | 7,434M | 7,434M | 7,434M | 7,434M |
| **Pro Forma EPS** | **~$10.29** | **~$10.36** | **~$10.43** | **~$10.43** |
| Standalone EPS | $10.29 | $10.29 | $10.29 | $10.29 |
| **Accretion / (Dilution)** | **~+0.1%** | **~+0.7%** | **~+1.4%** | **~+1.4%** |

The deal is **accretive** across all years, though modestly so in Year 1. The verdict is heavily dependent on synergy realisation — without synergies, the deal would be meaningfully dilutive.

---

## Sensitivity Analysis

The sensitivity table stress-tests the Year 1 accretion/dilution % across two key variables: the blended interest rate on new debt (rows) and total gross synergies realised in Year 1 (columns).

|  | $0M syn | $500M | $1,000M | $1,500M | $2,000M | $2,500M | $3,500M |
|---|---|---|---|---|---|---|---|
| **3.0%** | Dilutive | → | → | → | Slightly + | + | ++ |
| **4.0%** | Dilutive | → | → | Slightly + | + | ++ | +++ |
| **4.6% (base)** | Dilutive | → | → | Slightly + | + | ++ | +++ |
| **5.5%** | Dilutive | → | → | → | Slightly + | + | ++ |
| **6.5%** | Dilutive | → | → | → | → | Slightly + | + |

**Key takeaway:** The financing rate has limited impact on the verdict — a 3.5pp swing in rates (3.0% → 6.5%) shifts Year 1 EPS by only ~$0.19/share. The real swing factor is **synergy realisation**: at base rate (4.6%) with zero synergies the deal is ~0.8% dilutive; with $2.1B run-rate synergies it is ~1.4% accretive; and with $3.5B of synergies it reaches ~2.9% accretion. This makes synergy credibility the single most important assumption to challenge in an interview or board presentation.

---

## EPS Bridge (Year 1)

```
MSFT Standalone EPS          $10.29
(+) ATVI earnings contrib.   +$0.48    ATVI net income / MSFT share count
(−) After-tax interest        -$0.35    $50.5B debt × 4.6% blended, after tax shield
(−) After-tax PPA amort.      -$0.21    $1,957M GAAP charge × (1 − 18.8%)
(+) Year 1 synergies          +$0.09    40% of $2,100M run-rate, after tax
                             ──────
Pro Forma EPS (Year 1)       ~$10.29    +0.1% accretion
```

---

## Deal Memo: Strategic Rationale & Conclusion

**Strategic rationale.** Microsoft acquired Activision to accelerate its gaming strategy across three dimensions: (1) content — adding Call of Duty, World of Warcraft, Overwatch, and Candy Crush to its Game Pass subscription service; (2) mobile — establishing a major footprint in mobile gaming via King (Candy Crush) and Activision Mobile; (3) metaverse/cloud — securing IP and an 300M+ active-user base ahead of the next platform cycle. The deal also represented a defensive move, with Microsoft competing directly against Sony and Apple for gaming wallet share.

**Deal structure.** The 100% all-cash structure at $95.00/share represented a ~45% premium to ATVI's unaffected share price. Microsoft funded the deal using $20B of existing cash and $50.5B of new investment-grade debt — feasible given its Aaa/AAA credit rating and access to debt markets at historically tight spreads. No new shares were issued, keeping the EPS denominator fixed and simplifying the accretion/dilution math.

**Financial conclusion.** The deal is modestly accretive from Year 1 and reaches approximately **+1.4% EPS accretion by Year 3** at base-case synergy assumptions ($2.1B run-rate). At 70% synergy realisation (Year 2 trajectory), the deal is approximately **accretive by ~0.7%**. The financing cost (~$2.6B after-tax annually) and GAAP amortisation (~$1.6B after-tax annually) are significant headwinds; ATVI's earnings contribution ($3.6B) and synergies are the primary offsets. The deal makes financial sense at base case but hinges on synergy delivery — particularly cost synergies (Azure migration and headcount reduction), which are high-confidence and largely within Microsoft's control.

---

## Files

| File | Description |
|---|---|
| `MSFT_ATVI_Accretion_Dilution_Model.xlsx` | Full 6-tab accretion/dilution model |
| `generate_model.py` | Python script that regenerates the Excel workbook from scratch |
| `README.md` | This file |

---

## CV Summary

> Built an accretion/dilution model for Microsoft's $69B Activision acquisition; modelled $2.1B in blended cost and revenue synergies with a 3-year realisation ramp; concluded the deal is approximately 1.4% EPS accretive at full synergy run-rate and ~0.7% accretive at 70% synergy realisation, with financing cost and PPA amortisation as the primary EPS headwinds.
