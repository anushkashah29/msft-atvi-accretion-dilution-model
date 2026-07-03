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

## Advanced Analysis

A second workbook (`MSFT_ATVI_Advanced_Analysis.xlsx`) extends the core model with six additional modules, generated by `generate_advanced_analysis.py`. Each module addresses a distinct analytical question not covered by the base accretion/dilution model.

---

### Module 1 — One-Page Deal Memo

A formatted executive summary covering strategic rationale, deal structure, financial verdict, and three complex analytical questions (break-even synergy, financing structure effects, primary risk drivers). Intended as a board-ready or interview-ready single-page output.

---

### Module 2 — Break-even Synergy Solver

**Question:** What is the minimum synergy realization required for the acquisition to be EPS accretive?

**Methodology:** Set pro forma EPS equal to MSFT standalone EPS and solve algebraically for the required after-tax synergy contribution. Convert to pre-tax run-rate synergy at each year's ramp percentage.

**Key findings:**

| Year | Ramp | Required Run-Rate Synergy | Safety Margin vs. Base ($2,100M) |
|---|---|---|---|
| Year 1 | 40% | ~$1,950M | +7.7% |
| Year 2 | 70% | ~$1,114M | +88.5% |
| Year 3 | 100% | ~$780M | +169% |

The algebraic break-even is **~$780M pre-tax run-rate** ($633M after-tax). The base case of $2,100M sits **63% above** break-even, providing a significant margin of safety. Critically, cost synergies alone ($1,100M — Azure migration + headcount overlap) exceed the full-run-rate break-even threshold, meaning revenue synergies represent pure upside with no structural dependence for accretion at steady state.

---

### Module 3 — Financing Structure Optimizer

**Question:** How does the debt / equity / cash mix affect EPS accretion and leverage?

**Methodology:** Model seven financing scenarios ranging from 0% equity (actual: all new debt) to 100% equity (all new MSFT shares at $330/share close price). For each scenario, recalculate new share count, new debt, after-tax interest cost, and pro forma EPS.

**Key finding — the P/E paradox:**

| Financing Mix | New Shares (M) | AT Financing Cost ($M) | Year-1 EPS | Year-1 Accretion |
|---|---|---|---|---|
| 0% equity — all debt (actual) | 0 | $2,616M | $10.292 | +0.06% |
| 25% equity | 38M | $2,144M | $10.304 | +0.18% |
| 50% equity | 77M | $1,674M | $10.314 | +0.28% |
| 100% equity | 153M | $731M | $10.334 | +0.47% |

MSFT's implied earnings yield (E/P = $10.29 / $330 = **3.1%**) is lower than its after-tax cost of debt (4.6% × 81.2% = **3.7%**). This means issuing equity is *less costly* to EPS than borrowing — a counterintuitive result that holds for any high-P/E acquirer. Microsoft chose all-debt for strategic reasons: deal certainty, the interest tax shield, preserving share count optionality, and Aaa-rated access to debt markets at tight spreads.

---

### Module 4 — M&A Premium Sensitivity

**Question:** How does the purchase price paid affect accretion/dilution, and at what price does the deal flip?

**Methodology:** Vary the offer price from $65 to $115/share in $5 increments. Higher price → more new debt → higher interest cost → EPS drag. PPA amortization is held fixed (identified intangible fair values do not change with price; only non-amortized goodwill absorbs the incremental premium).

**Key findings:**

| Price / Share | Premium to Unaffected | Goodwill ($M) | Year-1 EPS | Year-1 Accretion |
|---|---|---|---|---|
| $70 | +7% | ~$35,200M | $10.382 | +0.93% |
| $80 | +22% | ~$42,400M | $10.346 | +0.58% |
| $90 | +37% | ~$49,600M | $10.310 | +0.23% |
| **$95 (actual)** | **+45%** | **~$46,000M** | **$10.292** | **+0.06%** |
| $100 | +53% | ~$60,800M | $10.274 | −0.12% |
| $110 | +68% | ~$68,000M | $10.237 | −0.48% |

The algebraic **break-even purchase price is ~$96.70/share** at base-case synergies. The actual $95.00 price sits $1.70 below break-even — the deal was priced within ~2% of EPS neutrality, making synergy delivery structurally non-optional. Without synergies, the break-even price falls to ~$71.48/share, confirming the deal cannot stand on ATVI's standalone earnings alone at the price paid.

---

### Module 5 — Tornado / Driver Analysis

**Question:** Which assumptions are the primary drivers of transaction value and risk?

**Methodology:** For each of seven key variables, hold all others at base case and compute Year-1 EPS at a plausible low and high. Rank by absolute EPS swing (low → high). WACC is modeled via a PV-annuity adjustment on the synergy stream to reflect how discount rate assumptions affect the time-value of synergy benefits.

**Ranked results (Year-1 EPS swing, largest to smallest):**

| Rank | Driver | Low Case | High Case | EPS Swing |
|---|---|---|---|---|
| 1 | Interest rate on new debt | 3.0% | 6.5% | ~$0.19/share |
| 2 | Purchase price per ATVI share | $70/sh | $110/sh | ~$0.14/share |
| 3 | PPA amortization life (game titles) | 4 years | 10 years | ~$0.14/share |
| 4 | Run-rate synergies | $700M | $3,500M | ~$0.12/share |
| 5 | Foregone yield on cash deployed | 2.5% | 5.5% | ~$0.07/share |
| 6 | WACC (PV of synergy stream) | 7.0% | 11.0% | ~$0.07/share |
| 7 | Effective tax rate | 14% | 24% | ~$0.06/share |

**Conclusions:**
- The **interest rate environment** is the single largest mechanical driver — a reflection of how much leverage was used. The rate cycle between deal announcement (Jan 2022, ~0% FF) and close (Oct 2023, 5.25–5.50% FF) directly raised financing costs.
- **Purchase price** and **PPA amortization life** are tied for second — both reflect judgments made at signing that cannot be revised post-close.
- **Synergy realization** is fourth mechanically but ranks first as a *qualitative* risk because it is an estimate subject to execution uncertainty; every other driver is observable from market data or GAAP rules.
- **Tax rate** has the smallest impact and acts inversely — a lower tax rate reduces the value of interest and PPA deductions, making it mildly negative for EPS despite common intuition.

An embedded horizontal bar chart (tornado-style) is auto-generated in the Excel tab.

---

### Module 6 — Historical Interest Rate Scenarios (2021–2024)

**Question:** How would deal economics have differed had Microsoft closed in a different rate environment?

**Methodology:** Model MSFT's hypothetical blended new-debt rate and foregone cash yield for five rate environments: 2021 (ZIRP), 2022-H1 (hiking begins), 2022-H2 (rapid acceleration), 2023-Q4 (actual close, peak cycle), and 2024 (cuts begin). All other assumptions held at base case.

| Scenario | FF Rate | MSFT Blended Rate | AT Financing Cost | Year-1 EPS | Year-1 Accretion | vs. Actual |
|---|---|---|---|---|---|---|
| 2021 (ZIRP) | 0.00–0.25% | ~1.6% | ~$998M | ~$10.43 | +1.40% | +$0.14/share |
| 2022-H1 (hiking) | 0.25–2.50% | ~3.0% | ~$1,834M | ~$10.36 | +0.72% | +$0.07/share |
| 2022-H2 (accelerating) | 2.50–4.50% | ~4.0% | ~$2,404M | ~$10.29 | +0.28% | +$0.03/share |
| **2023-Q4 (actual)** | **5.25–5.50%** | **~4.6%** | **~$2,616M** | **~$10.29** | **+0.06%** | **—** |
| 2024 (cuts begin) | 4.25–4.50% | ~4.3% | ~$2,453M | ~$10.30 | +0.20% | +$0.01/share |

**Conclusions:**
- The Fed's hiking cycle imposed approximately **$0.14/share of annual EPS cost** relative to what the deal would have achieved in the 2021 ZIRP environment. Over a 5-year hold, this compounds to ~$0.70/share of cumulative EPS foregone.
- The **2022-H1 window** (deal announcement month) would have been the optimal close — rates had not yet risen materially, and deal terms were already set. The 21-month FTC litigation delay forced the close into the most expensive rate environment in 40 years.
- The 2024 post-cut scenario offers only marginal improvement (~$0.01/share) over actual, as FF rates settled at 4.25–4.50% — still historically elevated relative to the 2010–2021 era.
- This analysis illustrates why deal timeline management (regulatory approvals, close speed) carries real financial cost: every month of delay in 2022–2023 meant more interest accrual at rising rates.

---

## Files

| File | Description |
|---|---|
| `MSFT_ATVI_Accretion_Dilution_Model.xlsx` | Core 6-tab accretion/dilution model |
| `MSFT_ATVI_Advanced_Analysis.xlsx` | Advanced analysis workbook (6 modules) |
| `generate_model.py` | Python script that regenerates the core Excel model |
| `generate_advanced_analysis.py` | Python script that generates the advanced analysis workbook |
| `README.md` | This file |
