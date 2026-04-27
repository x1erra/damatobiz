## Manual Scope

Use this process to calculate the APP statement look-through sections manually in Excel.

| Section | Denominator | SMA Treatment |
|---------|-------------|---------------|
| **Portfolio Composition** | Full portfolio MV (AMA + SMA) | Included |
| **Portfolio Breakdown** | AMA MV only | Excluded |
| **Portfolio Diversification** | AMA MV only | Excluded |

Work left to right through the workbook. Do not move numbers into `TOTAL SUMMARY` until the source section ties to its denominator.

## Before You Start – Gather Your Inputs

Have these ready (save copies in a dedicated folder):

| Input File                              | Purpose                                                                 | Where to Get It |
|-----------------------------------------|-------------------------------------------------------------------------|-----------------|
| eCISS portfolio holdings export         | Fund codes, descriptions, market values (CAD), SAA/TAA/SMA flags       | eCISS system |
| `Get Factset Model Codes.csv`           | Maps eCISS Fund Code → FactSet mandate/support file code               | Cinchy Tiles |
| `Asset Class Grouping For SMA.csv`      | Tells you how to classify each SMA holding (Income, Equity, etc.)      | Cinchy Tiles |
| Monthly FactSet mandate files (.XLSX)   | Detailed composite components + Port. Weight % for each AMA holding    | FactSet (monthly files) |


---

## Workbook Setup – One Controlled File

Create **one Excel workbook** per portfolio. Name tabs clearly so anyone can follow the flow:

**Recommended Tab Order (left → right):**

1. **IPS** – Paste raw eCISS holdings here (source of truth for all MV totals)
2. **25001 COMP**, **25016 COMP**, etc. – One tab per AMA mandate file for Portfolio Composition
3. **TOTAL COMP** – Summary of Portfolio Composition (final numbers)
4. **25001 BREAKD**, **25016 BREAKD**, etc. – Detailed breakdown tabs
5. **TOTAL BREAKD** – Portfolio Breakdown summaries
6. **25001 BREAKD (D)**, **25016 BREAKD**, etc. – Detailed breakdown diversification tabs
7. **TOTAL BREAKD (D)** – Portfolio Breakdown Diversification summaries
8. **TOTAL SUMMARY** – Final numbers, should mirror the statement output to 3 basis points variance on % 
9. **Formulas** – Tab with the four formulas for quick reference



---

## Part A – Bring Portfolio Values into the IPS Tab

### What You’re Doing
Establish the single source of truth for every market value that will appear in the final report.

### Actions
1. Create/open the workbook and rename the first tab **IPS**.
2. Paste the eCISS holdings.
3. Keep only these columns (delete the rest for cleanliness):
   - Fund Code
   - Fund Description
   - Total MV (CAD)
   - saa_taa (SAA / TAA / SMA)
4. Add two helper columns:
   - **Mandate Code** (look up in Get Factset Model Codes.csv)
   - **Is AMA?** (formula: `=IF(OR(D2="SAA",D2="TAA"),"Yes","No")`)
5. Filter out any row where Total MV (CAD) = 0.
6. At the bottom, add SUM formulas:
   - Total Portfolio MV = `=SUM(C6:C20)` (adjust range)
   - AMA MV = `=SUMIFS(C6:C20,F6:F20,"Yes")`
   - SMA MV = `=SUMIFS(C6:C20,F6:F20,"No")`
7. Name the three total cells so the rest of the workbook can use clean formulas:
   - Name the Total Portfolio MV cell `TOTAL_PORTFOLIO_MV`
   - Name the AMA MV cell `AMA_MV`
   - Name the SMA MV cell `SMA_MV`

### Before Moving On – Quick Check
- Does IPS Total MV match the portfolio total on the client statement? (Yes = good)
- Are all holdings correctly flagged SAA/TAA/SMA?
- Are the AMA and SMA subtotals clearly visible? (You’ll need them for denominators later)

---

## Part B – Identify & Open the Correct Mandate Files

### What You’re Doing
For every AMA holding, find the exact FactSet file that contains its look-through details.

### Actions
1. Open `Get Factset Model Codes.csv`.
2. For each row in your IPS tab where Is AMA? = “Yes”:
   - Find the Fund Code in the `sales_charge_code` column.
   - Confirm the `saa_taa` matches.
   - Copy the corresponding `mandate_code` (this is the file name you need, e.g. `25017`).
3. Download/open the matching monthly FactSet .XLSX file (use the date closest to your reporting period, usually month-end).
4. For each SMA holding, open `Asset Class Grouping For SMA.csv`, find the fund code, and note the `Portfolio Composition` classification. SMA rows do not need mandate files and are added directly to Portfolio Composition using their full eCISS market value.

**Example Mapping (Keep in Mind)**
| eCISS Fund Code | FactSet Mandate Code |
|-----------------|----------------------|
| 27017           | 25017                |
| 28016           | 25016                |

---

## Part C – Build Portfolio Composition (Full Portfolio View)

### What You’re Doing
Classify the *entire* portfolio (AMA + SMA) into the 8 high-level statement categories.

**Denominator:** Total Portfolio MV (from IPS)  
**Result must equal 100% of the whole portfolio.**

### C1. Create Composition Tabs for Each AMA Mandate
1. For each AMA mandate file, import the tab from FactSet tile into your main book and rename it accordingly **XXXXX COMP** (e.g. `25017 COMP`).
2. Open the grouping to the high-level composition view. In the FactSet outline, this is the level where broad asset-class rows are visible but the individual security-level detail remains collapsed.
   - Equity
   - Fixed Income
   - Cash & Equivalents
   - Currency Forwards
   - Derivatives
   - Preferred
   - [Cash]
   - FDS Outlier
   - etc.

### C2. Apply the First Classification Pass (Asset-Class Detection)
1. Add a new column at the far right called **Composition Class**.
2. In the first visible data row, paste **Formula A** from the Formula Appendix. The workbook template also keeps the formula on the **Formulas** tab so the analyst does not need to retype it.
3. Fill the formula down. The formula looks at column A (the row label) and assigns the correct high-level bucket (Income, Equity, Cash, Other, Liquid Alternatives, etc.).

### C3. Apply the Second Pass (Funds & Alternatives)
Some specific funds need special treatment (Private Alternatives or Liquid Alternatives funds). When first reviewing each mandate file, scan the visible rows for **Funds** and **Alternatives** groupings. Highlight those grouping rows in red so it is obvious that they need to be expanded for the second pass.

1. Open the grouping one level deeper so the individual fund names appear.
2. Stay in the same **Composition Class** column used for Formula A.
3. Paste **Formula B** directly over the Formula A result for the specific fund/alternative rows only.
4. Fill down visible rows only for that expanded fund/alternative section.
5. Do not create a second classification column for this pass. The final **Composition Class** column should contain Formula A results for normal asset-class rows and Formula B results only where Formula B is needed.

### C4. Copy “Visible Cells Only” (Critical Step!)
FactSet files contain hidden rows. If you copy normally you’ll pull in junk data in-between.

After Formula A has been applied and Formula B has overwritten the relevant fund/alternative rows:

1. Select the final result range. (Grabbing column A, B and the single **Composition Class** column)
2. Press `Ctrl + G` → **Special** → **Visible cells only** → OK.
3. `Ctrl + C` (Copy).
4. Go to your **TOTAL COMP** tab and `Ctrl + V` (Paste Values) into the appropriate category column.

### C5. Add SMA Holdings
1. Open `Asset Class Grouping For SMA.csv`.
2. For every SMA row in your IPS tab:
   - Find the Fund Code → read the “Portfolio Composition” column.
   - Add that MV to the matching category in TOTAL COMP.
3. No weighting needed — SMA MV is used as-is.

### C6. Complete TOTAL COMP
At this point, `TOTAL COMP` should combine the classified AMA composition rows from the mandate tabs plus the SMA rows from the SMA reference. For each category, sum the market value assigned to that category.

For AMA rows, use the classified composition output from the mandate tabs. For SMA rows, use the full SMA market value from eCISS and the Portfolio Composition category from the SMA CSV.

Build one clean working table on `TOTAL COMP` and name it `COMP_WORK`:

| Component | Port. Weight | Composition Class | Parent Holding MV | Weighted MV |
|-----------|--------------|-------------------|-------------------|-------------|
| Fixed Income | 42.57 | Income | 21722.45 | formula below |
| SMA holding | 100.00 | Equity | 15000.00 | formula below |

For AMA rows, `Parent Holding MV` comes from the matching eCISS holding and `Port. Weight` comes from the mandate file. For SMA rows, use `100.00` as the Port. Weight because the SMA holding is not being looked through.

In the `Weighted MV` column, paste:

```excel
=[@[Parent Holding MV]]*[@[Port. Weight]]/100
```

Create a clean summary table:

| Category            | Market Value (CAD) | % of Portfolio |
|---------------------|--------------------|----------------|
| Income              | formula below | formula below |
| Equity              | formula below | formula below |
| Balanced            | formula below | formula below |
| Liquid Alternatives | formula below | formula below |
| Sector              | formula below | formula below |
| Cash                | formula below | formula below |
| Other               | formula below | formula below |
| Private Alternatives| formula below | formula below |
| **Total**           | formula below | formula below |

Paste this in `B2`:

```excel
=SUMIFS(COMP_WORK[Weighted MV],COMP_WORK[Composition Class],$A2)
```

Paste this in `C2`:

```excel
=B2/TOTAL_PORTFOLIO_MV
```

Copy `B2:C2` down through the eight categories.

Paste this in `B10`:

```excel
=SUM(B2:B9)
```

Paste this in `C10`:

```excel
=SUM(C2:C9)
```

**Control Check:** Total MV must exactly equal the IPS Total Portfolio MV.

---

## Part D – Build Portfolio Breakdown (Actively Managed Detail)

### What You’re Doing
Show the *actively managed* portion in more detail (sub-asset classes + SAA vs TAA split).

**Denominator:** AMA MV only (from IPS)  
**Result must equal 100% of AMA MV.**

### D1. Prepare Breakdown Tabs
1. Create one **XXXXX BREAKD** tab for each AMA mandate. You can duplicate the matching COMP tab, but it is cleaner to re-copy the FactSet mandate data fresh so the tab does not carry over old helper formulas or highlights.
2. Open the grouping to the breakdown level. This is the level where the broad `Equity` row is expanded into Canadian, US, and International Equity, and where Fixed Income, Cash, Alternatives, and Other rows needed for the statement are visible.
3. Typical rows you’ll see:
   - Equity – Canadian Equities
   - Equity – US Equities
   - Equity – International Equities
   - Fixed Income (and sub-types)
   - Cash & Equivalents
   - Individual private alternative fund names
   - Other / FDS Outlier rows

### D2. Apply the Breakdown Formula
1. Use the first empty column to the right of the FactSet data as a helper column. Name it **Breakdown Class**. If there is no empty column beside the data, insert one.
2. Paste **Formula C** into the first visible row.
3. Fill down **visible rows only**.
4. Copy visible cells only → Paste Values into TOTAL BREAKD.

### D3. Calculate Weighted Market Values
This is where the “Port. Weight” column from FactSet becomes powerful.

Before calculating weighted market values, build the working table directly. Each row should be one retained visible row copied from a mandate tab. After the table is built, convert it to an Excel Table and name it `BREAKD_WORK`.

| Component | Port. Weight | Breakdown Class | Parent Fund Code | Parent Fund Description | Parent Holding MV | saa_taa | Weighted MV | SAA MV | TAA MV |
|-----------|--------------|-----------------|------------------|-------------------------|-------------------|---------|-------------|--------|--------|
| Fixed Income | 42.57 | Income | 27001 | Fixed Income Managed Class F | 21722.45 | SAA | formula below | formula below | formula below |

Then calculate:

- **Weighted MV** = `=[@[Parent Holding MV]]*[@[Port. Weight]]/100`
- **SAA MV** = `=IF([@[saa_taa]]="SAA",[@[Weighted MV]],0)`
- **TAA MV** = `=IF([@[saa_taa]]="TAA",[@[Weighted MV]],0)`

### D4. Summarize in TOTAL BREAKD
Use `SUMIFS` to roll up the working table from D3. In plain English, each row is asking Excel:

> “Find every working-table row where **Breakdown Class** equals this category, then add the matching Weighted MV, SAA MV, or TAA MV.”

Build `TOTAL BREAKD` in this format:

| Category | Total Weighted MV | Strategic % (SAA) | Tactical % (TAA) | Portfolio % | Control Check |
|----------|-------------------|-------------------|------------------|-------------|---------------|
| Income | formula below | formula below | formula below | formula below | formula below |
| International Equity | formula below | formula below | formula below | formula below | formula below |
| US Equity | formula below | formula below | formula below | formula below | formula below |
| Canadian Equity | formula below | formula below | formula below | formula below | formula below |
| Cash | formula below | formula below | formula below | formula below | formula below |
| Other | formula below | formula below | formula below | formula below | formula below |
| Alternatives | formula below | formula below | formula below | formula below | formula below |
| **TOTAL (must = 100%)** | formula below | formula below | formula below | formula below | formula below |

Paste this in `B2`:

```excel
=SUMIFS(BREAKD_WORK[Weighted MV],BREAKD_WORK[Breakdown Class],$A2)
```

Paste this in `C2`:

```excel
=SUMIFS(BREAKD_WORK[SAA MV],BREAKD_WORK[Breakdown Class],$A2)/AMA_MV
```

Paste this in `D2`:

```excel
=SUMIFS(BREAKD_WORK[TAA MV],BREAKD_WORK[Breakdown Class],$A2)/AMA_MV
```

Paste this in `E2`:

```excel
=C2+D2
```

Paste this in `F2`:

```excel
=C2+D2-E2
```

Copy `B2:F2` down through the seven categories.

In the total row, paste these formulas:

`B9`
```excel
=SUM(B2:B8)
```

`C9`
```excel
=SUM(C2:C8)
```

`D9`
```excel
=SUM(D2:D8)
```

`E9`
```excel
=SUM(E2:E8)
```

`F9`
```excel
=C9+D9-E9
```

Use **AMA MV from the IPS tab** as the denominator for Strategic %, Tactical %, and Portfolio %. Do not use Total Portfolio MV here.

**Control Check:** Strategic % + Tactical % must equal Portfolio % for every row.

---

## Part E – Build Portfolio Diversification (Sector & Fixed-Income Detail)

### What You’re Doing
Take the Breakdown one level deeper — sectors (Financials, IT, Energy…) and fixed-income quality buckets (Government, Investment Grade, High Yield).

**Denominator:** Still AMA MV only.

### E1. Prepare Diversification Tabs
Re-import the FactSet mandate data fresh for each diversification tab whenever possible. This avoids carrying over old breakdown formulas, hidden filters, or highlighted rows from the earlier pass.

Rename the tabs clearly, such as **25017 BREAKD (D)**.

Open the grouping to the sector / fixed-income detail level. At this level, you should be able to see rows such as:

- Government
- Investment Grade
- High Yield
- Financials
- Information Technology
- Industrials
- Consumer Discretionary
- Health Care
- Energy
- Communication Services
- Materials
- Consumer Staples
- Real Estate
- Utilities
- Cash
- Cash & Equivalents
- Alternatives
- Private alternative fund names, where applicable

### E2. Apply the Diversification Formula
1. Add helper column **Diversification Class**.
2. Paste **Formula D** (visible rows only).
3. Copy visible cells only → Paste Values.

### E3. Calculate Diversification Weighted Market Values
Build the diversification working table directly from the visible rows copied out of the mandate tabs. After the table is built, convert it to an Excel Table and name it `DIVERSIFICATION_WORK`.

| Component | Port. Weight | Diversification Class | Parent Fund Code | Parent Fund Description | Parent Holding MV | saa_taa | Weighted MV | SAA MV | TAA MV |
|-----------|--------------|-----------------------|------------------|-------------------------|-------------------|---------|-------------|--------|--------|
| Government | 23.00 | Government Bond | 27017 | Tactical Asset Allocation Conservative Bal Class F | 20423.86 | TAA | formula below | formula below | formula below |

Then calculate:

- **Weighted MV** = `=[@[Parent Holding MV]]*[@[Port. Weight]]/100`
- **SAA MV** = `=IF([@[saa_taa]]="SAA",[@[Weighted MV]],0)`
- **TAA MV** = `=IF([@[saa_taa]]="TAA",[@[Weighted MV]],0)`

### E4. Apply the “Small Sector Rule”
After you have preliminary % numbers:
1. Any category whose Portfolio % is **below 2% of AMA MV** → move its MV into the **Other** bucket.
2. Recalculate the affected rows by adding the below-threshold category MV into `Other`, setting the original category MV to zero, and then refreshing the Strategic %, Tactical %, and Portfolio % formulas using AMA MV as the denominator.

This keeps the report clean and client-friendly.

### E5. Finalize TOTAL BREAKD (or a dedicated Diversification section)
Create a clean diversification summary table on `TOTAL_BREAKD_D` using the working table you built in E3. This is the table that will feed the Diversification section on `TOTAL SUMMARY`.

Build the table directly in this format:

| Diversification Category | Total Weighted MV | Strategic % | Tactical % | Portfolio % | Control Check |
|--------------------------|-------------------|-------------|------------|-------------|---------------|
| Government Bond | formula below | formula below | formula below | formula below | formula below |
| Investment Grade Bond | formula below | formula below | formula below | formula below | formula below |
| High Yield Bond | formula below | formula below | formula below | formula below | formula below |
| Financials | formula below | formula below | formula below | formula below | formula below |
| Information Technology | formula below | formula below | formula below | formula below | formula below |
| Industrials | formula below | formula below | formula below | formula below | formula below |
| Consumer Discretionary | formula below | formula below | formula below | formula below | formula below |
| Other | formula below | formula below | formula below | formula below | formula below |
| Health Care | formula below | formula below | formula below | formula below | formula below |
| Energy | formula below | formula below | formula below | formula below | formula below |
| Communication Services | formula below | formula below | formula below | formula below | formula below |
| Materials | formula below | formula below | formula below | formula below | formula below |
| Consumer Staples | formula below | formula below | formula below | formula below | formula below |
| Real Estate | formula below | formula below | formula below | formula below | formula below |
| Utilities | formula below | formula below | formula below | formula below | formula below |
| Cash | formula below | formula below | formula below | formula below | formula below |
| Alternatives | formula below | formula below | formula below | formula below | formula below |
| **TOTAL (must = 100%)** | formula below | formula below | formula below | formula below | formula below |

Paste this in `B2`:

```excel
=SUMIFS(DIVERSIFICATION_WORK[Weighted MV],DIVERSIFICATION_WORK[Diversification Class],$A2)
```

Paste this in `C2`:

```excel
=SUMIFS(DIVERSIFICATION_WORK[SAA MV],DIVERSIFICATION_WORK[Diversification Class],$A2)/AMA_MV
```

Paste this in `D2`:

```excel
=SUMIFS(DIVERSIFICATION_WORK[TAA MV],DIVERSIFICATION_WORK[Diversification Class],$A2)/AMA_MV
```

Paste this in `E2`:

```excel
=C2+D2
```

Paste this in `F2`:

```excel
=C2+D2-E2
```

Copy `B2:F2` down through the 17 categories.

In the total row, paste these formulas:

`B19`
```excel
=SUM(B2:B18)
```

`C19`
```excel
=SUM(C2:C18)
```

`D19`
```excel
=SUM(D2:D18)
```

`E19`
```excel
=SUM(E2:E18)
```

`F19`
```excel
=C19+D19-E19
```

Use **AMA MV from the IPS tab** as the denominator. Do not use Total Portfolio MV here.

After the first summary is built, apply the small-sector rule from E4:
1. Identify any diversification category where Portfolio % is below 2%.
2. Move that category’s Total Weighted MV into `Other`.
3. Set the below-threshold category’s Strategic %, Tactical %, and Portfolio % to zero.
4. Recalculate `Other` and the total row.

The final `TOTAL_BREAKD_D` table should pass these checks before it is linked to `TOTAL SUMMARY`:

- Total Weighted MV equals AMA MV.
- Strategic % + Tactical % equals 100%.
- Portfolio % equals 100%.
- Every row’s Control Check equals 0.00%.

---

## Part F – Create the Final Summary Tab

Pull the finished numbers into **TOTAL SUMMARY** so it looks like a real client statement section. The skeleton workbook includes a completed summary layout for Portfolio Composition, Portfolio Breakdown, and Portfolio Diversification.

Typical layout (copy the structure from your practice test file):

- Header with statement title, date, IPS number
- Portfolio Composition table (MV, AMA MV, SMA MV, % of Portfolio)
- Portfolio Breakdown table (Strategic %, Tactical %, Portfolio %)
- Portfolio Diversification table (same % columns)

**Rule:** Never link a number into the final summary until the section has passed its denominator check. This means:

- Portfolio Composition must tie to Total Portfolio MV.
- Portfolio Breakdown must tie to AMA MV.
- Portfolio Diversification must tie to AMA MV.

If the denominator is wrong, the final percentages may look polished but still be wrong.

---

## Part G: Final Review Checklist

Before you hand the file over:

- [ ] IPS total MV = TOTAL COMP total MV (exact match)
- [ ] Portfolio Composition totals 100.00% of full portfolio
- [ ] AMA MV (from IPS) is used as denominator for Breakdown & Diversification
- [ ] Portfolio Breakdown totals exactly 100.00% of AMA MV
- [ ] Portfolio Diversification totals exactly 100.00% of AMA MV
- [ ] For every row: Strategic % + Tactical % = Portfolio %
- [ ] Equity tie-out: Canadian Equity + US Equity + International Equity ≈ Composition Equity (reasonable tolerance)
- [ ] No blank or “” classification results on real holdings
- [ ] All hidden-row copies used “Visible cells only”
- [ ] Workbook is saved with a clear name using the portfolio IPS number (e.g. `IPSNumber.xlsx`)

---

## Part H: Formula Appendix (Copy-Paste Ready)

**Formula A – Portfolio Composition Asset-Class Detection**  
*(High-level pass – paste into Composition Class column)*

```excel
=IF(INDIRECT("A"&ROW())="Equity","Equity",
IF(INDIRECT("A"&ROW())="Currency Forwards","Other",
IF(INDIRECT("A"&ROW())="Fixed Income","Income",
IF(INDIRECT("A"&ROW())="Cash","Cash",
IF(INDIRECT("A"&ROW())="Derivatives","Other",
IF(INDIRECT("A"&ROW())="Cash & Equivalents","Cash",
IF(INDIRECT("A"&ROW())="Preferred","Other",
IF(INDIRECT("A"&ROW())="FDS Outlier","Other",
IF(INDIRECT("A"&ROW())="[Cash]","Cash",
IF(INDIRECT("A"&ROW())="CI Lawrence Park Alternative Investment Grade Credit Fund","Liquid Alternatives",
""))))))))))
```

**Formula B – Portfolio Composition Funds/Alternatives Detection**  
*(Second pass for specific funds)*

```excel
=IF(INDIRECT("A"&ROW())="Alate I LP, Restricted","Private Alternatives",
IF(INDIRECT("A"&ROW())="Avenue Europe Special Situations Fund V (U.S.), L.P.","Private Alternatives",
IF(INDIRECT("A"&ROW())="Axia U.S. Grocery Net Lease Fund I LP, Restricted","Private Alternatives",
IF(INDIRECT("A"&ROW())="CI Adams Street Global Private Markets Fund (Class I)","Private Alternatives",
IF(INDIRECT("A"&ROW())="CI Lawrence Park Alternative Investment Grade Credit Fund","Liquid Alternatives",
IF(INDIRECT("A"&ROW())="CI PM Growth Fund BL LP (Series I)","Private Alternatives",
IF(INDIRECT("A"&ROW())="CI Private Markets Growth Fund I","Private Alternatives",
IF(INDIRECT("A"&ROW())="HarbourVest Adelaide Feeder E LP","Private Alternatives",
IF(INDIRECT("A"&ROW())="HarbourVest Adelaide Feeder F LP","Private Alternatives",
IF(INDIRECT("A"&ROW())="HarbourVest Adelaide Feeder G LP","Private Alternatives",
IF(INDIRECT("A"&ROW())="HarbourVest Infrastructure Income Cayman Parallel Partnership L.","Private Alternatives",
IF(INDIRECT("A"&ROW())="Institutional Fiduciary Tr Money Mkt Ptf","Equity",
IF(INDIRECT("A"&ROW())="Monarch Capital Partners Offshore VI LP","Private Alternatives",
IF(INDIRECT("A"&ROW())="MSILF PRIME PORTFOLIO-INST","Equity",
IF(INDIRECT("A"&ROW())="T.RX Capital Fund I, LP.","Private Alternatives",
IF(INDIRECT("A"&ROW())="Whitehorse Liquidity Partners V LP","Private Alternatives",
""))))))))))))))))
```

**Formula C – Portfolio Breakdown**  
*(Detailed sub-asset class pass)*

```excel
=IF(OR(
INDIRECT("A"&ROW())="Alate I LP, Restricted",
INDIRECT("A"&ROW())="Avenue Europe Special Situations Fund V (U.S.), L.P.",
INDIRECT("A"&ROW())="Axia U.S. Grocery Net Lease Fund I LP, Restricted",
INDIRECT("A"&ROW())="CI Adams Street Global Private Markets Fund (Class I)",
INDIRECT("A"&ROW())="CI Lawrence Park Alternative Investment Grade Credit Fund",
INDIRECT("A"&ROW())="CI PM Growth Fund BL LP (Series I)",
INDIRECT("A"&ROW())="CI Private Markets Growth Fund I",
INDIRECT("A"&ROW())="Demopolis Equity Partners Fund I, L.P.",
INDIRECT("A"&ROW())="HarbourVest Adelaide Feeder E LP",
INDIRECT("A"&ROW())="HarbourVest Adelaide Feeder F LP",
INDIRECT("A"&ROW())="HarbourVest Adelaide Feeder G LP",
INDIRECT("A"&ROW())="HarbourVest Infrastructure Income Cayman Parallel Partnership L.",
INDIRECT("A"&ROW())="Monarch Capital Partners Offshore VI LP",
INDIRECT("A"&ROW())="T.RX Capital Fund I, LP.",
INDIRECT("A"&ROW())="Whitehorse Liquidity Partners V LP"),"Alternatives",
IF(OR(
INDIRECT("A"&ROW())="Cash & Equivalents",
INDIRECT("A"&ROW())="Institutional Fiduciary Tr Money Mkt Ptf",
INDIRECT("A"&ROW())="MSILF PRIME PORTFOLIO-INST",
INDIRECT("A"&ROW())="[Cash]"),"Cash",
IF(OR(
INDIRECT("A"&ROW())="Preferred",
INDIRECT("A"&ROW())="Currency Forwards",
INDIRECT("A"&ROW())="Derivatives",
INDIRECT("A"&ROW())="FDS OUTLIER"),"Other",
IF(INDIRECT("A"&ROW())="Equity - Canadian Equities","Canadian Equity",
IF(INDIRECT("A"&ROW())="Equity - International Equities","International Equity",
IF(INDIRECT("A"&ROW())="Equity - US Equities","US Equity",
IF(INDIRECT("A"&ROW())="Fixed Income","Income","")))))))
```

**Formula D – Portfolio Diversification**  
*(Sector & fixed-income detail pass)*

```excel
=IF(INDIRECT("A"&ROW())="Government","Government Bond",
IF(INDIRECT("A"&ROW())="Investment Grade","Investment Grade Bond",
IF(INDIRECT("A"&ROW())="High Yield","High Yield Bond",
IF(INDIRECT("A"&ROW())="Financials","Financials",
IF(INDIRECT("A"&ROW())="Information Technology","Information Technology",
IF(INDIRECT("A"&ROW())="Industrials","Industrials",
IF(INDIRECT("A"&ROW())="Consumer Discretionary","Consumer Discretionary",
IF(INDIRECT("A"&ROW())="Health Care","Health Care",
IF(INDIRECT("A"&ROW())="Energy","Energy",
IF(INDIRECT("A"&ROW())="Communication Services","Communication Services",
IF(INDIRECT("A"&ROW())="Materials","Materials",
IF(INDIRECT("A"&ROW())="Consumer Staples","Consumer Staples",
IF(INDIRECT("A"&ROW())="Real Estate","Real Estate",
IF(INDIRECT("A"&ROW())="Utilities","Utilities",
IF(INDIRECT("A"&ROW())="Cash","Cash",
IF(INDIRECT("A"&ROW())="Cash & Equivalents","Cash",
IF(INDIRECT("A"&ROW())="[Cash]","Cash",
IF(INDIRECT("A"&ROW())="Alternatives","Alternatives",
IF(INDIRECT("A"&ROW())="Commodities","Other",
IF(INDIRECT("A"&ROW())="Preferred","Other",
IF(INDIRECT("A"&ROW())="Derivatives","Other",
IF(INDIRECT("A"&ROW())="Currency Forwards","Other",
IF(INDIRECT("A"&ROW())="FDS OUTLIER","Other",
IF(INDIRECT("A"&ROW())="Other","Other",
""))))))))))))))))))))))))
```

---
