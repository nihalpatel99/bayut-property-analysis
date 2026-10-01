# Dubai Property Prices: Advanced Excel Practice (Answer Key)

These are worked answers to 25 advanced Excel practice questions, all using `area_prices_monthly.csv`, a monthly dataset of Dubai residential prices by community.

Each answer gives the formula, what it does, and check values, so you can confirm your own result.

---

## Dataset

| Item | Detail |
|---|---|
| Rows | 6,384 = 84 communities × 76 months |
| Period | Jan 2020 – Apr 2026 (2026 is a **partial year**: Jan–Apr only) |
| Communities / zones | 84 communities grouped into 52 zones |
| Freehold | 74 freehold communities, 10 leasehold |

| Column | Meaning |
|---|---|
| `year_month` | Month (text `"2020-01"` in the raw CSV) |
| `community`, `zone` | Location |
| `is_freehold` | TRUE/FALSE |
| `secondary_price_per_sqft_usd` | **Resale price** ("secondary market" means properties bought from an existing owner) |
| `secondary_price_per_m2_usd` | Resale price per m² (= sq ft price × 10.764) |
| `offplan_price_per_sqft_usd` | New units bought from the developer before completion. **Blank for leasehold communities** |
| `rental_price_per_sqft_annual_usd` | Annual rent per sq ft |
| `n_listings_secondary` / `_offplan` / `_rental` | Listing counts |
| `cbuae_base_rate_pct` | UAE Central Bank base rate. **Market-wide: the same for all 84 communities in a given month** |
| `avg_mortgage_rate_pct` | Always base rate + 1.5 |

**Things to know about this data**
- The rent column is set at roughly **6.5% of price** for every row (range 6.1–7.0%), so gross yields barely vary. The yield formulas below are correct, but comparisons between yields (Q7, Q22, Q25) don't reveal much. This looks like a generated practice dataset.
- Rows are sorted by community, then by month. Some shortcuts below rely on this, and each one says so.

## Conventions used in this answer key

- The data is an Excel Table named **`Table13`** (select the data and press Ctrl+T). Rename it if you like, but update the formulas to match.
- From **Q1 onward, `year_month` holds real dates** (the 1st of each month). Many formulas depend on this.
- Inside the table, `[@column]` means **this row's value** and `[column]` means the **whole column**. Mixing them up causes `#SPILL!`.
- Formulas are shown on one line, so they paste cleanly. Copying multi-line formulas can bring in invisible characters that cause *"The syntax of this name isn't correct"*.
- LET variables use descriptive names (`curMonth`, not `d`). Short names can clash with cell references.

---

## Section 1: Cleaning and helper columns

### Q1. Convert `year_month` to a real date
Add a helper column, then paste its values over `year_month`:
```
=DATEVALUE([@year_month]&"-01")
```
Copy the helper column, **Paste Special → Values** over `year_month`, delete the helper, and format the column as `mmm-yyyy`.

**Power Query alternative:** select the column → **Transform → Data Type → Date**.

**Check:** `=ISNUMBER(INDEX(Table13[year_month],1))` → **TRUE**.

> Formatting a text cell as a date does **not** convert it: the cell stays text. Use `ISNUMBER()` to check what a column really holds.

### Q2. Check that m² price = sq ft price × 10.764
Helper column:
```
=ABS([@secondary_price_per_m2_usd]/([@secondary_price_per_sqft_usd]*10.764)-1)
```
Count the rows off by more than 1%:
```
=COUNTIFS(Table13[m2_check],">0.01")
```
or, without a helper column:
```
=SUMPRODUCT(--(ABS(Table13[secondary_price_per_m2_usd]/(Table13[secondary_price_per_sqft_usd]*10.764)-1)>0.01))
```
**Answer: 0 rows.** The largest deviation is **0.45%** (International City, 2020-03: 109 per sq ft vs 1,168 per m²). There are 277 rows above 0.2% and 1,640 above 0.1%. The small gaps come from rounding: the m² figure was calculated from the unrounded sq ft price, so cheap communities show the biggest percentage differences.

### Q3. Gross yield and off-plan premium
```
Gross Yield:       =[@rental_price_per_sqft_annual_usd]/[@secondary_price_per_sqft_usd]
Off-plan Premium:  =IF([@is_freehold],[@offplan_price_per_sqft_usd]/[@secondary_price_per_sqft_usd]-1,"")
```
`IFERROR` would also hide the errors on leasehold rows, but it hides **every** kind of error. The `IF` version only skips the rows where a blank is expected.

**Check (Palm Jumeirah, Jan 2020):** yield **6.50%** (43 / 662), premium **13.6%** (752 / 662 − 1).

### Q4. Every community has 76 months
```
=LET(comms,UNIQUE(Table13[community]),SUM(--(COUNTIFS(Table13[community],comms)<>76)))
```
**Answer: 0.** No community has a missing month. `ROWS(Table13)` = 6,384 = 84 × 76.

---

## Section 2: Lookups and conditional aggregation

### Q5. Input panel with a two-criteria XLOOKUP
Community dropdown in **T2**, month dropdown in **T3** (Data Validation → List, source `=SORT(UNIQUE(Table13[community]))`).
```
Resale:    =XLOOKUP(1,(Table13[community]=$T$2)*(Table13[year_month]=$T$3),Table13[secondary_price_per_sqft_usd],"No data")
Rent:      same formula, return Table13[rental_price_per_sqft_annual_usd]
Listings:  same formula, return Table13[n_listings_secondary]
Yield:     =Rent/Resale
```
**How it works:** each comparison returns TRUE/FALSE for all 6,384 rows. Multiplying them turns TRUE/FALSE into 1/0, so a row scores 1 only when **both** conditions are true. `XLOOKUP(1, …)` then finds that row.

**Check (Palm Jumeirah, Jan 2020):** 662 · 43 · 6.5% · 26 listings.

> `T3` must be the same type as the column (date vs date). A date never equals the text `"2020-01"`, so a mismatch returns "No data".

### Q6. Year-on-year growth using EDATE
Shift the **month you search for**, not the result:
```
=XLOOKUP(1,(Table13[community]=$T$2)*(Table13[year_month]=EDATE($T$3,-12)),Table13[secondary_price_per_sqft_usd],"No data")
```
YoY % = `current / previous − 1`.

**Check (Palm Jumeirah, Jan 2021):** 682 vs 662 → **+3.0%**. Any month in 2020 returns "No data", because there's no earlier year in the dataset.

> ❌ `=EDATE(XLOOKUP(…),-12)` treats the **price** (e.g. 682) as a date serial number, which is wrong.

**Type-agnostic version** (works whether the column is text or dates):
```
=LET(prev,TEXT(EDATE(--(TEXT($T$3,"yyyy-mm")&"-01"),-12),"yyyy-mm"),XLOOKUP(1,(Table13[community]=$T$2)*(TEXT(Table13[year_month],"yyyy-mm")=prev),Table13[secondary_price_per_sqft_usd],"No data"))
```

### Q7. Zone with the highest average yield in 2025
```
=LET(zones,UNIQUE(Table13[zone]),yld,AVERAGEIFS(Table13[Gross Yield],Table13[zone],zones,Table13[year_month],">="&DATE(2025,1,1),Table13[year_month],"<="&DATE(2025,12,1)),TAKE(SORTBY(HSTACK(zones,yld),yld,-1),5))
```
**Answer:** Deira **6.56%**, then Reem and Dubai Studio City (6.54%). Lowest: Mirdif (6.44%). The spread is tiny because rent is set at about 6.5% of price throughout.

### Q8. Freehold vs leasehold average price, latest month
```
=AVERAGEIFS(Table13[secondary_price_per_sqft_usd],Table13[is_freehold],TRUE,Table13[year_month],DATE(2026,4,1))
```
**Answer (Apr 2026):** freehold **617.86** (74 communities), leasehold **311.90** (10). Freehold is about 2× the price per sq ft.

> ❌ Typing `2026-4` without quotes is arithmetic: it equals **2022**. Use `DATE(2026,4,1)` for dates, or `"2026-04"` if the column is text.

---

## Section 3: Dynamic arrays and LAMBDA

### Q9. Top 10 communities by growth, Jan 2020 to Apr 2026
```
=LET(c,UNIQUE(Table13[community]),s,XLOOKUP(c,Table13[community],Table13[secondary_price_per_sqft_usd]),e,XLOOKUP(c,Table13[community],Table13[secondary_price_per_sqft_usd],,0,-1),g,e/s-1,TAKE(SORTBY(HSTACK(c,s,e,g),g,-1),10))
```
- `s`: XLOOKUP returns the **first** match, which is Jan 2020 (relies on the sort order).
- `e`: the last argument `-1` searches **bottom-up**, so it returns the last match, Apr 2026.

| Rank | Community | Start | End | Growth |
|---|---|---|---|---|
| 1 | Remraam | 129 | 242 | **87.6%** |
| 2 | Green Community | 210 | 393 | 87.1% |
| 3 | Living Legends | 208 | 388 | 86.5% |
| 4 | Falcon City of Wonders | 189 | 351 | 85.7% |
| 5 | Mira | 221 | 410 | 85.5% |
| 6 | Majan | 165 | 305 | 84.8% |
| 7 | The Greens | 277 | 510 | 84.1% |
| 8 | Jumeirah 3 | 464 | 846 | 82.3% |
| 9 | Dubai Production City | 152 | 277 | 82.2% |
| 10 | World Islands | 683 | 1,244 | 82.1% |

**Lowest growth:** Pearl Jumeira **60.0%**, then Madinat Jumeirah Living (63.0%) and Emirates Hills (63.6%). All three are expensive, premium communities. Use `TAKE(…,-3)` to get the bottom rows.

### Q10. Off-plan premium over 20%, latest month
```
=LET(latest,INDEX(Table13[year_month],ROWS(Table13)),prem,Table13[offplan_price_per_sqft_usd]/Table13[secondary_price_per_sqft_usd]-1,keep,(Table13[year_month]=latest)*Table13[is_freehold]*(prem>0.2),SORTBY(FILTER(HSTACK(Table13[community],prem),keep,"None above threshold"),FILTER(prem,keep,0),-1))
```
**Answer: none.** The highest premium in Apr 2026 is 15.0%. The third argument of FILTER shows a message instead of `#CALC!` when nothing matches.

| Threshold | Result |
|---|---|
| > 20% | None |
| > 15% | 1: **Arabian Ranches** 15.0% (672 → 773) |
| > 10% | 42 communities, led by Arabian Ranches, Umm Suqeim (14.5%), La Mer (14.4%) |

### Q11. CAGR as a LAMBDA
**Formulas → Name Manager → New**, Name `CAGR`, Refers to:
```
=LAMBDA(start,end,years,(end/start)^(1/years)-1)
```
Apply it to every community with MAP:
```
=LET(c,UNIQUE(Table13[community]),s,XLOOKUP(c,Table13[community],Table13[secondary_price_per_sqft_usd]),e,XLOOKUP(c,Table13[community],Table13[secondary_price_per_sqft_usd],,0,-1),cagr,MAP(s,e,LAMBDA(a,b,CAGR(a,b,6.25))),SORTBY(HSTACK(c,cagr),cagr,-1))
```
Jan 2020 to Apr 2026 is 75 months = **6.25 years**. BYROW alternative: `BYROW(HSTACK(s,e),LAMBDA(r,CAGR(INDEX(r,1),INDEX(r,2),6.25)))`.

**Check:** `=CAGR(129,242,6.25)` → **10.59%** (Remraam, highest). Lowest is Pearl Jumeira **7.81%**, Palm Jumeirah is **10.01%**, and the average is about 9.3%.

### Q12. Community × year matrix of average price
Excel 365 with PIVOTBY:
```
=PIVOTBY(Table13[community],YEAR(Table13[year_month]),Table13[secondary_price_per_sqft_usd],AVERAGE)
```
Fallback without PIVOTBY:
```
=LET(c,SORT(UNIQUE(Table13[community])),y,SEQUENCE(1,7,2020),MAKEARRAY(ROWS(c),7,LAMBDA(r,k,AVERAGEIFS(Table13[secondary_price_per_sqft_usd],Table13[community],INDEX(c,r),Table13[year_month],">="&DATE(INDEX(y,k),1,1),Table13[year_month],"<="&DATE(INDEX(y,k),12,1)))))
```
**Check (Palm Jumeirah):** 2020 **653.92** · 2021 718.92 · 2022 925.25 · 2023 1,110.83 · 2024 1,164.75 · 2025 1,157.42 · 2026 1,178.50.

---

## Section 4: Time series

### Q13. 3-month rolling average without crossing community boundaries
New table column:
```
=LET(curMonth,[@year_month],startMonth,EDATE(curMonth,-2),cnt,COUNTIFS([community],[@community],[year_month],">="&startMonth,[year_month],"<="&curMonth),rollAvg,AVERAGEIFS([secondary_price_per_sqft_usd],[community],[@community],[year_month],">="&startMonth,[year_month],"<="&curMonth),IF(cnt=3,rollAvg,NA()))
```
- Selecting rows **by community and date**, not by position, stops the window from reaching into the previous community's rows.
- `cnt=3` returns a value only when all 3 months exist. `NA()` leaves gaps in charts instead of plotting zeros.

**Why `=AVERAGE(E77:E79)` fails at the boundaries:**

| Row | Month | Community | Price | Naive 3-row average |
|---|---|---|---|---|
| 77 | 2026-04 | Palm Jumeirah | 1,202 | 1,186 ✓ |
| 78 | 2020-01 | Jumeirah Bay Island | 989 | **1,125.33 ✗** |
| 79 | 2020-02 | Jumeirah Bay Island | 951 | **1,047.33 ✗** |
| 80 | 2020-03 | Jumeirah Bay Island | 1,009 | 983 ✓ |

**Check:** Palm Jumeirah Mar 2020 **677**, Apr 2020 **673.33**, Apr 2026 **1,186**. Jumeirah Bay Island Jan and Feb 2020 → `#N/A`, Mar 2020 → **983**.

Common errors:
- `#SPILL!`: you used `[year_month]` where you needed `[@year_month]`.
- *"Syntax of this name isn't correct"*: invisible characters from pasting, or a clash with a short LET name.

### Q14. Month of peak price for each community
```
=LET(comms,SORT(UNIQUE(Table13[community])),peak,MAXIFS(Table13[secondary_price_per_sqft_usd],Table13[community],comms),peakMonth,MAP(comms,peak,LAMBDA(cm,pk,XLOOKUP(1,(Table13[community]=cm)*(Table13[secondary_price_per_sqft_usd]=pk),Table13[year_month]))),HSTACK(comms,peak,peakMonth))
```
MAP is **required** here because the two-criteria test compares a whole column with one community at a time. Format the third column as `mmm-yyyy`.

| Community | Peak | Month |
|---|---|---|
| Palm Jumeirah | 1,223 | Dec 2023 |
| Jumeirah Bay Island | 1,777 | Oct 2023 |
| Remraam | 256 | Oct 2023 |
| Pearl Jumeira | 1,358 | Dec 2023 |

**Finding:** nearly every community peaked in late 2023 or 2024 (22 of the 84 peaked in Oct 2023), and prices have eased since. **Ties:** 5 communities hit their peak price in more than one month (e.g. Al Barsha South, 3 times). Add `,,0,-1` to the XLOOKUP to return the most recent one instead of the first.

### Q15. Month-on-month change and count of falls
MoM column:
```
=LET(prevPrice,SUMIFS([secondary_price_per_sqft_usd],[community],[@community],[year_month],EDATE([@year_month],-1)),IF(prevPrice=0,"",[@secondary_price_per_sqft_usd]/prevPrice-1))
```
With one row per community per month, SUMIFS returns that single price, and it returns 0 (not an error) when there's no earlier month.

Count the falls:
```
=LET(comms,SORT(UNIQUE(Table13[community])),falls,COUNTIFS(Table13[community],comms,Table13[MoM],"<0"),SORTBY(HSTACK(comms,falls),falls,-1))
```
**Check:** Palm Jumeirah MoM: Feb 2020 +0.6%, Mar +5.6%, Apr **−7.4%**. Palm Jumeirah has **30** falls out of 75 months. Most falls: Bukadra and Madinat Jumeirah Living (**36**). Fewest: Dubai Studio City and Dubai South (**25**). Average about 30.

**Finding:** every community grew 60–88%, but prices still fell in about 40% of months. The monthly series is noisy.

### Q16. Index to Jan 2020 = 100 and chart 5 communities
Index column:
```
=[@secondary_price_per_sqft_usd]/XLOOKUP([@community],[community],[secondary_price_per_sqft_usd])*100
```
Chart grid (months down the side, communities across the top):
- `H1:L1`: five community dropdowns
- `G2`: `=SORT(UNIQUE(Table13[year_month]))`
- `H2` (copy across to L2): `=SUMIFS(Table13[Index],Table13[community],H$1,Table13[year_month],$G2#)`

Select G1:L77 → **Insert → Line chart**.

| Community | Mar 2020 | Peak (month) | Apr 2026 |
|---|---|---|---|
| Remraam | 103.9 | 198.4 (Oct 2023) | **187.6** |
| Palm Jumeirah | 106.2 | 184.7 (Dec 2023) | **181.6** |
| Arabian Ranches | 100.3 | 178.5 (Sep 2024) | **174.1** |
| Jumeirah Bay Island | 102.0 | 179.7 (Oct 2023) | **170.2** |
| Pearl Jumeira | 100.3 | 178.2 (Dec 2023) | **160.0** |

---

## Section 5: PivotTables and Power Pivot

### Q17. Average price by zone × year, with slicer and timeline
Rows `zone`, Columns year (group the dates → Years), Values `secondary_price_per_sqft_usd` set to **Average**. Then **Insert Slicer → is_freehold** and **Insert Timeline → year_month**.

**Check (zone Jumeirah):** 2020 **580.25** · 2021 646.09 · 2022 831.36 · 2023 1,012.03 · 2024 1,046.56 · 2025 1,031.56 · 2026 1,028.62.

### Q18. Show Values As
Values: `n_listings_secondary` as **Sum**.
- **% of Column Total:** each zone's share of the year's listings.
- **Difference From:** Base field = Year, Base item = **(previous)**.

| Zone Jumeirah | 2020 | 2021 | 2022 | 2023 | 2024 | 2025 | 2026 |
|---|---|---|---|---|---|---|---|
| Sum | 9,382 | 9,058 | 8,774 | 9,054 | 9,484 | 9,127 | 3,186 |
| % of column | 9.75% | 9.50% | 8.91% | 9.32% | 9.82% | 9.35% | 9.64% |
| Diff. from prev. | (blank) | −324 | −284 | +280 | +430 | −357 | **−5,941** |

Grand total per year: about 96–98k (2020: **96,232**).

> ⚠️ The **−5,941** for 2026 is caused by the partial year (Jan–Apr only), not by a real fall. Filter out 2026, or compare Jan–Apr 2026 with Jan–Apr 2025. % of Column Total isn't affected.

### Q19. Date table, relationship and DAX measures
1. Build **DimDate**: in A2, `=SEQUENCE(DATE(2026,12,31)-DATE(2020,1,1)+1,1,DATE(2020,1,1))` (2,557 days). Copy it and **Paste Values in the same place**, add the `Date` header, then add `Year =YEAR([@Date])`, `MonthNum =MONTH([@Date])` and `Month =TEXT([@Date],"mmm")`. Press Ctrl+T, name the table `DimDate`, and add it to the Data Model. The alternative is Power Pivot → Design → Date Table → New.
2. **Mark as Date Table**, and set **Month → Sort by Column → MonthNum**.
3. Diagram View: drag `Table13[year_month]` onto `DimDate[Date]` (one-to-many).

```
Avg Price := AVERAGE(Table13[secondary_price_per_sqft_usd])

Price YoY % :=
VAR PrevPrice = CALCULATE([Avg Price], SAMEPERIODLASTYEAR(DimDate[Date]))
RETURN DIVIDE([Avg Price] - PrevPrice, PrevPrice)

Avg Yield := AVERAGEX(Table13, DIVIDE(Table13[rental_price_per_sqft_annual_usd], Table13[secondary_price_per_sqft_usd]))
```
- `DIVIDE(cur, prev) - 1` returns **−100%** when there's no prior year (BLANK − 1). `DIVIDE(cur - prev, prev)` returns blank instead.
- `AVERAGEX` averages the row-level yields, so every community-month counts equally. `DIVIDE(SUM(rent), SUM(price))` gives more weight to expensive communities.
- *Circular dependency* error means Avg Yield was created as a **calculated column**. Delete the column and create it as a **measure** (Power Pivot → Measures → New Measure).

| Year | Avg Price | Price YoY % | Avg Yield |
|---|---|---|---|
| 2020 | 325.38 | (blank) | 6.50% |
| 2021 | 360.89 | 10.9% | 6.50% |
| 2022 | 462.61 | 28.2% | 6.50% |
| 2023 | 565.84 | 22.3% | 6.50% |
| 2024 | 587.23 | 3.8% | 6.50% |
| 2025 | 576.52 | −1.8% | 6.50% |
| 2026 | 579.64 | 0.5% | 6.50% |

Month level: **Apr 2026 = 581.44, +0.6%** vs Apr 2025 (578.14).

### Q20. Summing the mortgage rate gives nonsense
The rate is repeated on all 84 community rows each month, so **Sum** adds it 84 times per month (2025 total = 5,871.6). The fix is a measure that takes the rate once per month, then averages across months:
```
Avg Mortgage Rate := AVERAGEX(VALUES(Table13[year_month]), CALCULATE(MAX(Table13[avg_mortgage_rate_pct])))
```
**Check:** 2025 → **5.825%** (base rate 4.325%).

---

## Section 6: Statistics and what-if

### Q21. Correlation between rates and prices, 2022–2024
Monthly summary:
- `A2`: `=SORT(UNIQUE(Table13[year_month]))`
- `B2`: `=XLOOKUP(A2#,Table13[year_month],Table13[cbuae_base_rate_pct])`
- `C2`: `=AVERAGEIFS(Table13[secondary_price_per_sqft_usd],Table13[year_month],A2#)`

CORREL, self-contained and type-agnostic:
```
=LET(ym,TEXT(Table13[year_month],"yyyy-mm"),months,UNIQUE(ym),yr,--LEFT(months,4),sel,FILTER(months,(yr>=2022)*(yr<=2024)),rate,--XLOOKUP(sel,ym,Table13[cbuae_base_rate_pct]),price,MAP(sel,LAMBDA(m,AVERAGE(FILTER(Table13[secondary_price_per_sqft_usd],ym=m)))),CORREL(rate,price))
```
Scatter chart: select the 2022–2024 rows of B:C → **Insert → Scatter** → **Trendline (Linear)**, with the equation and R² displayed.

**Answer:** CORREL = **+0.93**, trendline **y = 32.07x + 412.16**, **R² = 0.866**. The points form vertical columns because the rate only took a few values (0.4 → 1.4% Apr 2022 → 4.4% Oct 2022 → 5.4% Jul 2023 → 4.65% Jul 2024).

> **Interpretation:** a positive correlation does **not** mean higher rates caused higher prices. Both rose over the same period for different reasons: global inflation drove rates, and the post-COVID boom drove prices. This is a **spurious correlation**. A better test is to correlate the month-on-month **changes** in each series, which removes the shared trend.

### Q22. Regress yield on price
```
=LINEST(Table13[Gross Yield],Table13[secondary_price_per_sqft_usd],TRUE,TRUE)
```
**Answer:** slope ≈ **0** (4 × 10⁻⁹), intercept **6.50%**, R² ≈ **0.000002**. There's no relationship: expensive communities don't have lower yields in this data, because rent is set at about 6.5% of price.

### Q23. Buy-to-let model (1,000 sq ft, Palm Jumeirah, Apr 2026)
| Input | Value |
|---|---|
| Price | 1,202 × 1,000 = **$1,202,000** |
| Annual rent | 78 × 1,000 = **$78,000** ($6,500/month) |
| Down payment 25% | $300,500 → loan **$901,500** |
| Mortgage rate | **5.5%** (`avg_mortgage_rate_pct`) |

```
Monthly payment: =PMT(5.5%/12, 25*12, -901500)   →  $5,536
Monthly cash flow: =6500 - 5536                   →  +$964
```
- **Data Table:** put the rate in a column (e.g. 3.5–7%) and the down payment in a row (15–40%), with the payment formula in the corner cell. Then **What-If Analysis → Data Table**, with Row input = the down-payment cell and Column input = the rate cell.
- **Goal Seek:** set the cash-flow cell to 0 by changing the rent. Break-even rent is **$66,432/yr** ($66.43 per sq ft).

This ignores service charges, fees and vacancy. A real model would include them.

---

## Section 7: Presentation

### Q24. One-page dashboard controlled by a slicer
Slicers only control **PivotTables and PivotCharts**, so the dashboard is built on the Data Model. Extra measures:
```
Latest Price :=
VAR lastD = MAX(Table13[year_month])
RETURN CALCULATE([Avg Price], DimDate[Date] = lastD)

Latest YoY % :=
VAR lastD = MAX(Table13[year_month])
VAR cur  = CALCULATE([Avg Price], DimDate[Date] = lastD)
VAR prev = CALCULATE([Avg Price], DimDate[Date] = EDATE(lastD, -12))
RETURN DIVIDE(cur - prev, prev)

Latest Yield :=
VAR lastD = MAX(Table13[year_month])
RETURN CALCULATE([Avg Yield], DimDate[Date] = lastD)

Base Rate :=
VAR lastD = MAX(Table13[year_month])
RETURN CALCULATE(MAX(Table13[cbuae_base_rate_pct]), DimDate[Date] = lastD)

Price Index :=
VAR firstD = CALCULATE(MIN(Table13[year_month]), ALL(DimDate))
VAR base   = CALCULATE([Avg Price], ALL(DimDate), DimDate[Date] = firstD)
RETURN DIVIDE([Avg Price], base) * 100

Growth % :=
VAR firstD = CALCULATE(MIN(Table13[year_month]), ALL(DimDate))
VAR lastD  = CALCULATE(MAX(Table13[year_month]), ALL(DimDate))
VAR s = CALCULATE([Avg Price], ALL(DimDate), DimDate[Date] = firstD)
VAR e = CALCULATE([Avg Price], ALL(DimDate), DimDate[Date] = lastD)
RETURN DIVIDE(e - s, s)
```
PivotTables on a hidden "Calc" sheet:

| PivotTable | Rows | Columns | Values | Extra |
|---|---|---|---|---|
| KPI | none | none | Latest Price, Latest YoY %, Latest Yield, Base Rate | |
| Index | DimDate[Date] | zone | Price Index | Line PivotChart |
| Top10 | community | none | Growth % | Value Filters → Top 10, sort descending, Bar PivotChart |

- **KPI cards:** insert a shape, select it, and type `=Calc!B4` in the formula bar. The shape then shows that cell's value.
- **Slicer:** insert a `zone` slicer, then right-click it → **Report Connections** → tick all three PivotTables.
- **Tidy up:** hide gridlines and the chart field buttons, and hide the Calc sheet.

| Slicer | Latest price | YoY % | Yield | Base rate | Top bar |
|---|---|---|---|---|---|
| None selected | **581.44** | **+0.57%** | **6.50%** | **4.00%** | Remraam 87.6% |
| Zone = Jumeirah | **1,031.62** | **−0.36%** | **6.51%** | **4.00%** | Jumeirah 3 82.3% … Pearl Jumeira 60.0% |

### Q25. Conditional formatting
- **Colour scale:** select the Q12 matrix → Home → Conditional Formatting → Color Scales.
- **Flag below-median yield for the month:** add a helper column `MonthMedian = MEDIAN(FILTER([Gross Yield],[year_month]=[@year_month]))`, then select the table rows and add the formula rule `=$N2<$O2`, where N is Gross Yield and O is MonthMedian. Conditional formatting rules can't use structured references, so they need normal cell references. About half the rows are flagged, as you'd expect with a median.

---

## Key lessons

1. **Text vs date** caused most of the errors here. Check with `ISNUMBER()`, and remember that formatting doesn't convert a value's type.
2. **Select rows by criteria, not by position.** This prevents boundary errors (Q13, Q15) and survives re-sorting.
3. **Use `[@col]` for this row and `[col]` for the whole column.** Mixing them up causes `#SPILL!`.
4. **Watch for partial periods** when comparing years (Q18, Q19). Compare like with like.
5. **Repeated values need MAX or AVERAGE, not SUM**, in pivots and DAX (Q20, Q24).
6. **Correlation is not causation.** Series that trend together will always correlate (Q21).
7. **Measures, not calculated columns,** for aggregations in the Data Model (Q19).
