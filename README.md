# Bayut Property Prices Analysis


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



## Section 1: Cleaning and helper columns

### Q1. Convert `year_month` to a real date

=DATEVALUE(A2&"-01")

<img width="93" height="510" alt="image" src="https://github.com/user-attachments/assets/f4355ea2-5a14-45b2-b1e1-7134197e3a4c" />


### Q2. Check that m² price = sq ft price × 10.764
Helper column:
```
=ABS([@secondary_price_per_m2_usd]/([@secondary_price_per_sqft_usd]*10.764)-1)
```
<img width="874" height="641" alt="image" src="https://github.com/user-attachments/assets/cce13375-7ccc-481d-875a-be90c6499b2f" />

### Q3. Gross yield and off-plan premium
```
Gross Yield:       =[@rental_price_per_sqft_annual_usd]/[@secondary_price_per_sqft_usd]
Off-plan Premium:  =IF([@is_freehold],[@offplan_price_per_sqft_usd]/[@secondary_price_per_sqft_usd]-1,"")
```
<img width="334" height="600" alt="image" src="https://github.com/user-attachments/assets/53550895-a46c-49c6-bfb3-f70ad3d15330" />


### Q4. Every community has 76 months
```
=LET(comms,UNIQUE(Table13[community]),SUM(--(COUNTIFS(Table13[community],comms)<>76)))
```
<img width="265" height="646" alt="image" src="https://github.com/user-attachments/assets/7307824e-5818-4fbd-b1d6-5a6f74c4435d" />

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

<img width="485" height="204" alt="image" src="https://github.com/user-attachments/assets/e9e9c04e-4403-4bef-8611-5b005c563faa" />



**Check (Palm Jumeirah, Jan 2020):** 662 · 43 · 6.5% · 26 listings.



### Q6. Year-on-year growth using EDATE
Shift the **month you search for**, not the result:
```
=XLOOKUP(1,(Table13[community]=$T$2)*(Table13[year_month]=EDATE($T$3,-12)),Table13[secondary_price_per_sqft_usd],"No data")
```
YoY % = `current / previous − 1`.

<img width="442" height="184" alt="image" src="https://github.com/user-attachments/assets/6f6b42fd-31e4-4e2f-ba17-f6499b35ef8e" />


### Q7. Zone with the highest average yield in 2025
```
=LET(zones,UNIQUE(Table13[zone]),yld,AVERAGEIFS(Table13[Gross Yield],Table13[zone],zones,Table13[year_month],">="&DATE(2025,1,1),Table13[year_month],"<="&DATE(2025,12,1)),TAKE(SORTBY(HSTACK(zones,yld),yld,-1),5))
```
**Answer:** Deira **6.56%**, then Reem and Dubai Studio City (6.54%). Lowest: Mirdif (6.44%). The spread is tiny because rent is set at about 6.5% of price throughout.

<img width="430" height="460" alt="image" src="https://github.com/user-attachments/assets/4b91d80a-68ab-40fd-876d-6117ebba574a" />


### Q8. Freehold vs leasehold average price, latest month
```
=AVERAGEIFS(Table13[secondary_price_per_sqft_usd],Table13[is_freehold],TRUE,Table13[year_month],DATE(2026,4,1))
```
**Answer (Apr 2026):** freehold **617.86** (74 communities), leasehold **311.90** (10). Freehold is about 2× the price per sq ft.

<img width="469" height="59" alt="image" src="https://github.com/user-attachments/assets/5ebfa540-e441-42df-b8e4-1be003a5f890" />


## Section 3: Dynamic arrays and LAMBDA

### Q9. Top 10 communities by growth, Jan 2020 to Apr 2026
```
=LET(c,UNIQUE(Table13[community]),s,XLOOKUP(c,Table13[community],Table13[secondary_price_per_sqft_usd]),e,XLOOKUP(c,Table13[community],Table13[secondary_price_per_sqft_usd],,0,-1),g,e/s-1,TAKE(SORTBY(HSTACK(c,s,e,g),g,-1),10))
```
- `s`: XLOOKUP returns the **first** match, which is Jan 2020 (relies on the sort order).
- `e`: the last argument `-1` searches **bottom-up**, so it returns the last match, Apr 2026.

<img width="571" height="198" alt="image" src="https://github.com/user-attachments/assets/c288dfc1-b99e-439c-8823-7dbe601d56b4" />

**Lowest growth:** Pearl Jumeira **60.0%**, then Madinat Jumeirah Living (63.0%) and Emirates Hills (63.6%). All three are expensive, premium communities. Use `TAKE(…,-3)` to get the bottom rows.


### Q10. Community × year matrix of average price
Excel 365 with PIVOTBY:
```
=PIVOTBY(Table13[community],YEAR(Table13[year_month]),Table13[secondary_price_per_sqft_usd],AVERAGE)
```
Fallback without PIVOTBY:
```
=LET(c,SORT(UNIQUE(Table13[community])),y,SEQUENCE(1,7,2020),MAKEARRAY(ROWS(c),7,LAMBDA(r,k,AVERAGEIFS(Table13[secondary_price_per_sqft_usd],Table13[community],INDEX(c,r),Table13[year_month],">="&DATE(INDEX(y,k),1,1),Table13[year_month],"<="&DATE(INDEX(y,k),12,1)))))
```
**Check (Palm Jumeirah):** 2020 **653.92** · 2021 718.92 · 2022 925.25 · 2023 1,110.83 · 2024 1,164.75 · 2025 1,157.42 · 2026 1,178.50.

<img width="833" height="475" alt="image" src="https://github.com/user-attachments/assets/9a220a2e-6af2-449f-b912-d250bdd4604b" />


---

## Section 4: Time series


### Q11. Month of peak price for each community
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

<img width="267" height="497" alt="image" src="https://github.com/user-attachments/assets/2c19831d-d14e-497b-8ab1-67fbcb8ae16d" />


### Q12. Index to Jan 2020 = 100 and chart 5 communities
Index column:
```
=[@secondary_price_per_sqft_usd]/XLOOKUP([@community],[community],[secondary_price_per_sqft_usd])*100
```
Chart grid (months down the side, communities across the top):
- `H1:L1`: five community dropdowns
- `G2`: `=SORT(UNIQUE(Table13[year_month]))`
- `H2` (copy across to L2): `=SUMIFS(Table13[Index],Table13[community],H$1,Table13[year_month],$G2#)`

Select G1:L77 → **Insert → Line chart**.

<img width="409" height="521" alt="image" src="https://github.com/user-attachments/assets/612fd9a6-4dd2-4c82-a45b-c4245c5be6e0" />


## Section 5: PivotTables and Power Pivot

### Q13. Average price by zone × year, with slicer and timeline
Rows `zone`, Columns year (group the dates → Years), Values `secondary_price_per_sqft_usd` set to **Average**. Then **Insert Slicer → is_freehold** and **Insert Timeline → year_month**.

<img width="616" height="559" alt="image" src="https://github.com/user-attachments/assets/da4a7f66-9e39-4898-b211-f19912624955" />



### Q14. Date table, relationship and DAX measures

```
Avg Price := AVERAGE(Table13[secondary_price_per_sqft_usd])



Avg Yield := AVERAGEX(Table13, DIVIDE(Table13[rental_price_per_sqft_annual_usd], Table13[secondary_price_per_sqft_usd]))
```
- `DIVIDE(cur, prev) - 1` returns **−100%** when there's no prior year (BLANK − 1). `DIVIDE(cur - prev, prev)` returns blank instead.
- `AVERAGEX` averages the row-level yields, so every community-month counts equally. `DIVIDE(SUM(rent), SUM(price))` gives more weight to expensive communities.
- *Circular dependency* error means Avg Yield was created as a **calculated column**. Delete the column and create it as a **measure** (Power Pivot → Measures → New Measure).

<img width="257" height="184" alt="image" src="https://github.com/user-attachments/assets/b0b598ca-d749-45b6-8475-65de33713cbe" />


## Section 6: Statistics and what-if

### Q15. Correlation between rates and prices, 2022–2024
Monthly summary:
- `A2`: `=SORT(UNIQUE(Table13[year_month]))`
- `B2`: `=XLOOKUP(A2#,Table13[year_month],Table13[cbuae_base_rate_pct])`
- `C2`: `=AVERAGEIFS(Table13[secondary_price_per_sqft_usd],Table13[year_month],A2#)`

CORREL, self-contained and type-agnostic:
```
=LET(ym,TEXT(Table13[year_month],"yyyy-mm"),months,UNIQUE(ym),yr,--LEFT(months,4),sel,FILTER(months,(yr>=2022)*(yr<=2024)),rate,--XLOOKUP(sel,ym,Table13[cbuae_base_rate_pct]),price,MAP(sel,LAMBDA(m,AVERAGE(FILTER(Table13[secondary_price_per_sqft_usd],ym=m)))),CORREL(rate,price))
```

<img width="768" height="556" alt="image" src="https://github.com/user-attachments/assets/e40ab6bf-3b8f-4c72-ad31-b8fa6220885f" />



## Section 7: Presentation

### Q26. One-page dashboard controlled by a slicer
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

<img width="1171" height="621" alt="image" src="https://github.com/user-attachments/assets/f89b3be3-e7b2-41fc-82b2-d0854b1a6e4a" />

