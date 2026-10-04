# Bike Sharing Demand Analysis Using Excel

An Excel-based analysis of bike-sharing demand: hourly rental patterns, rider types, weather, calendar effects, correlation, forecasting and anomaly detection.

## Project Overview

This project analyzes **1,000 hourly bike-sharing records** from **1 January to 14 February 2011**. The work was done entirely in Microsoft Excel (Power Query, lookup tables, formulas, PivotTables, charts, dashboard) and summarised in a 17-slide presentation.

## Objectives

- Analyze overall demand and compare **casual vs registered** riders.
- Identify hourly, weekday/weekend, working-day and holiday patterns.
- Examine the relationship between rentals and weather, temperature and humidity.
- Measure relationships with Pearson correlation.
- Forecast upcoming daily demand and detect unusual days with the IQR method.

## Dataset Scope

| Metric | Value |
| --- | ---: |
| Hourly records | 1,000 |
| Total rentals | 58,304 |
| Average hourly rentals | 58.30 |
| Min / max hourly rentals | 1 / 249 |
| Period | 1 Jan 2011 – 14 Feb 2011 (45 days) |
| Season | Winter only |

> Hours with no rentals are not recorded as rows, so some days have fewer than 24 records. The data ends at 07:00 on 14 February, so that day is partial. Seasonal and full-month comparisons are outside the scope of this analysis.

## Methodology

1. **Clean** the three source datasets with Power Query (data types, duplicates, missing values).
2. **Merge and append** them into one 1,000-row hourly dataset.
3. **Add features** with lookup tables (VLOOKUP / INDEX-MATCH) and IF logic: day type, working day and weather labels.
4. **Analyze** with PivotTables, charts and CORREL.
5. **Forecast** daily demand with Excel's Forecast Sheet and **detect anomalies** with the IQR method.
6. **Present** the results in an Excel dashboard and a PowerPoint deck.

## Key Findings

- **Riders:** registered riders account for 91.6% of rentals; casual riders 8.4%.
- **Time of day:** demand peaks at 08:00 (128.3) and 17:00 (140.2 rentals/hour).
- **Day type:** working days average 64.14 rentals/hour vs 47.65 on non-working days; weekdays 63.34 vs weekends 48.09.
- **Weather:** clear weather averages 61.67 rentals/hour vs 33.20 in light rain.
- **Correlation:** temperature is more strongly associated with casual riders (+0.470) than registered riders (+0.169); humidity is negatively associated with total rentals (−0.254).
- **Forecast:** the first 7 forecast days average about 1,500 rentals/day (1,500.49).
- **Anomaly:** the IQR method flags 14 Feb (151 rentals), which is a partial day, not a real demand drop.

## Recommendations

- Prioritise bike availability around the 08:00 and 17:00 commuter peaks.
- Adjust availability for weekends and light-rain conditions, when demand is lower.
- Use the forecast for short-term planning and validate it with data from other seasons.

## Repository Structure

```
├── Assignment.pdf                       # Original assignment brief
├── Bike_Sharing_Demand_Analysis.xlsx    # Analysis workbook
├── Bike_Sharing_Demand.pptx             # Presentation
└── Datasets/
    ├── dataset_1.xlsx
    ├── dataset_2.xlsx
    └── dataset_3.xlsx
```

## Limitations

- Winter-only data covering 45 days, so results should not be generalised to other seasons.
- Correlation does not establish causation.
- Heavy Rain has only 1 observation, and holiday results come from a single holiday (17 Jan).
- The forecast is based on a short history and is indicative only.
- VBA macros are not included in this submission.

## Tools

Microsoft Excel (Power Query, PivotTables, charts, slicers/timeline, Forecast Sheet) and Microsoft PowerPoint.