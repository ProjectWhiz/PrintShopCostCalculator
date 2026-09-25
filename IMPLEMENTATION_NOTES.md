# Implementation Notes: Print Shop Calculator App

## What Was Built

A React + Vite frontend-only web app was created to reproduce the two requested workbook tabs:

1. Cost Calculator (from `1st_Step_Print_Shop_Cost_Calculator.xlsx`)
2. Job Cost Calculator (from `Print_Shop_Cost_Analysis_and_Efficiency_Workbook.xlsx`)

No backend was added. All calculations run in the browser using React state and derived values.

## Why Frontend-Only

You asked for a no-cost-friendly setup that can be hosted on Vercel without paid backend services.

This design works well because:

- The source Excel logic is formula-based and deterministic.
- No server data persistence is required for basic use.
- Static hosting keeps deployment simple and free-tier compatible.

## Formula Parity Approach

The two source tabs were parsed to extract cell values and formulas, then translated into JavaScript calculation functions.

### 1) Cost Calculator Formula Mapping

Implemented in `src/App.jsx` under `calculateCostCalculator(...)`.

Key equivalents:

- Extended direct cost rows: `include ? unitCost * qty : 0`
- Labor rows 1-4: `include ? rate * hours : 0`
- General overhead allocation:
  - `(sumDirect + sumLaborRows1to4) * overheadPercent`
- Waste allowance:
  - `(blankExtended + transferExtended) * wastePercent`
- Total true job cost:
  - `sum(direct rows) + sum(labor/overhead/waste rows)`
- True cost per item:
  - `totalTrueJobCost / quantity`
- Target selling price per item:
  - `trueCostPerItem / (1 - targetMargin - paymentProcessingPercent)`
- Recommended job total:
  - `targetSellingPricePerItem * quantity * (1 + rushPercentIfIncluded)`
- Expected gross profit:
  - `recommendedJobTotal - totalTrueJobCost - (recommendedJobTotal * paymentProcessingPercent)`
- Selling-price test block:
  - quoted total, payment fee, profit, margin, and status.

### 2) Job Cost Calculator Formula Mapping

Implemented in `src/App.jsx` under `calculateJobRow(...)`.

Per-row equivalents:

- Waste qty:
  - `ceil(qty * defaultWastePercent)`
- Blank + material:
  - `(qty + wasteQty) * (blankCost + printMaterialCost)`
- Setup labor:
  - `setupHours * setupRate`
- Production labor + overhead:
  - `productionHours * loadedLaborRate`
- Machine cost:
  - `productionHours * machineRateByMethod`
- QC/pack labor:
  - `qcHours * qcRate`
- Sales fees:
  - `quotedPrice * salesFeePercent`
- Rush surcharge:
  - `rush ? quotedPrice * rushFeePercent : 0`
- Total job cost:
  - `blankAndMaterial + setup + productionLaborOverhead + machine + qc + shipping + other + salesFees`
- Cost per good unit:
  - `totalJobCost / qty`
- Recommended price:
  - `((totalJobCost - salesFees) / (1 - targetMargin - salesFeePercent)) * rushMultiplier`
- Actual gross profit:
  - `quotedPrice - totalJobCost`
- Actual margin %:
  - `actualGrossProfit / quotedPrice`

Loaded labor rate is computed from assumptions exactly like the workbook helper logic:

- `loadedLaborRate = hourlyProductionLaborRate * (1 + laborBurdenPercent) + (monthlyFixedOverhead / monthlyProductiveShopHours)`

### Quick Quote Section

Also included from the Job Cost Calculator sheet:

- Total cost:
  - `qty * (blankCost + decorationCost) + setup + productionOverhead + shippingOther`
- Cost per unit
- Recommended selling price:
  - `totalCost / (1 - targetMargin)`
- Recommended price per unit

## UX and Design Decisions

- Two-tab interface keeps both calculators in one app.
- Inputs mirror spreadsheet sections so the migration feels familiar.
- Result cards highlight key outputs for quoting decisions.
- Responsive layout supports desktop and mobile.
- Styled with a distinct visual identity (custom typography, gradients, and motion), avoiding generic boilerplate look.

## Important Notes

- This app does not save history to a database (by design, frontend-only).
- If you want persistence later without a paid backend, local browser storage can be added.
- Formulas are implemented intentionally to match workbook behavior, including cases where the source workbook computes values not directly included in total cost columns.

## Main Files

- `src/App.jsx` - Calculator state, formulas, and UI structure
- `src/App.css` - Component-level styling and responsive layout
- `src/index.css` - Global theme, fonts, and page background
- `README.md` - Setup and deploy instructions
