# Three Statement Financial Model (Excel)

This project generates a fully linked 3-statement financial model workbook for a publicly listed company profile in INR bn.
Current prefill: Reliance Industries Limited (consolidated) historicals for FY2023-FY2025.

## Output

- `three_statement_model.xlsx`
- `index.html` (simple browser UI)

## Included sheets

1. `Control Panel`
- Scenario selector: Base / Bull / Bear
- Linked control assumptions for FY2026E: revenue growth, EBITDA margin, capex %, AR/Inventory/AP days, tax rate, interest rate, minimum cash
- Balance check monitor

2. `Historical Financials` (3 years)
- Income statement: Revenue, COGS, Gross Profit, Operating Expenses, EBITDA, Depreciation, EBIT, Interest, Tax, Net Income
- Balance sheet items: Cash, AR, Inventory, PPE, AP, Debt, Equity

3. `Assumptions`
- Scenario tables (5-year) for revenue growth, EBITDA margin, and capex %
- Global assumptions (5-year) for WC days, tax, interest, min cash, depreciation %, and OpEx %

4. `Working Capital Schedule`
- AR / Inventory / AP roll-forward and change in NWC

5. `PPE Schedule`
- Opening PPE, capex, depreciation, closing PPE

6. `Debt Schedule`
- Opening debt, interest expense, debt draw/(repay), closing debt
- Debt mechanism keeps cash above minimum cash target

7. `Projection Model` (5-year forecast)
- Integrated Income Statement, Balance Sheet, and Cash Flow Statement
- Fully linked formulas and balance check (`Assets - Liabilities - Equity`)

8. `Ratio Dashboard`
- EBITDA margin
- ROE
- ROCE
- Net Debt / EBITDA
- Free Cash Flow

9. `DCF Valuation`
- 5-year FCF build from projected statements
- WACC and terminal growth inputs
- Enterprise value, equity value, implied price per share

Control panel also includes:
- 2D sensitivity table (`Revenue Growth delta` vs `EBITDA Margin delta`)
- Output metric: stressed FY2030 Net Debt / EBITDA

## How to run

```bash
python3 create_three_statement_model.py
```

No external Python dependencies are required.

## Historical data source notes

- FY2025 and FY2024 consolidated line items: RIL Annual Report 2024-25 (Consolidated Financial Statements).
- FY2024 and FY2023 comparative line items: RIL media release (Audited financial results for quarter/year ended March 31, 2024).
- EBITDA in the historical sheet is a model-defined operational EBITDA (`Revenue - COGS - Operating Expenses`) using reported component lines.
