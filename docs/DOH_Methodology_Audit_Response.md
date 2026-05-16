# Inventory Days of Holding (DOH) — Analytical Methodology & Documentation

**Company:** Farmley (FMCG – Healthy Snacking)  
**Prepared By:** Supply Chain & Operations Team  
**Date:** May 2026  
**Document Version:** 1.0  
**Classification:** Internal – Audit Response  

---

## 1. Purpose & Scope

This document addresses the audit observation that Days of Holding (DOH) for Finished Goods (FG) and Raw Materials (RM) was previously determined based on judgment and historical experience, lacking analytically driven methodology, authorizations, and documentation.

This document provides:
- The analytical framework now used to determine DOH
- Mathematical formulations and their statistical basis
- Data sources and input parameters
- Authorization matrix and review cadence
- Deviation monitoring and exception handling

**Scope:** All 415+ SKUs across 5 manufacturing origins (Indore, Purnea, Udupi, Jaipur, Lucknow).

---

## 2. Methodology Overview

The DOH determination follows a **statistically-driven Min-Max inventory model** using demand variability analysis and service-level-based safety stock computation.

### 2.1 Framework

```
DOH = Inventory Level / Average Daily Demand
```

Where Inventory Level is computed analytically (not by judgment) using:
1. ABC Classification (Pareto analysis of dollar value)
2. Service Level assignment by class
3. Statistical Safety Stock computation
4. Min-Max Inventory band calculation

---

## 3. ABC Classification & Service Level Assignment

### 3.1 ABC Classification
SKUs are classified using standard Pareto (80-20) analysis based on cumulative dollar value contribution:

| Class | Cumulative Value | Typical SKU % | Service Level | Z-Score |
|-------|-----------------|---------------|---------------|---------|
| A     | 0–80%           | ~20%          | 95%           | 1.6449  |
| B     | 80–95%          | ~30%          | 90%           | 1.2816  |
| C     | 95–100%         | ~50%          | 85%           | 1.0364  |

### 3.2 Rationale
- **Class A** (high-value): Higher service level to prevent stockouts on revenue-critical items
- **Class B** (medium-value): Balanced service level
- **Class C** (low-value): Lower service level to avoid excess capital tied in slow-moving stock

### 3.3 Data Source
ABC classification is derived from the SKU master with annualized demand × unit price ranking.

---

## 4. Safety Stock Calculation

### 4.1 Formula

```
Safety Stock = Z × σ_daily × √(Lead Time)
```

Where:
- **Z** = Z-score corresponding to the target service level (from ABC class)
- **σ_daily** = Daily demand standard deviation = σ_monthly / √30
- **σ_monthly** = Standard deviation of monthly demand over rolling 6-month window
- **Lead Time** = 2 days (all origins — factory and warehouse are co-located)

### 4.2 Statistical Basis
This formula assumes demand follows a normal distribution and provides protection against demand variability during the replenishment lead time. The √(Lead Time) factor accounts for demand uncertainty accumulation over the lead time period.

### 4.3 Lead Time Justification
| Origin   | Lead Time (Days) | Rationale |
|----------|-----------------|-----------|
| Indore   | 2               | Production is continuous/daily; LT = process + pack + move to FG warehouse |
| Purnea   | 2               | Production is continuous/daily; LT = process + pack + move to FG warehouse |
| Udupi    | 2               | Production is continuous/daily; LT = process + pack + move to FG warehouse |
| Jaipur   | 2               | Production is continuous/daily; LT = process + pack + move to FG warehouse |
| Lucknow  | 2               | Production is continuous/daily; LT = process + pack + move to FG warehouse |

**Note:** While the full production cycle from raw material to finished good is 25-30 days, since production is a continuous daily pipeline, the effective FG replenishment lead time is only 2 days (time from production signal to FG available in warehouse). New FG exits the pipeline daily.

### 4.4 Order Cycle Justification
| Parameter | Value | Rationale |
|-----------|-------|-----------|
| Order Cycle | 1 day | Sales orders received daily for all SKUs; production and dispatch happen daily |

The 1-day order cycle means MAX inventory = MIN + 1 day of demand. This is appropriate because:
- Sales orders are received daily for all SKUs
- Production happens daily based on projected demand + actual orders
- FG is dispatched daily to fulfill orders
- Since replenishment happens every day, only 1 day of cycle stock is needed above the reorder point
- Safety stock (statistically computed) handles demand variability
- This results in DOH bands of approximately 3-5 days (lean operation with daily replenishment)

### 4.4 Demand Window
- Rolling 6-month historical data (currently Oct 2025 – Mar 2026)
- Updated monthly as new actuals become available

---

## 5. Min-Max Inventory Determination

### 5.1 Formulas

| Parameter | Formula | Purpose |
|-----------|---------|---------|
| **MIN (Reorder Point)** | (Daily Demand × Lead Time) + Safety Stock | Trigger point for replenishment |
| **MAX Inventory** | MIN + Order Quantity | Upper inventory limit |
| **Order Quantity** | Daily Demand × 1 day | Daily cycle stock (orders, production & dispatch all happen daily) |
| **AVG Inventory** | (MAX + MIN) / 2 | Expected steady-state inventory |

### 5.2 DOH Derivation

| DOH Metric | Formula | Interpretation |
|------------|---------|----------------|
| **DOH_MIN** | MIN Inventory / Daily Demand | Minimum days of cover before stockout risk |
| **DOH_MAX** | MAX Inventory / Daily Demand | Maximum days of cover post-replenishment |
| **DOH_AVG** | AVG Inventory / Daily Demand | Expected average holding period |

### 5.3 Dual-Basis Analysis
All calculations are performed in both:
- **Units (UNT):** For production planning
- **Weight (KGS):** For warehouse capacity and logistics planning

Conversion: `KGS = Units × Conversion Factor (kg/unit)` from Bill of Materials.

---

## 6. Actual DOH Computation & Variance Analysis

### 6.1 Actual DOH

```
Actual DOH = Current Closing Stock (KGS) / Daily Demand (KGS)
```

Data source: Latest month's closing inventory from stock ledger.

### 6.2 Variance

```
DOH Variance = Actual DOH − DOH_AVG (Recommended)
```

### 6.3 Status Classification

| Status | Condition | Action Required |
|--------|-----------|-----------------|
| **EXCESS** | Actual DOH > DOH_MAX | Review for slow-moving stock, consider promotions or production hold |
| **OPTIMAL** | DOH_MIN ≤ Actual DOH ≤ DOH_MAX | No action — within recommended band |
| **BELOW ROP** | Actual DOH < DOH_MIN | Urgent replenishment, risk of stockout |

---

## 7. Data Sources & Input Parameters

| Input | Source | Frequency |
|-------|--------|-----------|
| Monthly Sales (Units) | SAP/ERP Sales Report | Monthly |
| Sales Projection | Commercial Team Forecast | Monthly |
| ABC Classification | Finance/SCM Joint Review | Quarterly |
| Closing Stock (KGS) | Warehouse Management System | Monthly |
| Conversion Factor (kg/unit) | Bill of Materials / Production | As updated |
| Lead Time | Operations Team | Annual review |

---

## 8. Authorization & Governance

### 8.1 RACI Matrix

| Activity | Responsible | Accountable | Consulted | Informed |
|----------|-------------|-------------|-----------|----------|
| ABC Classification Review | SCM Analyst | Head of SCM | Finance, Sales | Plant Heads |
| Service Level Setting | Head of SCM | VP Operations | CFO, Sales Head | Auditors |
| Safety Stock Parameters | SCM Analyst | Head of SCM | Plant Heads | Finance |
| DOH Bands Approval | Head of SCM | VP Operations | CFO | Board |
| Monthly DOH Review | SCM Analyst | Head of SCM | Plant Heads | Finance |
| Exception Approval (>DOH_MAX) | Head of SCM | VP Operations | CFO | Auditors |

### 8.2 Review Cadence

| Review | Frequency | Participants | Output |
|--------|-----------|--------------|--------|
| DOH Monitoring | Weekly | SCM Team | Dashboard review |
| Inventory Review Meeting | Monthly | SCM + Finance + Sales | Action items for EXCESS/BELOW ROP |
| ABC Reclassification | Quarterly | SCM + Finance | Updated classifications |
| Parameter Review (SL, LT) | Semi-Annual | VP Ops + CFO | Approved parameter changes |
| Full Methodology Audit | Annual | Internal Audit + SCM | Compliance certification |

### 8.3 Change Control
Any modification to:
- Service level targets
- Lead time assumptions
- Safety stock formula
- ABC classification thresholds

Requires written approval from VP Operations and CFO, documented with effective date and rationale.

---

## 9. System Implementation

### 9.1 Automated Computation
- **Platform:** Web-based inventory analysis tool (Vercel-deployed)
- **Input:** Excel files (Sales data, Projections, ABC classification, Stock position)
- **Output:** Downloadable Excel with live formulas (auditable, traceable)
- **Methodology Tab:** Built into application for real-time reference

### 9.2 Excel Formula Transparency
The output Excel workbook contains:
- **Raw_Data sheet:** Source data with cross-references
- **Inventory_Units sheet:** All formulas visible (e.g., `=Z_score * STDEV/SQRT(30) * SQRT(LT)`)
- **Inventory_KGS sheet:** Weight-based formulas
- **DOH_Analysis sheet:** Actual vs Recommended with variance
- **Methodology_DOH sheet:** Formula documentation within the workbook

All computed cells use Excel formulas (not hardcoded values), enabling auditability.

---

## 10. Exception Handling

### 10.1 New SKU Launch (No Historical Data)
- Use category average demand ± 20% buffer
- Review after 3 months of actual data
- Classified as "B" until sufficient history

### 10.2 Seasonal Products
- Separate seasonal demand model (3-month pre-season buildup)
- DOH bands adjusted seasonally with documented approval

### 10.3 SKU Discontinuation
- DOH_MAX set to 0 (no new replenishment)
- Monitored until stock depleted or written off

---

## 11. Audit Trail & Documentation

| Document | Location | Retention |
|----------|----------|-----------|
| This Methodology Document | SharePoint / Compliance folder | Permanent |
| Monthly DOH Reports (Excel) | Shared Drive / Inventory Reports | 7 years |
| Parameter Change Approvals | Email / Document Management | 7 years |
| Exception Approvals | Signed physical + digital copy | 7 years |
| Quarterly ABC Review Minutes | Meeting repository | 5 years |
| Web Application Source Code | GitHub (version-controlled) | Permanent |

---

## 12. Summary — Addressing Audit Observation

| Audit Concern | Resolution |
|---------------|-----------|
| DOH based on judgment | Replaced with statistical formula: Z × σ_daily × √LT |
| Lacks analytical methodology | ABC-based service levels + demand variability model documented |
| No authorizations | RACI matrix defined; parameter changes require VP Ops + CFO sign-off |
| No documentation | This document + Methodology tab in tool + formula Excel outputs |

---

## 13. Approval & Sign-Off

| Role | Name | Signature | Date |
|------|------|-----------|------|
| Prepared By (SCM Analyst) | _________________ | _________ | ___/___/2026 |
| Reviewed By (Head of SCM) | _________________ | _________ | ___/___/2026 |
| Approved By (VP Operations) | _________________ | _________ | ___/___/2026 |
| Acknowledged (CFO) | _________________ | _________ | ___/___/2026 |
| Noted (Internal Audit) | _________________ | _________ | ___/___/2026 |

---

*Document Control: Version 1.0 | Effective Date: May 2026 | Next Review: November 2026*
