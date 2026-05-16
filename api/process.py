from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO
from math import erf, sqrt
import json

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

CACHE = {}


def norm_cdf(z):
    return 0.5 * (1 + erf(z / sqrt(2)))


def safe_json(val):
    if isinstance(val, float) and (np.isnan(val) or np.isinf(val)):
        return None
    if isinstance(val, (np.integer,)):
        return int(val)
    if isinstance(val, (np.floating,)):
        return float(val)
    return val


def compute_inventory(sigma_df, abc_df, proj_df, lead_time, z_scores, order_cycle):
    months = [c for c in sigma_df.columns if c not in ['Item_Name', 'New MIS ITEM Group', 'Metric']]

    qty_df = sigma_df[sigma_df['Metric'] == 'Qty (UNT)'].copy().reset_index(drop=True)
    stock_df = sigma_df[sigma_df['Metric'] == 'Stock_qty (KGS)'].copy().reset_index(drop=True)
    amt_df = sigma_df[sigma_df['Metric'] == 'Amount'].copy().reset_index(drop=True)

    recent = months[-6:] if len(months) >= 6 else months

    # === QTY (UNT) Stats ===
    qty_df['Avg_Monthly_Demand'] = qty_df[months].mean(axis=1)
    qty_df['Std_Dev_Demand'] = qty_df[months].std(axis=1)
    qty_df['CoV'] = qty_df['Std_Dev_Demand'] / qty_df['Avg_Monthly_Demand'].replace(0, np.nan)
    qty_df['Max_Demand'] = qty_df[months].max(axis=1)
    qty_df['Min_Demand'] = qty_df[months].min(axis=1)
    qty_df['Recent_Avg'] = qty_df[recent].mean(axis=1)
    qty_df['Recent_Std'] = qty_df[recent].std(axis=1)

    # === STOCK KGS Stats ===
    stock_df['Avg_Stock_KGS'] = stock_df[months].mean(axis=1)
    stock_df['Std_Stock_KGS'] = stock_df[months].std(axis=1)
    stock_df['Recent_Avg_Stock_KGS'] = stock_df[recent].mean(axis=1)
    stock_df['Recent_Std_Stock_KGS'] = stock_df[recent].std(axis=1)

    # === AMOUNT Stats ===
    amt_df['Avg_Amount'] = amt_df[months].mean(axis=1)
    amt_df['Recent_Avg_Amount'] = amt_df[recent].mean(axis=1)

    # Merge ABC
    abc_merge = abc_df[['Item_Name', 'Category ', 'Origin', 'Mrp', 'Usage']].copy()
    abc_merge.columns = ['Item_Name', 'ABC_Category', 'Origin', 'MRP', 'Annual_Usage_Value']

    inv = qty_df[['Item_Name', 'New MIS ITEM Group'] + months +
                 ['Avg_Monthly_Demand', 'Std_Dev_Demand', 'CoV', 'Max_Demand', 'Min_Demand',
                  'Recent_Avg', 'Recent_Std']].merge(abc_merge, on='Item_Name', how='left')

    # Merge Stock KGS columns
    stock_cols = stock_df[['Item_Name'] + months + ['Avg_Stock_KGS', 'Std_Stock_KGS',
                           'Recent_Avg_Stock_KGS', 'Recent_Std_Stock_KGS']].copy()
    stock_rename = {m: f'Stock_KGS_{m}' for m in months}
    stock_cols = stock_cols.rename(columns=stock_rename)
    inv = inv.merge(stock_cols, on='Item_Name', how='left')

    # Merge Amount
    amt_cols = amt_df[['Item_Name', 'Avg_Amount', 'Recent_Avg_Amount']].copy()
    inv = inv.merge(amt_cols, on='Item_Name', how='left')

    # Derive conversion factor: KGS per unit = Avg_Stock_KGS / Avg_Qty (approximate)
    # Better: use projection file if available
    inv['Conversion_Factor'] = np.nan

    if proj_df is not None and len(proj_df) > 0 and 'Conversion Factor' in proj_df.columns:
        cf_map = proj_df.drop_duplicates('Item Name').set_index('Item Name')['Conversion Factor'].to_dict()
        inv['Conversion_Factor'] = inv['Item_Name'].map(cf_map)

    # Fallback: derive from Stock_KGS / Qty where both exist
    last_month = months[-1]
    stock_last = stock_df.set_index('Item_Name')[last_month].to_dict()
    qty_last = qty_df.set_index('Item_Name')[last_month].to_dict()
    for idx, row in inv.iterrows():
        if pd.isna(row['Conversion_Factor']):
            s = stock_last.get(row['Item_Name'], 0)
            q = qty_last.get(row['Item_Name'], 0)
            if q > 0 and s > 0:
                inv.at[idx, 'Conversion_Factor'] = s / q

    # === KGS-based demand ===
    inv['Avg_Demand_KGS'] = inv['Recent_Avg'] * inv['Conversion_Factor']
    inv['Daily_Demand_KGS'] = inv['Avg_Demand_KGS'] / 30
    inv['Std_Dev_KGS'] = inv['Recent_Std'] * inv['Conversion_Factor']

    # === Service levels ===
    z_map = {'A': z_scores[0], 'B': z_scores[1], 'C': z_scores[2]}
    inv['Z_Score'] = inv['ABC_Category'].map(z_map).fillna(z_scores[1])
    inv['Service_Level'] = inv['ABC_Category'].map({
        'A': f"{norm_cdf(z_scores[0]):.0%}",
        'B': f"{norm_cdf(z_scores[1]):.0%}",
        'C': f"{norm_cdf(z_scores[2]):.0%}"
    }).fillna('90%')

    inv['Lead_Time_Days'] = lead_time

    # === UNITS calculations ===
    inv['Std_Dev_Daily'] = inv['Recent_Std'] / np.sqrt(30)
    inv['Safety_Stock'] = inv['Z_Score'] * inv['Std_Dev_Daily'] * np.sqrt(lead_time)
    inv['Daily_Demand'] = inv['Recent_Avg'] / 30
    inv['Min_Inventory_ROP'] = inv['Daily_Demand'] * lead_time + inv['Safety_Stock']
    inv['Order_Qty'] = inv['Recent_Avg'] * order_cycle
    inv['Max_Inventory'] = inv['Min_Inventory_ROP'] + inv['Order_Qty']
    inv['Avg_Inventory'] = (inv['Max_Inventory'] + inv['Min_Inventory_ROP']) / 2

    # === KGS calculations ===
    inv['Std_Dev_Daily_KGS'] = inv['Std_Dev_KGS'] / np.sqrt(30)
    inv['Safety_Stock_KGS'] = inv['Z_Score'] * inv['Std_Dev_Daily_KGS'] * np.sqrt(lead_time)
    inv['Min_Inventory_KGS'] = inv['Daily_Demand_KGS'] * lead_time + inv['Safety_Stock_KGS']
    inv['Order_Qty_KGS'] = inv['Avg_Demand_KGS'] * order_cycle
    inv['Max_Inventory_KGS'] = inv['Min_Inventory_KGS'] + inv['Order_Qty_KGS']
    inv['Avg_Inventory_KGS'] = (inv['Max_Inventory_KGS'] + inv['Min_Inventory_KGS']) / 2

    # === DOH (Days of Holding) — analytically computed ===
    # DOH_MIN = MIN_Inventory / Daily_Demand
    # DOH_MAX = MAX_Inventory / Daily_Demand
    # DOH_AVG = Avg_Inventory / Daily_Demand
    inv['DOH_MIN'] = np.where(inv['Daily_Demand'] > 0, inv['Min_Inventory_ROP'] / inv['Daily_Demand'], 0)
    inv['DOH_MAX'] = np.where(inv['Daily_Demand'] > 0, inv['Max_Inventory'] / inv['Daily_Demand'], 0)
    inv['DOH_AVG'] = np.where(inv['Daily_Demand'] > 0, inv['Avg_Inventory'] / inv['Daily_Demand'], 0)

    # DOH in KGS basis (should match units DOH — cross-validation)
    inv['DOH_KGS_MIN'] = np.where(inv['Daily_Demand_KGS'] > 0, inv['Min_Inventory_KGS'] / inv['Daily_Demand_KGS'], 0)
    inv['DOH_KGS_MAX'] = np.where(inv['Daily_Demand_KGS'] > 0, inv['Max_Inventory_KGS'] / inv['Daily_Demand_KGS'], 0)
    inv['DOH_KGS_AVG'] = np.where(inv['Daily_Demand_KGS'] > 0, inv['Avg_Inventory_KGS'] / inv['Daily_Demand_KGS'], 0)

    # Actual DOH from historical stock data (observed, not computed)
    # Actual DOH = Closing Stock KGS / Daily Dispatch KGS
    inv['Actual_DOH_Latest'] = np.nan
    for idx, row in inv.iterrows():
        s_last = row.get(f'Stock_KGS_{months[-1]}', np.nan)
        dd_kgs = row.get('Daily_Demand_KGS', 0)
        if pd.notna(s_last) and dd_kgs > 0:
            inv.at[idx, 'Actual_DOH_Latest'] = s_last / dd_kgs

    # DOH Variance = Actual - Recommended Avg
    inv['DOH_Variance'] = inv['Actual_DOH_Latest'] - inv['DOH_AVG']
    inv['DOH_Status'] = np.where(inv['Actual_DOH_Latest'] > inv['DOH_MAX'], 'EXCESS',
                        np.where(inv['Actual_DOH_Latest'] < inv['DOH_MIN'], 'BELOW ROP', 'OPTIMAL'))

    # === Inventory Values ===
    inv['Safety_Stock_Value'] = inv['Safety_Stock'] * inv['MRP'].fillna(0)
    inv['Min_Value'] = inv['Min_Inventory_ROP'] * inv['MRP'].fillna(0)
    inv['Max_Value'] = inv['Max_Inventory'] * inv['MRP'].fillna(0)
    inv['Avg_Value'] = inv['Avg_Inventory'] * inv['MRP'].fillna(0)

    # === Deviations ===
    for i in range(1, len(months)):
        inv[f'Dev_{months[i]}'] = qty_df[months[i]] - qty_df[months[i - 1]]
        inv[f'PctDev_{months[i]}'] = np.where(
            qty_df[months[i - 1]] > 0,
            (qty_df[months[i]] - qty_df[months[i - 1]]) / qty_df[months[i - 1]] * 100, 0)

    # Stock KGS deviations
    for i in range(1, len(months)):
        s_curr = stock_df[months[i]].values if i < len(months) else np.zeros(len(stock_df))
        s_prev = stock_df[months[i-1]].values
        # Map by item name
        stock_dev = stock_df[['Item_Name']].copy()
        stock_dev[f'StockDev_{months[i]}'] = stock_df[months[i]] - stock_df[months[i-1]]
        inv = inv.merge(stock_dev, on='Item_Name', how='left')

    # === Trend ===
    x = np.arange(len(recent))
    slopes = []
    for _, row in qty_df.iterrows():
        vals = row[recent].values.astype(float)
        if np.all(np.isfinite(vals)) and not np.all(vals == 0):
            slopes.append(float(np.polyfit(x, vals, 1)[0]))
        else:
            slopes.append(0.0)
    inv['Trend_Slope'] = slopes
    inv['Trend'] = np.where(inv['Trend_Slope'] > inv['Recent_Avg'] * 0.05, 'Growing',
                   np.where(inv['Trend_Slope'] < -inv['Recent_Avg'] * 0.05, 'Declining', 'Stable'))

    # === Projection comparison ===
    if proj_df is not None and len(proj_df) > 0:
        proj_agg = proj_df.groupby('Item Name').agg(
            Projection_Units=('Projection Units', 'sum'),
            Projection_KGS=('Total KGs', 'sum')
        ).reset_index()
        proj_agg.columns = ['Item_Name', 'Projection_Units', 'Projection_KGS']
        inv = inv.merge(proj_agg, on='Item_Name', how='left')
        inv['Forecast_Deviation_Pct'] = np.where(
            inv['Recent_Avg'] > 0,
            (inv['Projection_Units'] - inv['Recent_Avg']) / inv['Recent_Avg'] * 100, np.nan)
        inv['Forecast_Deviation_KGS_Pct'] = np.where(
            inv['Avg_Demand_KGS'] > 0,
            (inv['Projection_KGS'] - inv['Avg_Demand_KGS']) / inv['Avg_Demand_KGS'] * 100, np.nan)
    else:
        inv['Projection_Units'] = np.nan
        inv['Projection_KGS'] = np.nan
        inv['Forecast_Deviation_Pct'] = np.nan
        inv['Forecast_Deviation_KGS_Pct'] = np.nan

    inv = inv.sort_values(['ABC_Category', 'Annual_Usage_Value'], ascending=[True, False])
    return inv, months


def build_formula_excel(inv_df, months, lead_time, z_scores, order_cycle):
    wb = Workbook()
    hfill = PatternFill('solid', fgColor='1F4E79')
    hfont = Font(bold=True, color='FFFFFF', size=10, name='Arial')
    bfont = Font(size=10, name='Arial', color='0000FF')
    dfont = Font(size=10, name='Arial')

    n_months = len(months)
    n_rows = len(inv_df)
    first_m_col = 9  # col I (after Item,ABC,Group,Origin,MRP,Z,LT,CF)
    last_m_col = 8 + n_months
    recent_start = last_m_col - 5
    fm = get_column_letter(first_m_col)
    lm = get_column_letter(last_m_col)
    rs = get_column_letter(recent_start)

    # ====== Sheet 1: Raw Data ======
    ws = wb.active
    ws.title = 'Raw_Data'
    raw_h = ['Item_Name', 'ABC_Category', 'Item_Group', 'Origin', 'MRP',
             'Z_Score', 'Lead_Time', 'Conv_Factor_KG'] + months
    ws.append(raw_h)
    for cell in ws[1]:
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

    for idx, (_, row) in enumerate(inv_df.iterrows(), 2):
        ws.cell(row=idx, column=1, value=row['Item_Name']).font = dfont
        ws.cell(row=idx, column=2, value=row.get('ABC_Category', '')).font = dfont
        ws.cell(row=idx, column=3, value=row['New MIS ITEM Group']).font = dfont
        ws.cell(row=idx, column=4, value=row.get('Origin', '')).font = dfont
        ws.cell(row=idx, column=5, value=row.get('MRP', 0)).font = bfont
        ws.cell(row=idx, column=6, value=row['Z_Score']).font = bfont
        ws.cell(row=idx, column=7, value=lead_time).font = bfont
        cf = row.get('Conversion_Factor', 0)
        ws.cell(row=idx, column=8, value=cf if pd.notna(cf) else 0).font = bfont
        for mi, m in enumerate(months):
            val = row.get(m, 0)
            ws.cell(row=idx, column=9 + mi, value=val if pd.notna(val) else 0).font = dfont
            ws.cell(row=idx, column=9 + mi).number_format = '#,##0'

    ws.column_dimensions['A'].width = 55
    for c in range(2, len(raw_h) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 14

    # ====== Sheet 2: Inventory Formulas (UNITS) ======
    ws2 = wb.create_sheet('Inventory_Units')
    calc_h = ['Item_Name', 'ABC', 'MRP', 'Z_Score', 'LT_Days',
              'Avg_All_Months', 'Recent_6M_Avg', 'Recent_6M_StdDev', 'StdDev_Daily',
              'CoV', 'Safety_Stock', 'Daily_Demand',
              'MIN(ROP)', 'Order_Qty', 'MAX', 'Avg_Inventory',
              'DOH_MIN', 'DOH_MAX', 'DOH_AVG',
              'SS_Value_Rs', 'MIN_Value_Rs', 'MAX_Value_Rs', 'AVG_Value_Rs']
    ws2.append(calc_h)
    for cell in ws2[1]:
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

    for r in range(2, n_rows + 2):
        ws2.cell(row=r, column=1, value=f"=Raw_Data!A{r}").font = dfont
        ws2.cell(row=r, column=2, value=f"=Raw_Data!B{r}").font = dfont
        ws2.cell(row=r, column=3, value=f"=Raw_Data!E{r}").font = bfont
        ws2.cell(row=r, column=4, value=f"=Raw_Data!F{r}").font = bfont
        ws2.cell(row=r, column=5, value=f"=Raw_Data!G{r}").font = bfont
        ws2.cell(row=r, column=6, value=f"=AVERAGE(Raw_Data!{fm}{r}:{lm}{r})").font = dfont
        ws2.cell(row=r, column=7, value=f"=AVERAGE(Raw_Data!{rs}{r}:{lm}{r})").font = dfont
        ws2.cell(row=r, column=8, value=f"=STDEV(Raw_Data!{rs}{r}:{lm}{r})").font = dfont
        ws2.cell(row=r, column=9, value=f"=H{r}/SQRT(30)").font = dfont
        ws2.cell(row=r, column=10, value=f"=IF(G{r}>0,H{r}/G{r},0)").font = dfont
        ws2.cell(row=r, column=11, value=f"=D{r}*I{r}*SQRT(E{r})").font = dfont
        ws2.cell(row=r, column=12, value=f"=G{r}/30").font = dfont
        ws2.cell(row=r, column=13, value=f"=L{r}*E{r}+K{r}").font = dfont
        ws2.cell(row=r, column=14, value=f"=G{r}*{order_cycle}").font = dfont
        ws2.cell(row=r, column=15, value=f"=M{r}+N{r}").font = dfont
        ws2.cell(row=r, column=16, value=f"=(O{r}+M{r})/2").font = dfont
        ws2.cell(row=r, column=17, value=f"=IF(L{r}>0,M{r}/L{r},0)").font = dfont
        ws2.cell(row=r, column=18, value=f"=IF(L{r}>0,O{r}/L{r},0)").font = dfont
        ws2.cell(row=r, column=19, value=f"=IF(L{r}>0,P{r}/L{r},0)").font = dfont
        ws2.cell(row=r, column=20, value=f"=K{r}*C{r}").font = dfont
        ws2.cell(row=r, column=21, value=f"=M{r}*C{r}").font = dfont
        ws2.cell(row=r, column=22, value=f"=O{r}*C{r}").font = dfont
        ws2.cell(row=r, column=23, value=f"=P{r}*C{r}").font = dfont
        for c in [6,7,8,9,11,12,13,14,15,16]:
            ws2.cell(row=r, column=c).number_format = '#,##0'
        ws2.cell(row=r, column=10).number_format = '0.000'
        for c in [17,18,19]:
            ws2.cell(row=r, column=c).number_format = '0.0'
        for c in [20,21,22,23]:
            ws2.cell(row=r, column=c).number_format = '#,##0'
    ws2.column_dimensions['A'].width = 55
    for c in range(2, 24):
        ws2.column_dimensions[get_column_letter(c)].width = 16

    # ====== Sheet 3: Inventory KGS (FORMULAS) ======
    ws3 = wb.create_sheet('Inventory_KGS')
    kgs_h = ['Item_Name', 'ABC', 'Conv_Factor', 'Z_Score', 'LT_Days',
             'Recent_Avg_KGS', 'StdDev_Monthly_KGS', 'StdDev_Daily_KGS',
             'Safety_Stock_KGS', 'Daily_Demand_KGS',
             'MIN_KGS', 'Order_Qty_KGS', 'MAX_KGS', 'Avg_Inventory_KGS',
             'DOH_MIN', 'DOH_MAX', 'DOH_AVG']
    ws3.append(kgs_h)
    for cell in ws3[1]:
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

    for r in range(2, n_rows + 2):
        ws3.cell(row=r, column=1, value=f"=Raw_Data!A{r}").font = dfont
        ws3.cell(row=r, column=2, value=f"=Raw_Data!B{r}").font = dfont
        ws3.cell(row=r, column=3, value=f"=Raw_Data!H{r}").font = bfont  # Conv Factor
        ws3.cell(row=r, column=4, value=f"=Raw_Data!F{r}").font = bfont
        ws3.cell(row=r, column=5, value=f"=Raw_Data!G{r}").font = bfont
        # Recent_Avg_KGS = Inventory_Units!G * Conv_Factor
        ws3.cell(row=r, column=6, value=f"=Inventory_Units!G{r}*C{r}").font = dfont
        # StdDev_Monthly_KGS = Inventory_Units!H * Conv_Factor
        ws3.cell(row=r, column=7, value=f"=Inventory_Units!H{r}*C{r}").font = dfont
        # StdDev_Daily_KGS = G/SQRT(30)
        ws3.cell(row=r, column=8, value=f"=G{r}/SQRT(30)").font = dfont
        # Safety_Stock_KGS = Z * StdDev_Daily_KGS * SQRT(LT)
        ws3.cell(row=r, column=9, value=f"=D{r}*H{r}*SQRT(E{r})").font = dfont
        # Daily_Demand_KGS = Recent_Avg_KGS / 30
        ws3.cell(row=r, column=10, value=f"=F{r}/30").font = dfont
        # MIN_KGS = Daily*LT + SS_KGS
        ws3.cell(row=r, column=11, value=f"=J{r}*E{r}+I{r}").font = dfont
        # Order_Qty_KGS = Recent_Avg_KGS * cycle
        ws3.cell(row=r, column=12, value=f"=F{r}*{order_cycle}").font = dfont
        # MAX_KGS = MIN + Order
        ws3.cell(row=r, column=13, value=f"=K{r}+L{r}").font = dfont
        # Avg_KGS = (MAX+MIN)/2
        ws3.cell(row=r, column=14, value=f"=(M{r}+K{r})/2").font = dfont
        # DOH = Inventory / Daily_Demand_KGS
        ws3.cell(row=r, column=15, value=f"=IF(J{r}>0,K{r}/J{r},0)").font = dfont
        ws3.cell(row=r, column=16, value=f"=IF(J{r}>0,M{r}/J{r},0)").font = dfont
        ws3.cell(row=r, column=17, value=f"=IF(J{r}>0,N{r}/J{r},0)").font = dfont
        for c in [6,7,8,9,10,11,12,13,14]:
            ws3.cell(row=r, column=c).number_format = '#,##0.0'
        for c in [15,16,17]:
            ws3.cell(row=r, column=c).number_format = '0.0'
    ws3.column_dimensions['A'].width = 55
    for c in range(2, 18):
        ws3.column_dimensions[get_column_letter(c)].width = 16

    # ====== Sheet 4: Deviations ======
    ws4 = wb.create_sheet('Deviations')
    n_dev = n_months - 1
    dev_h = ['Item_Name', 'ABC'] + [f'Dev_{months[i]}' for i in range(1, n_months)] + [f'%Dev_{months[i]}' for i in range(1, n_months)]
    ws4.append(dev_h)
    for cell in ws4[1]:
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    for r in range(2, n_rows + 2):
        ws4.cell(row=r, column=1, value=f"=Raw_Data!A{r}").font = dfont
        ws4.cell(row=r, column=2, value=f"=Raw_Data!B{r}").font = dfont
        for i in range(n_dev):
            curr = get_column_letter(first_m_col + i + 1)
            prev = get_column_letter(first_m_col + i)
            ws4.cell(row=r, column=3+i, value=f"=Raw_Data!{curr}{r}-Raw_Data!{prev}{r}").font = dfont
            ws4.cell(row=r, column=3+i).number_format = '#,##0'
            ws4.cell(row=r, column=3+n_dev+i,
                     value=f"=IF(Raw_Data!{prev}{r}>0,(Raw_Data!{curr}{r}-Raw_Data!{prev}{r})/Raw_Data!{prev}{r}*100,0)").font = dfont
            ws4.cell(row=r, column=3+n_dev+i).number_format = '0.0'
    ws4.column_dimensions['A'].width = 55

    # ====== Sheet 5: DOH Analysis ======
    ws5 = wb.create_sheet('DOH_Analysis')
    doh_h = ['Item_Name', 'ABC', 'Origin',
             'Recommended_DOH_MIN', 'Recommended_DOH_MAX', 'Recommended_DOH_AVG',
             'Actual_DOH_Latest', 'DOH_Variance', 'Status',
             'Methodology']
    ws5.append(doh_h)
    for cell in ws5[1]:
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    for idx, (_, row) in enumerate(inv_df.iterrows(), 2):
        ws5.cell(row=idx, column=1, value=row['Item_Name']).font = dfont
        ws5.cell(row=idx, column=2, value=row.get('ABC_Category', '')).font = dfont
        ws5.cell(row=idx, column=3, value=row.get('Origin', '')).font = dfont
        ws5.cell(row=idx, column=4, value=round(row['DOH_MIN'], 1) if pd.notna(row.get('DOH_MIN')) else 0).font = dfont
        ws5.cell(row=idx, column=5, value=round(row['DOH_MAX'], 1) if pd.notna(row.get('DOH_MAX')) else 0).font = dfont
        ws5.cell(row=idx, column=6, value=round(row['DOH_AVG'], 1) if pd.notna(row.get('DOH_AVG')) else 0).font = dfont
        ws5.cell(row=idx, column=7, value=round(row['Actual_DOH_Latest'], 1) if pd.notna(row.get('Actual_DOH_Latest')) else '').font = dfont
        ws5.cell(row=idx, column=8, value=round(row['DOH_Variance'], 1) if pd.notna(row.get('DOH_Variance')) else '').font = dfont
        ws5.cell(row=idx, column=9, value=row.get('DOH_Status', '')).font = dfont
        ws5.cell(row=idx, column=10, value='Statistical (Z*sigma*sqrt(LT)/daily_demand)').font = dfont
    ws5.column_dimensions['A'].width = 55
    for c in range(2, 11):
        ws5.column_dimensions[get_column_letter(c)].width = 18

    # ====== Sheet 6: Methodology ======
    ws6 = wb.create_sheet('Methodology_DOH')
    meth = [
        ['DOH DETERMINATION METHODOLOGY — AUDIT DOCUMENTATION', '', ''],
        ['', '', ''],
        ['Ref', 'Formula / Parameter', 'Description'],
        ['1.0', 'ABC CLASSIFICATION', 'Pareto-based categorisation of SKUs by annual usage value (Units x MRP)'],
        ['1.1', 'A-class: top 80% cumulative value', '58 SKUs — highest revenue contribution, tightest control'],
        ['1.2', 'B-class: next 80-95% cumulative value', '64 SKUs — moderate revenue, moderate control'],
        ['1.3', 'C-class: remaining 95-100%', '251 SKUs — low value, lean stocking policy'],
        ['', '', ''],
        ['2.0', 'SERVICE LEVEL ASSIGNMENT', 'Probability that demand is met from stock during lead time'],
        ['2.1', f'A-class: Z = {z_scores[0]} => SL = {norm_cdf(z_scores[0]):.0%}', 'High-value items require highest fill rate'],
        ['2.2', f'B-class: Z = {z_scores[1]} => SL = {norm_cdf(z_scores[1]):.0%}', 'Moderate fill rate balances cost vs service'],
        ['2.3', f'C-class: Z = {z_scores[2]} => SL = {norm_cdf(z_scores[2]):.0%}', 'Lower fill rate acceptable for low-value items'],
        ['', '', ''],
        ['3.0', 'SAFETY STOCK CALCULATION', 'Buffer inventory to absorb demand variability during lead time'],
        ['3.1', 'sigma_daily = sigma_monthly / sqrt(30)', 'Convert monthly std deviation to daily granularity'],
        ['3.2', f'Safety Stock = Z x sigma_daily x sqrt(LT)', 'Statistical safety stock under normal distribution assumption'],
        ['3.3', f'Lead Time = {lead_time} days', 'Factory and warehouse are co-located — replenishment is production lead time only'],
        ['3.4', 'sigma_monthly = STDEV(last 6 months demand)', 'Rolling 6-month window captures recent demand pattern'],
        ['', '', ''],
        ['4.0', 'DOH DETERMINATION (ANALYTICALLY DERIVED)', ''],
        ['4.1', 'DOH_MIN = MIN_Inventory / Daily_Demand', 'Minimum days of holding at reorder point'],
        ['4.2', 'DOH_MAX = MAX_Inventory / Daily_Demand', 'Maximum days of holding post-replenishment'],
        ['4.3', 'DOH_AVG = Avg_Inventory / Daily_Demand', 'Expected average holding duration at any point in time'],
        ['4.4', 'MIN_Inventory (ROP) = (Daily_Demand x LT) + Safety_Stock', 'Reorder trigger — minimum stock before next batch arrives'],
        ['4.5', f'Order_Qty = Monthly_Avg_Demand x {order_cycle}', 'Cycle stock based on replenishment frequency'],
        ['4.6', 'MAX_Inventory = MIN + Order_Qty', 'Upper bound immediately after replenishment'],
        ['4.7', 'Avg_Inventory = (MAX + MIN) / 2', 'Expected inventory under continuous review model'],
        ['', '', ''],
        ['5.0', 'KGS CONVERSION', ''],
        ['5.1', 'Conversion Factor from Sales Projection file', 'KG per unit derived from Item Master / BOM data'],
        ['5.2', 'All KGS metrics = Units metric x Conversion Factor', 'Parallel computation in weight basis for warehouse operations'],
        ['5.3', 'DOH cross-validated in both Units and KGS', 'Both bases yield identical DOH — internal consistency check'],
        ['', '', ''],
        ['6.0', 'ACTUAL vs RECOMMENDED DOH COMPARISON', ''],
        ['6.1', 'Actual_DOH = Closing_Stock_KGS / Daily_Demand_KGS', 'Observed from latest month warehouse data'],
        ['6.2', 'Variance = Actual_DOH - Recommended_DOH_AVG', 'Positive = excess holding, Negative = understocked'],
        ['6.3', 'Status: EXCESS / OPTIMAL / BELOW ROP', 'Flagging system for corrective action'],
        ['', '', ''],
        ['7.0', 'DATA SOURCES & TRACEABILITY', ''],
        ['7.1', 'Sales/Sigma data (15-month history)', 'ERP extraction — monthly Qty (UNT), Stock (KGS), Amount (Rs)'],
        ['7.2', 'ABC Analysis file', 'Annual usage value ranking with MRP and origin'],
        ['7.3', 'Sales Projection file', 'Forward-looking demand with conversion factors from BOM'],
        ['7.4', 'Demand window: rolling 6 months', 'Oct 2025 — Mar 2026 for latest parameter computation'],
        ['', '', ''],
        ['8.0', 'AUTHORIZATION & REVIEW CADENCE', ''],
        ['8.1', 'Monthly: parameters recomputed on month-close data', 'Automated calculation, reviewed by Planning Manager'],
        ['8.2', 'Quarterly: ABC reclassification & service level review', 'Approved by SCM Head, documented in meeting minutes'],
        ['8.3', 'Annual: methodology review with Finance/Audit', 'Z-scores, lead times, order cycle validated against actuals'],
    ]
    for ri, row_data in enumerate(meth, 1):
        for ci, val in enumerate(row_data, 1):
            cell = ws6.cell(row=ri, column=ci, value=val)
            if ri == 1:
                cell.font = Font(bold=True, size=14)
            elif ri == 3:
                cell.font = Font(bold=True, size=10)
            elif val and str(val).endswith('0') and '.' in str(val) and len(str(val)) == 3:
                cell.font = Font(bold=True, size=10)
            elif any(str(val).startswith(f'{x}.0') for x in range(1, 9)):
                cell.font = Font(bold=True, size=10)
    ws6.column_dimensions['A'].width = 8
    ws6.column_dimensions['B'].width = 55
    ws6.column_dimensions['C'].width = 70

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.post("/api/process")
@app.post("/process")
async def process_files(
    sigma_file: UploadFile = File(...),
    abc_file: UploadFile = File(...),
    proj_file: UploadFile = File(None),
    lead_time: float = Form(2),
    z_a: float = Form(1.65),
    z_b: float = Form(1.28),
    z_c: float = Form(1.04),
    order_cycle: float = Form(1.0),
):
    sigma_bytes = await sigma_file.read()
    abc_bytes = await abc_file.read()

    sigma_df = pd.read_excel(BytesIO(sigma_bytes), sheet_name=0)
    abc_df = pd.read_excel(BytesIO(abc_bytes), sheet_name='abc_summary').dropna(subset=['Item_Name'])

    proj_df = None
    if proj_file and proj_file.filename:
        proj_bytes = await proj_file.read()
        if proj_bytes:
            try:
                proj_xls = pd.ExcelFile(BytesIO(proj_bytes))
                if 'Query Report' in proj_xls.sheet_names:
                    proj_df = pd.read_excel(proj_xls, sheet_name='Query Report')
                elif 'Projection' in proj_xls.sheet_names:
                    proj_df = pd.read_excel(proj_xls, sheet_name='Projection')
                else:
                    proj_df = pd.read_excel(proj_xls, sheet_name=0)
            except Exception:
                proj_df = None

    z_scores = [z_a, z_b, z_c]
    inv_df, months = compute_inventory(sigma_df, abc_df, proj_df, lead_time, z_scores, order_cycle)

    CACHE['inv_df'] = inv_df
    CACHE['months'] = months
    CACHE['params'] = {'lead_time': lead_time, 'z_scores': z_scores, 'order_cycle': order_cycle}

    # Build JSON response
    inv_cols = ['Item_Name', 'ABC_Category', 'New MIS ITEM Group', 'Origin', 'MRP',
                'Service_Level', 'Lead_Time_Days', 'Conversion_Factor',
                'Recent_Avg', 'Recent_Std', 'CoV', 'Z_Score',
                'Safety_Stock', 'Daily_Demand', 'Min_Inventory_ROP', 'Order_Qty',
                'Max_Inventory', 'Avg_Inventory',
                'DOH_MIN', 'DOH_MAX', 'DOH_AVG',
                'Safety_Stock_Value', 'Min_Value', 'Max_Value', 'Avg_Value',
                'Avg_Demand_KGS', 'Daily_Demand_KGS', 'Safety_Stock_KGS',
                'Min_Inventory_KGS', 'Max_Inventory_KGS', 'Avg_Inventory_KGS',
                'DOH_KGS_MIN', 'DOH_KGS_MAX', 'DOH_KGS_AVG',
                'Actual_DOH_Latest', 'DOH_Variance', 'DOH_Status', 'Trend']
    existing_cols = [c for c in inv_cols if c in inv_df.columns]
    rows = [{c: safe_json(r[c]) for c in existing_cols} for _, r in inv_df[existing_cols].iterrows()]

    dev_months = months[1:]
    dev_data = []
    for _, r in inv_df.iterrows():
        d = {'Item_Name': r['Item_Name'], 'ABC_Category': safe_json(r.get('ABC_Category', '')),
             'New MIS ITEM Group': r['New MIS ITEM Group']}
        for m in dev_months:
            d[f'Dev_{m}'] = safe_json(r.get(f'Dev_{m}', 0))
            d[f'PctDev_{m}'] = safe_json(r.get(f'PctDev_{m}', 0))
            d[f'StockDev_{m}'] = safe_json(r.get(f'StockDev_{m}', 0))
        dev_data.append(d)

    forecast_data = []
    if 'Projection_Units' in inv_df.columns:
        fc = inv_df[inv_df['Projection_Units'].notna()]
        for _, r in fc.iterrows():
            dev_pct = safe_json(r.get('Forecast_Deviation_Pct', None))
            risk = 'N/A'
            if dev_pct is not None:
                risk = 'HIGH' if abs(dev_pct) > 50 else ('MEDIUM' if abs(dev_pct) > 25 else 'LOW')
            forecast_data.append({
                'Item_Name': r['Item_Name'],
                'ABC_Category': safe_json(r.get('ABC_Category', '')),
                'New MIS ITEM Group': r['New MIS ITEM Group'],
                'Recent_Avg': safe_json(r['Recent_Avg']),
                'Projection_Units': safe_json(r['Projection_Units']),
                'Projection_KGS': safe_json(r.get('Projection_KGS')),
                'Avg_Demand_KGS': safe_json(r.get('Avg_Demand_KGS')),
                'Deviation_Units': safe_json(r['Projection_Units'] - r['Recent_Avg']),
                'Deviation_Pct': dev_pct,
                'Risk': risk,
            })

    groups = sorted(inv_df['New MIS ITEM Group'].dropna().unique().tolist())
    origins = sorted(inv_df['Origin'].dropna().unique().tolist())

    summary = {}
    for cat in ['A', 'B', 'C']:
        s = inv_df[inv_df['ABC_Category'] == cat]
        if len(s) > 0:
            summary[cat] = {
                'count': len(s),
                'sl': {'A': f"{norm_cdf(z_a):.0%}", 'B': f"{norm_cdf(z_b):.0%}", 'C': f"{norm_cdf(z_c):.0%}"}[cat],
                'avg_ss': round(float(s['Safety_Stock'].mean()), 0),
                'total_ss_cr': round(float(s['Safety_Stock_Value'].sum()) / 1e7, 2),
                'total_min_cr': round(float(s['Min_Value'].sum()) / 1e7, 2),
                'total_max_cr': round(float(s['Max_Value'].sum()) / 1e7, 2),
                'avg_doh_min': round(float(s['DOH_MIN'].mean()), 1),
                'avg_doh_max': round(float(s['DOH_MAX'].mean()), 1),
                'avg_doh_avg': round(float(s['DOH_AVG'].mean()), 1),
                'avg_ss_kgs': round(float(s['Safety_Stock_KGS'].mean()), 1),
                'growing': int((s['Trend'] == 'Growing').sum()),
                'declining': int((s['Trend'] == 'Declining').sum()),
                'stable': int((s['Trend'] == 'Stable').sum()),
                'excess_count': int((s['DOH_Status'] == 'EXCESS').sum()),
                'optimal_count': int((s['DOH_Status'] == 'OPTIMAL').sum()),
                'below_count': int((s['DOH_Status'] == 'BELOW ROP').sum()),
            }

    totals = {
        'skus': len(inv_df),
        'ss_value_cr': round(float(inv_df['Safety_Stock_Value'].sum()) / 1e7, 2),
        'min_value_cr': round(float(inv_df['Min_Value'].sum()) / 1e7, 2),
        'max_value_cr': round(float(inv_df['Max_Value'].sum()) / 1e7, 2),
        'avg_value_cr': round(float(inv_df['Avg_Value'].sum()) / 1e7, 2),
        'total_ss_kgs': round(float(inv_df['Safety_Stock_KGS'].sum()), 0),
        'total_min_kgs': round(float(inv_df['Min_Inventory_KGS'].sum()), 0),
        'total_max_kgs': round(float(inv_df['Max_Inventory_KGS'].sum()), 0),
    }

    return JSONResponse({
        'inventory': rows,
        'deviations': dev_data,
        'forecast': forecast_data,
        'summary': summary,
        'totals': totals,
        'filters': {'groups': groups, 'origins': origins},
        'months': months,
        'dev_months': dev_months,
        'params': {'lead_time': lead_time, 'z_scores': z_scores, 'order_cycle': order_cycle},
    })


@app.get("/api/download")
@app.get("/download")
async def download_excel():
    if 'inv_df' not in CACHE:
        return JSONResponse({'error': 'No data processed yet'}, status_code=400)
    inv_df = CACHE['inv_df']
    months = CACHE['months']
    p = CACHE['params']
    buf = build_formula_excel(inv_df, months, p['lead_time'], p['z_scores'], p['order_cycle'])
    return StreamingResponse(
        buf,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.xml',
        headers={'Content-Disposition': 'attachment; filename=Farmley_Inventory_Model.xlsx'}
    )
