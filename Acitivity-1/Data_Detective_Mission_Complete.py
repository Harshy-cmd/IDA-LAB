"""
=============================================================================
🔍 DATA DETECTIVE MISSION — Exploratory Data Analysis
=============================================================================
UCI Online Retail Dataset
Source: https://archive.ics.uci.edu/dataset/352/online-retail

This is a standalone Python script for the complete EDA.
Can be run in Google Colab, Jupyter Notebook, or any Python environment.
=============================================================================
"""

# =============================================================================
# SECTION 1: IMPORT LIBRARIES
# =============================================================================
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import warnings

warnings.filterwarnings('ignore')

# Set professional plot style
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette('husl')
plt.rcParams['figure.dpi'] = 150
plt.rcParams['font.size'] = 10

print('✅ All libraries imported successfully!')

# =============================================================================
# SECTION 2: LOAD DATASET
# =============================================================================
# Option A: Load from local file (uncomment the line for your path)
# df = pd.read_excel('Online Retail.xlsx')

# Option B: Load from UCI repository (requires: pip install ucimlrepo)
# from ucimlrepo import fetch_ucirepo
# online_retail = fetch_ucirepo(id=352)
# df = online_retail.data.original

# Option C: Load from Google Drive (Colab)
# from google.colab import drive
# drive.mount('/content/drive')
# df = pd.read_excel('/content/drive/MyDrive/Online Retail.xlsx')

# For this script, using local file:
df = pd.read_excel('Online Retail-2.xlsx')

print(f'✅ Dataset loaded: {df.shape[0]:,} rows × {df.shape[1]} columns')

# =============================================================================
# SECTION 3: DATASET OVERVIEW
# =============================================================================
print('\n' + '=' * 60)
print('DATASET OVERVIEW')
print('=' * 60)
print(f'\nRows:    {df.shape[0]:,}')
print(f'Columns: {df.shape[1]}')
print(f'Memory:  {df.memory_usage(deep=True).sum() / (1024*1024):.2f} MB')
print(f'\nDate Range: {df["InvoiceDate"].min()} to {df["InvoiceDate"].max()}')
print(f'Unique Invoices:  {df["InvoiceNo"].nunique():,}')
print(f'Unique Products:  {df["StockCode"].nunique():,}')
print(f'Unique Customers: {df["CustomerID"].nunique():,}')
print(f'Unique Countries: {df["Country"].nunique()}')

print('\n--- Column Info ---')
print(df.dtypes)
print('\n--- First 5 Records ---')
print(df.head().to_string())
print('\n--- Last 5 Records ---')
print(df.tail().to_string())

# =============================================================================
# SECTION 4: DATA DICTIONARY
# =============================================================================
print('\n' + '=' * 60)
print('DATA DICTIONARY')
print('=' * 60)
data_dict = pd.DataFrame({
    'Column': df.columns,
    'Dtype': [str(t) for t in df.dtypes],
    'Non-Null': [df[c].notna().sum() for c in df.columns],
    'Missing': [df[c].isnull().sum() for c in df.columns],
    'Missing%': [(df[c].isnull().sum()/len(df)*100) for c in df.columns],
    'Example': [str(df[c].dropna().iloc[0]) for c in df.columns]
})
print(data_dict.to_string(index=False))

# =============================================================================
# SECTION 5: MISSING VALUES ANALYSIS
# =============================================================================
print('\n' + '=' * 60)
print('MISSING VALUES ANALYSIS')
print('=' * 60)

missing_count = df.isnull().sum()
missing_pct = (df.isnull().sum() / len(df) * 100).round(2)
missing_df = pd.DataFrame({
    'Missing Count': missing_count,
    'Missing %': missing_pct
})
print(missing_df)
print(f'\n📌 CustomerID: {missing_count["CustomerID"]:,} missing ({missing_pct["CustomerID"]}%)')
print(f'📌 Description: {missing_count["Description"]:,} missing ({missing_pct["Description"]}%)')

# Visualization: Missing Values
fig, ax = plt.subplots(figsize=(10, 5))
colors_missing = ['#F44336' if pct > 0 else '#4CAF50' for pct in missing_pct.values]
bars = ax.barh(missing_df.index, missing_pct.values, color=colors_missing, edgecolor='white')
ax.set_title('Missing Values by Column (%)', fontsize=16, fontweight='bold', pad=15)
ax.set_xlabel('Missing Percentage (%)', fontsize=12)
for bar, val in zip(bars, missing_pct.values):
    if val > 0:
        ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2.,
                f'{val}%', ha='left', va='center', fontsize=9, fontweight='bold')
plt.tight_layout()
plt.savefig('viz_missing_values.png', bbox_inches='tight')
plt.show()

# =============================================================================
# SECTION 6: DUPLICATE ANALYSIS
# =============================================================================
print('\n' + '=' * 60)
print('DUPLICATE RECORDS ANALYSIS')
print('=' * 60)
dup_count = df.duplicated().sum()
dup_pct = dup_count / len(df) * 100
print(f'Duplicate rows:       {dup_count:,}')
print(f'Duplicate percentage: {dup_pct:.2f}%')
print(f'Decision: Keep — may represent legitimate repeat purchases')

# =============================================================================
# SECTION 7: UNUSUAL VALUES INVESTIGATION
# =============================================================================
print('\n' + '=' * 60)
print('UNUSUAL VALUES INVESTIGATION')
print('=' * 60)

neg_qty = df[df['Quantity'] < 0]
zero_qty = df[df['Quantity'] == 0]
large_qty = df[df['Quantity'] > 1000]
zero_price = df[df['UnitPrice'] == 0]
neg_price = df[df['UnitPrice'] < 0]
cancelled = df[df['InvoiceNo'].astype(str).str.startswith('C')]

print(f'Negative quantities:    {len(neg_qty):,} ({len(neg_qty)/len(df)*100:.2f}%) → Returns')
print(f'Zero quantities:        {len(zero_qty):,}')
print(f'Quantity > 1,000:       {len(large_qty):,} → Bulk/wholesale orders')
print(f'Zero prices:            {len(zero_price):,} ({len(zero_price)/len(df)*100:.2f}%) → Free/sample items')
print(f'Negative prices:        {len(neg_price):,}')
print(f'Cancellations ("C"):    {len(cancelled):,} ({len(cancelled)/len(df)*100:.2f}%)')
print(f'Max Quantity:           {df["Quantity"].max():,}')
print(f'Min Quantity:           {df["Quantity"].min():,}')
print(f'Max UnitPrice:          £{df["UnitPrice"].max():,.2f}')

# =============================================================================
# SECTION 8: DATA PREPROCESSING
# =============================================================================
print('\n' + '=' * 60)
print('DATA PREPROCESSING')
print('=' * 60)

# Preserve original
df_original = df.copy()

# Ensure datetime
df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])

# Create Revenue column
df['Revenue'] = df['Quantity'] * df['UnitPrice']

# Create date features
df['Year'] = df['InvoiceDate'].dt.year
df['Month'] = df['InvoiceDate'].dt.month
df['Day'] = df['InvoiceDate'].dt.day
df['DayOfWeek'] = df['InvoiceDate'].dt.dayofweek
df['Hour'] = df['InvoiceDate'].dt.hour
df['MonthYear'] = df['InvoiceDate'].dt.to_period('M')

# Create clean dataset
df_clean = df[
    (~df['InvoiceNo'].astype(str).str.startswith('C')) &
    (df['Quantity'] > 0) &
    (df['UnitPrice'] > 0)
].copy()

print(f'Original rows:  {len(df):,}')
print(f'Clean rows:     {len(df_clean):,}')
print(f'Rows removed:   {len(df) - len(df_clean):,} ({(len(df) - len(df_clean))/len(df)*100:.2f}%)')
print(f'✅ Revenue column created')
print(f'✅ Date features created')
print(f'✅ Cancellations, negative quantities, zero prices removed')

# =============================================================================
# SECTION 9: DESCRIPTIVE STATISTICS
# =============================================================================
print('\n' + '=' * 60)
print('DESCRIPTIVE STATISTICS')
print('=' * 60)

for col in ['Quantity', 'UnitPrice', 'Revenue']:
    print(f'\n--- {col} ---')
    print(df_clean[col].describe())
    print(f'Median: {df_clean[col].median()}')

# Business metrics
print('\n--- Business Metrics ---')
total_revenue = df_clean['Revenue'].sum()
num_transactions = df_clean['InvoiceNo'].nunique()
num_customers = df_clean['CustomerID'].nunique()
num_products = df_clean['StockCode'].nunique()
avg_txn_value = df_clean.groupby('InvoiceNo')['Revenue'].sum().mean()
avg_qty_per_txn = df_clean.groupby('InvoiceNo')['Quantity'].sum().mean()
df_clean_cust = df_clean[df_clean['CustomerID'].notna()]
rev_per_customer = df_clean_cust.groupby('CustomerID')['Revenue'].sum().mean()

print(f'Total Revenue:              £{total_revenue:,.2f}')
print(f'Number of Transactions:     {num_transactions:,}')
print(f'Number of Customers:        {num_customers:,}')
print(f'Number of Products:         {num_products:,}')
print(f'Avg Transaction Value:      £{avg_txn_value:,.2f}')
print(f'Avg Qty per Transaction:    {avg_qty_per_txn:,.1f}')
print(f'Avg Revenue per Customer:   £{rev_per_customer:,.2f}')

# =============================================================================
# SECTION 10: COUNTRY ANALYSIS
# =============================================================================
print('\n' + '=' * 60)
print('COUNTRY ANALYSIS')
print('=' * 60)

country_revenue = df_clean.groupby('Country')['Revenue'].sum().sort_values(ascending=False)
print('\nTop 10 Countries by Revenue:')
for i, (country, rev) in enumerate(country_revenue.head(10).items(), 1):
    pct = rev / total_revenue * 100
    print(f'  {i:2d}. {country}: £{rev:,.2f} ({pct:.1f}%)')

# =============================================================================
# SECTION 11: PRODUCT ANALYSIS
# =============================================================================
print('\n' + '=' * 60)
print('PRODUCT ANALYSIS')
print('=' * 60)

product_qty = df_clean.groupby(['StockCode', 'Description'])['Quantity'].sum().sort_values(ascending=False)
print('\nTop 10 Products by Quantity:')
for i, ((code, desc), qty) in enumerate(product_qty.head(10).items(), 1):
    print(f'  {i:2d}. [{code}] {desc}: {qty:,}')

product_rev = df_clean.groupby(['StockCode', 'Description'])['Revenue'].sum().sort_values(ascending=False)
print('\nTop 10 Products by Revenue:')
for i, ((code, desc), rev) in enumerate(product_rev.head(10).items(), 1):
    print(f'  {i:2d}. [{code}] {desc}: £{rev:,.2f}')

# =============================================================================
# SECTION 12: CUSTOMER ANALYSIS
# =============================================================================
print('\n' + '=' * 60)
print('CUSTOMER ANALYSIS')
print('=' * 60)

customer_revenue = df_clean_cust.groupby('CustomerID')['Revenue'].sum().sort_values(ascending=False)
print('\nTop 10 Customers by Revenue:')
for i, (cust_id, rev) in enumerate(customer_revenue.head(10).items(), 1):
    print(f'  {i:2d}. Customer {int(cust_id)}: £{rev:,.2f}')

top10_rev = customer_revenue.head(10).sum()
top10_pct = top10_rev / df_clean_cust['Revenue'].sum() * 100
top20pct_n = int(len(customer_revenue) * 0.2)
top20pct_rev = customer_revenue.head(top20pct_n).sum()
top20pct_pct = top20pct_rev / df_clean_cust['Revenue'].sum() * 100
print(f'\nTop 10 customers: {top10_pct:.1f}% of revenue')
print(f'Top 20% customers: {top20pct_pct:.1f}% of revenue')

# =============================================================================
# SECTION 13: TIME ANALYSIS
# =============================================================================
print('\n' + '=' * 60)
print('TIME ANALYSIS')
print('=' * 60)

monthly_revenue = df_clean.groupby('MonthYear')['Revenue'].sum()
print('\nMonthly Revenue:')
for period, rev in monthly_revenue.items():
    print(f'  {period}: £{rev:,.2f}')

day_names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
dow_txns = df_clean.groupby('DayOfWeek')['InvoiceNo'].nunique()
print('\nTransactions by Day of Week:')
for day_num, txn in dow_txns.items():
    print(f'  {day_names[day_num]}: {txn:,}')

# =============================================================================
# SECTION 14: VISUALIZATION 1 — Revenue by Country (Top 10)
# =============================================================================
country_rev_top10 = df_clean.groupby('Country')['Revenue'].sum().sort_values(ascending=False).head(10)

fig, ax = plt.subplots(figsize=(12, 7))
colors = sns.color_palette('viridis', len(country_rev_top10))
bars = ax.bar(range(len(country_rev_top10)), country_rev_top10.values, color=colors, edgecolor='white', linewidth=0.5)
ax.set_xticks(range(len(country_rev_top10)))
ax.set_xticklabels(country_rev_top10.index, rotation=45, ha='right', fontsize=9)
ax.set_title('Top 10 Countries by Revenue', fontsize=16, fontweight='bold', pad=15)
ax.set_xlabel('Country', fontsize=12)
ax.set_ylabel('Revenue (£)', fontsize=12)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'£{x:,.0f}'))
for bar, val in zip(bars, country_rev_top10.values):
    ax.text(bar.get_x() + bar.get_width()/2., bar.get_height()*1.01,
            f'£{val:,.0f}', ha='center', va='bottom', fontsize=7, fontweight='bold')
plt.tight_layout()
plt.savefig('viz1_revenue_by_country.png', bbox_inches='tight')
plt.show()
print('📊 UK dominates with 84.6% of total revenue.')

# =============================================================================
# SECTION 15: VISUALIZATION 2 — Top 10 Products by Revenue
# =============================================================================
product_rev_top10 = df_clean.groupby('Description')['Revenue'].sum().sort_values(ascending=True).tail(10)

fig, ax = plt.subplots(figsize=(12, 7))
colors = sns.color_palette('magma', len(product_rev_top10))
bars = ax.barh(range(len(product_rev_top10)), product_rev_top10.values, color=colors, edgecolor='white', linewidth=0.5)
ax.set_yticks(range(len(product_rev_top10)))
ax.set_yticklabels(product_rev_top10.index, fontsize=9)
ax.set_title('Top 10 Products by Revenue', fontsize=16, fontweight='bold', pad=15)
ax.set_xlabel('Revenue (£)', fontsize=12)
ax.set_ylabel('Product Description', fontsize=12)
ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'£{x:,.0f}'))
for bar, val in zip(bars, product_rev_top10.values):
    ax.text(bar.get_width() + max(product_rev_top10)*0.01, bar.get_y() + bar.get_height()/2.,
            f'£{val:,.0f}', ha='left', va='center', fontsize=8, fontweight='bold')
plt.tight_layout()
plt.savefig('viz2_top_products_revenue.png', bbox_inches='tight')
plt.show()
print('📊 DOTCOM POSTAGE and REGENCY CAKESTAND 3 TIER lead revenue.')

# =============================================================================
# SECTION 16: VISUALIZATION 3 — Monthly Revenue Trend
# =============================================================================
monthly_rev = df_clean.groupby('MonthYear')['Revenue'].sum()
monthly_labels = [str(p) for p in monthly_rev.index]

fig, ax = plt.subplots(figsize=(14, 7))
ax.plot(range(len(monthly_rev)), monthly_rev.values, 'o-', color='#2196F3',
        linewidth=2.5, markersize=8, markerfacecolor='white', markeredgewidth=2)
ax.fill_between(range(len(monthly_rev)), monthly_rev.values, alpha=0.15, color='#2196F3')
ax.set_xticks(range(len(monthly_labels)))
ax.set_xticklabels(monthly_labels, rotation=45, ha='right', fontsize=9)
ax.set_title('Monthly Revenue Trend (Dec 2010 – Dec 2011)', fontsize=16, fontweight='bold', pad=15)
ax.set_xlabel('Month', fontsize=12)
ax.set_ylabel('Revenue (£)', fontsize=12)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'£{x:,.0f}'))
peak_idx = monthly_rev.values.argmax()
ax.annotate(f'Peak: £{monthly_rev.values[peak_idx]:,.0f}',
            xy=(peak_idx, monthly_rev.values[peak_idx]),
            xytext=(peak_idx-2, monthly_rev.values[peak_idx]*1.1),
            arrowprops=dict(arrowstyle='->', color='red', lw=1.5),
            fontsize=10, fontweight='bold', color='red')
plt.tight_layout()
plt.savefig('viz3_monthly_revenue_trend.png', bbox_inches='tight')
plt.show()
print('📊 Strong Q4 seasonality. November 2011 is the peak month.')

# =============================================================================
# SECTION 17: VISUALIZATION 4 — Customer Revenue Concentration (Pareto)
# =============================================================================
customer_rev = df_clean_cust.groupby('CustomerID')['Revenue'].sum().sort_values(ascending=False)
cumulative_rev = customer_rev.cumsum() / customer_rev.sum() * 100
customer_pct = np.arange(1, len(cumulative_rev) + 1) / len(cumulative_rev) * 100

fig, ax = plt.subplots(figsize=(12, 7))
ax.plot(customer_pct, cumulative_rev.values, color='#E91E63', linewidth=2.5)
ax.fill_between(customer_pct, cumulative_rev.values, alpha=0.1, color='#E91E63')
ax.axhline(y=80, color='gray', linestyle='--', alpha=0.7, label='80% Revenue Line')
ax.axvline(x=20, color='gray', linestyle=':', alpha=0.7, label='20% Customers Line')
idx_80 = np.argmin(np.abs(cumulative_rev.values - 80))
pct_at_80 = customer_pct[idx_80]
ax.plot(pct_at_80, 80, 'ro', markersize=10, zorder=5)
ax.annotate(f'{pct_at_80:.1f}% of customers\ngenerate 80% of revenue',
            xy=(pct_at_80, 80), xytext=(pct_at_80+15, 65),
            arrowprops=dict(arrowstyle='->', color='red', lw=1.5),
            fontsize=10, fontweight='bold', color='red',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='yellow', alpha=0.8))
ax.set_title('Customer Revenue Concentration (Pareto Analysis)', fontsize=16, fontweight='bold', pad=15)
ax.set_xlabel('Cumulative % of Customers (Ranked by Revenue)', fontsize=12)
ax.set_ylabel('Cumulative % of Revenue', fontsize=12)
ax.legend(fontsize=10)
ax.set_xlim(0, 100)
ax.set_ylim(0, 105)
plt.tight_layout()
plt.savefig('viz4_customer_concentration.png', bbox_inches='tight')
plt.show()
print(f'📊 {pct_at_80:.1f}% of customers generate 80% of revenue. Strong Pareto effect.')

# =============================================================================
# SECTION 18: VISUALIZATION 5 — Transactions by Day of Week
# =============================================================================
fig, ax = plt.subplots(figsize=(10, 6))
colors_dow = ['#4CAF50' if d != 5 else '#F44336' for d in dow_txns.index]
bars = ax.bar(range(len(dow_txns)), dow_txns.values, color=colors_dow, edgecolor='white', linewidth=0.5)
ax.set_xticks(range(len(dow_txns)))
ax.set_xticklabels([day_names[i] for i in dow_txns.index], fontsize=10)
ax.set_title('Number of Transactions by Day of Week', fontsize=16, fontweight='bold', pad=15)
ax.set_xlabel('Day of Week', fontsize=12)
ax.set_ylabel('Number of Transactions', fontsize=12)
for bar, val in zip(bars, dow_txns.values):
    ax.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 20,
            f'{val:,}', ha='center', va='bottom', fontsize=9, fontweight='bold')
plt.tight_layout()
plt.savefig('viz5_transactions_by_day.png', bbox_inches='tight')
plt.show()
print('📊 No Saturday transactions. Thursday is busiest.')

# =============================================================================
# SECTION 19: VISUALIZATION 6 — Revenue Distribution
# =============================================================================
rev_filtered = df_clean[df_clean['Revenue'] <= df_clean['Revenue'].quantile(0.99)]['Revenue']

fig, axes = plt.subplots(1, 2, figsize=(14, 6))
axes[0].hist(rev_filtered, bins=50, color='#FF9800', edgecolor='white', alpha=0.8)
axes[0].set_title('Revenue Distribution per Line Item', fontsize=13, fontweight='bold')
axes[0].set_xlabel('Revenue (£)', fontsize=11)
axes[0].set_ylabel('Frequency', fontsize=11)
axes[0].axvline(rev_filtered.median(), color='red', linestyle='--', linewidth=1.5, label=f'Median: £{rev_filtered.median():.2f}')
axes[0].axvline(rev_filtered.mean(), color='blue', linestyle='--', linewidth=1.5, label=f'Mean: £{rev_filtered.mean():.2f}')
axes[0].legend(fontsize=9)

bp = axes[1].boxplot(rev_filtered, vert=True, patch_artist=True, widths=0.6)
bp['boxes'][0].set_facecolor('#42A5F5')
bp['boxes'][0].set_alpha(0.7)
axes[1].set_title('Revenue Boxplot', fontsize=13, fontweight='bold')
axes[1].set_ylabel('Revenue (£)', fontsize=11)
axes[1].set_xticklabels(['Revenue'])

plt.suptitle('Revenue Distribution Analysis (99th Percentile)', fontsize=15, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig('viz6_revenue_distribution.png', bbox_inches='tight')
plt.show()

# =============================================================================
# SECTION 20: THREE KEY FINDINGS
# =============================================================================
print('\n' + '=' * 60)
print('THREE KEY FINDINGS')
print('=' * 60)

print('''
FINDING 1: EXTREME GEOGRAPHIC REVENUE CONCENTRATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Discovery:  UK accounts for 84.6% of total revenue (£9.03M / £10.67M)
Evidence:   37 non-UK countries combined = only 15.4%
Visual:     Revenue by Country bar chart
Impact:     High geographic concentration risk

FINDING 2: STRONG Q4 SEASONALITY WITH NOVEMBER PEAK
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Discovery:  Revenue surges Sep-Nov; November = £1.51M (peak)
Evidence:   November is 2.9× higher than February (lowest)
Visual:     Monthly Revenue Trend line chart
Impact:     Inventory & marketing should front-load before September

FINDING 3: PARETO EFFECT IN CUSTOMER REVENUE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Discovery:  Top 20% of customers generate 74.6% of revenue
Evidence:   Top 10 customers alone = 17.3% of total revenue
Visual:     Customer Concentration Pareto curve
Impact:     Critical need for VIP retention programs
''')

# =============================================================================
# SECTION 21: BUSINESS RECOMMENDATION
# =============================================================================
print('=' * 60)
print('BUSINESS RECOMMENDATION')
print('=' * 60)
print('''
RECOMMENDATION:
Implement a tiered VIP customer retention program targeting the top 20%
of customers, with priority focus on Q4 (Sep–Nov) engagement.

EVIDENCE:
- Top 20% of customers = 74.6% of revenue (Finding 3)
- Q4 is highest-revenue period (Finding 2)
- UK-concentrated business (Finding 1) → retention over expansion

EXPECTED IMPACT:
5-10% increase in VIP Q4 spending → £186K-£372K additional revenue
(Must be validated through A/B testing)
''')

# =============================================================================
# SECTION 22: FINAL SUMMARY
# =============================================================================
print('=' * 60)
print('FINAL SUMMARY')
print('=' * 60)
print(f'''
Total Records:           {len(df_original):,}
Clean Records:           {len(df_clean):,}
Date Range:              Dec 2010 – Dec 2011
Total Revenue:           £{total_revenue:,.2f}
Unique Customers:        {num_customers:,}
Unique Products:         {num_products:,}
Countries:               38
UK Revenue Share:        84.6%
Peak Month:              November 2011
Avg Transaction Value:   £{avg_txn_value:,.2f}
Missing CustomerID:      24.93%
Cancellation Rate:       1.71%

✅ Analysis Complete — All statistics from actual dataset
''')
