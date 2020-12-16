import pandas as pd
import numpy as np

# Reading Files

## SKU Master
sku = pd.read_excel('BHUS SKU Master.xlsx', sheet_name='Master', header=1)
sku = sku.dropna(axis=1, how='all')
sku = sku.dropna(axis=0, how='all')

## Inventory
inv = pd.read_csv('All Brands - Current Inventory by Brand(New).csv', header=2)

## Disallow
disallow = pd.read_excel('Disallow.xlsx', header=0, usecols=[0])

# Customized Inputs
stock_threshold = 5

# Data Wrangling

## Determining brand through the Inventory file
brand = str(inv.Brand[0])
if brand == 'BAL': brand = 'BHUS'

## Combining the SKU Master and Inventory
df = pd.merge(sku, inv, how='left', left_on='Product Code', right_on='SKU')
df['Stock'] = df['Stock'].replace(np.nan, 0)

## Add stocks of children into parent's
for i in df.index:
    if df.iloc[i]['Entity'] == 'PARENT':
        code = df.iloc[i]['Product Code']
        df.iloc[i, df.columns.get_loc('Stock')] += df[df['Parent Code']==code]['Stock'].sum()
    else:
        continue

## Convert Season column into a numpy array of [Christmas, Fall, Spring, Summer]
df['Season'] = df['Season'].str.split(',')
df['Season'] = df['Season'].replace(np.nan, 'Year-Round') # assumes products without Season indicated is year-round
season = pd.get_dummies(df['Season'].apply(pd.Series).stack()).sum(level=0)
### If a product has Year-Round in the list, assign True for all seasons
for i, row in season.iterrows():
    if (row['Year-Round'] == 1):
        row.loc[: ~row['Year-Round']] = 1
season = season.drop(columns=['Year-Round'])
df['Season'] = season.apply(lambda r: tuple(r), axis=1).apply(np.array)

## Drop unnecessary columns
df = df.drop(columns=['Drop-Shipped?', 'Include in Pricing?', 'Included in OLD File', 'Brand', 'SKU'])

# Filters

## Filter for Pool of Main Products
df = df[(df['Status']!='REMOVED') & (df['Category'].notnull())]
df = df.reset_index(drop=True)

## Filter for Pool of Cross-sells
df_filtered = df[df['Stock'] >= stock_threshold]
df_filtered = df_filtered[~df_filtered['Product Code'].isin(disallow['Product Code'])]

text = 'Out of %d products, %d will be used for recommendations.' %(df.shape[0], df_filtered.shape[0])
print(text)

## Entity Eligibilty
### Entity Relationships
###    * Standalone -> Standalone, Children (Soft-Cart)
###    * Standalone -> Standalone, Parent (Product Page)
###    * Parent -> Parent, Standalone (Product Page)
###    * Children -> Children, Standalone (Soft-Cart)
sc = df[df['Entity']!='PARENT']
for_sc = df_filtered[df_filtered['Entity']!='PARENT']
pdp = df[df['Entity']!='CHILDREN']
for_pdp = df_filtered[df_filtered['Entity']!='CHILDREN']

# Cross-sell
def reco(i, for_):
    # Removing the product from the recommendation pool
    output = for_
    if i in for_.index: output = for_.drop(i)
    # Cross-sell placeholder
    import random
    try:
        p = random.sample(list(output.index), 5)
    except:
        p = np.nan
    return p

## Applying cross-sell to soft cart and pdp
sc_recommendations = {i : reco(i, for_sc) for i in sc.index}
pdp_recommendations = {i : reco(i, for_pdp) for i in pdp.index}

# Convert to Readable
i_to_code = df['Product Code'].to_dict()
i_to_name = df['Product Name'].to_dict()
i_to_pimid = df['PIMID'].to_dict()

sc_recs = pd.DataFrame.from_dict(sc_recommendations, orient='index')
sc_recs.columns += 1
sc_recs['Place'] = 'Soft Cart'
pdp_recs = pd.DataFrame.from_dict(pdp_recommendations, orient='index')
pdp_recs.columns += 1
pdp_recs['Place'] = 'Product Page'

## Combine SC and PDP
results = pd.concat([sc_recs, pdp_recs])
results_code = results.rename(mapper=i_to_code).replace(i_to_code)
results_name = results.rename(mapper=i_to_name).replace(i_to_name)

## Excel Writer
writer = pd.ExcelWriter('cross-sell.xlsx', engine='xlsxwriter')
results_code.to_excel(writer, 'SKU')
results_name.to_excel(writer, 'Product Name')
worksheet = writer.sheets['Product Name']
for idx, col in enumerate(results_name):  
    series = results_name[col]
    worksheet.set_column(idx, idx, 30) # Sets default column width
writer.save()

# Override



# Convert to PIM
columns = ['Id', 'Relationship External Id', 'Relationship Type', 'From Entity External Id', 'From Entity Type',
           'From Entity Category Path', 'From Entity Container', 'From Entity Organization',
           'To Entity External Id', 'To Entity Type', 'To Entity Category Path', 'To Entity Container',
           'To Entity Organization', 'Cross-Sell Attributes//Cross-Sell Location',
           'Cross-Sell Attributes//Sort Order']
pim = pd.DataFrame(columns=columns)
results_pim = results.rename(mapper=i_to_pimid).replace(i_to_pimid)
data = pd.melt(results_pim.reset_index(), id_vars=['index', 'Place']).sort_values(by=['index','variable']).reset_index(drop=True)
pim.iloc[:, pim.columns.get_loc('From Entity External Id')] = data.iloc[:,0]
pim.iloc[:, pim.columns.get_loc('Cross-Sell Attributes//Sort Order')] = data.iloc[:,2]
pim.iloc[:, pim.columns.get_loc('To Entity External Id')] = data.iloc[:,3]
pim.iloc[:, pim.columns.get_loc('Cross-Sell Attributes//Cross-Sell Location')] = data.iloc[:,1]
pim['Relationship Type'] = 'Cross-Sell Relationship'

## Downloading Excel file for upload
import datetime
now = datetime.datetime.now()
file_name = brand + 'WebOutbound-AllEntities-StagingDelta_(' + now.strftime("%Y-%b-%dT%H.%M.%S.%f")[:-3] + ')_1of1.xlsx'
pim.to_excel(file_name)
