### GUI (Graphical User Interface) of the Cross-Sell Tool

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import numpy as np
import random
import os, sys, subprocess
import datetime

def open_file(filename):
    ### Opens file in its default application
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener ="open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])
## Taken from https://stackoverflow.com/questions/17317219/is-there-an-platform-independent-equivalent-of-os-startfile

def browse(title, entry):
    path = filedialog.askopenfilename(title=title)
    entry.delete(0, tk.END)
    entry.insert(0, path)     
  
class App:
    def __init__(self, window):
        frame_1 = tk.Frame(window, padx=10, pady=10)
        frame_1.grid(column=0, row=0)
        text='Welcome to the Cross-Sell application. Please upload the following files to proceed.'
        label = tk.Label(frame_1, text=text, font=('None', 15))
        label.grid(column=0, row=0, columnspan=6, sticky='W', pady=1)
        
        #### Uploading the required files

        ### SKU Master
        label = tk.Label(frame_1, text='Select the SKU Master file:')
        label.grid(column=0, row=1, sticky='W')
        self.entry1 = tk.Entry(frame_1, text='', width="40")
        self.entry1.grid(column=1, row=1, columnspan=3, sticky='W')

        def sel_sku():
            browse('Select the SKU Master file', self.entry1)

        btn = tk.Button(frame_1, text='Browse', command=sel_sku, width=15)
        btn.grid(column=4, row=1, sticky='W', columnspan=2)        
        
        ### Inventory
        label = tk.Label(frame_1, text='Select the Inventory file:')
        label.grid(column=0, row=2, sticky='W')
        self.entry2 = tk.Entry(frame_1, text='', width="40")
        self.entry2.grid(column=1, row=2, columnspan=3)

        def sel_inv():
            browse('Select the Inventory file', self.entry2)
     
        btn = tk.Button(frame_1, text='Browse', command=sel_inv, width=15)
        btn.grid(column=4, row=2, sticky='W', columnspan=2)

        ### Disallow
        label = tk.Label(frame_1, text='Select the Disallow file:')
        label.grid(column=0, row=3, sticky='W')
        self.entry3 = tk.Entry(frame_1, text='', width="40")
        self.entry3.grid(column=1, row=3, columnspan=3)

        def sel_dis():
            browse('Select the Disallow file', self.entry3)
     
        btn = tk.Button(frame_1, text='Browse', command=sel_dis, width=15)
        btn.grid(column=4, row=3, sticky='W', columnspan=2)

        ### Set stock threshold
        label = tk.Label(frame_1, text='Select the stock threshold:')
        label.grid(column=0, row=4,sticky='W')

        self.stock_threshold = tk.IntVar()
        self.entry4 = tk.Spinbox(frame_1, text=self.stock_threshold, from_=1, to=5, width='10')
        self.entry4.grid(column=1, row=4, columnspan=3, sticky='W')
        self.stock_threshold.set(5)

        #### Submit button
        self.btn = tk.Button(frame_1, text="Submit", command=lambda:self.clicked(frame_1), width=15)
        self.btn.grid(column=4, row=5, sticky='W', columnspan=2)


    def clicked(self, frame_1):
        sku = self.entry1.get()
        sku = pd.read_excel(sku, sheet_name='Master', header=1)
        #sku = sku.dropna(axis=1, how='all')
        sku = sku.dropna(axis=0, how='all')
        inv = self.entry2.get()
        inv = pd.read_csv(inv, header=0)
        brand = str(inv.Brand[0])
        if brand == 'BAL': brand = 'BHUS'
        disallow = self.entry3.get()
        disallow = pd.read_excel(disallow, header=0, usecols=[0])
        sku = sku[sku['Product Code'].notna()] #remove rows with blank SKU
        sku['Product Code'] = sku['Product Code'].astype(str)
        #For products without Sub-Category, inherit its Category
        sku['Sub-Category'] = sku['Sub-Category'].fillna(sku['Category'])
        inv['SKU'] = inv['SKU'].astype(str)
        df = pd.merge(sku, inv, how='left', left_on='Product Code', right_on='SKU')
        df['Stock'] = df['Stock'].replace(np.nan, 0)
        #df = df.drop(columns=['Drop-Shipped?', 'Brand', 'SKU'])
        # Removed 'Include in Pricing?' 'Included in OLD File']
        
        # For mapping purposes
        # df['Product Code']= df['Product Code'].apply(pd.to_numeric, errors='ignore')
        i_to_code = df['Product Code'].to_dict()
        i_to_name = df['Product Name'].to_dict()
        i_to_pimid = df['PIMID'].to_dict()
        # Counting stock status for PARENT entities
        for i in df.index: 
            if df.iloc[i]['Entity'] == 'PARENT':
                code = df.iloc[i]['Product Code']
                df.iloc[i, df.columns.get_loc('Stock')] += df[df['Parent Code']==code]['Stock'].sum()
            else:
                continue
        # Converting Season column to numpy array
        df['Season'] = df['Season'].str.split(',')
        df['Season'] = df['Season'].replace(np.nan, 'Year-Round') # assumes products without Season indicated is year-round
        season = pd.get_dummies(df['Season'].apply(pd.Series).stack()).sum(level=0)
        for i, row in season.iterrows():
            if (row['Year-Round'] == 1):
                row.loc[: ~row['Year-Round']] = 1 #all SKUs with Year-Round entry automatically applies for all seasons
        season = season.drop(columns=['Year-Round'])
        df['Season'] = season.apply(lambda r: tuple(r), axis=1).apply(np.array)
        
        # Filter products to be processed
        df = df[(df['Status']!='REMOVED') & (df['Category'].notnull())]
        # Filter products to be used as cross-sell
        stock_threshold = self.stock_threshold.get()
        df_filtered = df[df['Stock'] >= stock_threshold]
        df_filtered = df_filtered[~df_filtered['Product Code'].isin(disallow['Product Code'])]
        
        #SC and PDP
        sc = df[df['Entity']!='PARENT']
        for_sc = df_filtered[df_filtered['Entity']!='PARENT']
        pdp = df[df['Entity']!='CHILDREN']
        for_pdp = df_filtered[df_filtered['Entity']!='CHILDREN']
        
        ##FRAME 2
        frame_1.grid_forget()
        frame_2 = tk.Frame(window, padx=10, pady=10)
        frame_2.grid(column=0, row=0)
        # Display number of products eligible for cross-sells
        text = 'Out of %d products, %d will be used for recommendations.' %(df.shape[0], df_filtered.shape[0])
        tk.Label(frame_2, text=text, fg='#00B2B9').grid(column=0, row=0, sticky='W', columnspan=2)
    
        #Create Rules File
        subcats = df.groupby(['Category', 'Sub-Category'])['Stock'].apply(lambda x: (x > stock_threshold).sum())
        subcats = pd.DataFrame(subcats.rename('Available SKU Count'))
        exclude_subcats = subcats[subcats['Available SKU Count']==0].index.get_level_values(1).tolist()
        rules_df = pd.concat([subcats, pd.DataFrame(columns=['Slot 1', 'Slot 2', 'Slot 3', 'Slot 4', 'Slot 5',
                    'Slot 6', 'Slot 7', 'Slot 8', 'Slot 9', 'Slot 10'])])
        

        def rules():
            # Opens Excel application
            filename_r = 'Rules_' + str(brand) + '.xlsx'
            rulespath = filedialog.asksaveasfile(mode='w', initialfile=filename_r)
            writer = pd.ExcelWriter(rulespath.name, engine='xlsxwriter')
            rules_df.to_excel(writer, sheet_name='Rules')
            workbook = writer.book
            worksheet = writer.sheets['Rules']
            worksheet.set_column('A:A', 20)
            worksheet.set_column('B:B', 25)
            worksheet.set_column('D:M', 25)
            worksheet.freeze_panes(1,3)
            worksheet.data_validation('D2:M'+str(rules_df.shape[0]+1), {'validate': 'list', 'source': '$B$2:$B$100'})
            worksheet.write('C1', 'Available SKU Count', workbook.add_format({'text_wrap': True, 'bold': True}))
            format1 = workbook.add_format({'bg_color': '#FFC7CE'})
            for i in exclude_subcats:
                worksheet.conditional_format('D2:M'+str(rules_df.shape[0]+1), {'type': 'text',
                                                                'criteria': 'containing',
                                                                'value': i,
                                                                'format': format1})
            writer.save()
            # Opens Rules file
            open_file(writer)
            ### Selecting the filled out Rules file
            label = tk.Label(frame_2, text='Select the filled out Rules file:')
            label.grid(column=0, row=3, sticky='W')
            self.entryrules = tk.Entry(frame_2, text='', width="40")
            self.entryrules.grid(column=1, row=3, columnspan=3, sticky='W')

            def sel_rules():
                browse('Select the filled out Rules file', self.entryrules)

            btn = tk.Button(frame_2, text='Browse', command=sel_rules, width=15)
            btn.grid(column=4, row=3, sticky='W', columnspan=2) 

            text = 'Click Run to proceed with the cross-sell selection using filled out Rules file.'
            tk.Label(frame_2, text=text).grid(column=0, row=4, columnspan=2)   
            btn = tk.Button(frame_2, text="Run", command=cross_sell, width=15)
            btn.grid(column=4, row=4, sticky='W', columnspan=2)   

        def cross_sell():
            #FRAME 3
            frame_2.grid_forget()
            self.frame_3 = tk.Frame(window, padx=10, pady=10)
            self.frame_3.grid(column=0, row=0)
            def reco(i, for_):
                output = for_.copy()
                if i in for_.index: output = output.drop(i)
                i_sub = df.loc[i, 'Sub-Category']
                output.loc[:, 'Score'] = 1
                collection = df.loc[i, 'Collection']
                output.loc[output['Collection']==collection, 'Score'] += 1
                season = df.loc[i, 'Season']
                output.loc[output['Season'].apply(lambda x: x*season).apply(lambda x: 1 in x), 'Score'] += 1
                realism = df.loc[i, 'Realism']
                output.loc[output['Realism']==collection, 'Score'] += 1
                light = df.loc[i, 'Light Type']
                output.loc[output['Light Type']==collection, 'Score'] += 1
                p = [] 
                Rules = pd.read_excel(self.entryrules.get(), sheet_name='Rules')
                Rules.drop(columns=['Available SKU Count', 'Category'], inplace=True) 
                Rules.set_index(['Sub-Category'], inplace=True)
                combs = []
                for i in list(subcats.loc['Trees'].index):
                    combs.append(set([i, 'Tree Skirts']))
                for sub in Rules.loc[i_sub]:
                    pool = output.loc[output['Sub-Category']==sub]
                    pool_index = list(pool.index)
                    if pool_index: # if pool is not empty
                        s = np.random.choice(pool_index, p=pool['Score']/pool['Score'].sum()) # randomly selects from remaining pool
                        pool = pool.drop(s)
                        pool_index = list(pool.index)
                        while (s in p) & (len(pool_index) > 0) :  # repeats the loop if the random choice is already in the list to avoid duplicates                      
                            s = np.random.choice(pool_index, p=pool['Score']/pool['Score'].sum()) #should be updated based on remaining index
                            pool.drop(s, inplace=True) 
                            pool_index = list(pool.index)
                        if s not in p:
                            p.append(s)
                        else:
                            continue
                return p
            sc_recommendations = {i : reco(i, for_sc) for i in sc.index}
            pdp_recommendations = {i : reco(i, for_pdp) for i in pdp.index}
            sc_recs = pd.DataFrame.from_dict(sc_recommendations, orient='index')
            sc_recs.columns += 1
            sc_recs['Place'] = 'Soft Cart'
            pdp_recs = pd.DataFrame.from_dict(pdp_recommendations, orient='index')
            pdp_recs.columns += 1
            pdp_recs['Place'] = 'Product Page'
            self.results = pd.concat([sc_recs, pdp_recs])
            
            label = tk.Label(self.frame_3, text='Please download the initial output file.')
            label.grid(column=0, row=1, columnspan=4, sticky='W')  
            btn = tk.Button(self.frame_3, text="Download file", command=download, width=15)
            btn.grid(column=4, row=1, sticky='W', columnspan=2) 

        
        def download():   
            """ Downloads the initial cross-sell output """
            results_code = self.results.rename(mapper=i_to_code).replace(i_to_code)
            results_name = self.results.rename(mapper=i_to_name).replace(i_to_name)
            # Override Sheet
            override = pd.DataFrame(columns=['Product Code', 'Place', 'Slot', 'Current', 'Replace With', 'Duplicates', 
                                            'Invalid Entity', 'Out of Stock'])
            sku = df[['Product Code', 'Product Name', 'Entity', 'Category', 'Sub-Category', 'Stock']]
            #Excel Writer
            filename_cs = 'Cross-Sells_' + str(brand) + '.xlsx'
            path = filedialog.asksaveasfile(mode='w', initialfile=filename_cs)
            writer = pd.ExcelWriter(path.name, engine='xlsxwriter')
            results_code.to_excel(writer, sheet_name='Product Code')
            results_name.to_excel(writer, sheet_name='Product Name')
            override.to_excel(writer, sheet_name='Override', index=False)
            sku.to_excel(writer, sheet_name='Master', index=False)
            workbook = writer.book
            ws_name = writer.sheets['Product Name']
            ws_name.set_column('A:K', 25)
            ws_master = writer.sheets['Master']
            ws_master.set_column('B:B', 25)
            ws_master.set_column('E:E', 20)
            worksheet = writer.sheets['Override']
            center = workbook.add_format({'align': 'center'})
            worksheet.set_column('A:A', 20)
            worksheet.set_column('D:E', 20)
            worksheet.set_column('F:H', 15, center)
            worksheet.freeze_panes(1,0)
            worksheet.data_validation('B2:B1000', {'validate': 'list', 'source': ['Soft Cart', 'PDP']})
            worksheet.data_validation('C2:C1000', {'validate': 'integer',
                                                'criteria': 'between',
                                                'minimum': 1,
                                                'maximum': 10,
                                                'error_title': 'Input value is not valid.',
                                                'error_message': 'Enter a value from 1 to 10.'})
            t = self.results.shape[0]+1
            for i in range(2,100):
                formula1 = "{=INDEX('Product Code'!$B$2:$K$%d,MATCH(A%d&B%d,'Product Code'!$A$2:$A$%d&'Product Code'!$L$2:$L$%d,0), C%d)}" % (t, i, i, t, t, i)
                location1 = 'D%d' % i
                worksheet.write_array_formula(location1, formula1)
                formula2 = "{=COUNTIF(INDEX('Product Code'!$B$2:$K$%d,MATCH(A%d&B%d,'Product Code'!$A$2:$A$%d&'Product Code'!$L$2:$L$%d,0),),E%d)>0}" % (t, i, i, t, t, i)
                location2 = 'F%d' % i
                worksheet.write_array_formula(location2, formula2)
                formula3 = '=IF(B2="PDP", IF(VLOOKUP(E%d,Master!A2:F2968,3, FALSE)="CHILDREN", "Invalid Entity",""), IF(VLOOKUP(E%d,Master!A2:F2968,3, FALSE)="PARENT", "TRUE","FALSE"))' % (i, i)
                location3 = 'G%d' % i
                worksheet.write_formula(location3, formula3)
                formula4 = '=IF(VLOOKUP(E%d,Master!A2:F2968,6, FALSE)=0, "TRUE", "FALSE")' % i
                location4 = 'H%d' % i
                worksheet.write_formula(location4, formula4)

            redfill = workbook.add_format({'bg_color': '#FFC7CE'})
            worksheet.conditional_format('F2:H%d' %t , {'type': 'text',
                                                        'criteria': 'containing',
                                                        'value': 'TRUE',
                                                        'format': redfill})
            writer.save()
            # Opens Excel application
            open_file(writer)     
            # Override
            # QA and Output Process
            label = tk.Label(self.frame_3, text='Please manually update the Override sheet of the output file if necessary.')
            label.grid(column=0, row=2, sticky='W', columnspan=3)
            ### Selecting the filled out cross-sells file
            label = tk.Label(self.frame_3, text='Select the filled out Cross-Sells file:')
            label.grid(column=0, row=3, sticky='W')
            self.entrycross = tk.Entry(self.frame_3, text='', width="40")
            self.entrycross.grid(column=1, row=3, columnspan=3, sticky='W')

            def sel_cross():
                browse('Select the filled out Cross-Sells file', self.entrycross)

            btn = tk.Button(self.frame_3, text='Browse', command=sel_cross, width=15)
            btn.grid(column=4, row=3, sticky='W', columnspan=2) 
            
            # label = tk.Label(window, text='Click Override once ready.')
            # label.grid(column=0, row=4, sticky='W', columnspan=3)
            def override():
                #FRAME 4
                self.frame_3.grid_forget()
                self.frame_4 = tk.Frame(window, padx=10, pady=10)
                self.frame_4.grid(column=0, row=0)
                overrides = pd.read_excel(path.name, sheet_name='Override')
                overrides.dropna(axis=0, inplace=True)
                overrides.iloc[:,:5] = overrides.iloc[:,:5].astype(dtype = int, errors = 'ignore')
                overrode = results_code.set_index([results_code.index, 'Place']).sort_index()
                for i in range(len(overrides)):
                    baseproduct = overrides['Product Code'][i].astype(str)
                    place = overrides['Place'][i]
                    slot = overrides['Slot'][i]
                    new = overrides['Replace With'][i]
                    overrode.loc[(baseproduct, place), slot] = new
                # Convert to index
                code_to_i = {v: k for k, v in i_to_code.items()}
                self.overrode = overrode.rename(mapper=code_to_i).replace(code_to_i)
                label = tk.Label(self.frame_4, text='Override successful.', fg='#00B2B9')
                label.grid(column=5, row=3, sticky='W', columnspan=2) 
                label = tk.Label(self.frame_4, text='Download output file')
                label.grid(column=0, row=4, sticky='W')             
                btn = tk.Button(self.frame_4, text="Download file", command=PIM, width=15)
                btn.grid(column=4, row=4, sticky='W')
            btn = tk.Button(self.frame_3, text="Submit", command=override, width=15)
            btn.grid(column=4, row=12, sticky='W')
                
                
        def PIM():
            #convert final dataframe to PIM-ready upload
            now = datetime.datetime.now()
            file_name = brand + 'WebOutbound-AllEntities-StagingDelta_(' + now.strftime("%Y-%b-%dT%H.%M.%S.%f")[:-3] + ')_1of1.xlsx'
            #path = filedialog.asksaveasfile(mode='w', initialfile=file_name)
            columns = ['Id', 'Relationship External Id', 'Relationship Type', 'From Entity External Id', 'From Entity Type',
                   'From Entity Category Path', 'From Entity Container', 'From Entity Organization',
                   'To Entity External Id', 'To Entity Type', 'To Entity Category Path', 'To Entity Container',
                   'To Entity Organization', 'Cross-Sell Attributes//Cross-Sell Location',
                   'Cross-Sell Attributes//Sort Order']
            pim = pd.DataFrame(columns=columns)   
            final = self.overrode.copy()
            final = final.reset_index().set_index('level_0')
            final.index.names = [None]    
            results_pim = final.rename(mapper=i_to_pimid).replace(i_to_pimid)
            data = results_pim.reset_index()
            data = pd.melt(data, id_vars=['index', 'Place']).sort_values(by=['index','variable']).reset_index(drop=True)
            pim.iloc[:, pim.columns.get_loc('From Entity External Id')] = data.iloc[:,0]
            pim.iloc[:, pim.columns.get_loc('Cross-Sell Attributes//Sort Order')] = data.iloc[:,2]
            pim.iloc[:, pim.columns.get_loc('To Entity External Id')] = data.iloc[:,3]
            pim.iloc[:, pim.columns.get_loc('Cross-Sell Attributes//Cross-Sell Location')] = data.iloc[:,1]
            pim['Relationship Type'] = 'Cross-Sell Relationship'
            pim = pim[pim['To Entity External Id'].notna()]
            writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            pim.to_excel(writer, index=False)
            writer.save()
            tk.Label(self.frame_4, text=file_name, fg='#00B2B9').grid(row=5)
            tk.Label(self.frame_4, text='has been saved to your computer. You may now upload it directly to Hybris.', fg='#00B2B9').grid(row=6)
        
        text = 'Please download the Rules template and fill out.'
        tk.Label(frame_2, text=text).grid(column=0, row=2, columnspan=2, sticky='W')
        btn = tk.Button(frame_2, text="Download file", command=rules, width=15)
        btn.grid(column=4, row=2, sticky='W', columnspan=2)
 
window = tk.Tk()
window.title('Cross-Sell')
window.geometry('700x500')
app = App(window)
window.mainloop()
