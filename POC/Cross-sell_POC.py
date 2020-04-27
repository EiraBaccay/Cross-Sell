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
    ### Opens file in default application
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener ="open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])
## https://stackoverflow.com/questions/17317219/is-there-an-platform-independent-equivalent-of-os-startfile
        
class App:
    def __init__(self, window):
        text='Welcome to the Cross-Sell application. Please upload the following files to proceed.'
        label = tk.Label(window, text=text, font=(None, 15))
        label.grid(column=0, row=0, columnspan=6, sticky='W')

        #### Uploading the required files
        def browse(title, entry):
            path = filedialog.askopenfilename(title=title)
            entry.delete(0, tk.END)
            entry.insert(0, path)
        
        ### SKU Master
        label = tk.Label(window, text='Select the SKU Master file:')
        label.grid(column=0, row=1, sticky='W')
        self.entry1 = tk.Entry(window, text='', width="40")
        self.entry1.grid(column=1, row=1, columnspan=3, sticky='W')

        def sel_sku():
            browse('Select the SKU Master file', self.entry1)

        btn = tk.Button(window, text='Browse', command=sel_sku, width=15)
        btn.grid(column=4, row=1, sticky='W', columnspan=2)        
        
        ### Inventory
        label = tk.Label(window, text='Select the Inventory file:')
        label.grid(column=0, row=2, sticky='W')
        self.entry2 = tk.Entry(window, text='', width="40")
        self.entry2.grid(column=1, row=2, columnspan=3)

        def sel_inv():
            browse('Select the Inventory file', self.entry2)
     
        btn = tk.Button(window, text='Browse', command=sel_inv, width=15)
        btn.grid(column=4, row=2, sticky='W', columnspan=2)

        ### Disallow
        label = tk.Label(window, text='Select the Disallow file:')
        label.grid(column=0, row=3, sticky='W')
        self.entry3 = tk.Entry(window, text='', width="40")
        self.entry3.grid(column=1, row=3, columnspan=3)

        def sel_dis():
            browse('Select the Disallow file', self.entry3)
     
        btn = tk.Button(window, text='Browse', command=sel_dis, width=15)
        btn.grid(column=4, row=3, sticky='W', columnspan=2)

        ### Set stock threshold
        label = tk.Label(window, text='Select the stock threshold:')
        label.grid(column=0, row=4,sticky='W')

        self.stock_threshold = tk.IntVar()
        self.entry4 = tk.Spinbox(window, text=self.stock_threshold, from_=1, to=5, width='10')
        self.entry4.grid(column=1, row=4, columnspan=3, sticky='W')
        self.stock_threshold.set(5)

        #### Submit button
        btn = tk.Button(window, text="Submit", command=self.clicked, width=15)
        btn.grid(column=4, row=5, sticky='W', columnspan=2)


    def clicked(self):
        sku = self.entry1.get()
        sku = pd.read_excel(sku, sheet_name='Master', header=1)
        sku = sku.dropna(axis=1, how='all')
        sku = sku.dropna(axis=0, how='all')
        inv = self.entry2.get()
        inv = pd.read_csv(inv, header=2)
        brand = str(inv.Brand[0])
        if brand == 'BAL': brand = 'BHUS'
        disallow = self.entry3.get()
        disallow = pd.read_excel(disallow, header=0, usecols=[0])
        df = pd.merge(sku, inv, how='left', left_on='Product Code', right_on='SKU')
        df['Stock'] = df['Stock'].replace(np.nan, 0)
        df = df.drop(columns=['Drop-Shipped?', 'Include in Pricing?', 'Included in OLD File', 'Brand', 'SKU'])
        # For mapping purposes
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
        # Filter products to be processed
        df = df[(df['Status']!='REMOVED') & (df['Category'].notnull())]
        # Filter products to be used as cross-sell
        stock_threshold = self.stock_threshold.get()
        df_filtered = df[df['Stock'] >= stock_threshold]
        df_filtered = df_filtered[~df_filtered['Product Code'].isin(disallow['Product Code'])]
        # Display number of products eligible for cross-sells
        text = 'Out of %d products, %d will be used for recommendations.' %(df.shape[0], df_filtered.shape[0])
        tk.Label(window, text=text, fg='#00B2B9').grid(column=0, row=6, columnspan=2)
        #SC and PDP
        sc = df[df['Entity']!='PARENT']
        for_sc = df_filtered[df_filtered['Entity']!='PARENT']
        pdp = df[df['Entity']!='CHILDREN']
        for_pdp = df_filtered[df_filtered['Entity']!='CHILDREN']
        # Placeholder Cross-Sell
        def reco(i, for_):
            output = for_
            if i in for_.index: output = for_.drop(i)
            p = random.sample(list(output.index), k=10)
            return p
        sc_recommendations = {i : reco(i, for_sc) for i in sc.index}
        pdp_recommendations = {i : reco(i, for_pdp) for i in pdp.index}
        sc_recs = pd.DataFrame.from_dict(sc_recommendations, orient='index')
        sc_recs.columns += 1
        sc_recs['Place'] = 'Soft Cart'
        pdp_recs = pd.DataFrame.from_dict(pdp_recommendations, orient='index')
        pdp_recs.columns += 1
        pdp_recs['Place'] = 'Product Page'
        results = pd.concat([sc_recs, pdp_recs])
        results_code = results.rename(mapper=i_to_code).replace(i_to_code)
        results_name = results.rename(mapper=i_to_name).replace(i_to_name)
        label = tk.Label(window, text='Please download the output file and don\'t change the file name.')
        label.grid(column=0, row=7, columnspan=4, sticky='W')   

        ### Frames
        override_frame = tk.Frame(window)
        download_frame = tk.Frame(window)
        final_frame = tk.Frame(window)
        
        def download():   
            # Downloads the initial cross-sell output
            path = filedialog.asksaveasfile(mode='w', initialfile='cross-sell.xlsx')
            writer = pd.ExcelWriter(path.name, engine='xlsxwriter')
            results_code.to_excel(writer, 'SKU')
            results_name.to_excel(writer, 'Product Name')
            worksheet = writer.sheets['Product Name']
            for idx, col in enumerate(results_name):  # Sets column width of the excel file
                series = results_name[col]
                worksheet.set_column(idx, idx, 30) 
            writer.save()
    
            tk.Label(window, text='File saved.', fg='#00B2B9').grid(column=4, row=8, sticky='W')
            # Opens Excel application
            open_file(writer)     
            # Override
            tk.Label(window, text='Would you like to override the results?').grid(column=0, row=10, sticky='W', columnspan=3)
            # QA and Output Process
            var = tk.IntVar()

            #Radio Button Selection
            def sel(var):         
                if var.get() == 1:
                    download_frame.grid_forget()
                    final_frame.grid_forget()
                    override_frame.grid(column=0, row=11, columnspan=6, sticky='WE')
                    override_frame.grid_columnconfigure(0, weight=1)
                    label = tk.Label(override_frame, text='Please manually override the SKU sheet of cross-sell.xlsx and save.')
                    label.grid(column=0, row=0, sticky='W', columnspan=3)
                    label = tk.Label(override_frame, text='Click Override once ready.')
                    label.grid(column=0, row=1, sticky='W', columnspan=3)
                    def override():
                        overrode = pd.read_excel(path.name, sheet_name='SKU', index_col=0, header=0)
                        code_to_i =  {y:x for x,y in i_to_code.items()}
                        self.overrode = overrode.rename(mapper=code_to_i).replace(code_to_i)
                        label = tk.Label(override_frame, text='Override successful.', fg='#00B2B9')
                        label.grid(column=5, row=2, sticky='W', columnspan=2) 
                        label = tk.Label(override_frame, text='Download output file')
                        label.grid(column=0, row=3, sticky='W')             
                        btn = tk.Button(override_frame, text="Download file", command=PIM, width=15)
                        btn.grid(column=5, row=3, sticky='W')
                    btn = tk.Button(override_frame, text="Override", command=override, width=15)
                    btn.grid(column=5, row=1, sticky='W')
                else:
                    override_frame.grid_forget()
                    final_frame.grid_forget()
                    download_frame.grid(column=0, row=11, columnspan=6, sticky='WE')
                    download_frame.grid_columnconfigure(0, weight=1)
                    label = tk.Label(download_frame, text='Download final file for upload.')
                    label.grid(column=0, row=0, sticky='W', columnspan=3)            
                    btn = tk.Button(download_frame, text="Download file", command=PIM, width=15)
                    btn.grid(column=5, row=0, sticky='W')
                
            rad1 = tk.Radiobutton(window, text='Yes', variable=var, value=1, command=lambda: sel(var))
            rad1.grid(column=4, row=10, sticky='W')
            rad2 = tk.Radiobutton(window, text='No', variable=var, value=0, command=lambda: sel(var))
            rad2.grid(column=5, row=10, sticky='W')  

        def PIM():
            #convert final dataframe to PIM-ready upload
            now = datetime.datetime.now()
            file_name = brand + 'WebOutbound-AllEntities-StagingDelta_(' + now.strftime("%Y-%b-%dT%H.%M.%S.%f")[:-3] + ')_1of1.csv'
            final_frame.grid(column=0, row=13, columnspan=6, sticky='WE', pady=10)
            final_frame.grid_columnconfigure(0, weight=1)
            path = filedialog.asksaveasfile(mode='w', initialfile=file_name)
            columns = ['Id', 'Relationship External Id', 'Relationship Type', 'From Entity External Id', 'From Entity Type',
                   'From Entity Category Path', 'From Entity Container', 'From Entity Organization',
                   'To Entity External Id', 'To Entity Type', 'To Entity Category Path', 'To Entity Container',
                   'To Entity Organization', 'Cross-Sell Attributes//Cross-Sell Location',
                   'Cross-Sell Attributes//Sort Order']
            pim = pd.DataFrame(columns=columns)   
            try:
                final = self.overrode.copy()
            except:
                final = results.copy()      
            results_pim = final.rename(mapper=i_to_pimid).replace(i_to_pimid)
            data = pd.melt(results_pim.reset_index(), id_vars=['index', 'Place']).sort_values(by=['index','variable']).reset_index(drop=True)
            pim.iloc[:, pim.columns.get_loc('From Entity External Id')] = data.iloc[:,0]
            pim.iloc[:, pim.columns.get_loc('Cross-Sell Attributes//Sort Order')] = data.iloc[:,2]
            pim.iloc[:, pim.columns.get_loc('To Entity External Id')] = data.iloc[:,3]
            pim.iloc[:, pim.columns.get_loc('Cross-Sell Attributes//Cross-Sell Location')] = data.iloc[:,1]
            pim['Relationship Type'] = 'Cross-Sell Relationship'
            pim.to_csv(path, index=False)
            tk.Label(final_frame, text=file_name, fg='#00B2B9').grid(row=0)
            tk.Label(final_frame, text='has been saved to your computer. You may now upload it directly to Hybris.', fg='#00B2B9').grid(row=1)
            
        btn = tk.Button(window, text="Download file", command=download, width=15)
        btn.grid(column=4, row=7, sticky='W', columnspan=2)
 
window = tk.Tk()
window.title('Cross-Sell')
window.geometry('700x500')
app = App(window)
window.mainloop()
