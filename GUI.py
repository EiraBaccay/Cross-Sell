### GUI (Graphical User Interface) of the Cross-Sell Tool

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import numpy as np
import random

window = tk.Tk()
window.title('Cross-Sell')
window.geometry('700x400')
label = tk.Label(window, text='Welcome to the cross-sell application. Please upload the following files to proceed.')
label.grid(column=0, row=0, columnspan=3, sticky='W')

## Uploading the required files
# SKU Master
label = tk.Label(window, text='Select the SKU Master file:')
label.grid(column=0, row=1, sticky='W')
entry = tk.Entry(window, text='', width="40")
entry.grid(column=1, row=1, columnspan=3, sticky='W')

def sel_sku():
    path = filedialog.askopenfilename(title='Select the SKU Master file')
    entry.delete(0, tk.END)
    entry.insert(0, path)
    global sku
    sku = path

button_widget = tk.Button(window, text='Browse', command=sel_sku, width=10)
button_widget.grid(column=4, row=1, sticky='W')

# Inventory
label = tk.Label(window, text='Select the Inventory file:')
label.grid(column=0, row=2, sticky='W')
entry2 = tk.Entry(window, text='', width="40")
entry2.grid(column=1, row=2, columnspan=3)

def sel_inv():
    path = filedialog.askopenfilename(title='Select the inventory file')
    entry2.delete(0, tk.END)
    entry2.insert(0, path)
    global inv
    inv = path

button_widget = tk.Button(window, text='Browse', command=sel_inv, width=10)
button_widget.grid(column=4, row=2, sticky='W')


import os, sys, subprocess
def open_file(filename):
    ### Opens csv file in default application
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener ="open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])

def download():   
    # Downloads the initial cross-sell output
    global path
    path = filedialog.asksaveasfile(mode='w', initialfile="cross-sell.xlsx")
    writer = pd.ExcelWriter(path.name)
    results_code.to_excel(writer, 'SKU')
    results_name.to_excel(writer, 'Product Name')
    writer.save()
    #initial.to_csv(output.name)   
    tk.Label(window, text='File saved.').grid(column=4, row=8, sticky='W')
    ##Opens Excel application
    open_file(writer)     
    #Need override?
    tk.Label(window, text='Are you satisfied?').grid(column=0, row=10, sticky='W', columnspan=3)
    #QA and Output Process
    var = tk.IntVar()
    #var.set(1) #default selects Yes
    #sel(var)
    rad1 = tk.Radiobutton(window, text='Yes', variable=var, value=1, command=lambda: sel(var))
    rad1.grid(column=4, row=10, sticky='W')
    rad2 = tk.Radiobutton(window, text='No', variable=var, value=0, command=lambda: sel(var))
    rad2.grid(column=5, row=10, sticky='W')  

    
def PIM():
    #convert final dataframe to PIM-ready upload
    brand = 'BHUS'
    import datetime
    now = datetime.datetime.now()
    file_name = brand + 'WebOutbound-AllEntities-StagingDelta_(' + now.strftime("%Y-%b-%dT%H.%M.%S.%f")[:-3] + ')_1of1.csv'
    final_frame.grid(column=0, row=13, columnspan=5, sticky='WE', pady=10)
    final_frame.grid_columnconfigure(0, weight=1)
    path = filedialog.asksaveasfile(mode='w', initialfile=file_name)
    columns = ['Id', 'Relationship External Id', 'Relationship Type', 'From Entity External Id', 'From Entity Type',
           'From Entity Category Path', 'From Entity Container', 'From Entity Organization',
           'To Entity External Id', 'To Entity Type', 'To Entity Category Path', 'To Entity Container',
           'To Entity Organization', 'Cross-Sell Attributes//Cross-Sell Location',
           'Cross-Sell Attributes//Sort Order']
    pim = pd.DataFrame(columns=columns)   
    if overrode is not None:
        final = overrode.copy()      
    else:
        final = results.copy()      
    results_pim = final.rename(mapper=i_to_pimid).replace(i_to_pimid)
    data = pd.melt(results_pim.reset_index(), id_vars='index').sort_values(by=['index','variable']).reset_index(drop=True)
    pim.iloc[:, pim.columns.get_loc('From Entity External Id')] = data.iloc[:,0]
    pim.iloc[:, pim.columns.get_loc('Cross-Sell Attributes//Sort Order')] = data.iloc[:,1]
    pim.iloc[:, pim.columns.get_loc('To Entity External Id')] = data.iloc[:,2]
    pim['Relationship Type'] = 'Cross-Sell Relationship'
    pim['Cross-Sell Attributes//Cross-Sell Location'] = 'Soft Cart'
    pim.to_csv(path, index=False)
    tk.Label(final_frame, text=file_name).grid(row=0)
    tk.Label(final_frame, text='has been saved to your computer.').grid(row=1)
    tk.Label(final_frame, text='You may now upload it directly to Hybris.').grid(row=2)
    
def override():
    global overrode
    overrode = pd.read_excel(path.name, sheet_name='SKU', index_col=0, header=0)
    code_to_i =  {y:x for x,y in i_to_code.items()}
    overrode = overrode.rename(mapper=code_to_i).replace(code_to_i)
    label = tk.Label(override_frame, text='Override successful.')
    label.grid(column=0, row=2, sticky='W') 
    label = tk.Label(override_frame, text='Download output file')
    label.grid(column=0, row=3, sticky='W')             
    btn = tk.Button(override_frame, text="Download file", command=PIM, width=10)
    btn.grid(column=5, row=3, sticky='W')

def sel(var):
    #Radio Button Selection    
    if var.get() == 0:
        download_frame.grid_forget()
        final_frame.grid_forget()
        override_frame.grid(column=0, row=11, columnspan=5, sticky='WE')
        override_frame.grid_columnconfigure(0, weight=1)
        label = tk.Label(override_frame, text='Please manually override the SKU sheet of cross-sell.xlsx and save.')
        label.grid(column=0, row=0, sticky='W', columnspan=3)
        label = tk.Label(override_frame, text='Click Override once ready.')
        label.grid(column=0, row=1, sticky='W', columnspan=3)
        btn = tk.Button(override_frame, text="Override", command=override, width=10)
        btn.grid(column=5, row=1, sticky='W')
    else:
        override_frame.grid_forget()
        final_frame.grid_forget()
        download_frame.grid(column=0, row=11, columnspan=5, sticky='WE')
        download_frame.grid_columnconfigure(0, weight=1)
        label = tk.Label(download_frame, text='Download final file for upload.')
        label.grid(column=0, row=0, sticky='W')            
        btn = tk.Button(download_frame, text="Download file", command=PIM, width=10)
        btn.grid(column=5, row=0, sticky='W')

def clicked():
    one = pd.read_excel(sku, sheet_name='Master', header=1)
    one = one.dropna(axis=1, how='all')
    one = one.dropna(axis=0, how='all')
    two = pd.read_csv(inv, header=2)
    df = pd.merge(one, two, how='left', left_on='Product Code', right_on='SKU')
    df['Stock'] = df['Stock'].replace(np.nan, 0)
    df = df.drop(columns=['Drop-Shipped?', 'Include in Pricing?', 'Included in OLD File', 'Brand', 'SKU'])
    global i_to_code
    i_to_code = df['Product Code'].to_dict()
    i_to_name = df['Product Name'].to_dict()
    global i_to_pimid
    i_to_pimid = df['PIMID'].to_dict()
    for i in df.index: # Counting stock staus for PARENT entities
        if df.iloc[i]['Entity'] == 'PARENT':
            code = df.iloc[i]['Product Code']
            df.iloc[i, df.columns.get_loc('Stock')] += df[df['Parent Code']==code]['Stock'].sum()
        else:
            continue
    #FILTER
    stock_threshold = 5
    df_filtered = df[(df['Status']!='REMOVED') & (df['Stock'] >= stock_threshold) & (df['Category'].notnull())]
    text = 'Out of %d products, %d will be used for recommendations.' %(df.shape[0], df_filtered.shape[0])
    tk.Label(window, text=text).grid(column=0, row=5, columnspan=2) # Show number of products eligible for cross-sells
    #SC and PDP
    sc = df[df['Entity']!='PARENT']
    for_sc = df_filtered[df_filtered['Entity']!='PARENT']
    pdp = df[df['Entity']!='CHILDREN']
    for_pdp = df_filtered[df_filtered['Entity']!='CHILDREN']
    # Placeholder Cross-Sell
    def reco(i, for_sc):
        output = for_sc
        if i in for_sc.index: output = for_sc.drop(i)
        p = random.sample(list(output.index), k=10)
        return p
    sc_recommendations = {i : reco(i, for_sc) for i in sc.index}
    global results
    results = pd.DataFrame.from_dict(sc_recommendations, orient='index')
    results.columns += 1
    global results_code
    results_code = results.rename(mapper=i_to_code).replace(i_to_code)
    global results_name
    results_name = results.rename(mapper=i_to_name).replace(i_to_name)
    label = tk.Label(window, text='Please download the output file and don\'t change the file name.')
    label.grid(column=0, row=7, columnspan=4, sticky='W')   
    btn = tk.Button(window, text="Download file", command=download, width=10)
    btn.grid(column=4, row=7, sticky='W')

override_frame = tk.Frame(window)
download_frame = tk.Frame(window)
final_frame = tk.Frame(window)
btn = tk.Button(window, text="Upload", command=clicked, width=10)
btn.grid(column=4, row=3, sticky='W')


window.mainloop()






