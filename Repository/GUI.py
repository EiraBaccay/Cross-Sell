### GUI (Graphical User Interface) of the Cross-Sell Tool

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd

window = tk.Tk()
window.title('Cross-Sell')
window.geometry('700x350')
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
    global output
    output = filedialog.asksaveasfile(mode='w', initialfile="cross-sell.csv")
    initial.to_csv(output.name)
    tk.Label(window, text='File saved.').grid(column=4, row=8, sticky='W')
    ##Opens Excel application
    open_file(output.name)     
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
    try:
        final.to_csv(path)      
    except NameError:
        initial.to_csv(path)
    tk.Label(final_frame, text=file_name).grid(row=0)
    tk.Label(final_frame, text='has been saved to your computer.').grid(row=1)
    tk.Label(final_frame, text='You may now upload it directly to Hybris.').grid(row=2)
    
def override():
    global final
    final = pd.read_csv(output.name)
    label = tk.Label(override_frame, text='Override successful.')
    label.grid(column=0, row=1, sticky='W') 
    label = tk.Label(override_frame, text='Download output file')
    label.grid(column=0, row=2, sticky='W')             
    btn = tk.Button(override_frame, text="Download file", command=PIM, width=10)
    btn.grid(column=5, row=2, sticky='W')

def sel(var):
    #Radio Button Selection    
    if var.get() == 0:
        download_frame.grid_forget()
        final_frame.grid_forget()
        override_frame.grid(column=0, row=11, columnspan=5, sticky='WE')
        override_frame.grid_columnconfigure(0, weight=1)
        label = tk.Label(override_frame, text='Please manually override cross-sell.csv, save and click Override once ready.')
        label.grid(column=0, row=0, sticky='W')
        btn = tk.Button(override_frame, text="Override", command=override, width=10)
        btn.grid(column=5, row=0, sticky='W')
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
    df = pd.read_csv(sku)
    df2 = pd.read_csv(inv)
    ##INSERT CROSS-SELL HERE
    m = df.shape[0]
    n = df2.shape[0]
    text = 'The SKU file has ' + str(m) + ' rows.'
    tk.Label(window, text=text).grid(column=0, row=5)
    text = 'The Stocks file has ' + str(n) + ' rows.'
    tk.Label(window, text=text).grid(column=0, row=6)
    global initial
    initial = df.append(df2)  
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






