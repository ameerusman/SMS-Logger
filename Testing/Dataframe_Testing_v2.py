
"""
19.05.19: This PDF exporting causes data to be mangled into each other if there is to much, thus
the data would need to be broken into chunks before exporting it via PDF

PDF_Generator_v2: 29/9/19
TO DO
- Needs exception handling of range if the date is not present then it gives an error
- The graph needs to be fixed

PDF_Generator_v3: 3/10/2019

- Making the graph size dynamic and variable to number of days
"""
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
import os
import tkinter 
from tkinter import ttk

if __name__ == '__main__':  
    
    #For testing Purpose only
    date_start_formatted = "2019/01/5"
    date_end_formatted = "2019/01/23"
    
    df_3 = pd.read_excel('stock_data_v2.xlsx',sheet_name='Sheet1')
    df_2 = df_3.set_index("Date",drop=False)
    df = (df_2.loc [ date_start_formatted : date_end_formatted ])
    
    df_export = df.loc[df['Terminal ID'] == 1 ]
#     for value in df:
        
    
    print (df_export)
    
    
    
    
    