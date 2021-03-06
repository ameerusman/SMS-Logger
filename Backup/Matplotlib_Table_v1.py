
"""
19.05.19: This PDF exporting causes data to be mangled into each other if there is to much, thus
the data would need to be broken into chunks before exporting it via PDF
"""
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np

df = pd.read_excel('stock_data_v2.xlsx',sheet_name='Sheet1')
print ( "Total rows: ",len(df.index))

# df_split = np.array_split(df, df.rows)

df_split = {}

i = 0
df_rows = len(df.index)
df_page_size = 15

while True:
    df_split[i] = pd.DataFrame()
    df_split[i] = df.iloc[df_page_size*(i+1) - df_page_size : df_page_size*(i+1)]
    df_rows = df_rows - df_page_size
    i = i+1
    if (df_rows < 0):
        break
 


cell_height = len(df.index)
cell_width = 15

with PdfPages(r'Graph_1.pdf') as export_pdf:
    #Prints table on a graph style
    fig = plt.figure(figsize=(cell_width, cell_height)) #first argument is for cell width and 2nd for height cell 
    ax = plt.subplot(111)
    ax.axis('off')
    ax.table(cellText=df.values, colLabels=df.columns, bbox=[0,0,1,1])
    export_pdf.savefig()
#     plt.show()
    plt.close()
     
    plt.rc('text', usetex=False)
    fig = plt.figure(figsize=(cell_width, cell_height)) #first argument is for cell width and 2nd for height cell 
    ax = plt.subplot(111)
    ax.axis('off')
    ax.table(cellText=df_split[1].values, colLabels=df.columns, bbox=[0,0,1,1])
    export_pdf.savefig()
#     plt.show()
    plt.close()
    
    
print(df_split[1])  
    
    
    
    
    