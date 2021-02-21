
"""
19.05.19: This PDF exporting causes data to be mangled into each other if there is to much, thus
the data would need to be broken into chunks before exporting it via PDF
"""
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
import os

df = pd.read_excel('stock_data_v2.xlsx',sheet_name='Sheet1')


page_height = round(len(df.index),10) / 2
page_width = 8

dir_name = 'Data'

if not os.path.exists(dir_name):
    os.mkdir(dir_name)
    print("Directory " , dir_name ,  " Created ")
else:    
    print("Directory " , dir_name ,  " already exists")


with PdfPages(r'Data/Graph_1.pdf') as export_pdf:
    #Prints table on a graph style
    
    fig = plt.figure(figsize=(page_width, page_height)) #first argument is for cell width and 2nd for height cell 
    ax = plt.subplot(111)
    ax.axis('off')
    ax.table(cellText=df.values, colLabels=df.columns, bbox=[0,0,1,1])
    plt.title('Voltage Data')
    export_pdf.savefig()
#     plt.show()
    plt.close()
     
    plt.rc('text', usetex=False)    #creating new page
    x=df['Date']
    y=df['Voltage']
    z=df['Time']
    fig, ax=plt.subplots()
    ax2=ax.twiny()
    
    ax.plot(x,y,color='r')
    ax2.plot(z,y,'bo')
    ax.set_ylabel('Voltage [v]')
    ax.set_xlabel('Date')
    ax2.set_xlabel('Time')
    
    plt.grid(True)
#     plt.title('Voltage Graph')
    export_pdf.savefig()
#     plt.show()
    plt.close()
    
    
print("Done Exporting!")  
    
    
    
    
    