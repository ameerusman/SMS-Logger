
"""
19.05.19: This PDF exporting causes data to be mangled into each other if there is to much, thus
the data would need to be broken into chunks before exporting it via PDF
"""
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
import os
import tkinter 
from tkinter import ttk

def graph_print():	#Gets Date from GUI and prints Data PDF on Data Folder
    date_start_formatted = str(date_start.get())
    date_end_formatted = str(date_end.get())
    
    df_3 = pd.read_excel('stock_data_v2.xlsx',sheet_name='Sheet1')
    df_2 = df_3.set_index("Date",drop=False)
#     df = (df_2.loc [ str(date_start.get()) : str(date_end.get()) ])
    df = (df_2.loc [ date_start_formatted : date_end_formatted ])
    

    
    
    dir_name = 'Data'  
    page_height = round(len(df.index),10) / 2
    page_width = 8
    
    if (page_height < 7): 
        page_height = 7
      
    
    
    
    if not os.path.exists(dir_name):
        os.mkdir(dir_name)
        print("Directory " , dir_name ,  " Created ")
    else:    
        print("Directory " , dir_name ,  " already exists")
    
    
#     with PdfPages(r'Data/' + date_start_formatted + ' - '+ date_start_formatted +'.pdf') as export_pdf:
    with PdfPages(r'Data/Voltage_Data.pdf') as export_pdf:
        #Prints table on a graph style
        
        fig = plt.figure(figsize=(page_width, page_height)) #first argument is for cell width and 2nd for height cell 
        ax = plt.subplot(111)
        ax.axis('off')
        ax.table(cellText=df.values, colLabels=df.columns, bbox=[0,0,1,1])
        plt.title('Voltage Data from '+ date_start_formatted + ' - ' + date_end_formatted) 
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
        
#         plt.grid(True)
    #     plt.title('Voltage Graph')
        export_pdf.savefig()
    #     plt.show()
        plt.close()
        
        
    print("Done Exporting!")  
    
def gui_input():    #GUI for Date input
    
    window = tkinter.Tk()
    window.title('Graph Input')
    # window.geometry("250x150+300+300")
    heading_1 = ttk.Label(window, text='Enter your range which you want to plot in format yyyy/mm/d', font = ("Helvetica", 12) )
    heading_1.grid(row=0, column=0)
    
    heading_2 = ttk.Label(window, text='For example 6 May 2019 will be 2019/05/6', font = ("Helvetica", 8), justify = 'left' )
    heading_2.grid(row=1, column=0)
    
    global date_start
    global date_end
    
    date_start = tkinter.StringVar()
    date_end = tkinter.StringVar()
    
    user_input_1 = tkinter.Entry(window, width = 26, textvariable = date_start)
    user_input_1.grid(row = 2, column = 0)
    
    user_input_2 = tkinter.Entry(window, width = 26, textvariable = date_end)
    user_input_2.grid(row = 3, column = 0)
    
    button_submit = ttk.Button(window, text= 'Enter', command = graph_print)
    button_submit.grid(row = 4, column = 1)    

    window.mainloop()
    

if __name__ == '__main__':
    gui_input()      
    
    
    