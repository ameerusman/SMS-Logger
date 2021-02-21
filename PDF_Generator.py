
"""
19.05.19: This PDF exporting causes data to be mangled into each other if there is to much, thus
the data would need to be broken into chunks before exporting it via PDF

PDF_Generator_v2: 29/9/19
TO DO
- Needs exception handling of range if the date is not present then it gives an error
- The graph needs to be fixed

PDF_Generator_v3: 3/10/2019

- Making the graph size dynamic and variable to number of days

PDF_Generator_v5:
Added serial number and filtering by terminal ID
"""
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
import os
import tkinter 
from tkinter import ttk
import numpy as np

def graph_print():    #Gets Date from GUI and prints Data PDF on Data Folder
    date_start_formatted = str(date_start.get())
    date_end_formatted = str(date_end.get())
    terminal_id_formatted = str(terminal_id.get())
    
    #For testing Purpose only
#     terminal_id_formatted = '2'
#     date_start_formatted = "2019/01/5"
#     date_end_formatted = "2019/01/23"
    
    
    
    df_excel = pd.read_excel('stock_data_v2.xlsx',sheet_name='Sheet1')
    
#     df_3[date]
    df_date_index = df_excel.set_index("Date",drop=False)
    df_date_index_2 = df_excel.set_index("Date",drop=False)
    df_date_index['Date'] = pd.to_datetime(df_date_index['Date'])
    
    date_mask = (df_date_index['Date'] > date_start_formatted) & (df_date_index['Date'] <= date_end_formatted) 
    
    df_filtered = df_date_index_2.loc[date_mask] #creating a filter
    
    df_export = df_filtered.loc[df_filtered['Terminal ID'] == int(terminal_id_formatted) ]
    
    df_export.insert(loc=0, column='Sr. No.', value=np.arange(len(df_export)))

    
    
    dir_name = 'Data'  
    page_height = round(len(df_export.index),10) / 2 #round it by 10 and divide by 2
    page_width = 8
    
    if (page_height < 7): 
        page_height = 7
        
    if not os.path.exists(dir_name):
        os.mkdir(dir_name)
        print("Directory " , dir_name ,  " Created ")
    else:    
        print("Directory " , dir_name ,  " already exists")
    
        
#     path = 'Data\\Voltage' +    date_start.get() + ' - '+ date_end.get() + '.pdf'
#     path = 'Data//Voltage_Data.pdf'
#     path = os.path.join(r"Data/Voltage ",    format(date_start_formatted) + " - "+ format(date_start_formatted) + ".pdf")

#     path = r"Data/Voltage_{}_-_{}.pdf".format(date_start_formatted, date_start_formatted) 
#     path = "Data\\Voltage {} - {} .pdf".format(date_start.get(), date_end.get()) 
#     print (path)
#     path = "/Home"
#     os.path.join(path, "Data\\Voltage", date_start_formatted)
#     print (path)
    
    
    path = "Data\\Voltage Data.pdf"
    
    with PdfPages(path) as export_pdf:
#     with PdfPages(r'Data/Voltage_Data.pdf') as export_pdf:    #Working!
        #Prints table on a graph style
        
        fig = plt.figure(figsize=(page_width, page_height)) #first argument is for cell width and 2nd for height cell 
        ax = plt.subplot(111)
        ax.axis('off')
        ax.table(cellText=df_export.values, colLabels=df_export.columns, bbox=[0,0,1,1])
        plt.title('Voltage Data from '+ date_start_formatted + ' - ' + date_end_formatted) 
        export_pdf.savefig()
    #     plt.show()
        plt.close()
         
        plt.rc('text', usetex=False)    #creating new page
        x=df_export['Date']
        y=df_export['Voltage']
#         z=df['Time']    #this part is for labelling the top axis for time
        fig, ax=plt.subplots()
        
#         ax2=ax.twiny() #this part is for labelling the top axis for time
        
        ax.plot(x,y,color='r')
#         ax2.plot(z,y,'bo') #this part is for labelling the top axis for time
        ax.set_ylabel('Voltage [v]')
        ax.set_xlabel('Date')
        
        figure_size = plt.gcf()
        figure_size.set_size_inches(page_height,7)   #First argument is length and second is height
        
        plt.grid(color='black', linestyle='-.', linewidth=0.1)
        
        plt.xticks(rotation=45) #rotates the x-axis values

#         ax2.set_xlabel('Time')    #this part is for labelling the top axis for time
        
        
        
#         plt.grid(True)
    #     plt.title('Voltage Graph')
    
        export_pdf.savefig()
#         plt.show()
        plt.close()
        
        
    print("Done Exporting!")  
    
def gui_input():    #GUI for Date input
    
    window = tkinter.Tk()
    window.title('Graph Input')
    # window.geometry("250x150+300+300")
    heading_1 = ttk.Label(window, text='Enter your range which you want to plot in format yyyy/mm/dd', font = ("Helvetica", 12) )
    heading_1.grid(row=0, column=1)
    
    heading_2 = ttk.Label(window, text='For example 6 May 2019 will be 2019-05-06', font = ("Helvetica", 8), justify = 'left' )
    heading_2.grid(row=1, column=1)
    
    
    heading_3 = ttk.Label(window, text='Terminal ID: ', font = ("Helvetica", 8), justify = 'left' )
    heading_3.grid(row=2, column=0)
    
    heading_4 = ttk.Label(window, text='Start Date: ', font = ("Helvetica", 8), justify = 'left' )
    heading_4.grid(row=3, column=0)
    
    heading_5 = ttk.Label(window, text='End Date: ', font = ("Helvetica", 8), justify = 'left' )
    heading_5.grid(row=4, column=0)
    
    global date_start
    global date_end
    global terminal_id
    
    date_start = tkinter.StringVar()
    date_end = tkinter.StringVar()
    terminal_id = tkinter.StringVar()
    
    user_input_3 = tkinter.Entry(window, width = 15, textvariable = terminal_id)
    user_input_3.grid(row = 2, column = 1)
    
    user_input_1 = tkinter.Entry(window, width = 26, textvariable = date_start)
    user_input_1.grid(row = 3, column = 1)
    
    user_input_2 = tkinter.Entry(window, width = 26, textvariable = date_end)
    user_input_2.grid(row = 4, column = 1)
    
    button_submit = ttk.Button(window, text= 'Enter', command = graph_print)
    button_submit.grid(row = 5, column = 1)    

    window.mainloop()
    

if __name__ == '__main__':
    gui_input()      
    
    
    