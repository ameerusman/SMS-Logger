'''
Created on 06-May-2019

@author: Ameer Usman
'''
import pandas as pd
import time
import serial
import sys
import os.path
from threading import Thread
import matplotlib.pyplot as plt



def arduino_serial():   
    port_name = input("Please enter Arduino COM Port: ")    #Input
    
    data_current = "x"
    
    try:    #this is for opening COM port
        ser = serial.Serial(
            port='COM' + port_name  #variable COM port
    #         port='COM6'
            )
            
    except serial.SerialException:  #serial exception    
        print ('---COM Port already being used or invalid. Please release COM port and open again---')
        sys.exit()
        
    while True:
        try:    #for keyboard exception   
            while True:     #keep looping
                print ("Reading through Arduino")
                data_current = ser.readline()
                print ("Data from Arduino: ",data_current)
                
                if (data_current[0] == 36):
                    print ("Found Data!")
                    break
             
    #         print ("---Closing COM Port---")
    #         ser.close()
            
            data_string = data_current.decode(encoding='utf-8')
            data_string = data_string.replace('$','')
            data_string = data_string.replace('\n','')
            print ("data_string: ",data_string)
            
            data_list = data_string.split(",")
            print ("data_list: ", data_list)
            
            print (data_current)
            
            df_data = pd.DataFrame({'terminal_id':['{}'.format(data_list[0])],
                                    'date':['{}'.format(data_list[1])],
                                    'time':['{}'.format(data_list[2])],
                                    'voltage':['{}'.format(data_list[3])] })
            
    #         writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')
    #         df_data.to_excel(writer, index=False)
    #         df_data.to_excel(writer, sheet_name='Sheet1',startrow=len(df_data)+1, index=False, header=None)
    #         writer.save()
    
              #For trying out the exception rule  
    #         try:    
    #             df_1 = pd.read_excel('stock_data.xlsx', sheet_name='Sheet1')
    #               
    #             writer = pd.ExcelWriter('stock_data.xlsx', engine='xlsxwriter')
    #             df_1.to_excel(writer, index=False)
    #             df_data.to_excel(writer, sheet_name='Sheet1',startrow=len(df_1)+1, index=False, header=None)
    #             writer.save()
    # #         print(df_1)    #prints existing data
    #   
    #         except FileNotFoundError:
    #             print ("Writing a new File")
    #             writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')
    #             df_1.to_excel(writer, index=False)
    #             writer.save()
                
                
            
            file_exists = os.path.isfile("stock_data.xlsx") 
             
            if file_exists:
                # do something
                df_1 = pd.read_excel('stock_data.xlsx', sheet_name='Sheet1')
                  
                writer = pd.ExcelWriter('stock_data.xlsx', engine='xlsxwriter')
                df_1.to_excel(writer, index=False)
                df_data.to_excel(writer, sheet_name='Sheet1',startrow=len(df_1)+1, index=False, header=None)
                writer.save()
                
            else:
                # do something else
                print ("Writing a new File")
                writer = pd.ExcelWriter('stock_data.xlsx', engine='xlsxwriter')
                df_data.to_excel(writer, index=False)
                writer.save()
                
    #         writer = pd.ExcelWriter('stock_data.xlsx', engine='xlsxwriter')
    #         df_1.to_excel(writer, index=False)
    #         df_data.to_excel(writer, sheet_name='Sheet1',startrow=len(df_1)+1, index=False, header=None)
    #         writer.save()
            
            df_2=pd.read_excel('stock_data.xlsx', sheet_name='Sheet1')
    #         print(df_2)     #prints with new data
            
    #         total_rows=df_data.read_xlsx('pandas_simple.xlsx',)
    #         rows = df_data.book.sheet_by_index(0).nrows   
            
            
        except (KeyboardInterrupt):
            print ("Closing COM Port and Exiting") 
            ser.close()
            sys.exit()   

def gui_input():
    df = pd.read_excel('stock_data.xlsx',sheet_name='Sheet1')
    df_2 = df.set_index("date",drop=False)
    
    # df_3=(df_2.loc['2018/05/2':'2018/05/6'])
    
    df_3 = (df_2.loc['2018/05/2':'2018/05/6'])
    
    x=df_3['date']
    y=df_3['voltage']
    z=df_3['time']
    fig,ax=plt.subplots()
    ax2=ax.twiny()
    
    ax.plot(x,y,color='r')
    ax2.plot(z,y,'bo')
    ax.set_ylabel('Voltage [v]')
    ax.set_xlabel('Date')
    ax2.set_xlabel('Time')
    
    plt.title('Voltage Recorder')
    plt.show()

if __name__ == '__main__':   
    print ("Initiated Voltage Recorder")
    Thread(target = arduino_serial).start()
    Thread(target = gui_input).start()

