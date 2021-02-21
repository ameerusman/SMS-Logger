'''
Created on 06-May-2019

@author: Ameer Usman
'''

#26.9.19:    Modified it into an SIM900A sms monitor using the serial port
#            Includes AT commands and sms deletion

import pandas as pd
# import datetime
import serial
import sys
import os.path

date_start_global = ""
date_end_global = ""

def arduino_serial(): 
    port_name = input("Please enter Arduino COM Port: ")    #Input
    
    data_current = "x"  #the serial data is stored here
    
    try:    #this is for opening COM port
        ser = serial.Serial(
#             timeout = 10,           #timeout in seconds
            port='COM' + port_name  #variable COM port
    #         port='COM6'
            )
            
    except serial.SerialException:  #serial exception    
        print ('---COM Port already being used or invalid. Please release COM port and open again---')
        sys.exit()
        
    while True:
        try:    #for keyboard exception   
            while True:     #keep looping
                print ("Waiting for SMS.....")
                sms_data_1 = ser.readline()
                sms_data_2 = ser.readline()
                
                #This is for handling timeout
#                 try:
#                     data_current = ser.readline()
#                 except ser.SerialTimeoutException:
#                     print('Data could not be read')

                print ("Got an SMS!: ",sms_data_2)
                sms_decoded_data = sms_data_2.decode(encoding='utf-8')    #removing the <b'> thing
            
                
                
                if (sms_decoded_data[0] == '+'):
                    print ("Found valid String.")
                    break
             
    #         print ("---Closing COM Port---")
    #         ser.close()
            ser.write(b'AT+CMGR=1\r\n')
            
            sms_data_1 = ser.readline()
            sms_data_2 = ser.readline()
            sms_data_3 = ser.readline()
            sms_data_4 = ser.readline()
            
            print(sms_data_3)    #For Viewing 
            print(sms_data_4)    #For Viewing
            
            sms_string = sms_data_4.decode(encoding='utf-8')
            sms_string = sms_string.replace('$','')
            sms_string = sms_string.replace('\n','')
            ser.write(b'AT+CMGD=1\r\n')
            
            temp = ser.readline()
            temp = ser.readline()
            temp = ser.readline()
            temp = ser.readline()
#             print ("data_string: ",data_string)
            
            data_list = sms_string.split(",")
#             print ("data_list: ", data_list)
            
#             print (data_current)
            
            df_data = pd.DataFrame({'Terminal ID':['{}'.format(data_list[0])],  #column heading, insert this,
                                    'Date':['{}'.format(data_list[1])],         #format is an inserting method
                                    'Time':['{}'.format(data_list[2])],
                                    'Voltage':['{}'.format(data_list[3])] })
            
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




if __name__ == '__main__':   
    print ("Initiated Voltage Recorder")
    arduino_serial()

