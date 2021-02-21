'''
Created on 06-May-2019

@author: Ameer Usman
'''

#
#    Version 5: SMS Based. This is working but I believe it cannot cater universal messages
#    Version 6: Adding whole string check to the message reception (+CMTI: "SM")

import pandas as pd
import time
import serial
import sys
import os.path

DEBUG_MESSAGES = 0 #1 is Yes, 0 is No

date_start_global = ""
date_end_global = ""

def arduino_serial(): 
    port_name = input("Please enter Serial COM Port: ")    #Input
    
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
                if (ser.readline() == b'\r\n'):
                    sms_data_2 = ser.readline()
                    
                    #This is for handling timeout
    #                 try:
    #                     data_current = ser.readline()
    #                 except ser.SerialTimeoutException:
    #                     print('Data could not be read')
                    if (DEBUG_MESSAGES == 1):   
#                         print(sms_data_1)
                        print(sms_data_2)
                    
                    sms_decoded_data = sms_data_2.decode(encoding='utf-8')    #removing the <b'> thing
                    
                    if (DEBUG_MESSAGES == 1):
                        print(sms_decoded_data)
                        
                    sms_decoded_data_list = sms_decoded_data.split(",")
                    
                    if (sms_decoded_data_list[0] == '+CMTI: "SM"'):
                        print('Received an SMS')
                        time.sleep(1)
                        ser.flushInput()
                        
                        ser.write(b'AT+CMGR=') #reading sms at index 1
                        ser.write(sms_decoded_data_list[1].encode("ascii"))
                        ser.write(b'\r\n')
                        
                        sms_data_1 = ser.readline()
                        sms_data_2 = ser.readline()
                        sms_data_3 = ser.readline()
                        sms_data_4 = ser.readline()
                        sms_data_5 = ser.readline()
                        
                        if (DEBUG_MESSAGES == 1):
                            print(sms_data_1)
                            print(sms_data_2)
                            print(sms_data_3)    #For Viewing 
                            
                        
                        print(sms_data_4)    #For Viewing
                        print(sms_data_5)    #For Viewing
                        
                        ser.write(b'AT+CMGD=1,4\r\n')   #deleting all sms
                        
                        if (DEBUG_MESSAGES == 1):
                            print('Deleting all SMS')
                            
                        time.sleep(2)
                        ser.flushInput()
                        
                        sms_string = sms_data_5.decode(encoding='utf-8')
                        if(sms_string[0] == '$'):
                            print ("Verified as a Valid SMS")
                            break
                        else:
                            print('Invalid SMS')
                    
                
                
                    
                
                
             
    #         print ("---Closing COM Port---")
    #         ser.close()
            
            
            
#             sms_data_4 = sms_data_4.replace(27,'')
#             sms_data_4 = sms_data_4.replace('/n','')
            
                
            sms_string = sms_string.replace('$','')
            sms_string = sms_string.replace('\n','')
            sms_string = sms_string.replace('/n','')
            print(sms_string)
            
            ser.write(b'AT+CMGD=1,4\r\n')   #deleting all sms
            
            time.sleep(2)
            
            ser.flushInput()
#             ser.flushOutput()
#             ser.flush()

            
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

# def graph_print():
#         df = pd.read_excel('stock_data.xlsx',sheet_name='Sheet1')
#         df_2 = df.set_index("date",drop=False)
#         
#         # df_3=(df_2.loc['2018/05/2':'2018/05/6'])
#         
# #         df_3 = (df_2.loc[str(date_start_global.get()) : str(date_end_global.get())])
#         df_3 = (df_2.loc [ str(date_start.get()) : str(date_end.get()) ])
# 
#         
#         x=df_3['date']
#         y=df_3['voltage']
#         z=df_3['time']
#         fig,ax=plt.subplots()
#         ax2=ax.twiny()
#         
#         ax.plot(x,y,color='r')
#         ax2.plot(z,y,'bo')
#         ax.set_ylabel('Voltage [v]')
#         ax.set_xlabel('Date')
#         ax2.set_xlabel('Time')
#         
#         plt.title('Voltage Recorder')
#         plt.show    

# def gui_input():
#     
#     window = tkinter.Tk()
#     window.title('Graph Input')
#     # window.geometry("250x150+300+300")
#     heading = ttk.Label(window, text='Enter your range which you want to plot', font = 'red' )
#     heading.grid(row=0, column=0)
#     
#     global date_start
#     global date_end
#     
#     date_start = tkinter.StringVar()
#     date_end = tkinter.StringVar()
#     
#     user_input_1 = tkinter.Entry(window, width = 26, textvariable = date_start)
#     user_input_1.grid(row = 1, column = 0)
#     
#     user_input_2 = tkinter.Entry(window, width = 26, textvariable = date_end)
#     user_input_2.grid(row = 2, column = 0)
#     
#     button_submit = ttk.Button(window, text= 'Enter', command = graph_print)
#     button_submit.grid(row = 3, column = 1)    
# 
#     window.mainloop()
    


if __name__ == '__main__':   
    print ("Initiated SMS Monitor")
    arduino_serial()

