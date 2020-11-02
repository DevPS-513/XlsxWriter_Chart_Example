
import numpy as np
import pandas as pd
import xlsxwriter


# Get our input data
AMZN_data=pd.read_excel('./Stock Prices.xlsx',sheet_name='AMZN')
print(AMZN_data.head())

# Initialize the excel output file
excel_file_path='./Output_Stock_Prices_Chart.xlsx'
workbook=xlsxwriter.Workbook(excel_file_path)
AMZN_worksheet=workbook.add_worksheet()
date_format=workbook.add_format({'num_format': 'dd/mm/yy'})


for i,col_name in enumerate(AMZN_data.columns):
    AMZN_worksheet.write(0,i,col_name)
    if(i==0):
        AMZN_worksheet.write_column(1,i,AMZN_data[col_name],date_format)
    else:
        AMZN_worksheet.write_column(1, i, AMZN_data[col_name])

        chart=workbook.add_chart({'type':'scatter','subtype':'straight'})
        col_letter=xlsxwriter.utility.xl_col_to_name(i)


        chart.add_series({'categories':'=Sheet1!$A$2:$A$'+str(2+len(AMZN_data['Close']-1)),
                          'values':'=Sheet1!$'+col_letter+'$2:$'+col_letter+'$755',
                          'name': col_name})

        peak=np.max(AMZN_data[col_name])
        peak_index=np.argmax(AMZN_data[col_name])

        chart.add_series({'categories':'=Sheet1!$A$'+str(2+peak_index),
                          'values':'=Sheet1!$'+col_letter+'$'+str(2+peak_index),
                          'name': 'Peak',
                          'marker':{'type':'circle'}})

        chart.set_x_axis({'name':'Date'})
        chart.set_title({'name': col_name})
        chart.set_y_axis({'name':col_name,'min':0,'max':1.2*peak})

        AMZN_worksheet.insert_chart('B'+str(17*(i)),chart)
workbook.close()