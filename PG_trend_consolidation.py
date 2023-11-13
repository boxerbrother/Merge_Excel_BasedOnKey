import pandas as pd
import time
import os

if __name__=='__main__':
    s1 = time.time()    

    filename= input('Enter the file path of PD_HGPD merged output: ')
    #Check if the input file exists or not, if file doesn't exists in the provided path then exit
    if not (os.path.isfile(filename)):
        print("The entered file doesn't exists")
        print("Then entered file doesn't exist. Please re-run with correct file.")
        time.sleep(300)
        exit()
    else:
        print("PD File provided is valid %s", filename)

    #filename = r'C:\Users\sonikar\Documents\Scheme Formulation\Process_automation_output\PG_trend_analysis\PD_HDPD_Merged_Output 2023 Sep.xlsx'
    print("Begin reading the PD file...")

    #Read monthly data excel file
    monthly_data= pd.read_excel(filename)

    #Create Empty List
    column_name = list()
   
    #Add 'Part Number' 
    column_name.append(monthly_data.columns.values[0])
    #Add 'Qty'
    column_name.append(monthly_data.columns.values[4])
    #Add 'Value'
    column_name.append(monthly_data.columns.values[-1])
    
    masterfile_path = r'C:\Users\sonikar\Documents\ProcessAutomation\PG_Trend_Consolidation\PG_List_consolidated.xlsx'
    #Open PG_List_Consolidated excel file in WRITE made
    masterfile_data= pd.read_excel(masterfile_path)

    #Vlookup on 'Part Number' and whereever 'Part Number' matches append 'qty' & 'value' columns for that row
    merged_data = pd.merge(masterfile_data,monthly_data[column_name], on='Part Number',how='left')
    
    #Write Excel file without indexing
    merged_data.to_excel(masterfile_path,index=False)

    print('The consolidated file has been updated.\nThe file is present at ',masterfile_path)
    print("Total Time taken in secods is : ", time.time()- s1)
