import pandas as pd
import numpy as np
import os
import glob
import openpyxl as pyxl
from datetime import date

def doProcess(INPATH=".",OUTPATH="."):
    #INPATH=r"C:\Nomination"
    global dt
    dt= date.today().strftime("%d-%b-%Y")

    Input_fetch=glob.glob(os.path.join(INPATH, 'Input*.xlsx'))

    for inputs in Input_fetch:
        print(inputs)
        data=pd.read_excel(inputs)
        print(data)
        mbnl_id=data["MBNL ID"].values[0]
        TNS_id=data["Site ID"].values[0]
        easting=data["Physical Easting"].values[0]
        northing=data["Physical Northing"].values[0]
        Azimuth=data["Azimuth"].values[0]
        Height=data["C/L Height"].values[0]
        Antenna_type=data["Antenna_Type"].values[0]
        H3G_ID=data["H3G ID"].values[0]
        Radio_Planner=data["Radio Planner"].values[0]
        Contact=data["Contact"].values[0]
        Final_Option=data["Final Option"].values[0]
        Comment=data["Comment"].values[0]
        if ('/') in Azimuth:
            print(Azimuth)
            Azimuth_A=data["Azimuth"].values[0].split('/')[0]
            print(Azimuth_A)
            Azimuth_B=data["Azimuth"].values[0].split('/')[1]
            print(Azimuth_B)
            Azimuth_C=data["Azimuth"].values[0].split('/')[2]
            print(Azimuth_C)
        else:
            Azimuth_A=data["Azimuth"].values[0]
            Azimuth_B=data["Azimuth"].values[0]
            Azimuth_C=data["Azimuth"].values[0]
            print("No difference in azimuth")

        if ('/') in Height:
            print(Height)
            Height_A=data["C/L Height"].values[0].split('/')[0]
            print(Height_A)
            Height_B=data["C/L Height"].values[0].split('/')[1]
            print(Height_B)
            Height_C=data["C/L Height"].values[0].split('/')[2]
            print(Height_C)
        else:
            Height_A=data["C/L Height"].values[0]
            Height_B=data["C/L Height"].values[0]
            Height_C=data["C/L Height"].values[0]
            print("No difference in Height")
        
    Nomination_name = glob.glob(os.path.join(INPATH, 'NOMINATION*.xlsx'))
    print(Nomination_name)
    for file in Nomination_name:
        
            print(file)
            wb = pyxl.load_workbook(file) 
            sheet = wb.active
            if sheet['k5'].value is None :
                sheet['k5'].value=mbnl_id

            if sheet['M6'].value is None:
                sheet['M6'].value=TNS_id
                
            if sheet['D8'].value is None:
                sheet['D8'].value=easting

            if sheet['D9'].value is None:
                sheet['D9'].value=northing


            if sheet['k4'].value is None:
                sheet['k4'].value=H3G_ID
                sheet['C17'].value=H3G_ID
                sheet['C18'].value=H3G_ID
                sheet['C19'].value=H3G_ID
                sheet['C23'].value=H3G_ID
                sheet['C24'].value=H3G_ID
                sheet['C25'].value=H3G_ID
                sheet['C29'].value=H3G_ID
                sheet['C30'].value=H3G_ID
                sheet['C31'].value=H3G_ID
                sheet['C35'].value=H3G_ID
                sheet['C36'].value=H3G_ID
                sheet['C37'].value=H3G_ID
                sheet['P17'].value=Antenna_type
                sheet['P18'].value=Antenna_type
                sheet['P19'].value=Antenna_type
                sheet['P23'].value=Antenna_type
                sheet['P24'].value=Antenna_type
                sheet['P25'].value=Antenna_type
                sheet['P29'].value=Antenna_type
                sheet['P30'].value=Antenna_type
                sheet['P31'].value=Antenna_type
                sheet['P35'].value=Antenna_type
                sheet['P36'].value=Antenna_type
                sheet['P37'].value=Antenna_type

            if sheet['D5'].value is None:
                sheet['D5'].value=Radio_Planner
                
            if sheet['D6'].value is None:
                sheet['D6'].value=Contact

            if sheet['D11'].value is None:
                sheet['D11'].value=Final_Option


            if sheet['D13'].value is None:
                sheet['D13'].value=Comment

            if sheet['F17'].value is None or not None:
                sheet['F17'].value=Height_A
                sheet['F18'].value=Height_B
                sheet['F19'].value=Height_C
                sheet['F23'].value=Height_A
                sheet['F24'].value=Height_B
                sheet['F25'].value=Height_C
                sheet['F29'].value=Height_A
                sheet['F30'].value=Height_B
                sheet['F31'].value=Height_C
                sheet['F35'].value=Height_A
                sheet['F36'].value=Height_B
                sheet['F37'].value=Height_C

            if sheet['G17'].value is None or not None:
                sheet['G17'].value=Azimuth_A
                sheet['G18'].value=Azimuth_B
                sheet['G19'].value=Azimuth_C
                sheet['G23'].value=Azimuth_A
                sheet['G24'].value=Azimuth_B
                sheet['G25'].value=Azimuth_C
                sheet['G29'].value=Azimuth_A
                sheet['G30'].value=Azimuth_B
                sheet['G31'].value=Azimuth_C
                sheet['G35'].value=Azimuth_A
                sheet['G36'].value=Azimuth_B
                sheet['G37'].value=Azimuth_C


    original_filename = os.path.splitext(os.path.basename(file))[0] 
    wb.save(os.path.join(OUTPATH, f"{original_filename}_"+mbnl_id+dt+".xlsx"))
    print(original_filename, '      Saved')      
         

doProcess(INPATH=r"C:\Nomination" , OUTPATH=r"C:\Nomination")
