# -*- coding: utf-8 -*-
## Instructions about script:
## Edit these name of instance for concrete instance, if your model parts use different nomenclature  
## Input files provide this information  
## You can use another script with node details to copy instance name as well 
## Abaqus library paths vary—update accordingly (e.g., mine is c:/SIMULIA/Abaqus/6.14-1/code/python2.7/lib/abaqus_plugins/excelUtilities)  
## Check your laptop’s path and modify it in the code  
## If print issues arise, your Abaqus may use Python >2.7—update print statements and fix any errors  

## Input required
# Define part name and list of node numbers
# The idea is this, wether you want to work on extensiometer nodes or the sensors, it will extract displacement histories for you from ABAQUS file which you can process later.
# For example, look at the part name here 'Bridge Pier with Slabs-1', for your own file you will need to check assembly to see on which part the sensors are placed.
# Edit this part_name variable and also find the node number (It is recommended to do partitions for it, it will help you alot in post processing.
# If you cannot use this script, alternative way is to use history outputs (No worries :P)
# For any kind of information related to this script, feel free to ask me on my WhatsApp: +923440907874 or email: Tufail_mabood@yahoo.com

part_name = 'Bridge Pier with Slabs-1'  # You will get the name from PrintNodes.py i provided with this file. If the name is incorrect, the script will not work.
node_numbers = [243, 77, 233, 67, 224, 58, 546, 544]  # Nodes where sensors are located (I have included the nodes, however if you are doing Mesh Sensitivity, you will need to change the node tags accordingly)

## All libraries required
from abaqus import *
from abaqusConstants import *
import sys
import os
import time
import win32com.client
import subprocess
from caeModules import *
from driverUtils import executeOnCaeStartup

## Set Viewport
executeOnCaeStartup()
vp = session.viewports['Viewport: 1']
vp.makeCurrent(), vp.maximize(), vp.partDisplay.geometryOptions.setValues(referenceRepresentation=ON)

## Search for ODB file in the current directory
cwd = os.getcwd()
results_dir = os.path.join(cwd, "Results")  # Define results folder path

## Create "Results" folder if it doesn't exist (I have done it to get your data output clean)
if not os.path.exists(results_dir):
    os.makedirs(results_dir)
odb_files = [f for f in os.listdir(cwd) if f.endswith('.odb')]
if not odb_files:
    raise RuntimeError("No .odb file found in the current directory.")
odb_path = os.path.join(cwd, odb_files[0])  # Open the first found ODB file

## Get print message for odb file in ABAQUS Console
print "Opening ODB file:", odb_path
o1 = session.openOdb(odb_path)
vp = session.viewports['Viewport: 1']
vp.setValues(displayedObject=o1)
vp.odbDisplay.display.setValues(plotState=(CONTOURS_ON_DEF,))
odb = session.odbs[odb_path]

## Define displacement variables (If you are looking for other vairables, change these however for this you need to understand the variables nomenclature from ABAQUS documentations)
# Since i have done scripting for geometry too, the script have shared coordinate system with ABAQUS assembly so, note that:
# x - Cyclic Loading Direction
# y - Along Centre line of Stem of the assembly (In this direction the lumped mass gravity is acting in -ve direction)
# z - Out of Plan Direction

variables = [('U1', ('U', NODAL, ((COMPONENT, 'U1'),))),
             ('U2', ('U', NODAL, ((COMPONENT, 'U2'),))),
             ('U3', ('U', NODAL, ((COMPONENT, 'U3'),)))]

## Iterate over each node number (This is for saving the sensors data you have defined in the list in beginning of this script)
for node_number in node_numbers:
    for var_name, var_tuple in variables:
        unique_name = "Disp_at_Sensor_of_Node_{}_{}".format(node_number, var_name)  
        xyList = xyPlot.xyDataListFromField(odb=odb, outputPosition=NODAL, variable=(var_tuple,), nodeLabels=((part_name, (str(node_number),)),))

        # Ensure unique XYPlot (A warning appears, if this is not used because of ABAQUS limitation for deleting the XY Data from the ABAQUS visualization module)
        if unique_name in session.xyPlots:
            del session.xyPlots[unique_name]
            
        xyp = session.XYPlot(unique_name)
        chartName = xyp.charts.keys()[0]  # Fix indexing issue
        chart = xyp.charts[chartName]
        curveList = session.curveSet(xyData=xyList)
        chart.setValues(curvesToPlot=curveList)
        session.viewports['Viewport: 1'].setValues(displayedObject=xyp)

        # Define the save path for text file inside "Results" folder
        text_file_path = os.path.join(results_dir, unique_name + ".txt")

        # Save XY Data to text file
        with open(text_file_path, "w") as txt_file:
            txt_file.write("Part: " + part_name + "\n")
            txt_file.write("Node: " + str(node_number) + "\n")
            txt_file.write(var_name + "\n")
            for xyData in xyList:
                for point in xyData:
                    txt_file.write("{:.6f} {:.6f}\n".format(point[0], point[1]))
        print "Text file saved at:", text_file_path

        # Import Excel utility
        # Note that i have checked various abaqus installation of researchers and they prefer to change it accordingly during installation so, update this according to your installation
        # You can get this address directly from ABAQUS Tools and Plugins in GUI
        # Do not intermix forward slash and backward slash, only use forward slash
        sys.path.insert(9, r'c:/SIMULIA/Abaqus/6.14-1/code/python2.7/lib/abaqus_plugins/excelUtilities')
        import abq_ExcelUtilities.excelUtilities

        # Export to Excel (this will create a new Excel workbook)
        abq_ExcelUtilities.excelUtilities.XYtoExcel(xyDataNames=unique_name, trueName=unique_name)

        time.sleep(5)  # Allow time for Excel to open
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")  # Attach to the existing Excel session
        except:
            excel = win32com.client.Dispatch("Excel.Application")

        excel.Visible = False
        excel.DisplayAlerts = False

        # Get the most recently opened workbook (since Abaqus creates a new one dynamically)
        try:
            wb = excel.Workbooks(excel.Workbooks.Count)  # Select the last opened workbook
            print "Using Excel Workbook:", wb.Name
        except:
            print "Error: No active Excel workbook found!"
            excel.Quit()
            continue

        # Ensure the existing sheet remains unchanged
        ws = wb.ActiveSheet  

        # Define the intended save path inside "Results" folder
        save_path = os.path.join(results_dir, "Disp_at_Sensor_of_Node_{}_{}.xlsx".format(node_number, var_name))

        try:
            # Save and close the workbook without renaming sheets
            wb.SaveAs(save_path)  # Save it with the correct name
            wb.Close(SaveChanges=True)
            excel.Quit()
            time.sleep(2)
            subprocess.call("taskkill /F /IM excel.exe", shell=True)
            print "Excel file saved at:", save_path
        except Exception as e:
            print "Error interacting with Excel:", str(e)

        # Final confirmation
        time.sleep(3)
        if os.path.exists(save_path):
            print "Final confirmation: File successfully saved at:", save_path
        else:
            print "Warning: File was not found at expected location!"
