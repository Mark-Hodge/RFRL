# A simple file conversion program with additional functionality able to be
#   implemented at a later date.
# Developed for use by the Raspet Flight Research Laboratory
# @Author Mark Hodge        @contact mrh598@msstate.edu
# Created: 01 May 2019
# Last Modified: 07 May 2019
# Version 1.0.0
# ----------------------------------------------------------------------------------


# from VO_parser import  *
import pandas
import numpy as np
import os
import sys
if sys.version_info[0] >= 3:
    import PySimpleGUI as sg
else:
    import PySimpleGUI27 as sg

# Execute GUI and handle data passed in via 'Convert'.
def ExecuteInterface():

    # Defines the functional elements and look of GUI.
    layout = [[sg.Text("Select folder to export '.csv' files to:"), sg.Text('', key='_OUTPUT_') ],
              [sg.Input(key='_IN_'), sg.FolderBrowse()],
              [sg.Text('Select files to convert:')],
              [sg.Input(key='_WORK_'), sg.FilesBrowse()],
              [sg.Radio('Visual Observer files', "RADIO1", default=True, key='_VORAD_'),
               sg.Radio('Boundary files', "RADIO1", key='_BDRAD_')],
              [sg.Button('Convert'), sg.Button('Exit')]]

    # Call GUI window using the defined layout above.
    window = sg.Window('RFRL File Converter', layout)

    # Loop while program is running (until user terminates via 'Exit').
    while True:

      event, values = window.Read()             # Read in data from GUI.
      # print(event, values) <--- This is used for testing expected input vs. actual input.


      if event is None or event == 'Exit':      # Break loop and terminate program if user clicks 'Exit'.
          break
      if event == 'Convert':                    # Read in data and process if user clicks 'Convert'.
          dictString = values['_WORK_']         # Import all working files into a raw string.
          filesList = dictString.split(';')     # Split files in dictString by ';' and append to fileList(array).

          # Check if filetype button for 'Visual Observer files' selected.
          if (values['_VORAD_'] == True):

              # Iterate through elements in array.
              for VO_file in filesList:
                exportDirectory = values['_IN_']                        # Read in a single file path.
                # print(VO_file) <-- this line is used for testing purposes.

                drive, path_and_file = os.path.splitdrive(VO_file)      # Get location info form OS.
                path, file = os.path.split(path_and_file)               # Get filename info from OS.
                file = file.rstrip(".xlsx")                             # Strip off '.xlsx' type designator.

                exportDirectory = (exportDirectory + "/" + file + ".csv")   # Use original filename, using '.csv' type designator.
                VO_FileConversion(VO_file, exportDirectory)                 # Call function to process current file.


          # Check if filetype button for 'Boundary files' selected.
          if (values['_BDRAD_'] == True):

              # Iterate through elements in array.
              for (BD_file) in filesList:
                exportDirectory = values['_IN_']                      # Read in a single file path.
                # print(BD_file) <-- this line is used for testing purposes.

                drive, path_and_file = os.path.splitdrive(BD_file)    # Get location info form OS.
                path, file = os.path.split(path_and_file)             # Get filename info form OS.
                file = file.rstrip(".xlsx")                           # Stripp off '.xlsx' type designator.

                exportDirectory = (exportDirectory + "/" + file + ".csv")   # Use original filename, using '.csv' type designator.
                BD_FileConversion(BD_file, exportDirectory)                 # Call function to process current file.

          window.FindElement('_WORK_').Update('')       # Clear 'Select files to convert field'.

    # Close window object.
    window.Close()

# Convert single visual observer file and write finished product to user's desired directory.
def VO_FileConversion(user_workFile, user_exportDir):
    workFile = user_workFile
    exportDir = user_exportDir

    # initialize  Dataframe object and extract useful columns.
    df = pandas.read_excel(workFile, usecols=(0,3,4))

    # Rename index column and data columns.
    df.rename(columns={"Name": "POINT_ID"}, inplace=True)
    df.rename(columns={"Location/LatLonAlt/LatitudeDeg": "LatitudeDeg"}, inplace=True)
    df.rename(columns={"Location/LatLonAlt/LongitudeDeg": "LongitudeDeg"}, inplace=True)

    # Write converted data to a '.csv' file of the same name.
    df.to_csv(exportDir, index=False)

# Convert single Boundary file and write finished product to user's desired directory.
def BD_FileConversion(user_workFile, user_exportDir):
    workFile = user_workFile
    exportDir = user_exportDir

    # Initialize Dataframe object and extract useful data columns.
    df = pandas.read_excel(workFile, usecols=(0,1))
    df.insert(0, 'POINT_ID', np.arange(1, 1 + len(df)))     # Insert a column at position '0' and index rows.

    # Write converted data to a '.csv' file of the same name.
    df.to_csv(exportDir, index=False)

# Initialize program and call main function on startup.
if __name__ == '__main__':
    ExecuteInterface()