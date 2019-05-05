# from VO_parser import  *
import pandas
import numpy as np
import os
import sys
if sys.version_info[0] >= 3:
    import PySimpleGUI as sg
else:
    import PySimpleGUI27 as sg


def ExecuteInterface():
    layout = [[sg.Text("Select folder to export '.csv' files to:"), sg.Text('', key='_OUTPUT_') ],
              [sg.Input(key='_IN_'), sg.FolderBrowse()],
              [sg.Text('Select files to convert:')],
              [sg.Input(key='_WORK_'), sg.FilesBrowse()],
              [sg.Radio('Visual Observer Files', "RADIO1", default=True, key='_VORAD_'),
               sg.Radio('Boundary files', "RADIO1", key='_BDRAD_')],
              [sg.Button('Convert'), sg.Button('Exit')]]

    window = sg.Window('RFRL File Converter', layout)

    while True:                 # Event Loop
      event, values = window.Read()
      print(event, values)
      if event is None or event == 'Exit':
          break
      if event == 'Convert':
          dictString = values['_WORK_']
          filesList = dictString.split(';')

          if (values['_VORAD_'] == True):

              for VO_file in filesList:
                exportDirectory = values['_IN_']
                print(VO_file)
                drive, path_and_file = os.path.splitdrive(VO_file)
                path, file = os.path.split(path_and_file)
                file = file.rstrip(".xlsx")

                exportDirectory = (exportDirectory + "/" + file + ".csv")
                VO_FileConversion(VO_file, exportDirectory)

          # TODO: Implement file proccessing for BOUNDARY files
          #     also need to implement BOUND function in 'VO_parser.py' to pass to.
          if (values['_BDRAD_'] == True):

              for (BD_file) in filesList:
                  exportDirectory = values['_IN_']
                  print(BD_file)

                  drive, path_and_file = os.path.splitdrive(BD_file)
                  path, file = os.path.split(path_and_file)
                  file = file.rstrip(".xlsx")

                  exportDirectory = (exportDirectory + "/" + file + ".csv")
                  BD_FileConversion(BD_file, exportDirectory)

          window.FindElement('_WORK_').Update('')

    window.Close()


# TODO: ***Early Function Version: code inside is functional for multiple files
#   passed in via 'scratch1_persistantTEST.py'
def VO_FileConversion(user_workFile, user_exportDir):
    workFile = user_workFile
    exportDir = user_exportDir

    # TODO: FUNCTIONAL w/ args passed in from interface - make usable for BOUND files also
    #   seperate function for BOUND files???
    df = pandas.read_excel(workFile, usecols=(0,3,4))
    df.rename(columns={"Name": "POINT_ID"}, inplace=True)
    df.to_csv(exportDir, index=False)

def BD_FileConversion(user_workFile, user_exportDir):
    workFile = user_workFile
    exportDir = user_exportDir

    # TODO: FUNCTIONAL w/ args passed in from interface - make usable for BOUND files also
    #   seperate function for BOUND files???
    df = pandas.read_excel(workFile, usecols=(1,2))
    df.insert(0, 'POINT_ID', np.arange(1, 1 + len(df)))
    df.to_csv(exportDir, index=False)

if __name__ == '__main__':
    ExecuteInterface()