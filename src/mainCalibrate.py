
import os
import pandas as pd
import numpy as np
import colorama
import datetime as date
import pathlib

import config as cfg
import iecfg as iecfg
import ietcs as ie


startTotalTime = date.datetime.now()
numberofProcessedRegisters = 0
# ----------------------   CONFIG BLOCK - STARTS ----------------------------------------------------

# Filter for packets to compile
listRbcFilter = cfg.listRbcFilter
listJruFilter =  cfg.listJruFilter
listJruDiscard = cfg.listJruDiscard

# Variables for multiple files processing
boolMultipleFile = False    # To indicate if process multiple files in directory or a single file
# When boolMultipleFile False; the file to compile
inputSinglefile = 'JruC_JDRMDR Export 2026-06-16 12h 08m 42s-calibration.xlsx'

# Input and Output paths
inputPath = iecfg.inputPath
outputPath = iecfg.outputPath
calibrationPath = iecfg.calibrationPath

# CALIBRATION
nCalibrationIteractions = iecfg.nCalibrationIteractions
nameCalibrationFile = iecfg.nameCalibrationFile

# ----------------------   CONFIG BLOCK - ENDS ----------------------------------------------------

# ----------------------   USER INFO BLOCK - STARTS ----------------------------------------------------

# Print configuration info
print('ERTMS Insights. Version: ' + ie.ietcsVersion)
print('Using locale: ' + cfg.glLocale)
print('Processing multiple files: ' + str(boolMultipleFile))
print('Filter unknown balises: ' + str(ie.glboolFilterUnknownBalises))

# Print info for configured filters to be applied
outmessage = 'Applying filter for RBC packets. Processed packets: '
for x in range(len(listRbcFilter)):
    outmessage += (str(listRbcFilter[x]) + ',')
if (len(listRbcFilter) == 0) : outmessage += '*'
print(outmessage)

outmessage = 'Applying filter for JRU messages. Processed messages: '
for x in range(len(listJruFilter)):
    outmessage += (str(listJruFilter[x]) + ',')
if (len(listJruFilter) == 0) : outmessage += '*'
print(outmessage)

outmessage = 'Applying filter for JRU messages. Discarded messages: '
for x in range(len(listJruDiscard)):
    outmessage += (str(listJruDiscard[x]) + ',')
if (len(listJruDiscard) == 0) : outmessage += 'None'
print(outmessage)

print('Input path: ' + str(inputPath))
print('Output path: ' + str(outputPath))


# In case of multiple file processing
# Search for files in inputpath
listInputFiles = []
if boolMultipleFile == True:
    listInputFiles = os.listdir(str(inputPath))
else:
    listInputFiles = [inputSinglefile]

# First check: Only files that start by JruC_ (cfg.compiledFilePrefix) are allowed
for numseq in range(len(listInputFiles)):
    ix = None
    try:
        ix = listInputFiles[numseq].index(cfg.compiledFilePrefix)
    except:
        listInputFiles[numseq] = "not_allowed"

    if ix != 0: listInputFiles[numseq] = "not_allowed"

bool_not_allowed = True
while bool_not_allowed == True:
    try:
        listInputFiles.remove("not_allowed")
    except:
        bool_not_allowed = False


print('Files to calibrate: ' + str(listInputFiles))

print(colorama.Fore.GREEN + 'Calibration process: Antenna - frontend distance calibration' + colorama.Style.RESET_ALL)



userContinue = input("Continue? [y/n] Enter to continue. ")
if not((userContinue == 'y') | (userContinue == 'Y') | (userContinue == '')):
    print('Bye.')
    exit()

# ----------------------   USER INFO BLOCK - ENDS  ----------------------------------------------------

dfAggregatedAntennaDistances = pd.DataFrame() # Initialization
arrDfAntennaDistances = [] # Initialization
for numseq in range(len(listInputFiles)):

    #inputfile = inputpath + listInputFiles[0]
    inputfilePath = pathlib.Path.joinpath(inputPath, listInputFiles[numseq])
    strinputfile = str(inputfilePath)

    # Processing analytics
    arrEngine = []
    startTime = date.datetime.now()
           
    #  ---------------------- CALIBRATION ---------------------------
    
    t0 = date.datetime.now()
    print(colorama.Fore.GREEN + "Calibration starts" + colorama.Style.RESET_ALL)

    print("Loading compiled file: " + strinputfile)
    try:
        inputdf = pd.read_excel(strinputfile, 
                            sheet_name = 0,
                            dtype = {'BG_LINKS':str})
    except:
        print(colorama.Fore.YELLOW + "Can´t open compiled file" + colorama.Style.RESET_ALL)
        exit()

    # Removing unnamed columns using drop function
    inputdf.drop(inputdf.columns[inputdf.columns.str.contains(
        'unnamed', case=False)], axis=1, inplace=True)
        
    print("Preparing data and changing types")

    # Change DataFrame dtypes
    inputdf = ie.JruC_ComChangeDtypes(inputdf)
    dfMissionsSummary = pd.DataFrame()
    dfRoutes = pd.DataFrame()
    dfStatistics = pd.DataFrame()

    listAntInputFile = []
    listNidengine = []
    listIxMission = []
    listAntennaDistance = [] # Estimated antenna distance
    listErrorMean = [] # Estimated total error
    listDlinkMean = [] # Mean d_link distance
    listErrOdoMean = [] # Estimated odo error
    listOdoIntercept = [] # Estimated error due to antenna distance
    listErrDlinkMean = [] # Dlink error mean
    listErrDlinkStd = [] # Dlink error estandar deviation
    listErrDlinkMax = [] # Dlink error estandar deviation
    listErrDlinkMin = [] # Dlink error estandar deviation
    listDistance =  []
    listNumBG = []
    listOdoAccuracy = []

    dfAntennaDistances = pd.DataFrame() # Initialization
        
    # Get different nidengines
    df = inputdf[inputdf['NID_MESSAGE']!= 'STM INFORMATION [14]']
    arrEngine = np.array(df["NID_ENGINE"].unique())   
                
    for x in range (len(arrEngine)):

        nidEngine = str(arrEngine[x])

        # Calibration process: 
        ix = len(listAntennaDistance)

        listCalibratingDistance = [] # !Important
        distAntenna = 0 # initialDistAntenna # Initialization at first loop
        for i in range (nCalibrationIteractions): # max number of iteractions

            print ("Calibrating antenna distance for engine " + nidEngine + '. Loop number ' + str(i) + '. Estimated distance for antenna: ')
            
            if len(listCalibratingDistance) == 0:
                print(str(distAntenna))
            else:
                print(listCalibratingDistance)
            
            dfArrOutput = ie.JruC_SetRoutes(ie.glCalibrationProcess, inputdf, nidEngine, distAntenna, pd.DataFrame())
            dfMissionsSummary = dfArrOutput[0]
            dfRoutes = dfArrOutput[1]

            dfStatistics = ie.JruC_Statistics(dfMissionsSummary, dfRoutes)
            
            if(len(dfStatistics.axes[0]) == 0):
                print('No valid data in mission')
                break 

            listCalibratingDistance = [] # initialization each loop ! Important

            for n in range (len(dfStatistics.axes[0])):
                if i == 0:
                    listAntennaDistance.insert(ix+n, dfStatistics['intercept (eOdo)'][n]) # Initialization
                else:  
                    listAntennaDistance[ix+n] += dfStatistics['intercept (eOdo)'][n] 
                            
                listCalibratingDistance.insert(n, listAntennaDistance[ix+n])
                            
            distAntenna = listAntennaDistance[ix+n]
            if ((len(listAntennaDistance) == 1) & (str(distAntenna) == 'nan')):
                print('No valid data in mission')
                break
            print('Calibration accuracy:' )
            print(dfStatistics['errDlink (mean)'].max())
                        
            if(dfStatistics['errDlink (mean)'].max()) < iecfg.minCalibrationAccuracy:
                break 

        print ("Calibrating antenna distance for engine " + nidEngine + '. Loop number ' + str(i) + '. Estimated distance for antenna: ')
        if len(listCalibratingDistance) == 0:
                print(str(distAntenna))
        else:
            print(listCalibratingDistance)
                        
        for n in range(len(dfStatistics.axes[0])):
                    
            listNidengine.insert(ix+n, nidEngine)
            listIxMission.insert(ix+n, n)
            listAntInputFile.insert(ix+n, listInputFiles[numseq])
            listErrorMean.insert(ix+n, dfStatistics['error (mean)'][n])
            listErrOdoMean.insert(ix+n, dfStatistics['errOdo'][n])
            listDlinkMean.insert(ix+n, dfStatistics['dlink (mean)'][n])
            listOdoIntercept.insert(ix+n, dfStatistics['intercept (eOdo)'][n])
            listErrDlinkMean.insert(ix+n, dfStatistics['errDlink (mean)'][n])
            listErrDlinkStd.insert(ix+n, dfStatistics['errDlink (std)'][n])
            listErrDlinkMax.insert(ix+n, dfStatistics['errDlink (max)'][n])
            listErrDlinkMin.insert(ix+n, dfStatistics['errDlink (min)'][n])
            listDistance.insert(ix+n, dfStatistics['Distance'][n])
            listNumBG.insert(ix+n, dfStatistics['NumBG'][n])
            listOdoAccuracy.insert(ix+n, dfStatistics['Odo Accuracy'][n])
                    
                
        
    serNidengine = pd.Series(listNidengine, dtype = str)
    serIxMission = pd.Series(listIxMission, dtype = str)
    serAntInputFile = pd.Series(listAntInputFile, dtype = str)
    serAntennaDistance = pd.Series(listAntennaDistance, dtype = float)
    serErrorMean = pd.Series(listErrorMean, dtype = float)
    serOdoError = pd.Series(listErrOdoMean, dtype = float)
    serDlinkMean = pd.Series(listDlinkMean, dtype = float)
    serOdoIntercept = pd.Series(listOdoIntercept, dtype = float)
    serErrDlinkMean = pd.Series(listErrDlinkMean, dtype = float)
    serErrDlinkStd = pd.Series(listErrDlinkStd, dtype = float)
    serErrDlinkMax = pd.Series(listErrDlinkMax, dtype = float)
    serErrDlinkMin = pd.Series(listErrDlinkMin, dtype = float)
    serDistance = pd.Series(listDistance, dtype = float)
    serNumBG = pd.Series(listNumBG, dtype = int)
    serOdoAccuracy = pd.Series(listOdoAccuracy, dtype = float)       
            
            
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'File', serAntInputFile)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Engine', serNidengine)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Mission', serIxMission)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'ErrPos', serErrorMean)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Antenna', serAntennaDistance)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'ErrOdo', serOdoError)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'AccOdo', serOdoAccuracy)
    # dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'intercept (eOdo)', serOdoIntercept)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'meanErrDlink', serErrDlinkMean)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'stdErrDlink', serErrDlinkStd)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'maxErrDlink', serErrDlinkMax)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'minErrDlink', serErrDlinkMin)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Distance', serDistance)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'NumBG', serNumBG)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Dlink', serDlinkMean)
    dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'AccCal', serErrDlinkMean)
    
    arrDfAntennaDistances.insert(len(arrDfAntennaDistances), dfAntennaDistances)
            

    t1 = date.datetime.now()
    delta = t1-t0
    delta = delta.total_seconds()
    print(colorama.Fore.GREEN + "File calibration ends" + '. Partial time: ' + str(delta) + ' secs' + colorama.Style.RESET_ALL)
            

# Aggregate antenna distances
for n in range(len(arrDfAntennaDistances)):
        
    dfAntennaDistances = arrDfAntennaDistances[n] 
    dfAggregatedAntennaDistances = pd.concat([dfAggregatedAntennaDistances, dfAntennaDistances], ignore_index = True)

# Save calibration file
pathCalibratedAntdistFile = pathlib.Path.joinpath(calibrationPath, nameCalibrationFile)
outputCalibratedAntdistFile = str(pathCalibratedAntdistFile)    
print("Saving calibrated antena distance file. " + outputCalibratedAntdistFile)
# Create output folder if not exists
pathlib.Path(calibrationPath).mkdir(parents=True, exist_ok=True)
# Excel write (end of Phase 2)
try:
    with pd.ExcelWriter(outputCalibratedAntdistFile) as writer:
        wb = writer.book
        dfAggregatedAntennaDistances.to_excel(writer, sheet_name='CALIBRATION')
                            
except Exception as ex:
    print(colorama.Fore.YELLOW + "Can´t save calibration file. Error: " +str(ex) + colorama.Style.RESET_ALL)    


endTotalTime = date.datetime.now()
delta = endTotalTime - startTotalTime
delta = delta.total_seconds()
print("Total process time: " + str(delta) + ' secs')
