import os
import pandas as pd
import numpy as np
import colorama
import datetime as date
import pathlib

import config as cfg
import ietcs as ie
import iecfg as iecfg

global glversion

startTotalTime = date.datetime.now()
numberofProcessedRegisters = 0
# ----------------------   CONFIG BLOCK - STARTS ----------------------------------------------------

# Filter for packets to compile

listRbcFilter = cfg.listRbcFilter


listJruFilter =  cfg.listJruFilter

listJruDiscard = cfg.listJruDiscard


# Defining phases to execute
boolPhase1 = True   # To jump phase 2 (set routes, odo errors and statistics)
boolPhase2 = True   # To jump phase 3 (splitting routes and preparing charts)


# Variables for multiple files processing
boolMultipleFile = False    # To indicate if process multiple files in directory or a single file
# When boolMultipleFile False; the file to compile
#inputSinglefile = 'JruC_ESC S-103 CoMa 22-24 jun 2024 C1 (5).xlsx'
inputSinglefile = 'JruC_JDRMDR Export 2026-06-16 12h 08m 42s.xlsx'

# Input and Output paths

inputPath = iecfg.inputPath
outputPath = iecfg.outputPath
statisticsPath = iecfg.statisticsPath


# ESTIMATION
# boolProcess = ie.glEstimationProcess # To estimate routes, distances and odoerror
fileStatistics = 'StatisticsFile.xlsx'

# distAntenna = 5.0
# distAntenna = 9.0 # 8.98846075444001 # Cab 23 Isla Valencia
# distAntenna = 9.0 # 9.08745103043642 # Cab 24 Isla Valencia
# distAntenna = 18.9935265786265 # Israel 2024 Reverse 
# distAntenna = 1.0 # Israel  Nominal
# distAntenna = 19.0 # Israel  Reverse
#
# distAntenna = 6.73 # 673 cm según info de CEDEX # 5.17808485579857 # 5.97298442991297 # 5.97554642364505  # CoMa C1
# calOdoSlope = 1.00 - 0.999452402168725 # Use the value calculated during calibration

# distAntenna = 2.31822980987364 # 2.45178132582484 # CoMa C8
# distAntenna = 9.24 # 9.2877572294039 # S112 Isla Murcia

distAntenna = 2.00 # 1.23 # 2.00 # León Guardo

fileAntennaDistances = '' # To be implemented. Use a file with different antenna distances



# ----------------------   CONFIG BLOCK - ENDS ----------------------------------------------------

# ----------------------   USER INFO BLOCK - STARTS ----------------------------------------------------

# Print configuration info
print('ERTMS Insights. Version: ' + ie.ietcsVersion)
print(colorama.Fore.YELLOW + 'Using locale: ' + cfg.glLocale + colorama.Style.RESET_ALL)
print('Processing multiple files: ' + str(boolMultipleFile))
print(colorama.Fore.YELLOW + 'Filter unknown balises: ' + str(cfg.glboolFilterUnknownBalises) + colorama.Style.RESET_ALL)

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



print('Files to set routes: ' + str(listInputFiles))

if cfg.dmFile != '':
    print('Digital map file: ' + str(cfg.dmFile))

print(colorama.Fore.GREEN + 'Estimation process: Route, distance and odo error estimation' + colorama.Style.RESET_ALL)

dmFile = cfg.dmFile



userContinue = input("Continue? [y/n] Enter to continue. ")
if not((userContinue == 'y') | (userContinue == 'Y') | (userContinue == '')):
    print('Bye.')
    exit()

# ----------------------   USER INFO BLOCK - ENDS  ---------------------------------------------------

dfAggregatedAntennaDistances = pd.DataFrame() # Initialization
arrDfAntennaDistances = [] # Initialization

# For each comppiled file to be processed
for numseq in range(len(listInputFiles)):
    
    #inputfile = inputpath + listInputFiles[0]
    inputfilePath = pathlib.Path.joinpath(inputPath, listInputFiles[numseq])
    strinputfile = str(inputfilePath)

    # Get inputfilename without prefix and extension
    # Prefix = JRU_C (variable: compiledFilePrefix)
    # Extension = .xlsx

    # Remove prefix
    strAux = strinputfile
    iAux = strAux.index(cfg.compiledFilePrefix)
    strinputfileNoprefix = strAux[iAux + len(cfg.compiledFilePrefix):len(strAux)]

    # Remove extension
    strAux = strinputfile
    ix = 0
    n = 0
    while True:
        try:   
            iAux = strAux.index('.')
            strAux = strAux[iAux +1:len(strAux)]
            if n == 0:
                ix += iAux
            else:
                ix += (iAux + 1)
            n += 1
        except:
            break
    inputFileNameNoextension = strinputfileNoprefix[0:ix].lstrip() 

    # Processing analytics
    arrEngine = []
    startTime = date.datetime.now()
     
    #  ---------------------- PHASE 1 - SETTING ROUTES ---------------------------
    if boolPhase1 == True:

        t0 = date.datetime.now()
        print(colorama.Fore.GREEN + "Routes setting phase starts." + colorama.Style.RESET_ALL)

        print("Loading compiled file. " + strinputfile)
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

        if (fileAntennaDistances != ''):
            print("Getting reference antenna distances from " + fileAntennaDistances)
            # To be implemented
        
        # Get different nidengines
        df = inputdf # [inputdf['NID_MESSAGE']!= 'STM INFORMATION [14]']
        arrEngine = np.array(df["NID_ENGINE"].unique())

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
        listUslope = []
        listUslopeMax = []
        listOslope = []
        listOslopeMax = []
        listUsafe = []
        listOsafe = []
        listOverCorrVel =  []
        listUnderCorrVel = []

        dfAntennaDistances = pd.DataFrame() # Initialization
        dfDigitalMap = pd.DataFrame() # Initialization

        # Loading Digital Map File
        dmFilePath = cfg.dmFilePath
        strDigitalMapFile = cfg.dmFile
        if cfg.dmFile != '':

            dmFilePath = pathlib.Path.joinpath(cfg.dmFilePath, cfg.dmFile)
            strDmFile = str(dmFilePath)

            print("Loading digital map file. " + strDmFile)
            try:
                dfDigitalMap = pd.read_excel(strDmFile, 
                                    sheet_name = 0) 
                
            except:
                print(colorama.Fore.YELLOW + "Can´t open digital map file" + colorama.Style.RESET_ALL)
                dfDigitalMap = pd.DataFrame() # Reset
                
        for x in range (len(arrEngine)):
            nidEngine = str(arrEngine[x])

            # Output
            dfArrOutput = []
            pathRoutesFile = pathlib.Path.joinpath(outputPath, 'R_' + inputFileNameNoextension + '_' + nidEngine + '.xlsx')
            outputRoutesFile = str(pathRoutesFile)

               
            # Estimation process: routes setting and statistics
            print ("Setting routes and odometry accuracy for " + nidEngine) 

            # Get reference antenna distances
            if fileAntennaDistances != '':
                distAntenna = 0.0 # To be implemented; get distAntenna from file

            dfArrOutput = ie.JruC_SetRoutes(ie.glEstimationProcess, inputdf, nidEngine, distAntenna, dfDigitalMap)
            dfMissionsSummary = dfArrOutput[0]
            dfRoutes = dfArrOutput[1]

            print ("Setting statistics for " + nidEngine) 
            dfStatistics = ie.JruC_Statistics(dfMissionsSummary, dfRoutes)
            dfMissionsSummary = pd.concat([dfMissionsSummary, dfStatistics], axis = 1)
                
            for n in range(len(dfStatistics.axes[0])):
                    
                listNidengine.insert(n, nidEngine)
                listIxMission.insert(n, n)
                listAntInputFile.insert(n, listInputFiles[numseq])
                listErrorMean.insert(n, dfStatistics['error (mean)'][n])
                listErrOdoMean.insert(n, dfStatistics['errOdo'][n])
                listDlinkMean.insert(n, dfStatistics['dlink (mean)'][n])
                listOdoIntercept.insert(n, dfStatistics['intercept (eOdo)'][n])
                listErrDlinkMean.insert(n, dfStatistics['errDlink (mean)'][n])
                listErrDlinkStd.insert(n, dfStatistics['errDlink (std)'][n])
                listErrDlinkMax.insert(n, dfStatistics['errDlink (max)'][n])
                listErrDlinkMin.insert(n, dfStatistics['errDlink (min)'][n])
                listDistance.insert(n, dfStatistics['Distance'][n])
                listNumBG.insert(n, dfStatistics['NumBG'][n])
                listOdoAccuracy.insert(n, dfStatistics['Odo Accuracy'][n])
                listAntennaDistance.insert(n, distAntenna)
                listUslope.insert(n, dfStatistics['underslope (mean)'][n])
                listUslopeMax.insert(n, dfStatistics['underslope (max)'][n])
                listOslope.insert(n, dfStatistics['overslope (mean)'][n])
                listOslopeMax.insert(n, dfStatistics['overslope (max)'][n])
                listUsafe.insert(n, dfStatistics['usafeerror (max)'][n])
                listOsafe.insert(n, dfStatistics['osafeerror (max)'][n])
                listUnderCorrVel.insert(n, dfStatistics['undercorr (vel)'][n])
                listOverCorrVel.insert(n, dfStatistics['overcorr (vel)'][n])
                
            print("Saving routes file. " + outputRoutesFile)
            # Create output folder if not exists
            pathlib.Path(outputPath).mkdir(parents=True, exist_ok=True)
            # Excel write (end of Phase 2)
            try:
                with pd.ExcelWriter(outputRoutesFile) as writer:
                    wb = writer.book
                    dfMissionsSummary.to_excel(writer, sheet_name='MISSIONS')
                    dfRoutes.to_excel(writer, sheet_name='ROUTES')
                    bool = ie.JruC_SetRoutesFileFormat(wb, dfMissionsSummary)
                        
            except Exception as ex:
                print(colorama.Fore.YELLOW + "Can´t save routes file. Error: " +str(ex) + colorama.Style.RESET_ALL)
        
        # Aggregating statistics
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
        serUslope = pd.Series(listUslope, dtype = float)
        serUslopeMax = pd.Series(listUslopeMax, dtype = float)
        serOslope = pd.Series(listOslope, dtype = float)
        serOslopeMax = pd.Series(listOslopeMax, dtype = float)
        serUsafe = pd.Series(listUsafe, dtype = float)
        serOsafe = pd.Series(listOsafe, dtype = float)
        serUcorrVel = pd.Series(listUnderCorrVel, dtype = float)
        serOcorrVel = pd.Series(listOverCorrVel, dtype = float)
            
            
            
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
        dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Uslope', serUslope)
        dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Uslope (max)', serUslopeMax)
        dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Oslope', serOslope)
        dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Oslope (max)', serOslopeMax)
        dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Usafe', serUsafe)
        dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'Osafe', serOsafe)
        dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'UcorrVel', serUcorrVel)
        dfAntennaDistances.insert(len(dfAntennaDistances.columns), 'OcorrVel', serOcorrVel)
       
        arrDfAntennaDistances.insert(len(arrDfAntennaDistances), dfAntennaDistances)
            
        

        t1 = date.datetime.now()
        delta = t1-t0
        delta = delta.total_seconds()
        print(colorama.Fore.GREEN + "Routes setting phase ends" + '. Partial time: ' + str(delta) + ' secs' + colorama.Style.RESET_ALL)
            
    else:
        print(colorama.Fore.GREEN + "Jumping phase 2" + colorama.Style.RESET_ALL)

    #  ---------------------- PHASE 2 - Creating Charts ---------------------------
    if boolPhase2 == True:

        if boolPhase1 == False:  
            print(colorama.Fore.GREEN + "Routes setting phase was not executed. Chart phase jumped." + colorama.Style.RESET_ALL)
            
        else:
            t0 = date.datetime.now()
            print(colorama.Fore.GREEN + "Chart phase starts" + colorama.Style.RESET_ALL)
            # arrEngine = ['0']  # !!!!!!!!!!!!!!
            for x in range (len(arrEngine)):
                nidEngine = str(arrEngine[x])
                pathRoutesFile = pathlib.Path.joinpath(outputPath, 'R_' + inputFileNameNoextension + '_' + nidEngine + '.xlsx')
                inputRoutesFile = str(pathRoutesFile)
                print("Loading file: " + inputRoutesFile)
                try:

                    dfMissionsSummary = pd.DataFrame()
                    dfMissions = pd.DataFrame()
                    dfBook = pd.DataFrame()

                    dfBook = pd.read_excel(inputRoutesFile, 
                            sheet_name = None)
                    
                    dfMissionsSummary = dfBook['MISSIONS']
                    # Removing unnamed columns using drop function
                    dfMissionsSummary.drop(dfMissionsSummary.columns[dfMissionsSummary.columns.str.contains(
                        'unnamed', case=False)], axis=1, inplace=True)
                    
                    dfMissions = dfBook['ROUTES']
                    # Removing unnamed columns using drop function
                    dfMissions.drop(dfMissions.columns[dfMissions.columns.str.contains(
                        'unnamed', case=False)], axis=1, inplace=True)

                except:
                    print(colorama.Fore.YELLOW + "Can´t read routes file for " + nidEngine + colorama.Style.RESET_ALL)

                # Splitting missions
                print('Splitting missions for ' + nidEngine + '. Adding charts')
                dfSingleMission = ie.R_SplitMission (dfMissions, dfMissionsSummary, nidEngine, outputPath, inputFileNameNoextension)
                    
            
                
                
        t1 = date.datetime.now()
        delta = t1-t0
        delta = delta.total_seconds()
        print(colorama.Fore.GREEN + "Chart phase ends" + '. Partial time: ' + str(delta) + ' secs' + colorama.Style.RESET_ALL )
        
    else:
        print(colorama.Fore.GREEN + "Jumping chart phase" + colorama.Style.RESET_ALL)     

    endTime = date.datetime.now()
    delta = endTime - startTime
    delta = delta.total_seconds()
    print("File process time: " + str(delta) + ' secs')

# Aggregate statistics
for n in range(len(arrDfAntennaDistances)):
        
    dfAntennaDistances = arrDfAntennaDistances[n] 
    dfAggregatedAntennaDistances = pd.concat([dfAggregatedAntennaDistances, dfAntennaDistances], ignore_index = True)

# Save statistics file
pathCalibratedAntdistFile = pathlib.Path.joinpath(statisticsPath, fileStatistics)
outputCalibratedAntdistFile = str(pathCalibratedAntdistFile)    
print("Saving statistics file. " + outputCalibratedAntdistFile)
# Create output folder if not exists
pathlib.Path(statisticsPath).mkdir(parents=True, exist_ok=True)
# Excel write (end of Phase 2)
try:
    with pd.ExcelWriter(outputCalibratedAntdistFile) as writer:
        wb = writer.book
        dfAggregatedAntennaDistances.to_excel(writer, sheet_name='STATISTICS')
                            
except Exception as ex:
    print(colorama.Fore.YELLOW + "Can´t save statistics file. Error: " +str(ex) + colorama.Style.RESET_ALL)    

endTotalTime = date.datetime.now()
delta = endTotalTime - startTotalTime
delta = delta.total_seconds()
print("Total process time: " + str(delta) + ' secs')
