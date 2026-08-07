import pandas
import numpy
import re
import sys
import colorama as color
import datetime
import xlsxwriter as xls
import scipy
import openpyxl
import warnings
import pathlib
import copy

import config as cfg
import iecfg as iecfg

# Version control
# Author: Manuel Cáceres Marzal
# 
# V2.0 Date 28/11/25
#   - New version. Splitting code from compilation to allow multiple formats compilation
#   - Functions for JRU compilation are removed from this module (ietcs) and added to specific ones for each jru format
#   - Bugs to be fixed:
#       - Take account unknown balises to figure out accum distance
#       - Reduce 5.000 m accumulated error with initial error
#   - Changes:
#       - Include Level 1
#       - Take into account Q_DIR from packets received form TS
# V 2.1 Date 06/02/2026
#   - Update version number to make it consistent across all modules
#
# V 3.0 Date 06/02/2026
#   - New functions in other modules
#
# V 3.1 Date 03/03/2026
#   - Fixed bugs in run distance meassurement
#   - Fixed bugs with SSP calculations
#   - Change comparisson of position error with over/under estimation (to check expected window for balise)
#   - Included in charts: EBC, SBC, START, DMIACK, DRIVERACK
#   - Included in charts: Safe connection established/interrupted
#
#
# V 3.1.1 Date 20/07/2026
#   - Fixed error in D_LINK_0:
#       Correction: included next lines in code
#       if len(dfTrackDistances.axes[0]) > 0:
#            dfTrackDistances = dfTrackDistances.sort_values(by="INIT_INDEX")
#            dfTrackDistances.reset_index(drop=True, inplace=True)
#
#       in JruC_SetLinkedRoutes()
# 
# V 4.0 Date 06/08/2026
# 
#   - Updated for pythin 3.14
# 
#  
# Required BUGs/Improvements to be fixed:
#
#   1) Bug: SSP calculation 
#   2) FIXED (20/07/26) - Review: Review D_LINK_0 column in cases where a balise error occurs and LRBG not coherent with D_LINK_0
#   3) Improvement: Include default QLOCACC at route start
#   4) Improvement: From calibration, get slope (m) of calibrated linear regression (y = mx + n)
#   5) Improvement: Recalculate usafe amd osafe statistics taking into account QLOCACC of final balise
#

ietcsVersion = '4.0'

# Global variables

# Set of global variables to figure out progress
glPh2Progress = 0.0
glPh2NumberRows = 0
glPh2Percentage = 0.0
glPh2Prevpercentage = 0.0
glPh2Functions = 0
glPh2FunctionNum = 0
glPh3Progress = 0.0
glPh3NumberMissionsPerc = 0.0
glPh3NumberRows = 0

# SS26 variables
glOdoAccuracyImpairmentTh = 250
glOdoAccuracySafetyTh = 1500
glSlopeSS41Th = 0.05 # SS041 over/under estimation threshold = 5 + 5%s
glOdoAccuracyDistance = 5000
glSS41Intercept = 5

# SetRoute process variables:
glCalibrationProcess = 'Calibration'
glEstimationProcess = 'Estimation'

# set of global variables to change process
# !!! Change; nonsense to define as global
glboolFilterUnknownBalises = cfg.glboolFilterUnknownBalises



# global variable glLocale for content format
# Values
# ENG for english number formats and texts
#   , for thousands
#   . for decimals
# SPA for spanish number format and texts
#   . for thousands
#   , for decimals

glLocale = cfg.glLocale

# -----------------------------------  JruC_ Functions    -------------------------------------------
# Functions whose inpput file is the normalized compiled JRU file (JruC_)
# Modules:
# COM 'common'
# naming: JruC_ModuleFunctionName
# JruC_ = Compiled JRU File

# JruC_ComChangeDtypes(JruC_Df)
# Compiled JRU File
# Set right dtypes

def JruC_ComChangeDtypes(JruC_Df):

    global glPh2Progress
    global glPh2NumberRows
    global glPh2Percentage
    global glPh2Prevpercentage
    
    glPh2Progress = 0.0
    glPh2NumberRows = 0
    glPh2Percentage = 0.0
    glPh2Prevpercentage = 0.0

    # We change original DataFrame, not a copy
    # change df['PACKET_RBC'] to int and then to str
    glPh2NumberRows = len(JruC_Df.axes[0])
    glPh2Percentage = 0.20
    glPh2Prevpercentage = 0.20
    
    serPacketRbc = pandas.Series(JruC_Df['PACKET_RBC'].apply(utl_ChangeTypeStrFromInt))
    #for x in range(len(JruC_Df["PACKET_RBC"])):
    #    try:
    #        serPacketRbc[x] = str(int(JruC_Df["PACKET_RBC"][x]))
    #    except:
    #        serPacketRbc[x] = ""
    #        glPh2Progress = glPh2Prevpercentage* 100 + glPh2Percentage*(x+1)/glPh2NumberRows*100
    #        sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
    #        sys.stdout.flush()
    #        continue
        
    glPh2Progress = glPh2Prevpercentage* 100 + glPh2Percentage #*(x+1)/glPh2NumberRows*100
    sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
    sys.stdout.flush()
    
    JruC_Df["PACKET_RBC"] = serPacketRbc  
    
         
    # change df['NID_ENGINE'] to int and then to str
    glPh2Prevpercentage += glPh2Percentage
    glPh2Percentage = 0.20    
       
    

    serNidengine = pandas.Series(JruC_Df['NID_ENGINE'].apply(utl_ChangeTypeStrFromInt) )
    #for x in range(len(JruC_Df["NID_ENGINE"])):
    #    try:
    #        serNidengine[x] = str(int(JruC_Df["NID_ENGINE"][x]))
    #    except:
    #        serNidengine[x] = ""

    glPh2Progress = glPh2Prevpercentage* 100 + glPh2Percentage #*(x+1)/glPh2NumberRows*100
    sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
    sys.stdout.flush()
        

    JruC_Df["NID_ENGINE"] = serNidengine  

    # change df['NID_LRBG'] to int and then to str
    glPh2Prevpercentage += glPh2Percentage
    glPh2Percentage = 0.20 
    
    serNidlrbg = pandas.Series(JruC_Df['NID_LRBG'].apply(utl_ChangeTypeStrFromInt))
    #for x in range(len(JruC_Df["NID_LRBG"])):
    #    try:
    #        serNidlrbg[x] = str(int(JruC_Df["NID_LRBG"][x]))
    #    except:
    #        serNidlrbg[x] = ""
            
    glPh2Progress = glPh2Prevpercentage* 100 + glPh2Percentage #*(x+1)/glPh2NumberRows*100
    sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
    sys.stdout.flush()
        

    JruC_Df["NID_LRBG"] = serNidlrbg  

    # format df['Date'] 
    glPh2Prevpercentage += glPh2Percentage
    glPh2Percentage = 0.20 
    serDate = pandas.Series(JruC_Df['Date'].apply(utl_ChangeDateFormat))
    #for x in range(len(JruC_Df["Date"])):
    #    try:
    #        strDate = str(JruC_Df['Date'][x])
    #        i = strDate.index('/')
    #        dd = strDate[0:i].rjust(2, '0')
    #        strDate = strDate[i+1:len(strDate)]
    #        i = strDate.index('/')
    #        mm = strDate[0:i].rjust(2, '0')
    #        yy = strDate[i+1:len(strDate)]
    #        strDate = dd + '/' + mm + '/' + yy
    #        serDate[x] = strDate
    #    except:
    #        serDate[x] = ""
            
    glPh2Progress = glPh2Prevpercentage* 100 + glPh2Percentage #*(x+1)/glPh2NumberRows*100
    sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
    sys.stdout.flush()
        

    JruC_Df["Date"] = serDate  

    # format df['Time'] 
    glPh2Prevpercentage += glPh2Percentage
    glPh2Percentage = 0.20 

    serTime = pandas.Series(JruC_Df['Time'].apply(utl_ChangeTimeFormat))
    #for x in range(len(JruC_Df['Time'])):
    #    try:
    #        strTime = str(JruC_Df['Time'][x])
    #        i = strTime.index(':')
    #        hour = strTime[0:i].rjust(2, '0')
    #        strTime = strTime[i+1:len(strTime)]
    #        i = strTime.index(':')
    #        minutes = strTime[0:i].rjust(2, '0')
    #        strTime = strTime[i+1:len(strTime)]
    #        try:
    #            i = strTime.index('.') # case seconds with decimals
    #            seconds = strTime[0:i].rjust(2, '0')
    #        except:
    #            seconds = strTime[0:len(strTime)].rjust(2, '0')
    #        strTime = hour + ':' + minutes + ':' + seconds # H:m:s
    #        serTime[x] = strTime
    #    except:
    #        serTime[x] = ""
            
    glPh2Progress = glPh2Prevpercentage* 100 + glPh2Percentage #*(x+1)/glPh2NumberRows*100
    sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
    sys.stdout.flush()
        

    JruC_Df["Time"] = serTime  


    sys.stdout.write("\n")
    sys.stdout.flush()
    
    return JruC_Df


# JruC_SetRoutes(JruC_Df)
# Compiled JRU File
# Get missions info: indexes SoM / EoM, route distance, ...
# distAntenna: if None, means the function must estimate
# Return a DF with a summary of missions and a second DF with detailed route info
def JruC_SetRoutes(boolProcess, JruC_Df, nidengine, distAntenna, dfDigitalMap):

    global glLocale

    global glPh2Progress
    global glPh2NumberRows
    global glPh2Percentage
    global glPh2Prevpercentage
    global glboolFilterUnknownBalises
    global glCalibrationProcess
    global glEstimationProcess
    
    glPh2Progress = 0.0
    glPh2NumberRows = 0
    glPh2Percentage = 0.0
    glPh2Prevpercentage = 0.0
    
    
    if boolProcess == glCalibrationProcess:
        pepe = 0
    elif boolProcess == glEstimationProcess:
        pepe = 0
    else:
        print(color.Fore.RED + 'Unknown process: '  +   boolProcess + 
                                    color.Style.RESET_ALL)  
        exit()

    #inputdf = pandas.DataFrame(JruC_Df) # local copy 
    inputdf = JruC_Df.copy() # local copy 

    # filter nidengine
    inputdf = inputdf[(inputdf['NID_ENGINE'] == nidengine)] 
    # reindex inputdf to avoid gaps in index
    inputdf.index = numpy.arange(0, len(inputdf["NID_ENGINE"]))
    
    # filter  messages:
    # MESSAGE FROM BALISE with NID_BG not null
    serFilterNidBG = inputdf["NID_BG"].str.contains('-', na=False, regex=False)
    serFilterMrsp = inputdf["V_PERM"].isnull()
    serFilterMrsp = ~serFilterMrsp
    serFilterRbcPck = pandas.Series()
    # PACKET RBC FILTER
    # Note: take the same filter as for compiled file
    # In case you want to make an extra rbc packets filtering use listFilterRbcPck below
    listFilterRbcPck = [] #numpy.array(['136', '157', '150', '3', '33'])
    for strPck in listFilterRbcPck:
        if serFilterRbcPck.size == 0:
            serFilterRbcPck = inputdf["PACKET_RBC"].str.fullmatch(strPck, na=False)
        else:
            serFilterRbcPck |= inputdf["PACKET_RBC"].str.fullmatch(strPck, na=False)
    if serFilterRbcPck.size == 0:
        serFilterRbcPck = inputdf["PACKET_RBC"].isnull()
        serFilterRbcPck = ~serFilterRbcPck
    # Remove unknown balise records
    # except those coming from a balise message (relevant info)
    serFilterUnknownBalise = inputdf["NID_LRBG"].str.contains('16383', na=False, regex=False)
    serFilterUnknownBalise = ~serFilterUnknownBalise
    serFilterUnknownBalise |= serFilterNidBG  # Ojo los telegramas de balizas con unknown LRBG se mantienen
    
    # Filter input dataframe
    serFilter = serFilterNidBG | serFilterMrsp | serFilterRbcPck
    # To filter/not-filter unknown LRBG (= 16383)
    if glboolFilterUnknownBalises == True:
        serFilter &= serFilterUnknownBalise
    inputdf = inputdf[serFilter]
    # reindex inputdf to avoid gaps in index
    inputdf.index = numpy.arange(0, len(inputdf["NID_ENGINE"]))

    # Replace D_LRBG values at unknown LRBG with 0 value to avoid wrong distances
    for x in range (len(inputdf.axes[0])):
        if (inputdf['NID_LRBG'][x] == '16383'): 
            inputdf.iloc[x, inputdf.columns.get_loc('D_LRBG')] = 0
           

    

    # ---------- Adding new columns   ------------------
    # Add a new column to trace original index (after reindex)
    initialarrayIndex = numpy.arange(0, len(inputdf["NID_ENGINE"]))
    inputdf.insert(0, "INIT_INDEX", initialarrayIndex)

    emptyList = []
    for x in range (len(inputdf.axes[0])):
        emptyList.insert(x, "")

    emptyFloatList = []
    for x in range (len(inputdf.axes[0])):
        emptyFloatList.insert(x, numpy.nan)


    # Add a new column for calculated DLRBG bases on current Balise 
    inputdf.insert(inputdf.columns.get_loc('D_LRBG')+1, 'SELF_DLRBG_E', emptyFloatList)

    # Add a new column for accumulated route distances
    inputdf.insert(inputdf.columns.get_loc('SELF_DLRBG_E')+1, "D_ONB_ACCUM-0", emptyFloatList)

    # Add a new column for accumulated route distances
    inputdf.insert(inputdf.columns.get_loc('D_ONB_ACCUM-0')+1, 'D_ONB_ACCUM', emptyFloatList)
    
    # Add a new column for extended NID_LRBG
    inputdf.insert(inputdf.columns.get_loc('NID_LRBG')+1, "LRBG_Ext", emptyList)

    # Add a new column for accumulated linked distances
    inputdf.insert(inputdf.columns.get_loc('D_ONB_ACCUM')+1, 'D_LINK_0', emptyFloatList)

    # Add a new column for accumulated linked distances
    inputdf.insert(inputdf.columns.get_loc('D_LINK_0')+1, 'D_LINK_ACCUM', emptyFloatList)

    # Add a new column for Q_LOCACC
    inputdf.insert(inputdf.columns.get_loc('D_LINK_ACCUM')+1, 'Q_LOCACC', emptyFloatList)

    # Add a new column for absolute odometry error
    inputdf.insert(inputdf.columns.get_loc('Q_LOCACC')+1, "LOC_ERROR", emptyFloatList)

    # Add a new column for absolute odometry accuracy (overestimation)
    inputdf.insert(inputdf.columns.get_loc('LOC_ERROR')+1, "OVER_ACCURACY", emptyFloatList)

    # Add a new column for absolute odometry accuracy (underestimation)
    inputdf.insert(inputdf.columns.get_loc('OVER_ACCURACY')+1, "UNDER_ACCURACY", emptyFloatList)

    # Add a new column for relative over estimation error (%)
    inputdf.insert(inputdf.columns.get_loc('UNDER_ACCURACY')+1, "OVEREST_ERROR", emptyFloatList)

    # Add a new column for relative under estimation error (%)
    inputdf.insert(inputdf.columns.get_loc('OVEREST_ERROR')+1, "UNDEREST_ERROR", emptyFloatList)

    # Add a new column for percentage odometry error
    inputdf.insert(inputdf.columns.get_loc('UNDEREST_ERROR')+1, "ODOERROR_PERC", emptyFloatList)

    # Add a new column for mission time
    inputdf.insert(inputdf.columns.get_loc('ODOERROR_PERC')+1, "T_ONB_ACCUM", emptyFloatList)
    
    # Add a new column for V_STATIC
    inputdf.insert(inputdf.columns.get_loc('V_TRAIN')+1, "V_STATIC", emptyFloatList)

    # Add a new column for V_STATIC
    inputdf.insert(inputdf.columns.get_loc('V_STATIC')+1, 'ACCEL', emptyFloatList)

    # Add a new column for SS41 max accuracy
    inputdf.insert(inputdf.columns.get_loc('L_DOUBTUNDER')+1, "SS41_MIN", emptyFloatList)

    # Add a new column for SS41 max accuracy
    inputdf.insert(inputdf.columns.get_loc('SS41_MIN')+1, "SS41_MAX", emptyFloatList)

    # Add a new column for overestimation slope
    inputdf.insert(inputdf.columns.get_loc('L_DOUBTOVER')+1, 'OVERSLOPE', emptyFloatList)
    
    # Add a new column for underestimation slope
    inputdf.insert(inputdf.columns.get_loc('L_DOUBTUNDER')+1, 'UNDERSLOPE', emptyFloatList)   

    # Add a new column for over safe error (odo_error outside L_DUBTOVER CI)
    inputdf.insert(inputdf.columns.get_loc('OVERSLOPE')+1, 'OSAFE_ERROR', emptyFloatList)
    
    # Add a new column for under safe error (odo_error outside L_DUBTUNDER CI)
    inputdf.insert(inputdf.columns.get_loc('UNDERSLOPE')+1, 'USAFE_ERROR', emptyFloatList)

    # Add a new column for under DLINKERR
    inputdf.insert(inputdf.columns.get_loc('USAFE_ERROR')+1, 'DLINKERR', emptyFloatList)

    # Add a new column for under ODOERR (relative to distance)
    inputdf.insert(inputdf.columns.get_loc('DLINKERR')+1, 'ODOERR', emptyFloatList)

    # Add a new column for spot overestimation slope
    inputdf.insert(inputdf.columns.get_loc('ODOERR')+1, 'SPOTORSLOPE', emptyFloatList)

    # Add a new column for spot underestimation slope
    inputdf.insert(inputdf.columns.get_loc('SPOTORSLOPE')+1, 'SPOTURSLOPE', emptyFloatList)

    # Add new columns for ERRORS
    inputdf.insert(len(inputdf.columns), "ERR_TYPE", emptyList)
    inputdf.insert(len(inputdf.columns), "ERR_MESSAGE", emptyList)

    # Add new columns for DIGITAL MAP
    inputdf.insert(len(inputdf.columns), "SIGNAL", emptyList)
    inputdf.insert(len(inputdf.columns), "PAN", emptyList)
    

    # Lists for output dataframes
    # new dataframes for output: summary of routes
    dfMissionsSummary = pandas.DataFrame() 
    listNidEngine = [] # To insert in outputdf: list of nid_engines in missions
    listSoM = [] # To insert in outputdf: list of index for SoM rows
    listEoM = [] # To insert in outputdf: list of index for EoM rows
    listMissionCompleted = [] # To insert in outputdf: list to indicate if Mission is ended with EoM message or not (yes/no)
    listDmission = [] # To insert in outputdf: accumulated total distance for the mission
    listOdoerrMission = []
    listNumBG = [] # To insert in outputdf: number of BG within mission
    listNumLRBG = [] # To inesrt in outputdf: number of different reference balises in mission
    listModeSh = [] # To insert in outputdf: indicates if there wer shunting movements during the mission
    listDate = [] # To insert in outputdf: date for each start of mission
    listTime = [] # To insert in outputdf: time for each start of mission
    listMissionTime = [] # To inserty in outputdf: total time of mission
    listMissionDistAntenna = [] # To insert in outputdf: antenna distance to fe used (input data or estimated)
  
  
    # --------------------------- Preparing dataframes for  missions -------------------------
    glPh2Functions = 8
    glPh2FunctionNum = 0

    # Loop for each SoM/EoM section (each mission)
    # filter RBC packet 157 for SoM
    # filter RBC packet 150 for EoM
    
    dfSoM = inputdf[(inputdf['PACKET_RBC'] == '157')]
    arrSoMIndexes = []
    arrEoMIndexes = []
    # Case there are no SoM in file search for first MA and consider it as SoM
    if(len(dfSoM.axes[0]) == 0):
        # search for first MA (packets 3 or 33)
        dfSoM = inputdf[(inputdf['PACKET_RBC'] == '3') | (inputdf['PACKET_RBC'] == '33')]
        if(len(dfSoM.axes[0]) > 0):
            dfSoM.index = numpy.arange(0, len(dfSoM.axes[0]))
            index = dfSoM['INIT_INDEX'][0]
            arrSoMIndexes= numpy.array([index]) # list of SoM messages
    else:
        #if len(dfSoM.axes[0]) > 0:
            arrSoMIndexes= numpy.array(dfSoM.index) # list of SoM messages
        #else:
        #    dfSoM.index = numpy.arange(0, len(dfSoM.axes[0]))
        #    index = dfSoM['INIT_INDEX'][0]
        #    arrSoMIndexes= numpy.array([index]) # list of SoM messages

    dfEoM = inputdf[(inputdf['PACKET_RBC'] == '150')]
    if len(dfEoM.axes[0]) > 0:
        arrEoMIndexes= numpy.array(dfEoM.index) # list of EoM messages
    #else:
    #        dfEoM.index = numpy.arange(0, len(dfEoM.axes[0]))
    #        index = dfEoM['INIT_INDEX'][0]
    #        arrEoMIndexes= numpy.array([index]) # list of EoM messages
    
    # Case first SoM/EoM record found is an EoM
    # In this case we consider first MA previous to this first EoM as SoM
    if (len(dfEoM.axes[0]) > 0) & (len(dfSoM.axes[0]) > 0):
        firstIxEoM = dfEoM.iloc[0, dfEoM.columns.get_loc('INIT_INDEX')]
        firstIxSoM = dfSoM.iloc[0, dfSoM.columns.get_loc('INIT_INDEX')]
        if firstIxSoM > firstIxEoM:
            # search for first MA (packets 3 or 33)
            dfSoM = inputdf[(inputdf['PACKET_RBC'] == '3') | (inputdf['PACKET_RBC'] == '33')]
            if(len(dfSoM.axes[0]) > 0):
                dfSoM.index = numpy.arange(0, len(dfSoM.axes[0]))
                index = dfSoM['INIT_INDEX'][0]
                if index < firstIxEoM:
                    arrSoMIndexes = numpy.insert(arrSoMIndexes, 0, index)


    x = 0 # initialize x for next loop
    if len(arrSoMIndexes) > 0:
        glPh2NumberRows = len(arrSoMIndexes)
        glPh2Percentage = 0.2
        glPh2Prevpercentage = 0.0
        for x in range (len(arrSoMIndexes) -1 ): # special treatment for last mission later
            
            # index for SoM message
            listSoM.insert(x, arrSoMIndexes[x])
            # index for EoM message
            listIndexEoM = numpy.where((arrEoMIndexes > arrSoMIndexes[x]) & (arrEoMIndexes < arrSoMIndexes[x+1]), arrEoMIndexes, None)
            indexEoM = None
            # Search for index of EoM message after SoM message
            for i in range(len(listIndexEoM)):
                if(listIndexEoM[i] != None): 
                    indexEoM = listIndexEoM[i]
                    break
            
            if (indexEoM != None):
                # EoM found
                listEoM.insert(x, indexEoM)
                listMissionCompleted.insert(x, 'yes')
            else:
                # EoM not found; we use last message before next SoM as end of mission
                listEoM.insert(x, arrSoMIndexes[x+1]-1) # row before SoM packet (see warning above)
                listMissionCompleted.insert(x, 'no')

            glPh2Prevpercentage += glPh2Percentage/glPh2NumberRows    
            glPh2Progress = glPh2Prevpercentage* 100 
        
            sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
            sys.stdout.flush()

        # last Mission: special treatment in case there is no EoM message to end mission
        if x > 0 : x += 1
        listSoM.insert(len(listSoM), arrSoMIndexes[len(arrSoMIndexes)-1])
        # index for EoM message
        listIndexEoM = numpy.where(arrEoMIndexes > arrSoMIndexes[len(arrSoMIndexes)-1], arrEoMIndexes, None)
        indexEoM = None
        # Search for index of EoM message after SoM message
        for i in range(len(listIndexEoM)):
            if(listIndexEoM[i] != None): 
                indexEoM = listIndexEoM[i]
                break
        
        if (indexEoM != None):
            # last EoM found
            listEoM.insert(len(listEoM), indexEoM)
            listMissionCompleted.insert(x, 'yes') 
        else:
            # EoM not found: last message is treated as EoM
            serNidengine = inputdf[(inputdf['NID_ENGINE'] == nidengine)]
            listEoM.insert(len(listEoM), len(serNidengine)-1)
            listMissionCompleted.insert(x, 'no')

        # Change dtype of listEoM to numpy.int64 to homogenize with listSoM
        listEoM = [numpy.int64(x) for x in listEoM]

        glPh2Prevpercentage += glPh2Percentage/glPh2NumberRows    
        glPh2Progress = glPh2Prevpercentage* 100
        
        sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
        sys.stdout.flush()
        
        # Here we have matched SoM - EoM records

        # Set accumulated distance in each mission
        # Set accumulated time in each mission
        dfSingleMission = pandas.DataFrame()
        dfNidlrbgInMission = pandas.DataFrame()
        serExtendedNidlrbg = pandas.Series()
        
        # --------------------------- For each mission --------------------------------------------
        
        numBG = 0
        numLRBG = 0
        glPh2Prevpercentage = glPh2Percentage
        glPh2Percentage = 0.8
        glPh2NumberRows = len(listSoM)
        for x in range (len(listSoM)):

            # Filter by each mission
            dfSingleMission = inputdf.iloc[listSoM[x]:listEoM[x]+1]
            # reindex is necessary for loop
            dfSingleMission.index = numpy.arange(0, len(dfSingleMission["NID_LRBG"]))

            # --------- Block for getting additional info for Mission summary -------------

            # Note: add here code to get list of nid_bg overpassed by train
            # for future routes comparisson
            serBalisesInMission = dfSingleMission[dfSingleMission["NID_BG"].str.contains('-0', na = False)]['NID_BG']
            #.apply utl_remove_balise extension
            # We get number of over passed balises in mission; take into account only first balise in group (-0)
            numBG = len(dfSingleMission[dfSingleMission["NID_BG"].str.contains('-0', na = False)]['NID_BG'])
            # We get number reference balises in mission; filter 16383 id balises (unknown)
            numLRBG = len(dfSingleMission[(dfSingleMission["NID_LRBG"] != '16383')]['NID_LRBG'].unique())
            listNumBG.insert(x, numBG)
            listNumLRBG.insert(x,numLRBG)
            
            # Find if the mission includes mode shunting (M_MODE = 3)
            if len(dfSingleMission[dfSingleMission["M_MODE"] == 3]['M_MODE']) > 0:
                listModeSh.insert(x, 'yes')
            else:
                listModeSh.insert(x, 'no')

            # Get nidengine
            listNidEngine.insert(x, nidengine)
            
            # Get date and time
            strDate = str(dfSingleMission["Date"][0]) 
            #ix = strDate.index('/')
            #dd = strDate[0:ix].rjust(2, '0')
            #strDate = strDate[ix+1:len(strDate)]
            #ix = strDate.index('/')
            #mm = strDate[0:ix].rjust(2, '0')
            #yy = strDate[ix+1:len(strDate)]
            #strDate = dd + '/' + mm + '/' + yy # format dd/mm/yy
            strTime = dfSingleMission["Time"][0] # format hh:mm:ss
            listDate.insert(x, strDate)
            listTime.insert(x, strTime)
            
            # To identify orientation changes in reference balises: add an extension in nid_lrbg                      
            # To figure out distance we need to identify every change in orientation of f.e. 
            # for each balise in the mission and all the changes in LRBG
            # To do it we define a temporal Series including the balise id + the sequence of appearance in the mission
            # i.e. balise 1234 is 1234-1 and 1234-2 for next orientation change in the same mission
            nidlrbgPrev = None
            nidlrbgCurr = None
            nidBgPrev = None
            nidBgCurr =  None
            ext = 1
            # Add extension to NID_LRBG
            dfNidlrbgInMission = dfSingleMission[['NID_LRBG', 'Q_DLRBG']]
            # reindex is necessary for loop
            dfNidlrbgInMission.index = numpy.arange(0, len(dfNidlrbgInMission.axes[0]))
            
            
            # Add '-seq' to LRBG id
            # Warning: take into account same balise change of direction (change of Q_DLRBG)
            for j in range (len(dfSingleMission['NID_LRBG'])):

                # Get 
                strAux = str(dfSingleMission['NID_BG'][j]) 
                if strAux != '' and strAux != 'nan':
                    nidBgPrev = nidBgCurr
                    tempNidBg = dfSingleMission['NID_BG'][j]
                    try:
                        iAux = tempNidBg.index('-')
                        tempNidBg = tempNidBg[0:iAux]
                        nidBgCurr = tempNidBg
                    except:
                        nidBgCurr = dfSingleMission['NID_BG'][j]
                else:
                    nidBgPrev = nidBgCurr


                if j > 0:
                    # increase extension for nidlrbg
                    nidlrbgCurr = dfNidlrbgInMission['NID_LRBG'][j]
                    qdlrbgCurr = dfNidlrbgInMission['Q_DLRBG'][j]
                    nidlrbgPrev = dfNidlrbgInMission['NID_LRBG'][j-1]
                    qdlrbgPrev = dfNidlrbgInMission['Q_DLRBG'][j-1]
                    
                    if nidlrbgCurr != nidlrbgPrev: 
                        ext += 1
                    elif qdlrbgCurr != qdlrbgPrev:
                        ext += 1
                    #elif nidBgPrev != nidBgCurr:
                    #    ext += 1
                            
                serExtendedNidlrbg[j] = dfNidlrbgInMission['NID_LRBG'][j] + '-' + str(ext)

            for m in range (len(dfSingleMission['LRBG_Ext'])):
                pos = dfSingleMission['INIT_INDEX'][m]
                dfSingleMission.iloc[m, dfSingleMission.columns.get_loc('LRBG_Ext')] = serExtendedNidlrbg[m]
                inputdf.iloc[pos, inputdf.columns.get_loc('LRBG_Ext')] = serExtendedNidlrbg[m]             

            # 1) Calculating distances first stage: Set accumulated distance as estimated on OB
            # It needs a later calibration with linked info
            glPh2Percentage = 0.10
            glPh2Prevpercentage += glPh2Percentage/glPh2NumberRows

            listReturn = []
            listReturn = JruC_SetObestimatedRoutes(dfSingleMission, inputdf, distAntenna) 
            dist = listReturn[0]
            # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function
            dfSingleMission = pandas.DataFrame(listReturn[1])
            


            glPh2Progress = glPh2Prevpercentage* 100 
            sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
            sys.stdout.flush()                    
            
            # 2) Calculating distances second stage: Set accumulated distance as configured on linking info
            # starting distance point based on point 1, as estimated OB
            glPh2Percentage = 0.10
            glPh2Prevpercentage += glPh2Percentage/glPh2NumberRows

            # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function
            dfSingleMission = pandas.DataFrame(JruC_SetLinkedRoutes(dfSingleMission, inputdf, distAntenna))
            
            
                
            glPh2Progress = glPh2Prevpercentage* 100 
            sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
            sys.stdout.flush()
                
            # 3) Calibrate estimated route distances (1) based on linking info (2) at balise overpassing
            # Estimation of odometry error
            glPh2Percentage = 0.10
            glPh2Prevpercentage += glPh2Percentage/glPh2NumberRows

            listReturn = []            
            listReturn = JruC_SetCalibratedRoutes(dfSingleMission, inputdf, distAntenna)

            dist = listReturn[0]
            odoerr = listReturn[1]
            mtime = listReturn[2]
            dfSingleMission = pandas.DataFrame(listReturn[3])
            
            # info for missions summary
            listDmission.insert(x, dist)
            listOdoerrMission.insert(x, odoerr)
            listMissionTime.insert(x, mtime)
            listMissionDistAntenna.insert(x, distAntenna)

            glPh2Progress = glPh2Prevpercentage* 100 
            sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
            sys.stdout.flush()

            # 4) Set SSP info for each mission
            glPh2Percentage = 0.10
            glPh2Prevpercentage += glPh2Percentage/glPh2NumberRows


            # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function
            dfSingleMission = pandas.DataFrame(JruC_SetSSP(dfSingleMission, inputdf))
                
            glPh2Progress = glPh2Prevpercentage* 100 
            sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
            sys.stdout.flush()
                
            # 5) Set VMRSP/V_PERM info for each mission
            glPh2Percentage = 0.10
            glPh2Prevpercentage += glPh2Percentage/glPh2NumberRows

            # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function
            dfSingleMission = pandas.DataFrame(JruC_SetVmrsp(dfSingleMission, inputdf))
                
            glPh2Progress = glPh2Prevpercentage* 100 
            sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
            sys.stdout.flush()

            # 6) Set Acceleration info for each mission
            glPh2Percentage = 0.10
            glPh2Prevpercentage += glPh2Percentage/glPh2NumberRows

            # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function
            dfSingleMission = pandas.DataFrame(JruC_SetAccel(dfSingleMission, inputdf))
            
            glPh2Progress = glPh2Prevpercentage* 100
            sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
            sys.stdout.flush()

            # 7) Set OVER/UNDER slope
            glPh2Percentage = 0.10
            glPh2Prevpercentage += glPh2Percentage/glPh2NumberRows

            # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function
            dfSingleMission = pandas.DataFrame(JruC_SetSlope(dfSingleMission, inputdf))
            
             # 8) Set Digital Map info
            glPh2Percentage = 0.10
            glPh2Prevpercentage += glPh2Percentage/glPh2NumberRows

            # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function
            dfSingleMission = pandas.DataFrame(JruC_SetDigitalMap(dfSingleMission, inputdf, dfDigitalMap))
            
            glPh2Progress = glPh2Prevpercentage* 100
            sys.stdout.write("Progress: %d%%   \r" % (glPh2Progress) )
            sys.stdout.flush()
        
        # ------------------- Add columns with V_TRAIN per M_MODE (starts) ------------------------
        # 0 - Full Supervision
        # 1 - On Sight
        # 2 - Staff Responsible
        # 3 - Shunting
        # 4 - Unfitted
        # 6 - Stand By
        # 11 - Non leading
        # 12 - Limited Supervision
        # 13 - National System
        # Others
            
        refColumn = "V_TRAIN"
        
        # New column for distances where in FS mode
        # to differnetiate modes in excel graphics
        maskFS = inputdf['M_MODE'].eq(0, fill_value = False) # mask for FS mode
        # Knime dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'], 'V_FS' : emptyList})
        dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'].copy(), 'V_FS' : emptyFloatList})
        dfVel = dfVel[maskFS]
        dfVel.index = numpy.arange(0, len(dfVel["INIT_INDEX"]))
        if len(dfVel.axes[0]) >0:
            # Knime inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_FS", emptyList)
            inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_FS", emptyFloatList)
            refColumn = "V_FS"
            for x in range(len(dfVel.axes[0])):
                try:
                    pos = dfVel['INIT_INDEX'][x]
                    inputdf.iloc[pos, inputdf.columns.get_loc('V_FS')] = inputdf["V_TRAIN"][pos]
                except Exception as ex:
                    print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)
                    continue

        # New column for distances where in OS mode
        # to differnetiate modes in excel graphics
        maskOS = inputdf['M_MODE'].eq(1, fill_value = False) # mask for OS mode
        # knime dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'], 'V_OS' : emptyList})
        dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'].copy(), 'V_OS' : emptyFloatList})
        dfVel = dfVel[maskOS]
        dfVel.index = numpy.arange(0, len(dfVel["INIT_INDEX"]))
        if len(dfVel.axes[0]) >0:
            # knime inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_OS", emptyList)
            inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_OS", emptyFloatList)
            refColumn = "V_OS"
            for x in range(len(dfVel.axes[0])):
                try:
                    pos = dfVel['INIT_INDEX'][x]
                    inputdf.iloc[pos, inputdf.columns.get_loc('V_OS')] = inputdf["V_TRAIN"][pos]
                except Exception as ex:
                    print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)
                    continue

        # New column for distances where in SR mode
        # to differnetiate modes in excel graphics
        maskSR = inputdf['M_MODE'].eq(2, fill_value = False) # mask for SR mode
        # knime dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'], 'V_SR' : emptyList})
        dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'].copy(), 'V_SR' : emptyFloatList})
        dfVel = dfVel[maskSR]
        dfVel.index = numpy.arange(0, len(dfVel["INIT_INDEX"]))
        if len(dfVel.axes[0]) >0:
            # knime inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_SR", emptyList)
            inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_SR", emptyFloatList)
            refColumn = "V_SR"
            for x in range(len(dfVel.axes[0])):
                try:
                    pos = dfVel['INIT_INDEX'][x]
                    inputdf.iloc[pos, inputdf.columns.get_loc('V_SR')] = inputdf["V_TRAIN"][pos]
                except Exception as ex:
                    print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)
                    continue

        
        # New column for distances where in SH mode
        # to differnetiate modes in excel graphics
        maskSH = inputdf['M_MODE'].eq(3, fill_value = False) # mask for SH mode
        # knimedfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'], 'V_SH' : emptyList})
        dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'].copy(), 'V_SH' : emptyFloatList})
        dfVel = dfVel[maskSH]
        dfVel.index = numpy.arange(0, len(dfVel["INIT_INDEX"]))
        if len(dfVel.axes[0]) >0:
            # knimeinputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_SH", emptyList)
            inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_SH", emptyFloatList)
            refColumn = "V_SH"
            for x in range(len(dfVel.axes[0])):
                try:
                    pos = dfVel['INIT_INDEX'][x]
                    inputdf.iloc[pos, inputdf.columns.get_loc('V_SH')] = inputdf["V_TRAIN"][pos]
                except Exception as ex:
                    print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)
                    continue

        # New column for distances where in SB mode
        # to differnetiate modes in excel graphics
        maskSB = inputdf['M_MODE'].eq(6, fill_value = False)
        # Knime dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'], 'V_SB' : emptyList})
        dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'].copy(), 'V_SB' : emptyFloatList})
        dfVel = dfVel[maskSB]
        dfVel.index = numpy.arange(0, len(dfVel["INIT_INDEX"]))
        if len(dfVel.axes[0]) >0:
            # Knime inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_SB", emptyList)
            inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_SB", emptyFloatList)
            refColumn = "V_SB"
            for x in range(len(dfVel.axes[0])):
                try:
                    pos = dfVel['INIT_INDEX'][x]
                    inputdf.iloc[pos, inputdf.columns.get_loc('V_SB')] = inputdf["V_TRAIN"][pos]
                except Exception as ex:
                    print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)
                    continue

        # New column for distances where in LS mode
        # to differnetiate modes in excel graphics
        maskLS = inputdf['M_MODE'].eq(12, fill_value = False)
        # knimedfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'], 'V_LS' : emptyList})
        dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'].copy(), 'V_LS' : emptyFloatList})
        dfVel = dfVel[maskLS]
        dfVel.index = numpy.arange(0, len(dfVel["INIT_INDEX"]))
        if len(dfVel.axes[0]) >0:
            # Knime inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_LS", emptyList)
            inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_LS", emptyFloatList)
            refColumn = "V_LS"
            for x in range(len(dfVel.axes[0])):
                try:
                    pos = dfVel['INIT_INDEX'][x]
                    inputdf.iloc[pos, inputdf.columns.get_loc('V_LS')] = inputdf["V_TRAIN"][pos]
                except Exception as ex:
                    print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)
                    continue

        # New column for distances where in National System mode
        # to differnetiate modes in excel graphics
        maskNS = inputdf['M_MODE'].eq(13, fill_value = False)
        # Knime dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'], 'V_NS' : emptyList})
        dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'].copy(), 'V_NS' : emptyFloatList})
        dfVel = dfVel[maskNS]
        dfVel.index = numpy.arange(0, len(dfVel["INIT_INDEX"]))
        if len(dfVel.axes[0]) >0:
            # Knime inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_NS", emptyList)
            inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_NS", emptyFloatList)
            refColumn = "V_NS"
            for x in range(len(dfVel.axes[0])):
                try:
                    pos = dfVel['INIT_INDEX'][x]
                    inputdf.iloc[pos, inputdf.columns.get_loc('V_NS')] = inputdf["V_TRAIN"][pos]
                except Exception as ex:
                    print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)
                    continue

        # New column for distances where in other modes
        # to differnetiate modes in excel graphics
        maskOthers = ~(maskFS | maskOS | maskSR | maskSH | maskSB | maskLS | maskNS)
        # Knime dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'], 'V_OTHER' : emptyList})
        dfVel = pandas.DataFrame({'INIT_INDEX': inputdf['INIT_INDEX'].copy(), 'V_OTHER' : emptyFloatList})
        dfVel = dfVel[maskOthers]
        dfVel.index = numpy.arange(0, len(dfVel["INIT_INDEX"]))
        if len(dfVel.axes[0]) >0:
            # Knime inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_OTHER", emptyList)
            inputdf.insert(inputdf.columns.get_loc(refColumn)+1, "V_OTHER", emptyFloatList)
            refColumn = "V_OTHER"
            for x in range(len(dfVel.axes[0])):
                try:
                    pos = dfVel['INIT_INDEX'][x]
                    inputdf.iloc[pos, inputdf.columns.get_loc('V_OTHER')] = inputdf["V_TRAIN"][pos]
                except Exception as ex:
                    print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)
                    continue
        
    
        
            # ------------------- Add columns with V_TRAIN per M_MODE (ends) ------------------------
            
    # Output
    dfMissionsSummary["Engine"] = numpy.array(listNidEngine, dtype=int)
    dfMissionsSummary["IxSoM"] = numpy.array(listSoM, dtype=int)
    dfMissionsSummary["IxEoM"] = numpy.array(listEoM, dtype=int)
    dfMissionsSummary["Date"] = numpy.array(listDate)
    dfMissionsSummary["Time"] = numpy.array(listTime)
    dfMissionsSummary["EoM"] = numpy.array(listMissionCompleted)
    dfMissionsSummary["SH"] = numpy.array(listModeSh)
    dfMissionsSummary["Distance"] = numpy.array(listDmission, dtype=float)
    dfMissionsSummary["LocError"] = numpy.array(listOdoerrMission, dtype=float)
    dfMissionsSummary["NumBG"] = numpy.array(listNumBG, dtype=int)
    dfMissionsSummary["NumLrbg"] = numpy.array(listNumLRBG, dtype=int)
    dfMissionsSummary["Mtime (secs)"] = numpy.array(listMissionTime, dtype=float)
    dfMissionsSummary["DistAnt"] = numpy.array(listMissionDistAntenna, dtype=float)

    # Drop columns
    # Note: one finished this code drop unused columns at the beginning to reduce size of data and reduce timing
    #inputdf.drop(['INIT_INDEX','D_REF', 'Rbc Packets', 'Q_LINK', 'Balise Packets'], axis=1, inplace=True)
    # inputdf.drop(['INIT_INDEX', 'Rbc Packets', 'Q_LINK', 'Balise Packets'], axis=1, inplace=True)
    # inputdf.drop(['Rbc Packets', 'Q_LINK'], axis=1, inplace=True)
    # shift column 'V_PERM' after V_STATIC
    vmrspcol = inputdf.pop('V_PERM') 
    inputdf.insert(inputdf.columns.get_loc('V_STATIC')+1, "V_PERM", vmrspcol)

    # shift error columns to last position 
    last_column = inputdf.pop('ERR_TYPE') 
    inputdf.insert(len(inputdf.columns), 'ERR_TYPE', last_column) 
    last_column = inputdf.pop('ERR_MESSAGE') 
    inputdf.insert(len(inputdf.columns), 'ERR_MESSAGE', last_column) 

    dfArrOutput = []
    dfArrOutput.insert(0, dfMissionsSummary)
    dfArrOutput.insert(1, inputdf)

    sys.stdout.write("\n")
    sys.stdout.flush()
    
    return dfArrOutput

# JruC_SetObestimatedRoutes(JruC_SetObestimatedRoutes)
# Set estimated route info from ONB estimations
def JruC_SetObestimatedRoutes(dfMission, dfRoutes, distAntenna):

    # dfInputMission = pandas.DataFrame(dfMission) # Make a copy of Mission
    # !Important: Keep dfRoutes as reference to insert new values

    if distAntenna == None:
        print(color.Fore.RED + 'Error. Antenna distance value not valid.' +
                              color.Style.RESET_ALL)
        exit()

    # First, get matrix with distance traveled for train from balise overpassing
    # and change of DLRBG
    selfDlrbg = 0.0
    for n in range(len(dfMission)):
        
        strNidBGAux = str(dfMission['NID_BG'][n])
        
        if ((strNidBGAux == '') | (strNidBGAux == 'nan')):
            continue

        if '-0' not in strNidBGAux: # not first balise in BG
            continue
        
        strLrbgAux = dfMission['NID_LRBG'][n]
        ixAux = dfMission['INIT_INDEX'][n]
        # posRecord = dfInputMission['INIT_INDEX'][n]
        posRecord = dfMission['INIT_INDEX'][n]
               
        dfAux = dfMission[(dfMission['NID_LRBG'] == strLrbgAux) & (dfMission['INIT_INDEX'] >= ixAux)]    
        # Case: same balise found more than once in mission
        # Find those related to processed record
        dfTemp = dfAux[dfAux['NID_BG'] == strNidBGAux ]
        dfTemp.reset_index(drop=True, inplace=True)

        if len(dfTemp.axes[0]) > 1: 
                            
            # Delete records in case of balise found more than once in mission
            # Select that with minimum distance to MA record
            closestPos = 0
            closestDistance = None
            recordsDistance = 0
            for m in range(len(dfTemp.axes[0])):
                recordsDistance = abs(dfTemp['INIT_INDEX'][m] - posRecord)
                if closestDistance == None:
                    closestDistance = recordsDistance
                    closestPos = dfTemp['INIT_INDEX'][m]
                elif closestDistance > recordsDistance:
                    closestDistance = recordsDistance
                    closestPos = dfTemp['INIT_INDEX'][m]
                            
            dfTemp = dfTemp[dfTemp['INIT_INDEX'] == closestPos]
            dfTemp.reset_index(drop=True, inplace=True)

        dfAux = dfAux[dfAux['LRBG_Ext'] == dfTemp['LRBG_Ext'][0]]
        dfAux.reset_index(drop=True, inplace=True)
        # In case of error in balise that is not taken as a LRBG
        # we have more than one BG in this cet of data with the same LRBG_EXT
        # So, we filter by the last one
        for m in reversed(range(len(dfAux.axes[0]))):

            strLastNidBGAux = str(dfAux['NID_BG'][m])
        
            if ((strLastNidBGAux == '') | (strLastNidBGAux == 'nan')):
                continue

            if '-0' not in strLastNidBGAux: # not first balise in BG
                continue

            strLrbgAux = dfAux['NID_LRBG'][m]
            ixAux = dfAux['INIT_INDEX'][m]
                
            dfAux = dfAux[(dfAux['NID_LRBG'] == strLrbgAux) & (dfAux['INIT_INDEX'] >= ixAux)]    
        
            break
        
        if len(dfAux.axes[0]) < 2: # Case balise-0 last in group. Undefined distance
            continue

        dfAux.reset_index(drop=True, inplace=True)
        selfDlrbg = dfAux['D_LRBG'][len(dfAux.axes[0]) - 1] - dfAux['D_LRBG'][0]

        initIndexMission = dfAux['INIT_INDEX'][len(dfAux.axes[0]) - 1] 
        posSelfDlrbg = dfMission[dfMission['INIT_INDEX'] == initIndexMission].index[0]
        dfMission.iloc[posSelfDlrbg, dfMission.columns.get_loc('SELF_DLRBG_E')] = selfDlrbg
        


    # Get unique lrbg-ext ids; this list is used to figure out the accumulated distance            
    # Remove those rows with LRBG = 16383 (unknown)
    # listExtendedNidlrbg = dfInputMission[~dfInputMission["LRBG_Ext"].str.contains('16383', na = False)]['LRBG_Ext'].unique() # v2.1
    #listExtendedNidlrbg = dfInputMission['LRBG_Ext'].unique()
    listExtendedNidlrbg = dfMission['LRBG_Ext'].unique()
    dist = 0.0
    mTime = 0
    prevNidLRBG = ''
    prevSelfDlrg = 0.0
    prevDlrbg = 0.0
    prevQlrbg = None
    prevmTime = None
    for i in range (len(listExtendedNidlrbg)):
        # For each reference balise
        # dfNidExtendedLrbgInMission = dfInputMission[dfInputMission["LRBG_Ext"].str.contains(listExtendedNidlrbg[i], na = False)]
        # dfNidExtendedLrbgInMission = dfInputMission[dfInputMission["LRBG_Ext"].str.fullmatch(listExtendedNidlrbg[i], na = False)]
        dfNidExtendedLrbgInMission = dfMission[dfMission["LRBG_Ext"].str.fullmatch(listExtendedNidlrbg[i], na = False)].copy()
        dfNidExtendedLrbgInMission.index = numpy.arange(0, len(dfNidExtendedLrbgInMission["LRBG_Ext"]))

        # First and last pos in global DF (inputdf) related to the list LRG_Ext
        firstInitPos = dfNidExtendedLrbgInMission["INIT_INDEX"][0]
        lastInitPos = dfNidExtendedLrbgInMission["INIT_INDEX"][len(dfNidExtendedLrbgInMission["INIT_INDEX"])-1]
                

        if(i>0): # Take into account last extended reference balise for accumulated distance
                    
            # Calculate distance from previous LRBG-Ext
            if prevNidLRBG == dfNidExtendedLrbgInMission['NID_LRBG'][0]:
                # Case same LRBG but orientation change 
                if(prevQlrbg != dfNidExtendedLrbgInMission['Q_DLRBG'][0]):
                            
                    if((prevQlrbg == 2) | (dfNidExtendedLrbgInMission['Q_DLRBG'][0] == 2)):
                        # Orientation change from or to value 2 (unknown)
                        # relative distance last - previous (hypothesis : no change of orientation)
                        dist += (dfNidExtendedLrbgInMission['D_LRBG'][0] - prevDlrbg)
                                
                    else:
                        # Real change of orientation to/from 1/0
                        # relative distance last + previous (change of orientation)
                        dist += (dfNidExtendedLrbgInMission['D_LRBG'][0] + prevDlrbg)
                        # Note: we could add a check to be sure that the balise has been overpassed
                                
            else:
                # Case jump from a different LRBG
                # Change of reference balise
                # First calculate distance moved since new balise reading
                dist += (dfNidExtendedLrbgInMission['D_LRBG'][0] - prevSelfDlrg) - distAntenna # pepito grillo

                     
            # Calculate time from previous LRBG-Ext
            mTime0 = prevmTime
            strDate = str(dfNidExtendedLrbgInMission["Date"][0]) # dd/mm/yy
            #ix = strDate.index('/')
            #dd = strDate[0:ix].rjust(2, '0')
            #strDate = strDate[ix+1:len(strDate)]
            #ix = strDate.index('/')
            #mm = strDate[0:ix].rjust(2, '0')
            #yy = strDate[ix+1:len(strDate)]
            #strDate = dd + '/' + mm + '/' + yy # to assure dd/mm/yy
            strTime = strDate + ' ' + dfNidExtendedLrbgInMission["Time"][0] # H:mm:ss
            mTime1= pandas.to_datetime(strTime, format='%d/%m/%y %H:%M:%S')
            incrTime = mTime1 - mTime0
            mTime += incrTime.seconds      


                
        # Set accumulated distance at first position
        dfRoutes.iloc[firstInitPos, dfRoutes.columns.get_loc('D_ONB_ACCUM')] = dist
        # set accumulated time at firtst position
        dfRoutes.iloc[firstInitPos, dfRoutes.columns.get_loc('T_ONB_ACCUM')] = mTime

        # Figure out incremental distance in extended lrbg block
        dist0= dfNidExtendedLrbgInMission['D_LRBG'][0]
        dist1= dfNidExtendedLrbgInMission['D_LRBG'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1]
        dist += abs(dist1 - dist0)
        # Figure out incremental time in extended lrbg block
        strDate = str(dfNidExtendedLrbgInMission["Date"][0]) #.rjust(8, '0') # dd/mm/yy
        #ix = strDate.index('/')
        #dd = strDate[0:ix].rjust(2, '0')
        #strDate = strDate[ix+1:len(strDate)]
        #ix = strDate.index('/')
        #mm = strDate[0:ix].rjust(2, '0')
        #yy = strDate[ix+1:len(strDate)]
        #strDate = dd + '/' + mm + '/' + yy # to assure dd/mm/yy
        strTime = strDate + ' ' + dfNidExtendedLrbgInMission["Time"][0] # H:mm:ss
        mTime0= pandas.to_datetime(strTime, format='%d/%m/%y %H:%M:%S')
        strDate = str(dfNidExtendedLrbgInMission["Date"][len(dfNidExtendedLrbgInMission["Date"])-1]) #.rjust(8, '0') # dd/mm/yy
        #ix = strDate.index('/')
        #dd = strDate[0:ix].rjust(2, '0')
        #strDate = strDate[ix+1:len(strDate)]
        #ix = strDate.index('/')
        #mm = strDate[0:ix].rjust(2, '0')
        #yy = strDate[ix+1:len(strDate)]
        #strDate = dd + '/' + mm + '/' + yy # to assure dd/mm/yy
        strTime = strDate + ' ' + dfNidExtendedLrbgInMission["Time"][len(dfNidExtendedLrbgInMission["Time"])-1] # H:mm:ss
        mTime1= pandas.to_datetime(strTime, format='%d/%m/%y %H:%M:%S')
        incrTime = mTime1 - mTime0
        mTime += incrTime.seconds      

                
        # Set accumulated distance at last position
        dfRoutes.iloc[lastInitPos, dfRoutes.columns.get_loc('D_ONB_ACCUM')] = dist
                # Set accumulated time at last position
        dfRoutes.iloc[lastInitPos, dfRoutes.columns.get_loc('T_ONB_ACCUM')] = mTime

        # Save values of current lrbg to compare in next loop
        prevNidLRBG = dfNidExtendedLrbgInMission['NID_LRBG'][0]
        strSelfDlrbg = str(dfNidExtendedLrbgInMission['SELF_DLRBG_E'][len(dfNidExtendedLrbgInMission.axes[0]) -1])
        if ((strSelfDlrbg == '') | (strSelfDlrbg == 'nan')):
            prevSelfDlrg = 0.0
        else:
            prevSelfDlrg = float(strSelfDlrbg)

        
        prevDlrbg = dfNidExtendedLrbgInMission['D_LRBG'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1]
        prevQlrbg = dfNidExtendedLrbgInMission['Q_DLRBG'][len(dfNidExtendedLrbgInMission["Q_DLRBG"])-1]
        prevmTime = mTime1


        # Figure out intermediate distances values between first and last position within the extended LRBG
        if(len(dfNidExtendedLrbgInMission['LRBG_Ext']) > 2):
            initialAccDist = dfRoutes['D_ONB_ACCUM'][dfNidExtendedLrbgInMission["INIT_INDEX"][0]]
            initialDlrbg = dfRoutes['D_LRBG'][dfNidExtendedLrbgInMission["INIT_INDEX"][0]]

            for j in range(1, len(dfNidExtendedLrbgInMission['NID_LRBG'])-1):
                intermediateInitPos = dfNidExtendedLrbgInMission["INIT_INDEX"][j] 
                moveDist = abs(dfRoutes['D_LRBG'][dfNidExtendedLrbgInMission["INIT_INDEX"][j]] - initialDlrbg)
                dfRoutes.iloc[intermediateInitPos, dfRoutes.columns.get_loc('D_ONB_ACCUM')] = initialAccDist + moveDist
                    

        # Figure out intermediate time values between first and last position within the extended LRBG
        if(len(dfNidExtendedLrbgInMission['LRBG_Ext']) > 2):
            initialAccDist = dfRoutes['D_ONB_ACCUM'][dfNidExtendedLrbgInMission["INIT_INDEX"][0]]
            initialDlrbg = dfRoutes['D_LRBG'][dfNidExtendedLrbgInMission["INIT_INDEX"][0]]

            initialAccTime = dfRoutes['T_ONB_ACCUM'][dfNidExtendedLrbgInMission["INIT_INDEX"][0]]
            initialStrDate = str(dfRoutes["Date"][dfNidExtendedLrbgInMission["INIT_INDEX"][0]]) #.rjust(8, '0') # dd/mm/yy
            #ix = initialStrDate.index('/')
            #dd = initialStrDate[0:ix].rjust(2, '0')
            #initialStrDate = initialStrDate[ix+1:len(initialStrDate)]
            #ix = initialStrDate.index('/')
            #mm = initialStrDate[0:ix].rjust(2, '0')
            #yy = initialStrDate[ix+1:len(initialStrDate)]
            #initialStrDate = dd + '/' + mm + '/' + yy # to assure dd/mm/yy
            initialStrTime = initialStrDate + ' ' + dfRoutes["Time"][dfNidExtendedLrbgInMission["INIT_INDEX"][0]] # H:mm:ss
            initialmTime = pandas.to_datetime(initialStrTime, format='%d/%m/%y %H:%M:%S')

            for j in range(1, len(dfNidExtendedLrbgInMission['NID_LRBG'])-1):
                intermediateInitPos = dfNidExtendedLrbgInMission["INIT_INDEX"][j] 
                moveDist = abs(dfRoutes['D_LRBG'][dfNidExtendedLrbgInMission["INIT_INDEX"][j]] - initialDlrbg)
                dfRoutes.iloc[intermediateInitPos, dfRoutes.columns.get_loc('D_ONB_ACCUM')] = initialAccDist + moveDist
                    
                strDate = str(dfRoutes["Date"][dfNidExtendedLrbgInMission["INIT_INDEX"][j]]) # .rjust(8, '0') # dd/mm/yy
                #ix = strDate.index('/')
                #dd = strDate[0:ix].rjust(2, '0')
                #strDate = strDate[ix+1:len(strDate)]
                #ix = strDate.index('/')
                #mm = strDate[0:ix].rjust(2, '0')
                #yy = strDate[ix+1:len(strDate)]
                #strDate = dd + '/' + mm + '/' + yy # to assure dd/mm/yy
                strTime = strDate + ' ' + dfRoutes["Time"][dfNidExtendedLrbgInMission["INIT_INDEX"][j]] # H:mm:ss
                mTime1= pandas.to_datetime(strTime, format='%d/%m/%y %H:%M:%S')
                incrTime = mTime1 - initialmTime
                moveTime = incrTime.seconds
                dfRoutes.iloc[intermediateInitPos, dfRoutes.columns.get_loc('T_ONB_ACCUM')] = initialAccTime + moveTime
    

    # Copy columnm -0 to keep initial values
    firstMissionIndex = dfMission['INIT_INDEX'][0]
    lastMissionIndex = dfMission['INIT_INDEX'][len(dfMission) - 1]
    for z in range(len(dfMission)):

        ix = dfMission['INIT_INDEX'][z]
        if z == 0: firstInitIndex = ix
        # KNIME dfRoutes.iloc[ix, dfRoutes.columns.get_loc('D_ONB_ACCUM-0')] = dfMission['D_ONB_ACCUM'][z]
        dfRoutes.iloc[ix, dfRoutes.columns.get_loc('D_ONB_ACCUM-0')] = dfRoutes['D_ONB_ACCUM'][ix] 
        dfRoutes.iloc[ix, dfRoutes.columns.get_loc('SELF_DLRBG_E')] = dfMission['SELF_DLRBG_E'][z]

    # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function
    dfReturnedMission = dfRoutes.iloc[firstMissionIndex:lastMissionIndex+1].copy()
    dfReturnedMission.index = numpy.arange(0, len(dfReturnedMission["INIT_INDEX"]))
    listReturn = [dist, dfReturnedMission]
    
       
    return listReturn

# JruC_SetLinkedRoutes(dfMission, dfRoutes)
# Compiled JRU File
# Set route distance info from linking data
def JruC_SetLinkedRoutes(dfMission, dfRoutes, distAntenna):
            
    # dfInputMission = pandas.DataFrame(dfMission) # Make a copy of Mission
    # !Important: Keep dfRoutes as reference to insert new values

    if distAntenna == None:
        print(color.Fore.RED + 'Error. Antenna distance value not valid.' +
                              color.Style.RESET_ALL)
        exit()
    distAntenna = 0

    dfTrackDistances = pandas.DataFrame() # initialization output df

    
    # Variables to create dataframe with accumulated linked distances
    # Initialization on each mission
    iLinkedAcc=0 # index of dataframe for all linked balises in Mission
    # Next lists to insert in output DF
    listTrackInitIndex = []
    listTrackDlinkAcc = []
    listTrackNidBG = []
    listTrackQlocacc = []
    listTrackBGplusIndex = []

    # get info: linked balises and distances, from packt 5 info
    # for x in range(len(dfInputMission['BG_LINKS'])):
    for x in range(len(dfMission['BG_LINKS'])):
        boolMAoverlapped = True
        #posRecordMA = dfInputMission['INIT_INDEX'][x]
        posRecordMA = dfMission['INIT_INDEX'][x]
        #strBgLinks = dfInputMission['BG_LINKS'][x]
        strBgLinks = dfMission['BG_LINKS'][x]
        if str(strBgLinks) == 'nan':
            strBgLinks = ""

        if strBgLinks != "":
            if '731' in strBgLinks:
                pepe = 0
            boolBGLink = True
            iLinking=0
            listNidBG = []
            listDlink = []
            listQlocacc = []
            strNextBG = strBgLinks.strip()
            # Variable to be used for distance reference when jumping between MAs
            # In that cases linking of both MAs overlaps
            accDrefFirstBalNewMA = 0.0 
            refNidBG = ''
            # format of strBgLinks
            # 1361:0:10-1367:66.0:10-1371:97.0:10-1375:226.0:10-1379:97.0:10-16382:28.0:0
            # nid_bg:distance:qlocacc-nid_bg:distance:qlocacc- ...
            # Translate link info into accumulated distance on linked balises

            while boolBGLink == True:
                strNextBG = strNextBG[0:len(strNextBG)]
                try:
                    indexBG = strNextBG.index(':')
                except:
                    boolBGLink = False
                    # To be processed in future; error in linking constructor
                    continue

                nidBG = strNextBG[0:indexBG]
                
                
                try:
                    strNextBG = strNextBG[indexBG+1:len(strNextBG)]
                    indexDlink = strNextBG.index(':')
                except:
                    boolBGLink = False 
                    # To be processed in future; error in linking constructor
                    continue

                dlink = strNextBG[0:indexDlink]
                
                try:
                    strNextBG = strNextBG[indexDlink+1:len(strNextBG)]
                    indexQlocacc = strNextBG.index('-')
                except:
                    boolBGLink = False # last balise
                    indexQlocacc = len(strNextBG)

                qlocacc = strNextBG[0:indexQlocacc]
                #if qlocacc == 0: # !Important: value used for reference balise in MA message; not taken into account
                #    continue

                listNidBG.insert(iLinking, nidBG)
                listDlink.insert(iLinking, dlink)
                listQlocacc.insert(iLinking, int(qlocacc))

                iLinking += 1

                try:
                    strNextBG = strNextBG[indexQlocacc+1:len(strNextBG)]
                except:
                    boolBGLink = False 
                    # To be processed in future; error in linking constructor
                    continue
                    
                

            serNidbg = pandas.Series(listNidBG, dtype=str)
            serDlink = pandas.Series(listDlink, dtype=float)
            serQlocacc = pandas.Series(listQlocacc, dtype=int)
            # Dataframe with info from packet 5
            dfLinkingInfo = pandas.DataFrame({'NIDBG': serNidbg, 'DLINK': serDlink, 
                                              'QLOCACC': serQlocacc})

            # Delete overlaps linking info from previous MAs
            # dfTrackDistances is a DF created from linking (packet 5) AND train movement (overpassed balises)
            # It has the accumulated distance from the SoM based on d_link info from packet 5
            # These are the reference distances configured at trackside and sent to OB via linking info
            if len(dfTrackDistances.axes[0] > 0):  # Not the first MA in mission
                # In case we had a previous MA
                # it is necesary to keep accumulated distance 
                # in the moment train gets a new MA with overlapped info
                firstBaliseNewMA = dfLinkingInfo['NIDBG'][0]  # first balise in the new MA

                #dfFirstBalNewMARecord = dfInputMission[dfInputMission['NID_BG'] == (firstBaliseNewMA + '-0')]
                dfFirstBalNewMARecord = dfMission[dfMission['NID_BG'] == (firstBaliseNewMA + '-0')].copy()
                dfFirstBalNewMARecord.index = numpy.arange(0, len(dfFirstBalNewMARecord["NID_BG"])).copy()  

                # Case: same balise found more than once in mission
                if len(dfFirstBalNewMARecord.axes[0]) > 1: 
                    
                    # Delete records in case of balise found more than once in mission
                    # Select that with minimum distance to MA record
                    closestPos = 0
                    closestDistance = None
                    recordsDistance = 0
                    for m in range(len(dfFirstBalNewMARecord.axes[0])):
                        recordsDistance = abs(dfFirstBalNewMARecord['INIT_INDEX'][m] - posRecordMA)
                        if closestDistance == None:
                            closestDistance = recordsDistance
                            closestPos = dfFirstBalNewMARecord['INIT_INDEX'][m]
                        elif closestDistance > recordsDistance:
                            closestDistance = recordsDistance
                            closestPos = dfFirstBalNewMARecord['INIT_INDEX'][m]
                    
                    dfFirstBalNewMARecord = dfFirstBalNewMARecord[dfFirstBalNewMARecord['INIT_INDEX'] == closestPos]
                    dfFirstBalNewMARecord.reset_index(drop=True, inplace=True) 

                try:
                    # search for this balise in the previous MA linking info
                    # df = dfTrackDistances[dfTrackDistances['NID_BG'] == firstBaliseNewMA]
                    strFilter = firstBaliseNewMA + '-0-' + str(dfFirstBalNewMARecord['INIT_INDEX'][0])
                    if '639' in strFilter:
                        pepe = 0
                    df = dfTrackDistances[dfTrackDistances['BGPLUSINDEX'] == strFilter]
                    df.index = numpy.arange(0, len(df.axes[0]))

                    # Case: same balise found more than once in mission
                    if len(df.axes[0]) > 1: 
                        
                        # Delete records in case of balise found more than once in mission
                        # Select that with minimum distance to MA record
                        closestPos = 0
                        closestDistance = None
                        recordsDistance = 0
                        for m in range(len(df.axes[0])):
                            recordsDistance = abs(df['INIT_INDEX'][m] - posRecordMA)
                            if closestDistance == None:
                                closestDistance = recordsDistance
                                closestPos = df['INIT_INDEX'][m]
                            elif closestDistance > recordsDistance:
                                closestDistance = recordsDistance
                                closestPos = df['INIT_INDEX'][m]
                        
                        df = df[df['INIT_INDEX'] == closestPos]
                        df.reset_index(drop=True, inplace=True)


                    # accumulated distance at the first balise in the new MA calculated from previous MA, in case of overlapping
                    accDrefFirstBalNewMA = df['D_LINKED_ACCUM'][0] 
                except:
                    # If no overlapping, there is no reference
                    # Take reference acc distance from that calculated onboard
                    print(color.Fore.YELLOW + 'Warning. Linked MAs. Balise '
                          +  firstBaliseNewMA + ' not found'
                          + color.Style.RESET_ALL)
                    
                    # add error info
                    #pos = dfInputMission['INIT_INDEX'][x]
                    pos = dfMission['INIT_INDEX'][x]
                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = 'Warning'
                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = 'Linked MAs. Balise ' + firstBaliseNewMA + ' not found'
                    try:
                        boolMAoverlapped = False
                        # set as ref distance that estimated onboard (warning: reference distance with odo_error)
                        accDrefFirstBalNewMA = dfFirstBalNewMARecord['D_ONB_ACCUM'][0] # keep here to raise exception whe balise is not found within SoM (dfFirstBalNewMARecord empty)
                        # Alternative: set to 0 to avoid unknown odo_error (later, when d_link_accum == 0, set odoerror to 0)
                        accDrefFirstBalNewMA = 0.0
                        print(color.Fore.YELLOW + 'Warning. No overlap in MAs. Setting distance reference to 0 at balise: '
                          +  firstBaliseNewMA 
                          + color.Style.RESET_ALL)
                        pos = dfFirstBalNewMARecord['INIT_INDEX'][0]
                        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = 'Warning'
                        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = 'No overlap in MAs. Setting distance reference to 0 for balise'
                    except:
                        # Case: when SoM after overpassing this balise; balise no within mission
                        # Exception raised due to empty dfFirstBalNewMARecord
                        accDrefFirstBalNewMA = 0.0 # Set dref distance to 0
                        print('Info. Distance reference set to 0 for balise ' +  firstBaliseNewMA)
                        
                # Then, remove overlapped linking info
                #for m in range(len(dfLinkingInfo.axes[0])):
                #    dfTrackDistances = dfTrackDistances[dfTrackDistances['NID_BG'] != dfLinkingInfo['NIDBG'][m]]
                #dfTrackDistances.index = numpy.arange(0, len(dfTrackDistances.axes[0]))
                # Initialize lists to new values of dataframe
                listTrackInitIndex = []
                listTrackDlinkAcc = []
                listTrackNidBG = []
                listTrackQlocacc = []
                listTrackBGplusIndex = []
                #for p in range(len(dfTrackDistances)):
                #    listTrackInitIndex.insert(p, dfTrackDistances['INIT_INDEX'][p])
                #    listTrackDlinkAcc.insert(p, dfTrackDistances['D_LINKED_ACCUM'][p])
                #    listTrackNidBG.insert(p, dfTrackDistances['NID_BG'][p])
                #    listTrackQlocacc.insert(p, dfTrackDistances['Q_LOCACC'][p])
                #    strBGplusIndex = str(dfTrackDistances['NID_BG'][p]) + '-' + str(dfTrackDistances['INIT_INDEX'][p])  
                #    listTrackBGplusIndex.insert(p, strBGplusIndex)
                #iLinkedAcc = len(dfTrackDistances.axes[0] - 1)
                iLinkedAcc = 0
                #if iLinkedAcc <0: 
                #    iLinkedAcc = 0 # protection
                
                
            # For each balise in linking message (packet 5) calculate accumulated linked distance
            prevNidBG = '' # reference from previous balise in linking message 
            for n in range(len(dfLinkingInfo.axes[0])):
                startNidBG = dfLinkingInfo['NIDBG'][n]
                if startNidBG == '731':
                    pepe =0
                if n == (len(dfLinkingInfo.axes[0])-1):
                    endNidBG = startNidBG
                else: 
                    endNidBG = dfLinkingInfo['NIDBG'][n + 1]
                
                startDlink = dfLinkingInfo['DLINK'][n] 
                startQlocacc = dfLinkingInfo['QLOCACC'][n] 
                
                #dfStartBaliseInfo = dfInputMission[dfInputMission['NID_BG'] == (startNidBG + '-0')]
                dfStartBaliseInfo = dfMission[dfMission['NID_BG'] == (startNidBG + '-0')].copy()
                dfStartBaliseInfo.index = numpy.arange(0, len(dfStartBaliseInfo["NID_BG"]))                
                
                
                # Case: same balise found more than once in mission
                if len(dfStartBaliseInfo.axes[0]) > 1: 
                    
                    # Delete records in case of balise found more than once in mission
                    # Select that with minimum distance to MA record
                    closestPos = 0
                    closestDistance = None
                    recordsDistance = 0
                    for m in range(len(dfStartBaliseInfo.axes[0])):
                        recordsDistance = abs(dfStartBaliseInfo['INIT_INDEX'][m] - posRecordMA)
                        if closestDistance == None:
                            closestDistance = recordsDistance
                            closestPos = dfStartBaliseInfo['INIT_INDEX'][m]
                        elif closestDistance > recordsDistance:
                            closestDistance = recordsDistance
                            closestPos = dfStartBaliseInfo['INIT_INDEX'][m]
                    
                    dfStartBaliseInfo = dfStartBaliseInfo[dfStartBaliseInfo['INIT_INDEX'] == closestPos]
                    dfStartBaliseInfo.reset_index(drop=True, inplace=True)
                    
                    pos = dfStartBaliseInfo['INIT_INDEX'][0]
                    
                    if(dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] == ''):

                        print('Info. Balise found more than once in mission: '
                            +  str(dfStartBaliseInfo['NID_BG'][0]))
                                                
                        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = 'Info'
                        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = 'Balise found more than once in mission'
    
                # Validate starting records between linked balises
                # A) train has not passed over the balise
                if len(dfStartBaliseInfo.axes[0]) == 0: 
                    prevNidBG = startNidBG # save previous balise in linking message
                    continue # train has not passed over this balise

                # B) either train orientation mismatched with MA (packet 5) or balise not found
                # Case B only take into account in Level 1 (M_LEVEL = 2 )
                if dfStartBaliseInfo['M_LEVEL'][0] != 3: # Not level 2
                    if dfStartBaliseInfo['INIT_INDEX'][0] < posRecordMA: # Record found before packet 5 that contains current balise

                        # Depending on release "TELEGRAM FROM BALISE [6]" or "MESSAGE FROM BALISE [6]"
                        # In both cases the word BALISE is within
                        # if 'BALISE' in dfInputMission[dfInputMission['INIT_INDEX'] == posRecordMA].iloc[0]['NID_MESSAGE']:
                        if 'BALISE' in dfMission[dfMission['INIT_INDEX'] == posRecordMA].iloc[0]['NID_MESSAGE']:

                            # In this case packet 5 comes from balise and we have to check that we have the right start balise info
                            # In ths case LRBG_Ext dont match, balise not found
                            # if dfStartBaliseInfo['LRBG_Ext'][0] != dfInputMission[dfInputMission['INIT_INDEX'] == posRecordMA].iloc[0]['LRBG_Ext']:
                            if dfStartBaliseInfo['LRBG_Ext'][0] != dfMission[dfMission['INIT_INDEX'] == posRecordMA].iloc[0]['LRBG_Ext']:
                                prevNidBG = startNidBG # save previous balise in linking message
                                continue # train has not passed over this balise
                        else:
                            pepe = 0
                            # Level 1 but a message from RBC received; may be in a level transition
                        
                
                # dfEndBaliseInfo = dfInputMission[dfInputMission['NID_BG'] == (endNidBG + '-0')]
                dfEndBaliseInfo = dfMission[dfMission['NID_BG'] == (endNidBG + '-0')].copy()
                dfEndBaliseInfo.index = numpy.arange(0, len(dfEndBaliseInfo["NID_BG"]))

                if len(dfEndBaliseInfo.axes[0]) > 1: 

                    # Delete records in case of balise found more than once in mission
                    # Select that with minimum distance to MA record
                    closestPos = 0
                    closestDistance = None
                    recordsDistance = 0
                    for m in range(len(dfEndBaliseInfo.axes[0])):
                        #recordsDistance = abs(dfEndBaliseInfo['INIT_INDEX'][m] - posRecordMA)
                        
                        recordsDistance = dfEndBaliseInfo['INIT_INDEX'][m] - posRecordMA # ! Important: positive (bigger index)
                        if recordsDistance < 0: continue # Only those records after linking balise
                        if closestDistance == None:
                            closestDistance = recordsDistance
                            closestPos = dfEndBaliseInfo['INIT_INDEX'][m]
                        elif closestDistance > recordsDistance:
                            closestDistance = recordsDistance
                            closestPos = dfEndBaliseInfo['INIT_INDEX'][m]
                    
                    dfEndBaliseInfo = dfEndBaliseInfo[dfEndBaliseInfo['INIT_INDEX'] == closestPos]
                    dfEndBaliseInfo.reset_index(drop=True, inplace=True)
                

                # Validate starting records between linked balises
                    # Case mismatch between train movement and linking info
                    
                if len(dfEndBaliseInfo.axes[0]) > 0:
                    # Messages from balises in opposite order than in linking (packet 5)
                    if dfStartBaliseInfo['INIT_INDEX'][0] > dfEndBaliseInfo['INIT_INDEX'][0]: 
                        # Insert error message instead of warning about distance estimation
                        # Error - linking in other direction; not take into account
                        print(color.Fore.RED + 'Error. Linking mismatch.' 
                              + 'From linked balises ' + str(dfStartBaliseInfo['NID_BG'][0])
                              + ' to ' + str(dfEndBaliseInfo['NID_BG'][0])
                              + color.Style.RESET_ALL)
                        pos = dfStartBaliseInfo['INIT_INDEX'][0]
                        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = 'Error'
                        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = 'Linking mismatch.'
                        # continue # Linking mismatch
                        break # Linking mismatch
                    
                   
                
                                
                # Add info to output data frame
                # Reference to INIT_INDEX in original dataframe
                listTrackInitIndex.insert(iLinkedAcc, dfStartBaliseInfo['INIT_INDEX'][0])
                # Remove extension -0 from first balise in BG
                zeroBalise = dfStartBaliseInfo['NID_BG'][0]
                indexZB = zeroBalise.index('-')
                zeroBalise = zeroBalise[0:indexZB]
                listTrackNidBG.insert(iLinkedAcc, zeroBalise)
                strBGplusIndex = str(dfStartBaliseInfo['NID_BG'][0]) + '-' + str(dfStartBaliseInfo['INIT_INDEX'][0])
                listTrackBGplusIndex.insert(iLinkedAcc, strBGplusIndex)
                listTrackQlocacc.insert(iLinkedAcc, startQlocacc)

                if (iLinkedAcc == 0) & (len(dfTrackDistances.axes[0]) == 0): # first balise in mission
                    # We adjust initial starting point as calculated on board
                    # Train estimated distance and linking distance matched at start 
                    dlinkAcc = dfStartBaliseInfo['D_ONB_ACCUM'][0]
                    # KNIME if dlinkAcc == '': dlinkAcc = 0.0 # set to 0
                    if (numpy.isnan(dlinkAcc)): dlinkAcc = 0.0 # set to 0
                    listTrackDlinkAcc.insert(iLinkedAcc, dlinkAcc)
                    #pepito grillo

                else:
                    
                    # Balises not the first one at mission
                    # Set as reference the last balise in JRU with linked reference distance
                    # refNidBG = listTrackNidBG[iLinkedAcc -1]
                    
                    # Case: mismatch of order in linking and balise messages
                    boolLinkedError = False
                    if (refNidBG != '') & (prevNidBG != '') & (refNidBG != prevNidBG): 
                                               
                        # In this case previous reference is not valid
                        # refNidBG: last linked balise over passed (train movement)
                        # prevNidBG: previous balise in linking (in MA)
                        # startNidBG: current balise at track (train movement) that matches with linking balise
                        # To find out the distance we have to calculate linking distance between refNidBG and current
                        
                        # accDrefFirstBalNewMA (neede when n=0 first balise in MA), is not different as missing balises error doesnt apply with first balise in MA
                        # Not apply here accDrefFirstBalNewMA = dfStartBaliseInfo['D_ONB_ACCUM'][0] 

                        # !! Important: Adjust startDlink (n >0)  with missing balises 
                        
                        
                        print(color.Fore.RED + 'Error. Linked balise not found. Using estimated distance for balise '
                          +  zeroBalise 
                          + color.Style.RESET_ALL)
                        pos = dfStartBaliseInfo['INIT_INDEX'][0]
                        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = 'Error'
                        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = 'Linked balise not found. Using estimated distance for balise'
                        print('Prev (link): ' + str(prevNidBG) + ' - Current (link): ' + str(startNidBG) + ' - Last reference (train): ' + str(refNidBG))
                        boolLinkedError = True
                        # listLinkedBGError[iLinkedAcc] = 'yes'
                        
                        # refNidBG = startNidBG # reset balise reference
                        
                    if n == 0:
                        
                        # First balise for the new MA
                        # Take reference distance from previous MA (overlapped info)
                        # or Onboard calculated distance in case of no overlapping
                        if boolMAoverlapped == False:
                            # In this case accDrefFirstBalNewMA = 0. See above
                            boolMAoverlapped = True
                        listTrackDlinkAcc.insert(iLinkedAcc, accDrefFirstBalNewMA)
                        #pepito grillo

                    else:
                        
                        # For the rest of balises in the MA we add d_link 
                        # to the accumulated distance from previous balise
                        # 
                        if boolLinkedError == True:
                            # In this case:
                            # if (refNidBG != '') & (prevNidBG != '') & (refNidBG != prevNidBG): 
                            # There is a jump in balise linking. For example:
                            # Error. Linked balise not found. Using estimated distance for balise 2717
                            # Prev (link): 2716 - Current (link): 2717 - Last reference (train): 2715
                            # In previous case there is a linking between balse 2716 and 2717
                            # but last reference balise onboard is 2715
                            # Then it is necesary to calculate relative distance between 2715 and 2717
                            try:
                                ixFirst = dfLinkingInfo[dfLinkingInfo['NIDBG'] == refNidBG].index[0]
                                ixLast = dfLinkingInfo[dfLinkingInfo['NIDBG'] == startNidBG].index[0]
                                if ixFirst > ixLast: # Error in linking
                                    raise Exception()
                                if (ixLast - ixFirst) < 2:
                                    raise Exception()
                                dfAux = dfLinkingInfo[ixFirst + 1:ixLast + 1] # 
                                startDlink = dfAux['DLINK'].sum() 

                                listTrackDlinkAcc.insert(iLinkedAcc, listTrackDlinkAcc[iLinkedAcc - 1] + startDlink + distAntenna)
                                #pepito grillo
                                print(color.Fore.RED + 'Linked balise : ' + str(prevNidBG) + ' not overpassed. Using balise: ' +
                                    str(refNidBG) + ' as reference.'   +   
                                    color.Style.RESET_ALL)           

                            except Exception as ex:
                                print(color.Fore.RED + 'Error. Not valid linking found. Set distance to 0 '
                                + color.Style.RESET_ALL)     
                                listTrackDlinkAcc.insert(iLinkedAcc, 0)
                                #pepito grillo
                            

                        elif boolMAoverlapped == False:  # !!! Remove; not used
                            
                            # If not overlapped MAs: 
                            # accDrefFirstBalNewMA set to 0 
                            # tipically when transitions from NS in middle of the mission
                            #listTrackDlinkAcc.insert(iLinkedAcc, listTrackDlinkAcc[iLinkedAcc - 1] + startDlink)
                            listTrackDlinkAcc.insert(iLinkedAcc, accDrefFirstBalNewMA)
                            #pepito grillo
                            boolMAoverlapped = True # For next balise

                        elif((boolLinkedError == False) & (boolMAoverlapped == True)):

                            # There is no problem neither with linking or MA eoverlappoing
                            if iLinkedAcc == 0:
                                # Insert error message instead of warning about distance estimation
                                # Error - linking in other direction; not take into account
                                print(color.Fore.RED + 'Error. Linking mismatch.' 
                                    + 'From linked balises ' + str(dfStartBaliseInfo['NID_BG'][0])
                                    + ' to ' + str(dfEndBaliseInfo['NID_BG'][0])
                                    + color.Style.RESET_ALL)
                                pos = dfStartBaliseInfo['INIT_INDEX'][0]
                                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = 'Error'
                                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = 'Linking mismatch.'
                                continue # Linking mismatch
                            listTrackDlinkAcc.insert(iLinkedAcc, listTrackDlinkAcc[iLinkedAcc - 1] + startDlink + distAntenna)
                            #pepito grillo
                            # Check if previous warning or error in balise
                            pos = dfStartBaliseInfo['INIT_INDEX'][0]
                            errType = dfRoutes['ERR_TYPE'][pos]
                            errMessage = dfRoutes['ERR_MESSAGE'][pos]
                            if (errType != '') & (not ('Balise found more than once in mission' in errMessage)) :
                                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = ''
                                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = ''
                                print(color.Fore.GREEN + 'Removing ' + errType + ' from balise ' + startNidBG + ' with new MA'
                                      + color.Style.RESET_ALL)                                                      
                
                # Set as reference the last balise in JRU with linked reference distance
                iLinkedAcc += 1
                if iLinkedAcc > 0: refNidBG = listTrackNidBG[iLinkedAcc-1]
                prevNidBG = startNidBG # save previous balise in linking message
                
            # Avoid to replace qlocacc with 0 m if previously was set to a different value
            # Cause: when compounding the list of linking balises, first balise is the one coming in
            # MA message header
            #       #NID_LRBG: 
            #            NID_C (10): Reserved [391]
            #            NID_BG (14): 1656
            # This balise is needed to calculate distances but has an invalid qlocacc (set to 0)
            # This code has the mission to avoid setting a wrong qlocacc in this balise
            if len(dfTrackDistances.axes[0] > 0):
                for iqloc in range(len(listTrackQlocacc)):
                    if listTrackQlocacc[iqloc] == 0: # Value when balise in packet 15 header; not valid
                        ibgplus = listTrackBGplusIndex[iqloc]
                        dfAux = dfTrackDistances[dfTrackDistances['BGPLUSINDEX'] == ibgplus] # locate same balise in array
                        if len(dfAux.axes[0]) > 0:
                            ixAux = dfAux.index[0]
                            listTrackQlocacc[iqloc] = dfAux['Q_LOCACC'][ixAux]

            serTrackInitIndex = pandas.Series(listTrackInitIndex, dtype=int)
            serTrackDlinkAcc = pandas.Series(listTrackDlinkAcc, dtype=float)
            serTrackNidBG = pandas.Series(listTrackNidBG, dtype=str)
            serTrackQlocacc = pandas.Series(listTrackQlocacc, dtype=int)
            serTrackBGplusIndex = pandas.Series(listTrackBGplusIndex, dtype=str)
            dfTDnew = pandas.DataFrame({'INIT_INDEX': serTrackInitIndex, 
                                        'D_LINKED_ACCUM': serTrackDlinkAcc, 
                                        'NID_BG': serTrackNidBG,
                                        'Q_LOCACC': serTrackQlocacc,
                                        'BGPLUSINDEX': serTrackBGplusIndex}) 

            # Delete repeated linked nidbg in new MA
            # To keep bg more than once in a mission we use BG + Index
            # Warning: in case of MA length reduction the way of removing links in next sentences
            # is not valid; this is not important only the fact that we can have
            # some negative values in D_LIN_0, but it is not used for distance calculation
            # In case of MA cut we have to remove properly old links
            if len(dfTrackDistances.axes[0]) > 0:
                dfTrackDistances.drop( dfTrackDistances[ dfTrackDistances['BGPLUSINDEX'].apply(lambda x: x in listTrackBGplusIndex) ].index, inplace=True)  
            # Concat linking data until complete the whole mission linking 
            #if len(dfTrackDistances.axes[0]) > 0:
            #dfTrackDistances.reset_index(drop=True, inplace=True)

            dfTrackDistances = pandas.concat([dfTrackDistances, dfTDnew], ignore_index = True)

    # sort dfTrackDistances for init_index
    if len(dfTrackDistances.axes[0]) > 0:
        dfTrackDistances = dfTrackDistances.sort_values(by="INIT_INDEX")
        dfTrackDistances.reset_index(drop=True, inplace=True)


    for p in range(len(dfTrackDistances.axes[0])):
        try:
            pos = int(dfTrackDistances["INIT_INDEX"][p])
        except:
            print(dfTrackDistances)
            exit()

        if p == 0:
            dlink_0 = 0.0
        else:
            dlink_0 = dfTrackDistances["D_LINKED_ACCUM"][p] - dfTrackDistances["D_LINKED_ACCUM"][p-1] - distAntenna
            
        
        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('D_LINK_ACCUM')] = dfTrackDistances["D_LINKED_ACCUM"][p]
        # pepito grillo
                
        
        # KNIME if dfTrackDistances["D_LINKED_ACCUM"][p] == 0.0: # Case reset of distance for any reason (mode transition)
        if float(dfTrackDistances["D_LINKED_ACCUM"][p]) == 0.0: # Case reset of distance for any reason (mode transition)
               dlink_0 = 0.0

        if dlink_0 < 0: # Case MA cut
            # In this case we should change the method of removing data links
            # in dfTrackDistances above (checking againts dfTDnew)
            pepe = 0

        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('D_LINK_0')] = dlink_0
        if dlink_0 < 0:
            pepe = "pepito grillo"
        
        # Set qlocacc to NIDLRBG
        nidlrbg = dfRoutes['NID_LRBG'][pos]
        dfAux = dfTrackDistances[dfTrackDistances['NID_BG'] == nidlrbg]
        dfAux.reset_index(drop=True, inplace=True)
        if len(dfAux.axes[0]) > 0:
            # Set qlocacc at balise registers (when D_LINK_ACCUM is set)
            # Later, at function SetCalibratedRoutes, this value is extended to all registers with same LRBG
            dfRoutes.iloc[pos, dfRoutes.columns.get_loc('Q_LOCACC')] = dfAux['Q_LOCACC'][0] 


    # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function 
    firstMissionIndex = dfMission['INIT_INDEX'][0]
    lastMissionIndex = dfMission['INIT_INDEX'][len(dfMission) - 1]
    dfReturnedMission = dfRoutes.iloc[firstMissionIndex:lastMissionIndex+1].copy()
    dfReturnedMission.index = numpy.arange(0, len(dfReturnedMission["INIT_INDEX"]))
        
          
    return dfReturnedMission


# JruC_SetObestimatedRoutes(JruC_SetObestimatedRoutes)
# Set estimated route info from ONB estimations
def JruC_SetCalibratedRoutes(dfMission, dfRoutes, distAntenna):

    
    global glOdoAccuracyDistance

    # dfInputMission = pandas.DataFrame(dfMission) # Make a copy of Mission
    # !Important: Keep dfRoutes as reference to insert new values
    # dfInputRoutes = pandas.DataFrame(dfRoutes) # Make a copy of Routes (for getting slices)

    if distAntenna == None:
        print(color.Fore.RED + 'Error. Antenna distance value not valid.' +
                              color.Style.RESET_ALL)
        exit()

    # Initialization
    firstMissionIndex = dfMission['INIT_INDEX'][0]
    lastMissionIndex = dfMission['INIT_INDEX'][len(dfMission) - 1]

    # 1.- Set distances

    # Get unique lrbg-ext ids; this list is used to figure out the accumulated distance            
    # Remove those rows with LRBG = 16383 (unknown)
    # listExtendedNidlrbg = dfInputMission[~dfInputMission["LRBG_Ext"].str.contains('16383', na = False)]['LRBG_Ext'].unique() # v2.1
    # Note: previous filter discarded; see filtering at the beginning of JruC_SetRoutes() function
    # listExtendedNidlrbg = dfInputMission['LRBG_Ext'].unique()
    listExtendedNidlrbg = dfMission['LRBG_Ext'].unique()
    dist = 0.0
    odoerr = 0.0
    usafeerror = 0.0
    osafeerror = 0.0
    mtime = 0.0
    prevNidLRBG = ''
    lastCalibratingBalise = ''
    lastCalibratedDistance = 0.0
    prevSelfDlrg = 0.0
    prevDlrbg = 0.0
    prevQlrbg = None
    # Calculate distance
    for i in range (len(listExtendedNidlrbg)):
        # For each reference balise
        # dfNidExtendedLrbgInMission = dfInputMission[dfInputMission["LRBG_Ext"].str.contains(listExtendedNidlrbg[i], na = False)]
        # dfNidExtendedLrbgInMission = dfInputMission[dfInputMission["LRBG_Ext"].str.fullmatch(listExtendedNidlrbg[i], na = False)]
        dfNidExtendedLrbgInMission = dfMission[dfMission["LRBG_Ext"].str.fullmatch(listExtendedNidlrbg[i], na = False)].copy()
        dfNidExtendedLrbgInMission.index = numpy.arange(0, len(dfNidExtendedLrbgInMission["LRBG_Ext"]))

        # First and last pos in global DF (inputdf) related to the list LRG_Ext
        firstInitPos = dfNidExtendedLrbgInMission["INIT_INDEX"][0]
  
        lastInitPos = dfNidExtendedLrbgInMission["INIT_INDEX"][len(dfNidExtendedLrbgInMission["INIT_INDEX"])-1]
        if lastInitPos == 901:
            pepe = 0

        if(i>0): # Take into account last extended reference balise for accumulated distance
                    
            # Calculate distance from previous LRBG-Ext
            if prevNidLRBG == dfNidExtendedLrbgInMission['NID_LRBG'][0]:
                # Case same LRBG but orientation change 
                if(prevQlrbg != dfNidExtendedLrbgInMission['Q_DLRBG'][0]):
                            
                    if((prevQlrbg == 2) | (dfNidExtendedLrbgInMission['Q_DLRBG'][0] == 2)):
                        # Orientation change from or to value 2 (unknown)
                        # relative distance last - previous (hypothesis : no change of orientation)
                        dist += (dfNidExtendedLrbgInMission['D_LRBG'][0] - prevDlrbg)
                                
                    else:
                        # Real change of orientation to/from 1/0
                        # relative distance last + previous (change of orientation)
                        dist += (dfNidExtendedLrbgInMission['D_LRBG'][0] + prevDlrbg)
                        # Note: we could add a check to be sure that the balise has been overpassed
                                
            else:
                # Case jump from a different LRBG
                dist += dfNidExtendedLrbgInMission['D_LRBG'][0] - prevSelfDlrg
                # This value could be corrected later if needed calibration with balise

                    
        # Set accumulated distance at first position
        # First make calibration with respect to the linking info
        # if(dfNidExtendedLrbgInMission['D_LINK_ACCUM'][0]) != '':
        if(not numpy.isnan(dfNidExtendedLrbgInMission['D_LINK_ACCUM'][0])):
                    # Then calibrate 
            # take into account that odo error = 0 when D_LINK_ACCUM = 0.0
            # at start or when not linked MAs (no reference)
            #if(float(dfNidExtendedLrbgInMission['D_LINK_ACCUM'][0]) == 0.0):
            if(float(dfNidExtendedLrbgInMission['D_LINK_0'][0]) == 0.0): # pepito grillo
                odoerr = 0.0
            else: 
                odoerr = dist - dfNidExtendedLrbgInMission['D_LINK_ACCUM'][0] - distAntenna # pepito grillo

            dfRoutes.iloc[firstInitPos, dfRoutes.columns.get_loc('LOC_ERROR')] = odoerr
            
            # Calculate if odoerr out of CI
            usafeerror = 0.0
            if odoerr > 0: # Applies to underestimation
                usafeerror = odoerr - dfRoutes['L_DOUBTUNDER'][firstInitPos]
            if usafeerror < 0: usafeerror = 0.0
            dfRoutes.iloc[firstInitPos, dfRoutes.columns.get_loc('USAFE_ERROR')] = usafeerror
            osafeerror = 0.0
            if odoerr < 0: # Applies to overestimation
                osafeerror = -odoerr - dfRoutes['L_DOUBTOVER'][firstInitPos]
            if osafeerror < 0: osafeerror = 0.0
            dfRoutes.iloc[firstInitPos, dfRoutes.columns.get_loc('OSAFE_ERROR')] = osafeerror

            dist = dfNidExtendedLrbgInMission['D_LINK_ACCUM'][0]  # pepito grillo

            lastCalibratingBalise = str(dfNidExtendedLrbgInMission['NID_BG'][0])
            indexBG = lastCalibratingBalise.index('-')
            lastCalibratingBalise = lastCalibratingBalise[0:indexBG]
            lastCalibratedDistance = dist

        else:  # take into account linked value of last overpassed balise
            if lastCalibratingBalise == str(dfNidExtendedLrbgInMission['NID_LRBG'][0]):
                # odoerr = dist - (lastCalibratedDistance + dfNidExtendedLrbgInMission['D_LRBG'][0])
                # dfRoutes.iloc[firstInitPos, dfRoutes.columns.get_loc('LOC_ERROR')] = odoerr
                dist = lastCalibratedDistance + dfNidExtendedLrbgInMission['D_LRBG'][0]
                
                
        dfRoutes.iloc[firstInitPos, dfRoutes.columns.get_loc('D_ONB_ACCUM')] = dist
                

        # Figure out incremental distance in extended lrbg block
        dist0= dfNidExtendedLrbgInMission['D_LRBG'][0]
        dist1= dfNidExtendedLrbgInMission['D_LRBG'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1]
        dist += abs(dist1 - dist0)
                
                
        # Set accumulated distance at last position
        # First make calibration with respect to the linking info
        # KNIME if(dfNidExtendedLrbgInMission['D_LINK_ACCUM'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1]) != '':
        if(not numpy.isnan(dfNidExtendedLrbgInMission['D_LINK_ACCUM'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1])):
                    # Then calibrate 
            # take into account that odo error = 0 when D_LINK_ACCUM = 0.0
            # at start or when not linked MAs (no reference)
            # if(float(dfNidExtendedLrbgInMission['D_LINK_ACCUM'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1]) == 0.0):
            if(float(dfNidExtendedLrbgInMission['D_LINK_0'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1]) == 0.0):   # pepito grillo
                odoerr = 0.0
            else: 
                odoerr = dist - dfNidExtendedLrbgInMission['D_LINK_ACCUM'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1] - distAntenna  # pepito grillo
            
            dfRoutes.iloc[lastInitPos, dfRoutes.columns.get_loc('LOC_ERROR')] = odoerr

            # Calculate if odoerr out of CI
            usafeerror = 0.0
            if odoerr > 0: # Applies to underestimation
                usafeerror = odoerr - dfRoutes['L_DOUBTUNDER'][lastInitPos]
            if usafeerror < 0: usafeerror = 0.0
            dfRoutes.iloc[lastInitPos, dfRoutes.columns.get_loc('USAFE_ERROR')] = usafeerror
            osafeerror = 0.0
            if odoerr < 0: # Applies to overestimation
                osafeerror = -odoerr - dfRoutes['L_DOUBTOVER'][lastInitPos]
            if osafeerror < 0: osafeerror = 0.0
            dfRoutes.iloc[lastInitPos, dfRoutes.columns.get_loc('OSAFE_ERROR')] = osafeerror
            
            dist = dfNidExtendedLrbgInMission['D_LINK_ACCUM'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1]

            lastCalibratingBalise = str(dfNidExtendedLrbgInMission['NID_BG'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1])
            indexBG = lastCalibratingBalise.index('-')
            lastCalibratingBalise = lastCalibratingBalise[0:indexBG]
            lastCalibratedDistance = dist
                
        dfRoutes.iloc[lastInitPos, dfRoutes.columns.get_loc('D_ONB_ACCUM')] = dist

                
        # Save values of current lrbg to compare in next loop
        prevNidLRBG = dfNidExtendedLrbgInMission['NID_LRBG'][0]
        strSelfDlrbg = str(dfNidExtendedLrbgInMission['SELF_DLRBG_E'][len(dfNidExtendedLrbgInMission.axes[0]) -1])
        if ((strSelfDlrbg == '') | (strSelfDlrbg == 'nan')):
            prevSelfDlrg = 0.0
        else:
            prevSelfDlrg = float(strSelfDlrbg)

        prevDlrbg = dfNidExtendedLrbgInMission['D_LRBG'][len(dfNidExtendedLrbgInMission['D_LRBG'])-1]
        prevQlrbg = dfNidExtendedLrbgInMission['Q_DLRBG'][len(dfNidExtendedLrbgInMission["Q_DLRBG"])-1]
                
        # Figure out intermediate distances values between first and last position within the extended LRBG
        if(len(dfNidExtendedLrbgInMission['LRBG_Ext']) > 2):
            initialAccDist = dfRoutes['D_ONB_ACCUM'][dfNidExtendedLrbgInMission["INIT_INDEX"][0]]
            initialDlrbg = dfRoutes['D_LRBG'][dfNidExtendedLrbgInMission["INIT_INDEX"][0]]

            for j in range(1, len(dfNidExtendedLrbgInMission['NID_LRBG'])):
                
                intermediateInitPos = dfNidExtendedLrbgInMission["INIT_INDEX"][j] 
                moveDist = abs(dfRoutes['D_LRBG'][intermediateInitPos] - initialDlrbg)
                # Calibrate in case of linking data
                calDist = dfRoutes.iloc[intermediateInitPos, dfRoutes.columns.get_loc('D_LINK_ACCUM')]
                # KNIME if calDist != '':
                if (not numpy.isnan(calDist)):
                        
                    # KNIME if(dfNidExtendedLrbgInMission['D_LINK_ACCUM'][j] != ''):
                    if(not numpy.isnan(dfNidExtendedLrbgInMission['D_LINK_ACCUM'][j])):
                                            # take into account that odo error = 0 when D_LINK_ACCUM = 0.0
                        # at start or when not linked MAs (no reference)
                        # if(float(dfNidExtendedLrbgInMission['D_LINK_ACCUM'][j]) == 0.0):
                        if(float(dfNidExtendedLrbgInMission['D_LINK_0'][j]) == 0.0): # pepito grillo
                            odoerr = 0.0
                        else: 
                            odoerr = initialAccDist + moveDist - calDist - distAntenna # pepito grillo

                           
                        dfRoutes.iloc[intermediateInitPos, dfRoutes.columns.get_loc('LOC_ERROR')] = odoerr

                        # Calculate if odoerr out of CI
                        usafeerror = 0.0
                        if odoerr > 0: # Applies to underestimation
                            usafeerror = odoerr - dfRoutes['L_DOUBTUNDER'][intermediateInitPos]
                        if usafeerror < 0: usafeerror = 0.0
                        dfRoutes.iloc[intermediateInitPos, dfRoutes.columns.get_loc('USAFE_ERROR')] = usafeerror
                        osafeerror = 0.0
                        if odoerr < 0: # Applies to overestimation
                            osafeerror = -odoerr - dfRoutes['L_DOUBTOVER'][intermediateInitPos]
                        if osafeerror < 0: osafeerror = 0.0
                        dfRoutes.iloc[intermediateInitPos, dfRoutes.columns.get_loc('OSAFE_ERROR')] = osafeerror

                    dist = calDist  # set accumulated distance
                    lastCalibratingBalise = str(dfRoutes.iloc[intermediateInitPos, dfRoutes.columns.get_loc('NID_BG')])
                    indexBG = lastCalibratingBalise.index('-')
                    lastCalibratingBalise = lastCalibratingBalise[0:indexBG]
                    lastCalibratedDistance = dist
                    # Set new reference record to calculate relative movements
                    initialAccDist = calDist
                    initialDlrbg = dfRoutes['D_LRBG'][intermediateInitPos]

                else:
                    dist = initialAccDist + moveDist # set accumulated distance
                    
                dfRoutes.iloc[intermediateInitPos, dfRoutes.columns.get_loc('D_ONB_ACCUM')] = dist
                            

    # At this point we have to update dfMission with changes on dfRoutes
    # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function 
    # Update dfMission with D_ONB_ACCUM values
    dfMission = dfRoutes.iloc[firstMissionIndex:lastMissionIndex+1].copy()
    dfMission.index = numpy.arange(0, len(dfMission["INIT_INDEX"]))

    # KNIME dfOdoErr = dfInputMission[dfInputMission['D_LINK_ACCUM'] != '']
    # dfOdoErr = dfInputMission[dfInputMission['D_LINK_ACCUM'].notna()]
    dfOdoErr = dfMission[dfMission['D_LINK_ACCUM'].notna()].copy()
    dfOdoErr.index = numpy.arange(0, len(dfOdoErr.axes[0]))

    # Calculate odo_error and percentage
    totalOdoErr = dfOdoErr['LOC_ERROR'].sum()
    #totalOdoErr = 0.0
    for i in range(1, len(dfOdoErr.axes[0])):
        pos = dfOdoErr['INIT_INDEX'][i]
        linkedDist = dfOdoErr['D_LINK_ACCUM'][i] - dfOdoErr['D_LINK_ACCUM'][i-1]
        odoerr = dfOdoErr['LOC_ERROR'][i]
        if linkedDist == 0:
            odoerrPer = 0.0
        else: 
            odoerrPer = odoerr/linkedDist
        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ODOERROR_PERC')] = odoerrPer
        #totalOdoErr += odoerr

    # Calculate total distance
    missionDistance = 0.0
    #dfDistance = dfRoutes[dfRoutes['D_ONB_ACCUM'] == 0] pepito grillo
    # KNIME dfDistance = dfInputMission[dfInputMission['D_ONB_ACCUM'] == 0]
    # dfDistance = dfInputMission[dfInputMission['D_ONB_ACCUM'] == 0.0]

    dfDistance = dfMission[dfMission['D_ONB_ACCUM'] == 0.0].copy()
    dfDistance.reset_index(drop=True, inplace=True)
    if len(dfDistance.axes[0]) > 1: # Case of reset of distance in the middle of mission
        try:
            for z in range (1, len(dfDistance.axes[0])):
                indexAux = dfDistance['INIT_INDEX'][z]
                indexDistance = dfRoutes['INIT_INDEX'][indexAux -1]
                missionDistance += dfRoutes['D_ONB_ACCUM'][indexDistance]
            missionDistance += dist # previosly calculated dist
        except:
            missionDistance = dist # previously calculated dist


    else:
        missionDistance = dist # previously calculated dist

    #dist = missionDistance # Returned variable

    

    # Calculate mission time
    # strDate0 = dfInputMission['Date'][0] # dd/mm/yy
    strDate0 = dfMission['Date'][0] # dd/mm/yy
    # strT0 = strDate0 + ' ' + dfInputMission['Time'][0] # H:mm:ss
    strT0 = strDate0 + ' ' + dfMission['Time'][0] # H:mm:ss
    t0 = pandas.to_datetime(strT0, format='%d/%m/%y %H:%M:%S')

    # strDate1 = dfInputMission['Date'][len(dfInputMission.axes[0]) - 1] # dd/mm/yy
    strDate1 = dfMission['Date'][len(dfMission.axes[0]) - 1] # dd/mm/yy
    #strT1 = strDate1 + ' ' + dfInputMission['Time'][len(dfInputMission.axes[0]) - 1] # H:mm:ss
    strT1 = strDate1 + ' ' + dfMission['Time'][len(dfMission.axes[0]) - 1] # H:mm:ss
    t1 = pandas.to_datetime(strT1, format='%d/%m/%y %H:%M:%S')
    incrTime = t1 - t0
    mtime = incrTime.seconds
    
    
    # dfInputMission = pandas.DataFrame(dfMission) # New copy of Mission to reset previous filtering
    # 2.- Set QLOCACC and SS41 accuracy
    
    # KNIMEdfReferenceLocacc = dfInputMission[(dfInputMission['Q_LOCACC']!='')  & (dfInputMission['D_LINK_ACCUM']!='')]  # !!!!!!!!!!!!!!!
    # dfReferenceLocacc = dfInputMission[(dfInputMission['Q_LOCACC'].notna())  & (dfInputMission['D_LINK_ACCUM'].notna())]  # !!!!!!!!!!!!!!!
    dfReferenceLocacc = dfMission[(dfMission['Q_LOCACC'].notna())  & (dfMission['D_LINK_ACCUM'].notna())]  # !!!!!!!!!!!!!!!
    dfReferenceLocacc.index = numpy.arange(0, len(dfReferenceLocacc.axes[0]))
    
    # Set qlocacc accuracy (extend to all related lrbgs)
    qlocaccref = 0 # last balise reference
    for x in range (len(dfReferenceLocacc.axes[0])):
        # Balise reference values
        posRecordQlocaccRef = dfReferenceLocacc['INIT_INDEX'][x] # records with qlocacc got from MA
        qlocaccref = dfReferenceLocacc['Q_LOCACC'][x]
        
        # Important: reference balise for qlocacc: nidlrbg
        strRefBalise = dfReferenceLocacc['NID_LRBG'][x]   
        # dfFilterMessages = dfInputMission[dfInputMission['NID_LRBG'] == strRefBalise]
        dfFilterMessages = dfMission[dfMission['NID_LRBG'] == strRefBalise].copy()
        dfFilterMessages.index = numpy.arange(0, len(dfFilterMessages.axes[0]))

        listAux = dfFilterMessages['LRBG_Ext'].unique()
        # Case: same balise found more than once in mission
        if len(listAux) > 1:
            # Delete invalid records in case of balise found more than once in mission
            # Select that with minimum distance to Qlocacc Reference record
            closestLrbgExt = ''
            closestDistance = None
            recordsDistance = 0
            for z in range(len(listAux)):
                dfAux = dfFilterMessages[dfFilterMessages['LRBG_Ext'] == listAux[z]]
                dfAux.reset_index(drop=True, inplace=True)
                #if(len(dfAux.axes[0]) == 0): # not possible
                #    continue
                recordsDistance = abs(dfAux['INIT_INDEX'][0] - posRecordQlocaccRef)
                if closestDistance == None:
                    closestDistance = recordsDistance
                    closestLrbgExt = dfAux['LRBG_Ext'][0]
                elif closestDistance > recordsDistance:
                    closestDistance = recordsDistance
                    closestLrbgExt = dfAux['LRBG_Ext'][0]
                        
            dfFilterMessages = dfFilterMessages[dfFilterMessages['LRBG_Ext'] == closestLrbgExt]
            dfFilterMessages.reset_index(drop=True, inplace=True)       
        
            
        
        for p in range(len(dfFilterMessages.axes[0])):
            pos = dfFilterMessages['INIT_INDEX'][p]
            dfRoutes.iloc[pos, dfRoutes.columns.get_loc('Q_LOCACC')] = qlocaccref

    # Update dfMission with Q_LOCACC values
    dfMission = dfRoutes.iloc[firstMissionIndex:lastMissionIndex+1].copy()
    dfMission.index = numpy.arange(0, len(dfMission["INIT_INDEX"]))
    # Set SS41 accuracy
    distRef = 0 # last balise distance reference
    for x in range (len(dfReferenceLocacc.axes[0])):
        # Balise reference values
        posDistRef = dfReferenceLocacc['INIT_INDEX'][x]
        distRef = dfReferenceLocacc['D_ONB_ACCUM'][x]
        

        # Important: reference balise for ss041: record for last overpassed balise 
        # to get distance reference at this point
        strRefBalise = str(dfReferenceLocacc['NID_BG'][x])   
        ix = strRefBalise.index('-')
        strRefBalise = strRefBalise[0:ix]
        # dfFilterMessages = dfInputMission[dfInputMission['NID_LRBG'] == strRefBalise]
        dfFilterMessages = dfMission[dfMission['NID_LRBG'] == strRefBalise].copy()
        dfFilterMessages.index = numpy.arange(0, len(dfFilterMessages.axes[0]))
        if len(dfFilterMessages.axes[0]) == 0:
            print(color.Fore.YELLOW + 'Warning. Balise: ' + str(strRefBalise) +
                  ' overpassed but not found as LRBG' +
                  color.Style.RESET_ALL)
            pos = dfReferenceLocacc['INIT_INDEX'][x]
            dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = 'Warning'
            dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = 'Balise not found as LRBG'
                    
        for p in range(len(dfFilterMessages.axes[0])):
            pos = dfFilterMessages['INIT_INDEX'][p]
            # ss41Accuracy = 5 + 5%s
            s = dfFilterMessages['D_ONB_ACCUM'][p] - distRef
            qlocaccref = dfFilterMessages['Q_LOCACC'][p]
            # KNIMEif qlocaccref == '':
            if (numpy.isnan(qlocaccref)):
                continue
            
            # Important: could use the algorithm above to eliminate invalid LRBG references
            # In this case with next simple check is sufficient
            if pos < posDistRef: # To avoid invalid references when duplicated balises
                continue
            if s >= 0:
                ss41Accuracy = float(5 + 0.05*s)
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('SS41_MAX')] = -(qlocaccref + ss41Accuracy)
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('SS41_MIN')] = (qlocaccref + ss41Accuracy)



    # 3 Set odometry accuracy in last 5.000 metros as per SS26 3.6.8.3
    distLimit = glOdoAccuracyDistance # 5km reference to accumulate odo error
    # for x in reversed(range(len(dfInputMission.axes[0]))):
    for x in reversed(range(len(dfMission.axes[0]))):
        # dist = dfInputMission['D_ONB_ACCUM'][x]
        dist = dfMission['D_ONB_ACCUM'][x]
        # currentIndex = dfInputMission['INIT_INDEX'][x]
        currentIndex = dfMission['INIT_INDEX'][x]
        
        
        # In case of distance reset in the middle of mission due a transition out and back into ERMS
        # First filter for following registers (duplicated ones) to be removed
        # df5000 = pandas.DataFrame(dfInputMission) # Copy
        df5000 = dfMission.copy(deep = True) # Copy
        df5000 = df5000[df5000['INIT_INDEX'] <= currentIndex]
        df5000.reset_index(drop=True, inplace=True)
        # Second filter
        # We need to discard duplicated info; we search for distance reset (accumulated distance set to 0)
        # and removed duplicated info previous to distance 0
        for ixn in reversed(range(len(df5000.axes[0]))):
            try:
                if float(df5000['D_ONB_ACCUM'][ixn]) == 0.0:
                    ixn = df5000.index[ixn]
                    break
            except:
                continue

        df5000 = df5000.iloc[ixn:len(df5000.axes[0])]
        df5000.reset_index(drop=True, inplace=True)

        # Filter data in the last 5.000 m
        
        df5000 = df5000[(df5000['D_ONB_ACCUM'] >= (dist-distLimit)) &
                                (df5000['D_ONB_ACCUM'] <= dist)]
        df5000.reset_index(drop=True, inplace=True)

        if len(df5000.axes[0]) == 0:
            continue
        # Get starting and ending points
        l = len(df5000.axes[0])
        try:
            startingQlocacc = float(df5000['Q_LOCACC'][0])
        except:
            startingQlocacc = 0.0

        try:
            endingQlocacc = float(df5000['Q_LOCACC'][l-1])
        except:
            endingQlocacc = 0.0
        
        try:
            startingUnderReading = df5000['L_DOUBTUNDER'][0] - startingQlocacc
        except:
            startingUnderReading = 0.0
        
        try:    
            endingUnderReading = df5000['L_DOUBTUNDER'][l-1] - endingQlocacc
        except:
            endingUnderReading = 0.0

        try:
            startingOverReading = df5000['L_DOUBTOVER'][0] - startingQlocacc
        except:
            startingOverReading = 0.0

        try:
            endingOverReading = df5000['L_DOUBTOVER'][l-1] - endingQlocacc
        except:
            endingOverReading = 0.0

        # Filter data only at balise points
        # KNIME df5000 = df5000[df5000['D_LINK_ACCUM'] != ''] # Case balise detection; point when max L_DOUBTOVER & L_DOUBTUNDER
        df5000 = df5000[df5000['D_LINK_ACCUM'].notna()] # Case balise detection; point when max L_DOUBTOVER & L_DOUBTUNDER
        df5000.reset_index(drop=True, inplace=True)
        

        # Second filter to avoid duplicated data in case of distance reset in the middle of the mission
        # In this case previous registers (duplicated ones) to be removed
            

        qlocacc = 0
        # pos = dfInputMission['INIT_INDEX'][x]
        pos = dfMission['INIT_INDEX'][x]
        accumOdoUnderestimation = 0
        accumOdoOverestimation = 0

        for y in range(len(df5000)):  # not using .sum() as there are some balises with non-defined qlocacc

            qlocacc = df5000['Q_LOCACC'][y]
            # KNIME if(qlocacc == ''):
            if(numpy.isnan(qlocacc)):
                qlocacc = 0.0

                #if dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] != 'QLOCACC not defined':
                #    print(color.Fore.YELLOW + 'Warning1. QLOCACC not defined for balise: ' + str(dfRoutes['NID_LRBG'][pos]) + color.Style.RESET_ALL)
                #dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = 'Warning1'
                #dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = 'QLOCACC not defined'

            accumOdoUnderestimation += (df5000['L_DOUBTUNDER'][y] - qlocacc)
            accumOdoUnderestimation = abs(accumOdoUnderestimation)     
            
            accumOdoOverestimation += (df5000['L_DOUBTOVER'][y] - qlocacc)
            accumOdoOverestimation = abs(accumOdoOverestimation)  

        # accumOdoUnderestimation += (endingUnderReading - startingUnderReading) # discarded, not relevant
        # accumOdoOverestimation += (endingOverReading -  startingOverReading) # discarded, not relevant
        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('OVER_ACCURACY')] = accumOdoOverestimation
        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('UNDER_ACCURACY')] = accumOdoUnderestimation


    # 4.- Set under/over estimation % error
    # KNIME dfReferencePoints = dfInputMission[dfInputMission['D_LINK_ACCUM'] != '']
    # dfReferencePoints = dfInputMission[dfInputMission['D_LINK_ACCUM'].notna()]
    dfReferencePoints = dfMission[dfMission['D_LINK_ACCUM'].notna()].copy()
        # reindex is necessary for loop
    dfReferencePoints.index = numpy.arange(0, len(dfReferencePoints.axes[0]))
        
    # set variables for ub¡nder/over relative error value
    overEstimation = 0.0
    overPercentage = 0.0
    underEstimation = 0.0
    underPercentage = 0.0
    refDistance = 0.0
    qlocacc = 0.0
    dist = 0.0
    dfRefSection = pandas.DataFrame() # Initialization
        
    for n in range(1, len(dfReferencePoints.axes[0])):

        # Values (relative under/over estimation) at each section between squential BGs
        # To corralate against average velocity within the section
        dist = dfReferencePoints['D_LINK_ACCUM'][n-1]
            

        # Get df with  values within the section: from balise-0 to next balise-0
        ixSofSection = dfReferencePoints['INIT_INDEX'][n-1]
        ixEofSection = dfReferencePoints['INIT_INDEX'][n]
        dfRefSection = dfRoutes.iloc[ixSofSection + 1:ixEofSection+1] # from balise-0 to balise-0 excluding first one
        dfRefSection.index = numpy.arange(0, len(dfRefSection.axes[0]))

        for m in range(len(dfRefSection.axes[0])):

            qlocacc = dfRefSection['Q_LOCACC'][m]
            # KNIME if qlocacc == '':
            if (numpy.isnan(qlocacc)):
                pos = dfRefSection['INIT_INDEX'][m]
                if dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] != 'QLOCACC not defined':
                    print(color.Fore.YELLOW + 'Warning2. QLOCACC not defined for balise: ' + str(dfRoutes['NID_LRBG'][pos]) + color.Style.RESET_ALL)
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = 'Warning2'
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = 'QLOCACC not defined'
                continue  # cant define over/under estimation
                
            refDistance = dfRefSection['D_ONB_ACCUM'][m] - dist
            if int(refDistance) <= qlocacc: 
                continue  # not reliable when distance less t

            overEstimation = dfRefSection['L_DOUBTOVER'][m] - qlocacc
            overPercentage = overEstimation/refDistance
            
            
            underEstimation = dfRefSection['L_DOUBTUNDER'][m] - qlocacc
            underPercentage = underEstimation/refDistance
            # Insert values in routes dataframe
            pos = dfRefSection['INIT_INDEX'][m]
            dfRoutes.iloc[pos, dfRoutes.columns.get_loc('OVEREST_ERROR')] = overPercentage
            dfRoutes.iloc[pos, dfRoutes.columns.get_loc('UNDEREST_ERROR')] = underPercentage

    # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function 
    dfReturnedMission = dfRoutes.iloc[firstMissionIndex:lastMissionIndex+1].copy()
    dfReturnedMission.index = numpy.arange(0, len(dfReturnedMission["INIT_INDEX"]))

    listReturn = [missionDistance, totalOdoErr, mtime, dfReturnedMission] # output with summary info
    
    return listReturn # return accumulated distance and odoerr for route summary info



# JruC_SetSSP(dfMission, dfRoutes)
# Compiled JRU File
# Set SSP vaues into de route info from packet 27
def JruC_SetSSP(dfMission, dfRoutes):
            
    global glLocale

    dfInputMission = pandas.DataFrame(dfMission) # Make a copy of Mission
    # !Important: Keep dfRoutes as reference to insert new values


    dfSSPInfoinMission = pandas.DataFrame() # df with ssp info for the whole mission)
    dfSSPInfoInMA = pandas.DataFrame() # df for ssp info for each MA in mission


    # get matrix with linked balises and distances, from packt 27 info
    for x in range(len(dfInputMission['SSP'])):
        strSSP = dfInputMission['SSP'][x]
        strBalisePackets =str(dfInputMission['Balise Packets'][x])
        if str(strSSP) == 'nan':
            strSSP = ""
        elif (len(strBalisePackets) > 0) and (strBalisePackets != 'nan'):
        
            if '136' in strBalisePackets:
                # Pending for Level 2; packets from RBC
                strSSP = "" # Infill data for reference balise; discarded

        if strSSP != "":
            boolSSPInMA = True
            iSSPInMA=0
            indexBG = 0
            indexDref = 0
            indexVref = 0
            startInitIndex = 0
            listStartInitIndexInMA = []
            listNidbgInMA = []
            listDrefInMA = []
            listVrefInMA = []
            
            strSSP = strSSP.strip()
            startInitIndex = dfInputMission['INIT_INDEX'][x]
            # Variable to be used for distance reference when jumping between MAs
            # In that cases SSP of both MAs overlaps
            Dref = 0.0 
            Vref = 0.0
            prevDref = 0.0
            # format of strBgLinks
            # 1342-0,00:160-3.437:120-291,00:160-4.650:127
            # where first value, 1342, is the NID_BG of reference balise
            # nid_bg-distance from NidBG:V_STATIC-distance from NidBG:V_STATIC

            try:
                indexBG = strSSP.index('-')
                nidBG = strSSP[0:indexBG]
                
            except:
                boolSSPInMA = False

            indexBG = indexBG + 1
                
            while boolSSPInMA == True:
                
                listStartInitIndexInMA.insert(iSSPInMA, startInitIndex)
                listNidbgInMA.insert(iSSPInMA, nidBG)
                
                try:
                    strSSP = strSSP[indexBG:len(strSSP)]
                    indexDref = strSSP.index(':')
                except:
                    boolSSPInMA = False
                    continue

                Dref = strSSP[0:indexDref]

                if glLocale == 'SPA': 
                    Dref = Dref.replace('.','')
                    Dref = float(Dref.replace(",", "."))
                elif glLocale == 'ENG': Dref = float(Dref.replace(",", ""))
                else:
                    print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
                    exit(1)

                
                Dref += prevDref
                prevDref = Dref
                listDrefInMA.insert(iSSPInMA, Dref)

                try:
                    indexVref = strSSP.index('-')
                except:
                    boolSSPInMA = False # last balise
                    indexVref = len(strSSP)

                Vref = strSSP[indexDref+1:indexVref]

                if glLocale == 'SPA': 
                    Vref = float(Vref.replace(",", "."))
                elif glLocale == 'ENG': Vref = float(Vref.replace(",", ""))
                else:
                    print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
                    exit(1)

                
                listVrefInMA.insert(iSSPInMA, Vref)
                iSSPInMA += 1
                indexBG = indexVref + 1

            serStartInitIndexInMA = pandas.Series(listStartInitIndexInMA, dtype = int)
            serNidbgInMA = pandas.Series(listNidbgInMA, dtype = str)
            serDrefInMA = pandas.Series(listDrefInMA, dtype = float)
            serVrefInMA = pandas.Series(listVrefInMA, dtype = float)

            # Dataframe with info from packet 27
            dfSSPInfoInMA = pandas.DataFrame({'INIT_INDEX': serStartInitIndexInMA,'SSPBG': serNidbgInMA, 'SSPDREF': serDrefInMA, 'SSPVREF': serVrefInMA})

            # concat ssp in MA with ssp in Mission
            dfSSPInfoinMission = pandas.concat([dfSSPInfoinMission, dfSSPInfoInMA], ignore_index=True)

    # Set V_STATIC  
    if len(dfSSPInfoinMission.axes[0] > 0):

        # Search for SSP info in balises
        listSspBalises = dfSSPInfoinMission['SSPBG'].unique()               

        for x in range(len(listSspBalises)):

            strNidLRBG = listSspBalises[x]
            if strNidLRBG == '523':
                pepe = 0

            # Flter SSP records with NID_LRBG
            dfTemp = dfSSPInfoinMission[dfSSPInfoinMission['SSPBG'] == strNidLRBG]
            dfTemp.index = numpy.arange(0, len(dfTemp.axes[0]))
            # Format
            #      INIT_INDEX SSPBG  SSPDREF  SSPVREF
            # 7         414   705      0.0     60.0
            # 8         414   705    226.0     45.0
            # 9         414   705    326.0     50.0
            # 10        414   705   5426.0     70.0
            
            # Check if balise more than once in Mission
            # numTimes = len(dfTemp['INIT_INDEX'].unique())
            # if numTimes > 1:
                # To check: make a loop with different blocks
                # for i in range(numTimes)
            listInitIndexDF1 = dfTemp['INIT_INDEX'].unique()
            

            for p in range(len(listInitIndexDF1)):

                df1 = dfTemp[dfTemp['INIT_INDEX'] == listInitIndexDF1[p]]
                df1.index = numpy.arange(0, len(df1.axes[0]))

                if len(df1.axes[0]) == 0:
                        break  # For p
                
                currentLrbgInitIndex = df1['INIT_INDEX'][0]

                df2 = dfInputMission[dfInputMission['INIT_INDEX'] >= currentLrbgInitIndex]
                df2.index = numpy.arange(0, len(df2.axes[0]))

                
                nextLrbgInitIndex = currentLrbgInitIndex
                if len(df2.axes[0]) > 1:
                    nextLrbgInitIndex = df2['INIT_INDEX'][len(df2.axes[0]) - 1]
                    # search for next SSP record
                    for i in range(1, len(df2.axes[0])):

                        if str(df2['SSP'][i]) == 'nan' or str(df2['SSP'][i]) == '' :
                            continue
                        elif '136' in str(df2['Balise Packets'][i]):
                            continue
                        else:

                            nextLrbgInitIndex = df2['INIT_INDEX'][i]
                            break

                df2 = df2[df2['INIT_INDEX'] < nextLrbgInitIndex]
                df2.index = numpy.arange(0, len(df2.axes[0]))

                if len(df2.axes[0]) > 0:
                    
                    prevIndexN = 0
                    refSspVstatic = 0
                    prevSspVstatic = 0
                    for m in range(len(df1.axes[0])):  # SSP
                        # Take as reference distance first record with this LRBG
                        startDistance = df2['D_ONB_ACCUM'][0] # - df2['D_LRBG'][0]
                        prevSspVstatic = refSspVstatic
                        startSspInitIndex = df1['INIT_INDEX'][m]
                        refSspDist = df1['SSPDREF'][m]
                        refSspVstatic = df1['SSPVREF'][m]

                        if df2['INIT_INDEX'][0] < startSspInitIndex: # Out of scope of this SSP
                            break
                        

                        if m > 0:
                            for n in range(prevIndexN, len(df2.axes[0])): # Mission              
                                
                                #if df2['D_ONB_ACCUM'][n] == 0.0: # distance reset for diferent reasosns (i.e. not linked MAs)
                                #    break  # For n

                                runDistance = df2['D_ONB_ACCUM'][n] - startDistance
                                prevIndexN = n + 1
                                                    
                                if runDistance <= refSspDist:
                                    pos = df2['INIT_INDEX'][n]
                                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc('V_STATIC')] = prevSspVstatic
                                
                                else:
                                
                                    pos = df2['INIT_INDEX'][n]
                                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc('V_STATIC')] = refSspVstatic
                                    prevSspVstatic = refSspVstatic
                                        
                                    if m < (len(df1.axes[0]) -1): # last SSP section
                                        break
                        elif len(df1.axes[0]) == 1:

                            for n in range(len(df2.axes[0])): # Mission      

                                #if df2['D_ONB_ACCUM'][n] == 0.0: # distance reset for diferent reasosns (i.e. not linked MAs)
                                #    break  # For n

                                pos = df2['INIT_INDEX'][n]
                                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('V_STATIC')] = refSspVstatic

    # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function 
    firstMissionIndex = dfMission['INIT_INDEX'][0]
    lastMissionIndex = dfMission['INIT_INDEX'][len(dfMission) - 1]
    dfReturnedMission = dfRoutes.iloc[firstMissionIndex:lastMissionIndex+1]
    dfReturnedMission.index = numpy.arange(0, len(dfReturnedMission["INIT_INDEX"]))

    return dfReturnedMission

# JruC_SetVmrsp(dfMission, dfRoutes)
# Compiled JRU File
# Set V_PERM values figured out on OB
def JruC_SetVmrsp(dfMission, dfRoutes):
    dfInputMission = pandas.DataFrame(dfMission) # Make a copy of Mission
    # !Important: Keep dfRoutes as reference to insert new values

    # get Vmrsp values from input data
    strVmrsp = ''
    strRefVmrsp = ''
    vmrsp = 0
    for x in range(len(dfInputMission['V_PERM'])):
        
        strVmrsp = str(dfInputMission['V_PERM'][x])
        pos = dfInputMission['INIT_INDEX'][x]
        if strVmrsp == 'nan':
            strVmrsp = ''
            if strRefVmrsp != '':
                # In case M_MODE = National System (M_MODE = 13), set to 0
                #if(str(dfRoutes.iloc[pos, dfRoutes.columns.get_loc('M_MODE')]) == '13'):vmrsp = 0
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('V_PERM')] = vmrsp

        elif strVmrsp == '1023.0':  # Unknown
            dfRoutes.iloc[pos, dfRoutes.columns.get_loc('V_PERM')] = None
            strRefVmrsp = ''
        else:
            strRefVmrsp = strVmrsp
            vmrsp = int(float(strVmrsp))
            continue               

    # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function 
    firstMissionIndex = dfMission['INIT_INDEX'][0]
    lastMissionIndex = dfMission['INIT_INDEX'][len(dfMission) - 1]
    dfReturnedMission = dfRoutes.iloc[firstMissionIndex:lastMissionIndex+1]
    dfReturnedMission.index = numpy.arange(0, len(dfReturnedMission["INIT_INDEX"]))

    return dfReturnedMission

# JruC_SetAccel(dfMission, dfRoutes)
# Compiled JRU File
# Figure out acceleration
def JruC_SetAccel(dfMission, dfRoutes):

    dfInputMission = pandas.DataFrame(dfMission) # Make a copy of Mission
    # !Important: Keep dfRoutes as reference to insert new values

    accel = 0.0
    for x in range(1, len(dfInputMission.axes[0])):
        
        v1 = float(dfInputMission['V_TRAIN'][x])
        v0 = float(dfInputMission['V_TRAIN'][x-1])
        t1 = float(dfInputMission['T_ONB_ACCUM'][x])
        t0 = float(dfInputMission['T_ONB_ACCUM'][x-1])
        pos = dfInputMission['INIT_INDEX'][x]
        try:
            accel = (v1-v0)*1000/3600/(t1-t0)
            dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ACCEL')] = accel
        except:
            dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ACCEL')] = 0.0
            continue               

    # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function 
    firstMissionIndex = dfMission['INIT_INDEX'][0]
    lastMissionIndex = dfMission['INIT_INDEX'][len(dfMission) - 1]
    dfReturnedMission = dfRoutes.iloc[firstMissionIndex:lastMissionIndex+1]
    dfReturnedMission.index = numpy.arange(0, len(dfReturnedMission["INIT_INDEX"]))

    return dfReturnedMission

# JruC_SetSlope(dfMission, dfRoutes)
# Compiled JRU File
# Figure out acceleration
def JruC_SetSlope(dfMission, dfRoutes):

    global glSS41Intercept

    dfInputMission = pandas.DataFrame(dfMission) # Make a copy of Mission
    # !Important: Keep dfRoutes as reference to insert new values
    # KNIME dfReferencePoints = dfInputMission[dfInputMission['D_LINK_ACCUM'] != '']
    dfReferencePoints = dfInputMission[dfInputMission['D_LINK_ACCUM'].notna()]
    dfReferencePoints.index = numpy.arange(0, len(dfReferencePoints.axes[0]))

    # 1) Value of slope in the whole section between balises
    for n in range(1, len(dfReferencePoints.axes[0])):

            overslope = 0.0
            underslope = 0.0
            # Get df with  values within the section: from balise-0 to next balise-0
            ixSofSection = dfReferencePoints['INIT_INDEX'][n-1]
            ixEofSection = dfReferencePoints['INIT_INDEX'][n]
            dfRefSection = pandas.DataFrame(dfRoutes.iloc[ixSofSection + 1:ixEofSection+1]).copy() # Copy - from balise-0 to balise-0 excluding first one
            dfRefSection.index = numpy.arange(0, len(dfRefSection.axes[0]))

            l = len(dfRefSection.axes[0])
            if l < 1:
                continue # Not valida data
            
            # CRITERIA: Take into account than as observed in hasler JRU, the over/under estimation gets
            # their maximum values when reaches the new balise but the reset occurs a bit later
            # so we use the last point at df for maximum over/under estimation and
            # search the minimun value as reference for intercept value             
            
            # Overslope
            ixreset = dfRefSection['L_DOUBTOVER'].idxmin() 

            # Get distance at end point of section (maximum over/under estimation)
            s1 = float(dfRefSection['D_ONB_ACCUM'][l-1])
            s0 = float(dfRefSection['D_ONB_ACCUM'][ixreset]) # read note above

            if ((s1-s0) == 0.0):
                for pos in range(ixSofSection + 1, ixEofSection + 1):
                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc('OVERSLOPE')] = 0.0
                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc('UNDERSLOPE')] = 0.0
                continue    # zero division
            
            # Get qlocacc at each point
            # Important! In case there is no qlocacc, the take the same value (0)
            try:
                qlocacc1 = float(dfRefSection['Q_LOCACC'][l-1])
                if numpy.isnan(qlocacc1):
                    qlocacc1 = 0.0
            except:
                qlocacc1 = 0.0
            
            try:
                qlocacc0 = float(dfRefSection['Q_LOCACC'][ixreset])
                if numpy.isnan(qlocacc0):
                    qlocacc0 = 0.0
            except:
                qlocacc0 = 0.0
            
            y1 = float(dfRefSection['L_DOUBTOVER'][l-1]) - qlocacc1
            y0 = float(dfRefSection['L_DOUBTOVER'][ixreset]) - qlocacc0
            overslope = (y1-y0)/(s1-s0)
        
            # Underslope
            # took same reset point 
            y1 = float(dfRefSection['L_DOUBTUNDER'][l-1]) - qlocacc1
            y0 = float(dfRefSection['L_DOUBTUNDER'][ixreset]) - qlocacc0
            underslope = (y1-y0)/(s1-s0)

            for pos in range(ixSofSection + 1, ixEofSection + 1):
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('OVERSLOPE')] = overslope
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('UNDERSLOPE')] = underslope  

    # 2) Value of slope at each point: spot slope
    for m in range(1, len(dfInputMission.axes[0])):

        spotOverReading = 0.0
        spotUnderReading = 0.0
        spotOrSlope = 0.0
        spotUrSlope = 0.0
        qlocacc = 0
        dLrbg = 0.0

        lDoubtOver = dfInputMission['L_DOUBTOVER'][m]
        lDoubtUnder = dfInputMission['L_DOUBTUNDER'][m]
        qlocacc = dfInputMission['Q_LOCACC'][m]
        dLrbg = dfInputMission['D_LRBG'][m]

        #if qlocacc != '' and dLrbg > 0:
        if (not numpy.isnan(qlocacc)) and dLrbg > 0:
            spotOverReading = lDoubtOver - qlocacc
            spotUnderReading = lDoubtUnder - qlocacc
            try:
                spotOrSlope = spotOverReading/dLrbg
                # Knime if spotOrSlope == 0.0 or spotOrSlope == 0: spotOrSlope = ''
                if spotOrSlope == 0.0 or spotOrSlope == 0: spotOrSlope = numpy.nan
            except:
                # Knime spotOrSlope = ''
                spotOrSlope = numpy.nan
            
            try:
                spotUrSlope = spotUnderReading/dLrbg
                # Knime if spotUrSlope == 0.0 or spotUrSlope == 0: spotUrSlope = ''
                if spotUrSlope == 0.0 or spotUrSlope == 0: spotUrSlope = numpy.nan
            except:
                # Knime spotUrSlope = ''
                spotUrSlope = numpy.nan
        
        pos = dfInputMission['INIT_INDEX'][m]
        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('SPOTORSLOPE')] = spotOrSlope
        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('SPOTURSLOPE')] = spotUrSlope 
        
    # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function 
    firstMissionIndex = dfMission['INIT_INDEX'][0]
    lastMissionIndex = dfMission['INIT_INDEX'][len(dfMission) - 1]
    dfReturnedMission = dfRoutes.iloc[firstMissionIndex:lastMissionIndex+1]
    dfReturnedMission.index = numpy.arange(0, len(dfReturnedMission["INIT_INDEX"]))
            

    return dfReturnedMission


# JruC_Statistics(dfMissionsSummary, dfRoutes)
# Compiled JRU File
# Get missions info: indexes SoM / EoM, route distance, ...
# Return a DF with a summary of statistics

def JruC_Statistics(dfMissionsSummary, dfRoutes):   
    global glLocale

    global glPh2Progress
    global glPh2NumberRows
    global glPh2Percentage
    global glPh2Prevpercentage
    
    glPh2Progress = 0.0
    glPh2NumberRows = 0
    glPh2Percentage = 0.0
    glPh2Prevpercentage = 0.0
    
    # Remove other nidengines
    #inputdf = pandas.DataFrame(dfRoutes) # local copy 
    #inputdf = dfRoutes.copy()

    # Lists for output dataframes
    # new dataframes for output
    dfStatistics = pandas.DataFrame() 
   
    outputListOverMean = []
    outputListOverslopeCorrVel = []
    #outputListOverCorrAccel = []
    # outputListUnderMean = []
    outputListUnderslopeCorrVel = []
    #outputListUnderCorrAccel = [] 
    outputListTotalErrorMean = []   
    outputListerrOdo = []
    outputlistDlinkMean = []
    outputListerrOdoIntercept = []
    outputListerrDlinkMean = []
    outputListerrDlinkStd = []
    outputListerrDlinkMax = []
    outputListerrDlinkMin = []
    outputListUsafeerrorMax = []
    outputListOsafeerrorMax = []
    #outputListUsafeerrorMean = []
    #outputListOsafeerrorMean = []
    outputListOverSlopeMean = []
    outputListOverSlopeMax = []
    outputListUnderSlopeMean = []
    outputListUnderSlopeMax = []
    outputListDistance = []
    outputListNumBG = []
    outputListOdoAccuracy = []

    for x in range(len(dfMissionsSummary.axes[0])):
        ixSoM = dfMissionsSummary['IxSoM'][x]
        ixEoM = dfMissionsSummary['IxEoM'][x]

        # Filter by each mission
        #dfSingleMission = inputdf.iloc[ixSoM:ixEoM+1].copy()
        dfSingleMission = dfRoutes.iloc[ixSoM:ixEoM+1].copy()
        
        # Figure out % of over/under reading with respect to the distance
        # Filter relevant data in single mission (those records with linked info)
        # KNIME dfReferencePoints = dfSingleMission[dfSingleMission['D_LINK_ACCUM'] != '']
        dfReferencePoints = dfSingleMission[dfSingleMission['D_LINK_ACCUM'].notna()]
        # reindex is necessary for loop
        dfReferencePoints.index = numpy.arange(0, len(dfReferencePoints.axes[0]))
        
        # set variables for odoerr  correlation
        overEstimation = 0.0
        overSlope = 0.0
        underEstimation = 0.0
        underSlope = 0.0
        refDistance = 0.0
        qlocacc = 0.0
        listOverE = []
        listOverSlope = []
        listUnderSlope = []
        listUnderE = []
        serOverSlope = pandas.Series()
        serUnderSlope = pandas.Series()
        listAccel = []
        #serAcceleration = pandas.Series()
        listVtrain = []
        #serVtrain = pandas.Series()
        listIx = 0
        dist = 0.0
        dfRefSection = pandas.DataFrame() # Initialization
        
        for n in range(1, len(dfReferencePoints.axes[0])):

            # Values (relative under/over estimation) at each section between squential BGs
            # To corralate against average velocity within the section
            dist = dfReferencePoints['D_LINK_ACCUM'][n-1]
            

            # Get df with  values within the section: from balise-0 to next balise-0
            ixSofSection = dfReferencePoints['INIT_INDEX'][n-1]
            ixEofSection = dfReferencePoints['INIT_INDEX'][n]
            #dfRefSection = inputdf.iloc[ixSofSection + 1:ixEofSection+1] # from balise-0 to balise-0 excluding first one
            dfRefSection = dfRoutes.iloc[ixSofSection + 1:ixEofSection+1].copy() # from balise-0 to balise-0 excluding first one
            #dfRefSection = dfRefSection[dfRefSection['D_LINK_ACCUM'] != ''] # !!!!!!!!!!! Only on balise registers
            dfRefSection.index = numpy.arange(0, len(dfRefSection.axes[0]))

            for m in range(len(dfRefSection.axes[0])):

                accel = dfRefSection['ACCEL'][m]
                vel = dfRefSection['V_TRAIN'][m]

                qlocacc = dfRefSection['Q_LOCACC'][m]
                # KNIME if qlocacc == '':
                if (numpy.isnan(qlocacc)):
                    pos = dfRefSection['INIT_INDEX'][m]
                    if dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] != 'QLOCACC not defined':
                        print(color.Fore.YELLOW + 'Warning3. QLOCACC not defined for balise: ' + str(dfRoutes['NID_LRBG'][pos]) + color.Style.RESET_ALL)
                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_TYPE')] = 'Warning3'
                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ERR_MESSAGE')] = 'QLOCACC not defined'
                    continue  # cant define over/under estimation
                
                refDistance = dfRefSection['D_ONB_ACCUM'][m] - dist
                if int(refDistance) <= qlocacc:
                    continue  # not reliable wen distance less than accuracy 

                overEstimation = dfRefSection['L_DOUBTOVER'][m] - qlocacc
                listOverE.insert(listIx, overEstimation)
                underEstimation = dfRefSection['L_DOUBTUNDER'][m] - qlocacc
                listUnderE.insert(listIx, underEstimation)

                overSlope = dfRefSection['OVERSLOPE'][m]
                listOverSlope.insert(listIx, overSlope)


                underSlope = dfRefSection['UNDERSLOPE'][m]
                listUnderSlope.insert(listIx, underSlope)

                listAccel.insert(listIx, accel)
                listVtrain.insert(listIx, vel)
            
                listIx += 1
        
        # For each mission next statistics:
        # 1) Ratio under/over estimation divided by distance
        #   a) mean
        #   b) correlation with vel
        #   c) correlation with acceleration

        

        # Create auxiliary series for correlation
        serOverSlope = pandas.Series(listOverSlope, dtype=float)
        serUnderSlope = pandas.Series(listUnderSlope, dtype=float)
        
        #serAcceleration = pandas.Series(listAccel, dtype=float)
        serVtrain = pandas.Series(listVtrain, dtype=float)

        # over/under estimation max, min, average
        #overpMean = serOverPercentage.mean()
        #underpMean = serUnderPercentage.mean()

        # overslope/underslope correlation 
        if (serVtrain.std() == 0.0): # in this case Pearson correlation nan, warning
            corrOverslopeVel = numpy.nan
            corrUnderslopeVel = numpy.nan
        else:
            
            # over correlation with vtrain
            if (serOverSlope.std() == 0.0): # in this case Pearson correlation nan, warning
                corrOverslopeVel = numpy.nan
            else:   
                corrOverslopeVel = serOverSlope.corr(serVtrain, method = 'pearson')

            # under correlation with vtrain
            if (serUnderSlope.std() == 0.0): # in this case Pearson correlation nan, warning
                corrUnderslopeVel = numpy.nan
            else:  
                corrUnderslopeVel = serUnderSlope.corr(serVtrain, method = 'pearson')

        # output lists
        #outputListOverMean.insert(x, overpMean)
        #outputListUnderMean.insert(x, underpMean)
        outputListOverslopeCorrVel.insert(x, corrOverslopeVel)
        #outputListOverCorrAccel.insert(x, corrOverAccel)
        outputListUnderslopeCorrVel.insert(x, corrUnderslopeVel)
        #outputListUnderCorrAccel.insert(x, corrUnderAccel)


        # 2) odo error, over slope, underslope mean, osafeerror max and usafe error max
        # KNIME dfTemp = dfSingleMission[dfSingleMission['LOC_ERROR'] != '']
        dfTemp = dfSingleMission[dfSingleMission['LOC_ERROR'].notna()]
        totalErrorMean = dfTemp['LOC_ERROR'].mean()
        # KNIME dfTemp = dfSingleMission[dfSingleMission['OVERSLOPE'] != '']
        dfTemp = dfSingleMission[dfSingleMission['OVERSLOPE'].notna()]
        overSlopeMean = dfTemp['OVERSLOPE'].mean()
        overSlopeMax = dfTemp['OVERSLOPE'].max()
        # KNIME dfTemp = dfSingleMission[dfSingleMission['UNDERSLOPE'] != '']
        dfTemp = dfSingleMission[dfSingleMission['UNDERSLOPE'].notna()]
        underSlopeMean = dfTemp['UNDERSLOPE'].mean()
        underSlopeMax = dfTemp['UNDERSLOPE'].max()
        # KNIME dfTemp = dfSingleMission[dfSingleMission['OSAFE_ERROR'] != '']
        dfTemp = dfSingleMission[dfSingleMission['OSAFE_ERROR'].notna()]
        #osafeerrorMean = dfTemp['OSAFE_ERROR'].mean()
        osafeerrorMax = dfTemp['OSAFE_ERROR'].max()
        # KNIME dfTemp = dfSingleMission[dfSingleMission['USAFE_ERROR'] != '']
        dfTemp = dfSingleMission[dfSingleMission['USAFE_ERROR'].notna()]
        #usafeerrorMean = dfTemp['USAFE_ERROR'].mean()
        usafeerrorMax = dfTemp['USAFE_ERROR'].max()

        # Odo error regression to get slope and correct antenna distance to frontend
        # Linear regression: https://www.geeksforgeeks.org/solving-linear-regression-in-python/
        # Pendiente = Sxy/Sxx   
        # donde Sxy y Sxx son la covarianza de la muestra y la varianza de la muestra respectivamente.
        # Sxy = (x-xmean)*(y-ymean)
        # Sxx = (x-xmean)^2
        # Intersección = media y – pendiente* x media
        
        # Regression: total error vs distances between balises
        # Extraction of eOdo increase due to distance travelled
        # KNIME dfTemp = dfSingleMission[dfSingleMission['LOC_ERROR'] != '']
        dfTemp = dfSingleMission[dfSingleMission['LOC_ERROR'] .notna()]
        dfTemp.reset_index(drop=True, inplace=True)
        if len(dfTemp.axes[0]) > 0:
            serY = pandas.Series(dfTemp['LOC_ERROR'], name = 'Y')
            serY.reset_index(drop=True, inplace=True)
            ymean = serY.mean()
            serX = pandas.Series(dfTemp['D_LINK_0'], name = 'X')
            serX.reset_index(drop=True, inplace=True)
            xmean = serX.mean()
            dlinKMean = xmean
            
            

            dfSerxy = pandas.DataFrame({'Y': serY, 'X': serX})
            serSxy = dfSerxy.apply(utl_GetSxy, args=[ymean, xmean], axis=1)
            serSxx = dfSerxy.apply(utl_GetSxx, args=[xmean], axis=1)
            sxy = serSxy.sum()
            sxx = serSxx.sum()

            if sxx != 0.0:
                errOdoSlope = sxy/sxx
                errOdoIntercept = ymean - errOdoSlope*xmean
                # Option 1 to calculate odo error; although it is a valid option, used as totalErrorMean is not an estimation but a measure
                errOdo = totalErrorMean - errOdoIntercept 
            else:
                errOdoSlope = numpy.nan
                errOdoIntercept = numpy.nan
                errOdo = numpy.nan
                dlinKMean = numpy.nan

            # Option 2 to calculate odo error; not used as this is not a measure but an estimation
            # Option 1 and optinon 2 are equivalent and the final result is the same
            #serErrOdo = serX.apply(utl_GetMxOdo, args=[errOdoSlope])
            #errOdo = serErrOdo.mean() 
        
        
        else:
            errOdoSlope = numpy.nan
            errOdoIntercept = numpy.nan
            errOdo = numpy.nan
            dlinKMean = numpy.nan

        

        # Second estimation: antenna + dlink error
        listY = []
        for z in range(len(dfTemp.axes[0])):
            pos = dfTemp['INIT_INDEX'][z]
            # remove eOdo from error (remove error due to distance travelled)
            if numpy.isnan(errOdoSlope):
                y = numpy.nan
                # Insert DLINKERR and ODOERR in routes
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('DLINKERR')] = numpy.nan
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ODOERR')] = numpy.nan
            else:
                y = dfTemp['LOC_ERROR'][z] - errOdoSlope * dfTemp['D_LINK_0'][z]
                # Insert DLINKERR and ODOERR in routes 
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('DLINKERR')] = y
                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('ODOERR')] = errOdoSlope * dfTemp['D_LINK_0'][z]
                
            listY.insert(z, y)

        
        serY = pandas.Series(listY, name = 'Y')
        serY.reset_index(drop=True, inplace=True)
        if ((len(serY) == 0) or (len(serY[numpy.isnan(serY)]) > 0)):
            errDlinkMean = numpy.nan
            errDlinkStd = numpy.nan
            errDlinkMax = numpy.nan
            errDlinkMin = numpy.nan
        else:
            errDlinkMean = serY.mean() # eDlink + eAntenna
            errDlinkStd = serY.std() # deviation due to dlink error
            errDlinkMax = serY.max()
            errDlinkMin = serY.min()
            

        # output lists
        outputListTotalErrorMean.insert(x,totalErrorMean)
        outputListerrOdo.insert(x, errOdo)
        outputlistDlinkMean.insert(x, dlinKMean)
        outputListerrOdoIntercept.insert(x, errOdoIntercept)
        outputListerrDlinkMean.insert(x,errDlinkMean)
        outputListerrDlinkStd.insert(x, errDlinkStd)
        outputListerrDlinkMax.insert(x, errDlinkMax)
        outputListerrDlinkMin.insert(x, errDlinkMin)
        outputListOverSlopeMean.insert(x, overSlopeMean)
        outputListOverSlopeMax.insert(x, overSlopeMax)
        outputListUnderSlopeMean.insert(x, underSlopeMean)
        outputListUnderSlopeMax.insert(x, underSlopeMax)
        #outputListOsafeerrorMean.insert(x, osafeerrorMean)
        outputListOsafeerrorMax.insert(x, osafeerrorMax)
        #outputListUsafeerrorMean.insert(x, usafeerrorMean)
        outputListUsafeerrorMax.insert(x, usafeerrorMax)
        outputListDistance.insert(x, dfMissionsSummary['Distance'][x])
        outputListNumBG.insert(x, dfMissionsSummary['NumBG'][x])
        
        if ((not numpy.isnan(dlinKMean)) and (dlinKMean != 0.0)):
            odoAccuracy = 1-abs(errOdo/dlinKMean)
        else:
            odoAccuracy = numpy.nan

        outputListOdoAccuracy.insert(x, odoAccuracy)


    # end for x
    
    # Create series to insert in dfStatistics
    #outputSerOvermean = pandas.Series(outputListOverMean, dtype=float)
    outputSerOverCorrVel = pandas.Series(outputListOverslopeCorrVel, dtype=float)
    #outputSerOverCorrAccel = pandas.Series(outputListOverCorrAccel, dtype=float)

    #outputSerUndermean = pandas.Series(outputListUnderMean, dtype=float)
    outputSerUnderCorrVel = pandas.Series(outputListUnderslopeCorrVel, dtype=float)
    #outputSerUnderCorrAccel = pandas.Series(outputListUnderCorrAccel, dtype=float)

    outputSerTotalErrorMean = pandas.Series(outputListTotalErrorMean, dtype = float)
    outputSereOdo = pandas.Series(outputListerrOdo, dtype = float)
    outputSerDlink = pandas.Series(outputlistDlinkMean, dtype=float)
    outputSereOdoIntercept = pandas.Series(outputListerrOdoIntercept, dtype = float)
    outputSereDlinkMean = pandas.Series(outputListerrDlinkMean, dtype = float)
    outputSereDlinkStd = pandas.Series(outputListerrDlinkStd, dtype = float)
    outputSereDlinkMax = pandas.Series(outputListerrDlinkMax, dtype = float)
    outputSereDlinkMin = pandas.Series(outputListerrDlinkMin, dtype = float)
    outputSerOverSlopeMean = pandas.Series(outputListOverSlopeMean, dtype = float)
    outputSerOverSlopeMax = pandas.Series(outputListOverSlopeMax, dtype = float)
    outputSerUnderSlopeMean = pandas.Series(outputListUnderSlopeMean, dtype = float)
    outputSerUnderSlopeMax = pandas.Series(outputListUnderSlopeMax, dtype = float)
    
    #outputSerOsafeerrorMean = pandas.Series(outputListOsafeerrorMean, dtype = float)
    outputSerOsafeerrorMax = pandas.Series(outputListOsafeerrorMax, dtype = float)
    #outputSerUsafeerrorMean = pandas.Series(outputListUsafeerrorMean, dtype = float)
    outputSerUsafeerrorMax = pandas.Series(outputListUsafeerrorMax, dtype = float)

    dfStatistics.insert(len(dfStatistics.columns), 'error (mean)', outputSerTotalErrorMean)
    dfStatistics.insert(len(dfStatistics.columns), 'errOdo', outputSereOdo)
    dfStatistics.insert(len(dfStatistics.columns), 'dlink (mean)', outputSerDlink)
    dfStatistics.insert(len(dfStatistics.columns), 'intercept (eOdo)', outputSereOdoIntercept)
    dfStatistics.insert(len(dfStatistics.columns), 'errDlink (mean)', outputSereDlinkMean)
    dfStatistics.insert(len(dfStatistics.columns), 'errDlink (std)', outputSereDlinkStd)
    dfStatistics.insert(len(dfStatistics.columns), 'errDlink (max)', outputSereDlinkMax)
    dfStatistics.insert(len(dfStatistics.columns), 'errDlink (min)', outputSereDlinkMin)
    # dfStatistics.insert(len(dfStatistics.columns), 'osafeerror (mean)', outputSerOsafeerrorMean)
    dfStatistics.insert(len(dfStatistics.columns), 'osafeerror (max)', outputSerOsafeerrorMax)
    # dfStatistics.insert(len(dfStatistics.columns), 'usafeerror (mean)', outputSerUsafeerrorMean)
    dfStatistics.insert(len(dfStatistics.columns), 'usafeerror (max)', outputSerUsafeerrorMax)
    dfStatistics.insert(len(dfStatistics.columns), 'overslope (mean)', outputSerOverSlopeMean)
    dfStatistics.insert(len(dfStatistics.columns), 'overslope (max)', outputSerOverSlopeMax)
    dfStatistics.insert(len(dfStatistics.columns), 'underslope (mean)', outputSerUnderSlopeMean)
    dfStatistics.insert(len(dfStatistics.columns), 'underslope (max)', outputSerUnderSlopeMax)
    dfStatistics.insert(len(dfStatistics.columns), 'overcorr (vel)', outputSerOverCorrVel)
    dfStatistics.insert(len(dfStatistics.columns), 'undercorr (vel)', outputSerUnderCorrVel)
    dfStatistics.insert(len(dfStatistics.columns), 'Distance', outputListDistance)
    dfStatistics.insert(len(dfStatistics.columns), 'NumBG', outputListNumBG)
    dfStatistics.insert(len(dfStatistics.columns), 'Odo Accuracy', outputListOdoAccuracy)
    
    #dfStatistics.insert(len(dfStatistics.columns), 'over/dist (mean)', outputSerOvermean)
    #dfStatistics.insert(len(dfStatistics.columns), 'under/dist (mean)', outputSerUndermean)
    
    # Found no added value in next correlation values
    
    # dfStatistics.insert(len(dfStatistics.columns), 'overcorr (accel)', outputSerOverCorrAccel)
    # dfStatistics.insert(len(dfStatistics.columns), 'undercorr (accel)', outputSerUnderCorrAccel)
    
    return dfStatistics


# JruC_SetRoutesFileFormat(wb)
# Set format in excel for routes file
def JruC_SetRoutesFileFormat(wb, dfSummary):

    ws = wb.get_worksheet_by_name('MISSIONS')

    f_float = wb.add_format({'num_format': '#,##0.00', 'bg_color': '#EBF1DE', 'text_wrap': True})
    f_percentage = wb.add_format({'num_format': '0.00%', 'bg_color': '#EBF1DE', 'text_wrap': True})
    

    firstcol = 1 + dfSummary.columns.get_loc('error (mean)')
    lastcol = 1 + dfSummary.columns.get_loc('usafeerror (max)')
    ws.set_column(firstcol, lastcol, None, f_float)

    firstcol = 1 + dfSummary.columns.get_loc('overslope (mean)')
    lastcol = 1 + dfSummary.columns.get_loc('undercorr (vel)')
    ws.set_column(firstcol, lastcol, None, f_percentage)

    return True

# JruC_SetDigitalMap(dfMission, dfRoutes)
# Compiled JRU File
# Set Digital Map info 
def JruC_SetDigitalMap(dfMission, dfRoutes, dfDigitalMap):

    # List of variables
    strPrevLRBG = ''
    strNextLrbg = ''

    if len(dfDigitalMap.axes[0]) == 0:
        print("SetDigitalMap-INFO-Digital Map not defined.")
        return dfMission

    #dfInputMission = pandas.DataFrame(dfMission) # Make a copy of Mission
    dfInputMission = dfMission.copy() # Make a copy of Mission

    for x in range (len(dfDigitalMap.axes[0])):

        refLrbg = dfDigitalMap['NID_BG'][x] # Reference BG
        elName = dfDigitalMap['EL_NAME'][x] # Element name
        elType = dfDigitalMap['EL_TYPE'][x] # Element type (signal, lx, ...)
        refQdirlrbg = dfDigitalMap['Q_DIRLRBG'][x] # Orientation of the train in relation to the direction of the LRBG
        refQdlrbg = dfDigitalMap['Q_DLRBG'][x] # Qualifier telling on which side of the LRBG the estimated front end is
        refDistance = dfDigitalMap['DIST'][x] # Distance from element to lrbg

        # Filter register with ref same lrbg, train orientation and train side
        dfTemp = dfInputMission[dfInputMission['NID_LRBG'].astype(int) == refLrbg]
        dfTemp = dfTemp[dfTemp['Q_DIRLRBG'].astype(int) == refQdirlrbg]
        dfTemp = dfTemp[dfTemp['Q_DLRBG'].astype(int) == refQdlrbg]
        dfTemp.index = numpy.arange(0, len(dfTemp.axes[0]))


        for z in range(len(dfTemp.axes[0])):

            distance = dfTemp['D_LRBG'][z]
            pos = dfTemp['INIT_INDEX'][z]
                
            if distance == refDistance:

                # Frontend just at element
                try: 
                    #dfRoutes.iloc[pos, dfRoutes.columns.get_loc(elType)] = 1
                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc(elType)] = '1'
                except:
                    if str(elType) == 'LABEL':
                        dfRoutes.iloc[pos, dfRoutes.columns.get_loc('NID_BG')] = str(elName) + '-0'   
                    else:    
                        print(color.Fore.YELLOW + 'Type: ' + elType + ' not found for element name: ' +  elName + color.Style.RESET_ALL)
                    
                break
                    
            elif distance > refDistance:

                if z == 0:
                    
                    if pos > 0:
                        strPrevLRBG = str(dfRoutes['NID_LRBG'][pos-1])
                        if strPrevLRBG != str(refLrbg):
                            # Frontend just passed the element
                            try: 
                                #dfRoutes.iloc[pos, dfRoutes.columns.get_loc(elType)] = 1
                                #dfRoutes.iloc[pos-1, dfRoutes.columns.get_loc(elType)] = 0
                                dfRoutes.iloc[pos, dfRoutes.columns.get_loc(elType)] = '1'
                                dfRoutes.iloc[pos-1, dfRoutes.columns.get_loc(elType)] = '0'
                            except:
                                if str(elType) == 'LABEL':
                                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc('NID_BG')] = str(elName) + '-0'    
                                else: 
                                    print(color.Fore.YELLOW + 'Type: ' + elType + ' not found for element name: ' +  elName + color.Style.RESET_ALL)
                                    break
                
                elif z > 0:

                    # Frontend just passed the element
                    try: 
                        #dfRoutes.iloc[pos, dfRoutes.columns.get_loc(elType)] = 1
                        dfRoutes.iloc[pos, dfRoutes.columns.get_loc(elType)] = '1'
                        if pos > 0:
                            #dfRoutes.iloc[pos-1, dfRoutes.columns.get_loc(elType)] = 0
                            dfRoutes.iloc[pos-1, dfRoutes.columns.get_loc(elType)] = '0'
                    except:

                        if str(elType) == 'LABEL':
                            dfRoutes.iloc[pos, dfRoutes.columns.get_loc('NID_BG')] = str(elName) + '-0'    
                        else: 
                            print(color.Fore.YELLOW + 'Type: ' + elType + ' not found for element name: ' +  elName + color.Style.RESET_ALL)
                            break

                    # Check previous register
                    prevPos = dfTemp['INIT_INDEX'][z-1]
                    if (pos - prevPos) == 1:
                        # Frontend was positioned before element in previous register
                        try: 
                            #dfRoutes.iloc[prevPos, dfRoutes.columns.get_loc(elType)] = 0
                            dfRoutes.iloc[prevPos, dfRoutes.columns.get_loc(elType)] = '0'
                        except:

                            if str(elType) == 'LABEL':
                                dfRoutes.iloc[pos, dfRoutes.columns.get_loc('NID_BG')] = str(elName) + '-0'    
                            else: 
                                print(color.Fore.YELLOW + 'Type: ' + elType + ' not found for element name: ' +  elName + color.Style.RESET_ALL)
                                break
                
                break

            elif distance < refDistance:

                # To be completed
                if z == (len(dfTemp.axes[0]) - 1):

                    if pos < (len(dfRoutes.axes[0]) -1):
                        strNextLrbg = str(dfRoutes['NID_LRBG'][pos+1])
                        if strNextLrbg != str(refLrbg):
                            # Frontend just passed the element
                            try: 
                                #dfRoutes.iloc[pos, dfRoutes.columns.get_loc(elType)] = 0
                                #dfRoutes.iloc[pos+1, dfRoutes.columns.get_loc(elType)] = 1
                                dfRoutes.iloc[pos, dfRoutes.columns.get_loc(elType)] = '0'
                                dfRoutes.iloc[pos+1, dfRoutes.columns.get_loc(elType)] = '1'
                            except:
                                if str(elType) == 'LABEL':
                                    dfRoutes.iloc[pos, dfRoutes.columns.get_loc('NID_BG')] = str(elName) + '-0'    
                                else: 
                                    print(color.Fore.YELLOW + 'Type: ' + elType + ' not found for element name: ' +  elName + color.Style.RESET_ALL)
                                    break   

    # From python 3.14 up we need to return dfMission as it ies treated as a copy and not updated outside the function 
    firstMissionIndex = dfMission['INIT_INDEX'][0]
    lastMissionIndex = dfMission['INIT_INDEX'][len(dfMission) - 1]
    dfReturnedMission = dfRoutes.iloc[firstMissionIndex:lastMissionIndex+1]
    dfReturnedMission.index = numpy.arange(0, len(dfReturnedMission["INIT_INDEX"]))
          

    return dfReturnedMission



# -----------------------------------  R_Functions    -------------------------------------------
# Functions whose input file is a normalized routes  file (R_file)
#
# Modules:
# COM 'common'
# naming: R_FunctionName
# R_ = Normalized Routes File
# R_SplitMission (dfMissions, dfMissionsSummary, nidEngine, path)
# Split single route and prepare charts
def R_SplitMission (dfMissions, dfMissionsSummary, nidEngine, outputPath, name):

    global glPh3Progress
    global glPh3NumberMissionsPerc
    global glPh3NumberRows

    global glOdoAccuracyDistance
    global glOdoAccuracyImpairmentTh
    global glOdoAccuracySafetyTh
    global glSlopeSS41Th

    dfMissionsSummary = pandas.DataFrame(dfMissionsSummary) # make a copy
    dfMissionsSummary.index = numpy.arange(0, len(dfMissionsSummary.axes[0]))
    dfSingleMission = pandas.DataFrame() # Initialization

    length = len(dfMissionsSummary.axes[0])
    n = int(0)
    glPh3NumberMissionsPerc = 0.0

    # Create output folder if not exists
    outputMissionsPath = pathlib.Path.joinpath(outputPath, 'missions')
    pathlib.Path(outputMissionsPath).mkdir(parents=True, exist_ok=True)
    
    while n < length:
        
        pathMissionFile = pathlib.Path.joinpath(outputMissionsPath, 'M_' + str(name) + '_' + nidEngine + '_' + str(n) + '.xlsx')
        outputMissionFile = str(pathMissionFile)
        isom = dfMissionsSummary['IxSoM'][n]
        range = dfMissionsSummary['IxEoM'][n] + 1
        glPh3NumberMissionsPerc = (n+1)/len(dfMissionsSummary.axes[0])
                
        df = dfMissions.iloc[isom:range]
        
        dfSingleMission = pandas.DataFrame(df) # make a copy
        dfSingleMission.index = numpy.arange(0, len(df.axes[0]))
        
        if len(dfSingleMission) > 0:

            if str(dfSingleMission['NID_ENGINE'][0]) != nidEngine:
                sys.stdout.write("\n")
                sys.stdout.flush()
                print(color.Fore.YELLOW + "Nid engine: " + nidEngine 
                + ' is not in mission' 
                + '-isom: ' + str(isom)
                + '-range: ' + str(range)
                + color.Style.RESET_ALL)
                
                return(pandas.DataFrame())
            somPacket = str(int(dfSingleMission['PACKET_RBC'][0]))
            # Take into account that we can consider SoM a 157 packet or any MA packet
            if (somPacket != '157') & (somPacket != '3') & (somPacket != '33'):
                sys.stdout.write("\n")
                sys.stdout.flush()
                print(color.Fore.YELLOW + "SoM packet not found for " + nidEngine 
                + ' .' 
                + ' - ixSoM: ' + str(isom)
                + ' - range: ' + str(range)
                + color.Style.RESET_ALL)
                
                return(pandas.DataFrame())
            
            # Drop unused columns
            dfSingleMission.drop(['NID_MESSAGE',
                                #'Record Id',
                                #'Date',
                                #'Time',
                                'Q_SCALE',
                                #'NID_LRBG',
                                'LRBG_Ext',
                                #'D_LRBG',
                                'Q_DIRLRBG',
                                'Q_DLRBG',
                                #'SSP',
                                #'BG_LINKS',
                                #'ERR_TYPE',
                                #'ERR_MESSAGE'
                                ], axis=1, inplace=True)
            
            listIth = []
            listSth = []
            listSlopeSS41Th = []
            listBGerrorDraw = []
            listMADraw = []
            listLevel1Draw = []
            listLevel2Draw = []
            listSBCDraw = []
            listEBCDraw = []
            listDmiAckDrawLS = []
            listDmiAckDrawSR = []
            listDmiAckDrawTR = []
            listDriverAckDraw = []
            listDriverStartDraw = []
            listDriverOverrideDraw = []
            listRadioUpDraw = []
            listRadioDownDraw = []
            listLxProtected = []
            listLxNotprotected = []
            listSignalPrev = []
            listSignalAfter = []
            listLxPrev = []
            listLxAfter = []
            listTcRadioHole = []
            fValueForDrawing1 = dfSingleMission['V_TRAIN'].max()*1.2
            fValueForDrawing1 +=5
            fValueForDrawing2 = fValueForDrawing1 + 10
            fValueForDrawing3 = fValueForDrawing2 + 10
            fValueForDrawing4 = fValueForDrawing1 - 10
            fValueForDrawing5 = fValueForDrawing4 - 10
            i = 0
            while i < len(dfSingleMission.axes[0]):
                listIth.insert(i, glOdoAccuracyImpairmentTh)
                listSth.insert(i, glOdoAccuracySafetyTh)
                listSlopeSS41Th.insert(i, glSlopeSS41Th)
                trainSpeed = dfSingleMission['V_TRAIN'][i]
                if float(trainSpeed) <= 1.0: trainSpeed = 5 # !important to visually differentiate from EB liberation represented at y-value of 0
                
                if not numpy.isnan(dfSingleMission['NID_ERRORBG'][i]):
                    listBGerrorDraw.insert(i, fValueForDrawing1)
                else:
                    listBGerrorDraw.insert(i, numpy.nan)

                if dfSingleMission['PACKET_RBC'][i] == 3 or dfSingleMission['PACKET_RBC'][i] == 33:
                    listMADraw.insert(i, fValueForDrawing1)
                else:
                    listMADraw.insert(i, numpy.nan)

                if dfSingleMission['M_LEVEL'][i] == 2:
                    listLevel1Draw.insert(i, fValueForDrawing1)
                else:
                    listLevel1Draw.insert(i, numpy.nan)

                if dfSingleMission['M_LEVEL'][i] == 3:
                    listLevel2Draw.insert(i, fValueForDrawing1)
                else:
                    listLevel2Draw.insert(i, numpy.nan)
                
                if dfSingleMission['SBC_STATE'][i] == 1:
                    listSBCDraw.insert(i, trainSpeed)
                elif dfSingleMission['SBC_STATE'][i] == 0:
                    listSBCDraw.insert(i, -5)
                else:
                    listSBCDraw.insert(i, numpy.nan)

                try: # As far as Hasler compilation doesn´t include nexts
                    if dfSingleMission['EBC_STATE'][i] == 1:
                        listEBCDraw.insert(i, trainSpeed)
                    elif dfSingleMission['EBC_STATE'][i] == 0:
                        listEBCDraw.insert(i, -5)
                    else:
                        listEBCDraw.insert(i, numpy.nan)
                    
                    # DMI LS ACK
                    if 'MO22' in str(dfSingleMission['MOSYMB'][i]) : 
                        listDmiAckDrawLS.insert(i, fValueForDrawing2)
                        listDmiAckDrawTR.insert(i, numpy.nan)
                        listDmiAckDrawSR.insert(i, numpy.nan)
                    # DMI TRIP ACK
                    elif 'MO05' in str(dfSingleMission['MOSYMB'][i]) : 
                        listDmiAckDrawTR.insert(i, fValueForDrawing2)
                        listDmiAckDrawLS.insert(i, numpy.nan)
                        listDmiAckDrawSR.insert(i, numpy.nan)
                    # DMI SR ACK
                    elif 'MO10' in str(dfSingleMission['MOSYMB'][i]) : 
                        listDmiAckDrawSR.insert(i, fValueForDrawing2)
                        listDmiAckDrawLS.insert(i, numpy.nan)
                        listDmiAckDrawTR.insert(i, numpy.nan)
                    else:
                        listDmiAckDrawLS.insert(i, numpy.nan)
                        listDmiAckDrawSR.insert(i, numpy.nan)
                        listDmiAckDrawTR.insert(i, numpy.nan)
                    
                    # Driver Action: LS Acknowledgment
                    if str(dfSingleMission['DRACT'][i]) == '13.0': 
                        listDriverAckDraw.insert(i, fValueForDrawing2)
                    # Driver Action: SR Acknowledgment
                    elif str(dfSingleMission['DRACT'][i]) == '3.0': 
                        listDriverAckDraw.insert(i, fValueForDrawing2)
                    # Driver Action: TRIP Acknowledgment
                    elif str(dfSingleMission['DRACT'][i]) == '2.0': 
                        listDriverAckDraw.insert(i, fValueForDrawing2)
                    else:
                        listDriverAckDraw.insert(i, numpy.nan)
                    if '19' in str(dfSingleMission['DRACT'][i]) : # Start driver ack
                        listDriverStartDraw.insert(i, fValueForDrawing2)
                    else:
                        listDriverStartDraw.insert(i, numpy.nan)
                    if '14' in str(dfSingleMission['DRACT'][i]) : # Override
                        listDriverOverrideDraw.insert(i, fValueForDrawing2)
                    else:
                        listDriverOverrideDraw.insert(i, numpy.nan)
                    
                    # RADIO UP
                    if dfSingleMission['ST03'][i] == 'ST03':
                        listRadioUpDraw.insert(i, fValueForDrawing3)
                    else:
                        listRadioUpDraw.insert(i, numpy.nan)
                    # RADIO DOWN
                    if dfSingleMission['ST04'][i] == 'ST04':
                        listRadioDownDraw.insert(i, fValueForDrawing3)
                    else:
                        listRadioDownDraw.insert(i, numpy.nan)

                    # RADIONHOLE
                    if dfSingleMission['TC12'][i] == 'TC12':
                        listTcRadioHole.insert(i, fValueForDrawing3)
                    else:
                        listTcRadioHole.insert(i, numpy.nan)

                    # LX STATUS PROTECTED
                    if dfSingleMission['LX'][i] == 0:
                        listLxProtected.insert(i, fValueForDrawing4)
                    else:
                        listLxProtected.insert(i, numpy.nan)

                    # LX STATUS NOT PROTECTED
                    if dfSingleMission['LX'][i] == 1:
                        listLxNotprotected.insert(i, fValueForDrawing4)
                    else:
                        listLxNotprotected.insert(i, numpy.nan)

                    # Digital MAP SIGNAL (prev)
                    if dfSingleMission['SIGNAL'][i] == 0:
                        listSignalPrev.insert(i, fValueForDrawing5)
                    else:
                        listSignalPrev.insert(i, numpy.nan)
                    
                    # Digital MAP SIGNAL (after)
                    if dfSingleMission['SIGNAL'][i] == 1:
                        listSignalAfter.insert(i, fValueForDrawing5)
                    else:
                        listSignalAfter.insert(i, numpy.nan)

                    # Digital MAP PAN (prev)
                    if dfSingleMission['PAN'][i] == 0:
                        listLxPrev.insert(i, fValueForDrawing5)
                    else:
                        listLxPrev.insert(i, numpy.nan)
                    
                    # Digital MAP PAN (after)
                    if dfSingleMission['PAN'][i] == 1:
                        listLxAfter.insert(i, fValueForDrawing5)
                    else:
                        listLxAfter.insert(i, numpy.nan)

                    


                except:
                    pepe = 1


                i += 1
            
            # Insert new columns for accuracy thresholds
            serIth = pandas.Series(listIth)
            serSth = pandas.Series(listSth)
            serSlopeth = pandas.Series(listSlopeSS41Th)
            
            dfSingleMission.insert(len(dfSingleMission.columns), 'IMP_TH', serIth)
            dfSingleMission.insert(len(dfSingleMission.columns), 'SAFETY_TH', serSth)
            dfSingleMission.insert(len(dfSingleMission.columns), 'SS41_UNDERSLOPE', serSlopeth)
            dfSingleMission.insert(len(dfSingleMission.columns), 'SS41_OVERSLOPE', serSlopeth)
            
            # insert new column for drawing BGs in error
            serBGerrorDraw = pandas.Series(listBGerrorDraw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'BGERROR', serBGerrorDraw)

            # insert new column for drawing onboard received MAs 
            serMADraw = pandas.Series(listMADraw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'MA', serMADraw)

            # insert new column for drawing Level 1 
            serLevel1Draw = pandas.Series(listLevel1Draw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'L1', serLevel1Draw)

            # insert new column for drawing Level 3 
            serLevel2Draw = pandas.Series(listLevel2Draw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'L2', serLevel2Draw)

            # insert new column for drawing Level 3service brake commanded 
            serSBCDraw = pandas.Series(listSBCDraw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'SBC', serSBCDraw)

            # insert new column for drawing Level emergency brake commanded 
            serEBCDraw = pandas.Series(listEBCDraw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'EBC', serEBCDraw)

            # insert new column for drawing Acknowledgment LS symbol in DMI
            serSymAckLsDraw = pandas.Series(listDmiAckDrawLS)
            dfSingleMission.insert(len(dfSingleMission.columns), 'ACKLS', serSymAckLsDraw)

            # insert new column for drawing Acknowledgment SR symbol in DMI
            serSymAckSrDraw = pandas.Series(listDmiAckDrawSR)
            dfSingleMission.insert(len(dfSingleMission.columns), 'ACKSR', serSymAckSrDraw)

            # insert new column for drawing Acknowledgment TR symbol in DMI
            serSymAckTrDraw = pandas.Series(listDmiAckDrawTR)
            dfSingleMission.insert(len(dfSingleMission.columns), 'ACKTRIP', serSymAckTrDraw)

            # insert new column for driver´s acknowledgments
            serDriverAckDraw = pandas.Series(listDriverAckDraw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'DRACK', serDriverAckDraw)

            # insert new column for driver´s start
            serDriverStart = pandas.Series(listDriverStartDraw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'DRSTART', serDriverStart)

            # insert new column for driver´s override
            serDriverOverride = pandas.Series(listDriverOverrideDraw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'DROVERRIDE', serDriverOverride)


            # insert new column for radio up
            serRadioUpDraw = pandas.Series(listRadioUpDraw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'RADIOUP', serRadioUpDraw)

            # insert new column for radio down
            serRadioDownDraw = pandas.Series(listRadioDownDraw)
            dfSingleMission.insert(len(dfSingleMission.columns), 'RADIODOWN', serRadioDownDraw)

             # insert new column for radio holeTC
            serTCRadioHoleDraw = pandas.Series(listTcRadioHole)
            dfSingleMission.insert(len(dfSingleMission.columns), 'RHTC', serTCRadioHoleDraw)

            # insert new column for lx protected
            serLxProtected = pandas.Series(listLxProtected)
            dfSingleMission.insert(len(dfSingleMission.columns), 'LXPROT', serLxProtected)

            # insert new column for lx not protected
            serLxNotprotected = pandas.Series(listLxNotprotected)
            dfSingleMission.insert(len(dfSingleMission.columns), 'LXNOTPROT', serLxNotprotected)

            # insert new column for digital map SIGNAL (prev)
            serSignalPrev = pandas.Series(listSignalPrev)
            dfSingleMission.insert(len(dfSingleMission.columns), 'SIGPREV', serSignalPrev)

            # insert new column for digital map SIGNAL (after)
            serSignalAfter = pandas.Series(listSignalAfter)
            dfSingleMission.insert(len(dfSingleMission.columns), 'SIGAFTER', serSignalAfter)

            # insert new column for digital map PAN (prev)
            serLxPrev = pandas.Series(listLxPrev)
            dfSingleMission.insert(len(dfSingleMission.columns), 'LXPREV', serLxPrev)

            # insert new column for digital map PAN (after)
            serLxAfter = pandas.Series(listLxAfter)
            dfSingleMission.insert(len(dfSingleMission.columns), 'LXAFTER', serLxAfter)

            


            # Drop empty vel columns
            listVel = ['V_FS', 'V_OS', 'V_SR', 'V_SH', 'V_SB', 'V_LS', 'V_NS', 'V_OTHER']
            for vel in listVel:
                try:
                    dfVel = dfSingleMission[dfSingleMission[vel].isnull()]
                    
                    if len(dfVel.axes[0]) == len(dfSingleMission.axes[0]): 
                        dfSingleMission.drop([vel], axis=1, inplace=True)
                except:
                    continue
            
            # Remove not-0 balises
            dfSingleMission['NID_BG'] = dfSingleMission[dfSingleMission['NID_BG'].str.contains('-0', na = False)]['NID_BG']
            # Set last value in NID_BG for charting
            if str(dfSingleMission['NID_BG'][len(dfSingleMission.axes[0])-1]) == 'nan':

                dfSingleMission.iloc[len(dfSingleMission.axes[0])-1, dfSingleMission.columns.get_loc('NID_BG')] = '-' # For charting



            glPh3Progress = glPh3NumberMissionsPerc*100
            sys.stdout.write("Progress: %d%%   \r" % (glPh3Progress) )
            sys.stdout.flush()
            
            
        else:
            sys.stdout.write("\n")
            sys.stdout.flush()
            print(color.Fore.YELLOW + "Empty mission for " + nidEngine 
                + '-isom: ' + str(isom)
                + '-range: ' + str(range)
                + color.Style.RESET_ALL)
            dfSingleMission = pandas.DataFrame() 
        
        # Set all Over estimatoin data as negative values
        #dfSingleMission['L_DOUBTOVER'] = dfSingleMission['L_DOUBTOVER'].apply(utl_changeSign)
        #dfSingleMission['OVEREST_ERROR'] = dfSingleMission['OVEREST_ERROR'].apply(utl_changeSign)
        #dfSingleMission['OVERSLOPE'] = dfSingleMission['OVERSLOPE'].apply(utl_changeSign)
        #dfSingleMission['SS41_OVERSLOPE'] = dfSingleMission['SS41_OVERSLOPE'].apply(utl_changeSign)

        # Delete L_DOUBTUNDER / L_DOUBTOVER values of 32767 (Unknown)
        dfSingleMission['L_DOUBTUNDER'] = dfSingleMission['L_DOUBTUNDER'].apply(utl_delUnknownValue)
        dfSingleMission['L_DOUBTOVER'] = dfSingleMission['L_DOUBTOVER'].apply(utl_delUnknownValue)

        # Set all Under estimatoin data as negative values
        dfSingleMission['L_DOUBTUNDER'] = dfSingleMission['L_DOUBTUNDER'].apply(utl_changeSign)
        dfSingleMission['UNDEREST_ERROR'] = dfSingleMission['UNDEREST_ERROR'].apply(utl_changeSign)
        dfSingleMission['UNDERSLOPE'] = dfSingleMission['UNDERSLOPE'].apply(utl_changeSign)
        dfSingleMission['SS41_UNDERSLOPE'] = dfSingleMission['SS41_UNDERSLOPE'].apply(utl_changeSign)
        # Delete L_DOUBTUNDER / L_DOUBTOVER values of 32767 (Unknown)
        
            
                        
        # Excel write (end of Phase 2)
        dfSingleMissionSummary = pandas.DataFrame(dfMissionsSummary[n:n+1])
        dfSingleMissionSummary.drop(['IxSoM', 'IxEoM'], axis=1, inplace=True)
                                
        try:
            
            writer = pandas.ExcelWriter(outputMissionFile, engine='xlsxwriter')
            dfSingleMissionSummary.to_excel(writer, sheet_name='MISSION')
            dfSingleMission.to_excel(writer, sheet_name='ROUTE')

        except:
            print(color.Fore.YELLOW + "Can´t save route file: " + outputMissionFile + color.Style.RESET_ALL)

        # Create worksheet with charts
        try:
            wb = writer.book
            ws = wb.add_worksheet("Charts")
            
            # Speed chart
            speedChart = R_DrawSpeedChart(wb, dfSingleMission)
            ws.insert_chart('A1', speedChart)

            # odo_err chart
            odoerrorChart = R_DrawOdoerrorChart(wb, dfSingleMission)
            ws.insert_chart('A25', odoerrorChart)

            # Create speed & acceleration chart
            accelChart = R_DrawAccelChart(wb, dfSingleMission)
            ws.insert_chart('A54', accelChart)

            # Create odometry accuracy chart
            overUnderAccuracyChart = R_DrawOverUnderAccuracyChart(wb, dfSingleMission)
            ws.insert_chart('A74', overUnderAccuracyChart)

            # Create overestimation error/underestimation error chart
            #overUnderestLineChart = R_DrawOverUnderErrorChart(wb, dfSingleMission)
            #ws.insert_chart('A85', overUnderestLineChart)

            # Create overestimation /underestimation slope chart
            #slopeLineChart = R_DrawOverUnderSlopeChart(wb, dfSingleMission)
            #ws.insert_chart('A110', slopeLineChart)

            wb.close()

        except Exception as ex:    
            print(color.Fore.YELLOW + "Can´t add charts to " + outputMissionFile + ' - Exception: ' + str(ex) + color.Style.RESET_ALL)

        
        n += 1

        
    sys.stdout.write("\n")
    sys.stdout.flush()      
    return dfSingleMission

# R_DrawSpeedChart(wb, dfSingleMission)
# Draw speed chart
def R_DrawSpeedChart(wb, dfSingleMission):

    dfInput = pandas.DataFrame(dfSingleMission)
    
            
    # Create speed chart
    speedChart = wb.add_chart({'type': 'line'})
    eventsScatterChart = wb.add_chart({'type': 'scatter'})
            
            
    # Format chart (# 1 cm= 37,79527559055 px)
    # default with 480 px (12,7 cm)
    lastrow = len(dfInput.axes[0]) # Last excel row for chart
    speedChartWidth = float(lastrow/10*37.79)
    if speedChartWidth < 480: speedChartWidth = 480
    speedChart.set_title({'name': 'V_TRAIN',
                                  'name_font': {'size': 14, 'bold': False}})
    speedChart.set_size({'width': speedChartWidth,
                         'y_scale': 1.5})
    speedChart.set_legend({'position': 'bottom'})
    speedChart.set_y_axis({
                        'major_gridlines': {
                        'visible': True,
                        'line': {'width': 0.75, 'color': '#D9D9D9'}
                        },
                        'line': {'none': True},
                        'num_font': {'size': 10},
                        'name': 'Vel (km/h)',
                        'name_font': {'size': 10, 'bold': False}
                        })
    speedChart.set_x_axis({'major_tick_mark': 'none',
                            'line': {'none': True},
                            'num_format': '#,##0.00,',
                            'num_font': {'size': 10, 'rotation': -90},
                            'name': 'Dist (km)',
                            'name_font': {'size': 10, 'bold': False}
                            })
    speedChart.set_y2_axis({'visible': False,})
    speedChart.set_x2_axis({ 'label_position': 'high',
                            'visible': True,
                            'line': {'none': True},
                            'num_font':  {'rotation': -70, 'size': 8},
                            })
            

                                        
    # Add speed values to the chart.
    listVel = ['V_FS', 'V_OS', 'V_SR', 'V_SH', 'V_SB', 'V_LS', 'V_NS', 'V_OTHER']
    lineWidth = 2.0
    for vel in listVel:
        try:
            valuecolnum = 1+ dfInput.columns.get_loc(vel)
            catcolnum = 1 + dfInput.columns.get_loc('D_ONB_ACCUM') 
            #     [sheetname, first_row, first_col, last_row, last_col]
            if vel == 'V_FS': lineWidth = 1.0
            else: lineWidth = 2.0
            speedChart.add_series({
                                'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                                'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                                'line':       {'width': lineWidth},
                                'name':       ['ROUTE', 0, valuecolnum]
                                })
        except:
            continue
            
    # Add static profile serie
    lineWidth = 1.5
    valuecolnum = 1+ dfInput.columns.get_loc('V_STATIC')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    speedChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': '#E46C0A',
                                            'dash_type': 'round_dot'},
                            'name': ['ROUTE', 0, valuecolnum],
                            'y2_axis': True,
                            'x2_axis': True,
                            })
    # Add mrsp serie
    lineWidth = 1.0
    valuecolnum = 1+ dfInput.columns.get_loc('V_PERM')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    speedChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': '#FF0000'},
                            'name':       ['ROUTE', 0, valuecolnum],
                            'y2_axis': True,
                            'x2_axis': True,
                            })
    

    # Add indication for BGs in error
    valuecolnum = 1+ dfInput.columns.get_loc('BGERROR')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'triangle', 
                                       'size': 5,
                                       'border': {'color': '#FF0000'},
                                       'fill':   {'color': '#FF0000'},}
                            })
    

    # Add indication for onboard MA received
    valuecolnum = 1+ dfInput.columns.get_loc('MA')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'diamond', 
                                       'size': 5,
                                       'border': {'color': '#00B050'},
                                       'fill':   {'color': '#00B050'},}
                            })
    
    # Add indication for Level 1
    valuecolnum = 1+ dfInput.columns.get_loc('L1')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'square', 
                                       'size': 10,                                                                             
                                       'border': {'none': True},
                                       'fill':   {'color': '#B7DEE8', 'transparency': 80},
                                       }
                            })
    
    # Add indication for Level 2
    valuecolnum = 1+ dfInput.columns.get_loc('L2')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'square', 
                                       'size': 10,
                                       'border': {'none': True},
                                       'fill':   {'color': '#E6B9B8', 'transparency': 80},
                                       }
                            })
    
    # Add indication for SBC
    valuecolnum = 1+ dfInput.columns.get_loc('SBC')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'diamond', 
                                       'size': 7,
                                       'border': {'color': '#FF0000'},
                                       'fill':   {'color': '#FFFF00'},}
                            })
    
    # Add indication for EBC
    valuecolnum = 1+ dfInput.columns.get_loc('EBC')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'diamond', 
                                    'size': 7,
                                    'border': {'color': '#FF0000'},
                                    'fill':   {'color': '#FF0000'},}
                                })
        
    # Add indication DMI LS acknowledgment
    valuecolnum = 1+ dfInput.columns.get_loc('ACKLS')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'square', 
                                    'size': 5,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#00B050'},}
                            })
    
    # Add indication DMI LS acknowledgment
    valuecolnum = 1+ dfInput.columns.get_loc('ACKSR')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'square', 
                                    'size': 5,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#FFC000'},}
                            })
    
    # Add indication DMI TRIP acknowledgment
    valuecolnum = 1+ dfInput.columns.get_loc('ACKTRIP')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'square', 
                                    'size': 5,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#FF0000'},}
                            })
        
    # Add drivers acknowledgments
    valuecolnum = 1+ dfInput.columns.get_loc('DRACK')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'square', 
                                    'size': 5,
                                    'border': {'color': '#FFC000'},
                                    'fill':   {'color': '#1F497D'},}
                            })
    
    # Add drivers start
    valuecolnum = 1+ dfInput.columns.get_loc('DRSTART')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'circle', 
                                    'size': 7,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#00B050'},}
                            })
    
    # Add drivers override
    valuecolnum = 1+ dfInput.columns.get_loc('DROVERRIDE')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'circle', 
                                    'size': 7,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#FF0000'},}
                            })
    
    # Add radio up
    valuecolnum = 1+ dfInput.columns.get_loc('RADIOUP')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'circle', 
                                    'size': 5,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#92D050'},}
                            })
    
    
    # Add radio down
    valuecolnum = 1 + dfInput.columns.get_loc('RADIODOWN')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'circle', 
                                    'size': 5,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#FFC000'},}
                            })
    
    # Add radio hole track condition
    valuecolnum = 1+ dfInput.columns.get_loc('RHTC')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'star', 
                                    'size': 5,
                                    'border': {'color': '#FF0000'},
                                    'fill':   {'color': '#FFFFFF'},}
                            })

    # Add lx protected
    valuecolnum = 1 + dfInput.columns.get_loc('LXPROT')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'star', 
                                    'size': 5,
                                    'border': {'color': '#FFFFFF'},
                                    'fill':   {'color': '#00B050'},}
                            })
    
    # Add lx not protected
    valuecolnum = 1 + dfInput.columns.get_loc('LXNOTPROT')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'star', 
                                    'size': 5,
                                    'border': {'color': '#FFFF00'},
                                    'fill':   {'color': '#FF0000'},}
                            })
    
    # Add digital map signal (prev)
    valuecolnum = 1 + dfInput.columns.get_loc('SIGPREV')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'circle', 
                                    'size': 5,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#FFFFFF'},}
                            })
    
    # Add digital map signal (after)
    valuecolnum = 1 + dfInput.columns.get_loc('SIGAFTER')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'circle', 
                                    'size': 5,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#1F497D'},}
                            })
    
    # Add digital map Lx (prev)
    valuecolnum = 1 + dfInput.columns.get_loc('LXPREV')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'diamond', 
                                    'size': 5,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#FFFFFF'},}
                            })
    
    # Add digital map Lx (after)
    valuecolnum = 1 + dfInput.columns.get_loc('LXAFTER')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    eventsScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'diamond', 
                                    'size': 5,
                                    'border': {'color': '#1F497D'},
                                    'fill':   {'color': '#00B050'},}
                            })
        
    speedChart.combine(eventsScatterChart)
    speedChart.show_blanks_as('gap')
    

    return speedChart

# R_DrawOdoerrorChart(wb, dfSingleMission)
# Draw odo_err chart
def R_DrawOdoerrorChart(wb, dfSingleMission):

    dfInput = pandas.DataFrame(dfSingleMission) 

    # check if odoerror data
    #dfTemp = dfInput[~dfInput['LOC_ERROR'].isna()]
    #if dfTemp.empty: 
    #    return False     
           
    # Create speed chart
    odoerrorLineChart = wb.add_chart({'type': 'line'})
    odoerrorScatterChart = wb.add_chart({'type': 'scatter'})
    
        
    # Format chart (# 1 cm= 37,79527559055 px)
    # default with 480 px (12,7 cm)
    lastrow = len(dfInput.axes[0]) # Last excel row for chart
    chartWidth = float(lastrow/10*37.79)
    if chartWidth < 480: chartWidth = 480
    odoerrorLineChart.set_title({'name': 'ODO_ERR',
                                  'name_font': {'size': 14, 'bold': False}})
    odoerrorLineChart.set_size({'width': chartWidth,
                         'y_scale': 1.8})
    odoerrorLineChart.set_legend({'position': 'bottom'})
    odoerrorLineChart.set_y_axis({
                        'major_gridlines': {
                        'visible': True,
                        'line': {'width': 0.75, 'color': '#D9D9D9'}
                        },
                        'line': {'none': True},
                        'num_font': {'size': 10},
                        'name': 'CI and odo error (m)',
                        'name_font': {'size': 10, 'bold': False}
                        })
    odoerrorLineChart.set_x_axis({'major_tick_mark': 'none',
                            'line': {'none': True},
                            'num_font': {'size': 8, 'rotation': -70},
                            'name': 'Balises',
                            'name_font': {'size': 10, 'bold': False},
                            'label_position': 'low'
                            })
    
           
    # Add L_DOUBTOVER values serie
    lineWidth = 1.0
    valuecolnum = 1+ dfInput.columns.get_loc('L_DOUBTOVER')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoerrorLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth},
                            'name': ['ROUTE', 0, valuecolnum]
                            })
    # Add L_DOUBTUNDER values serie
    lineWidth = 1.0
    valuecolnum = 1+ dfInput.columns.get_loc('L_DOUBTUNDER')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoerrorLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth},
                            'name':       ['ROUTE', 0, valuecolnum]
                            })
    
    # Add SS41_MIN values serie
    lineWidth = 1.0
    valuecolnum = 1+ dfInput.columns.get_loc('SS41_MIN')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoerrorLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth,
                                           'color': '#4A7EBB',
                                           'dash_type': 'square_dot'},
                            'name': ['ROUTE', 0, valuecolnum]
                            })
    
    
    # Add SS41_MAX values serie
    lineWidth = 1.0
    valuecolnum = 1+ dfInput.columns.get_loc('SS41_MAX')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoerrorLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth,
                                           'color': '#BE4B48',
                                           'dash_type': 'square_dot'},
                            'name': ['ROUTE', 0, valuecolnum]
                            })

    # Add TOTAL LOC_ERROR values serie
    valuecolnum = 1+ dfInput.columns.get_loc('LOC_ERROR')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoerrorScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'circle', 
                                       'size': 4,
                                       'border': {'none': True},
                                       'fill':   {'color': '#FFC000'},}
                            })
    
    odoerrorLineChart.combine(odoerrorScatterChart)
    odoerrorLineChart.show_blanks_as('gap')

    # Add dlink error values serie
    valuecolnum = 1+ dfInput.columns.get_loc('DLINKERR')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoerrorScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'triangle', 
                                       'size': 3,
                                       'border': {'none': True},
                                       'fill':   {'color': '#FF0000'},}
                            })
    
    odoerrorLineChart.combine(odoerrorScatterChart)
    odoerrorLineChart.show_blanks_as('gap')

    # Add odo(onboard) error values serie
    valuecolnum = 1+ dfInput.columns.get_loc('ODOERR')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoerrorScatterChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'name':       ['ROUTE', 0, valuecolnum],
                            'marker': {'type': 'diamond', 
                                       'size': 3,
                                       'border': {'none': True},
                                       'fill':   {'color': '#1F497D'},}
                            })
    
    odoerrorLineChart.combine(odoerrorScatterChart)
    odoerrorLineChart.show_blanks_as('gap')

    return odoerrorLineChart

# R_DrawAccelChart(wb, dfSingleMission)
# Draw acceleration chart
def R_DrawAccelChart(wb, dfSingleMission):

    dfInput = pandas.DataFrame(dfSingleMission)
    
           
    # Create speed chart
    accelChart = wb.add_chart({'type': 'line'})
           
            
    # Format chart (# 1 cm= 37,79527559055 px)
    # default with 480 px (12,7 cm)
    lastrow = len(dfInput.axes[0]) # Last excel row for chart
    chartWidth = float(lastrow/10*37.79)
    if chartWidth < 480: chartWidth = 480
    accelChart.set_title({'name': 'SPEED & ACCEL',
                                  'name_font': {'size': 14, 'bold': False}})
    accelChart.set_size({'width': chartWidth,
                         'y_scale': 1.2})
    accelChart.set_legend({'position': 'bottom'})
    accelChart.set_y_axis({
                        'major_gridlines': {
                        'visible': True,
                        'line': {'width': 0.75, 'color': '#D9D9D9'}
                        },
                        'line': {'none': True},
                        'num_font': {'size': 10},
                        'name': 'Speed (km/h)',
                        'name_font': {'size': 10, 'bold': False}
                        })
    accelChart.set_x_axis({'major_tick_mark': 'none',
                            'line': {'none': True},
                            'num_format': '#,##0.00,',
                            'num_font': {'size': 10, 'rotation': -90},
                            'name': 'Dist (km)',
                            'name_font': {'size': 10, 'bold': False}
                        })
    
    accelChart.set_y2_axis({'visible': True,
                            'name': 'Accel (m/s2)',
                            'name_font': {'size': 10, 'bold': False}})
    
           
    # Add L_DOUBTOVEV_TRAINR values serie
    lineWidth = 1.0
    valuecolnum = 1+ dfInput.columns.get_loc('V_TRAIN')
    catcolnum = 1 + dfInput.columns.get_loc('D_ONB_ACCUM') 
    accelChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth},
                            'name': ['ROUTE', 0, valuecolnum]
                            })
    # Add Accel values serie
    lineWidth = 1.5
    valuecolnum = 1+ dfInput.columns.get_loc('ACCEL')
    catcolnum = 1 + dfInput.columns.get_loc('D_ONB_ACCUM') 
    accelChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': '#C0504D',
                                            'dash_type': 'round_dot'},
                            'name':       ['ROUTE', 0, valuecolnum],
                            'y2_axis': True,
                            })
    
    
    return accelChart

# R_DrawOverUnderAccuracyChart(wb, dfSingleMission)
# Draw Over/Under estimation accuracy chart (accumulated in 5.000 mts)
def R_DrawOverUnderAccuracyChart(wb, dfSingleMission):

    dfInput = pandas.DataFrame(dfSingleMission)
            
    # Create speed chart
    odoAccChart = wb.add_chart({'type': 'line'})
            
    # Format chart (# 1 cm= 37,79527559055 px)
    # default with 480 px (12,7 cm)
    lastrow = len(dfInput.axes[0]) # Last excel row for chart        
    chartWidth = float(lastrow/10*37.79)
    if chartWidth < 480: chartWidth = 480
    odoAccChart.set_title({'name': 'OVER UNDER ESTIMATION ACCURACY (5.000 m)',
                                  'name_font': {'size': 14, 'bold': False}})
    odoAccChart.set_size({'width': chartWidth,
                         'y_scale': 1.2})
    odoAccChart.set_legend({'position': 'bottom'})
    odoAccChart.set_y_axis({
                        'major_gridlines': {
                        'visible': True,
                        'line': {'width': 0.75, 'color': '#D9D9D9'}
                        },
                        'line': {'none': True},
                        'num_font': {'size': 10},
                        'name': 'Odo accuracy (m)',
                        'name_font': {'size': 10, 'bold': False}
                        })
    odoAccChart.set_x_axis({'major_tick_mark': 'none',
                            'line': {'none': True},
                            'num_font': {'size': 8, 'rotation': -70},
                            'name': 'Balises',
                            'name_font': {'size': 10, 'bold': False}
                            })

            
    # Add safety treshold serie
    lineWidth = 1.25
    valuecolnum = 1+ dfInput.columns.get_loc('SAFETY_TH')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoAccChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': '#FF0000',
                                            'dash_type': 'round_dot'},
                            'name': ['ROUTE', 0, valuecolnum],
                            })
    
    # Add impairment treshold serie
    lineWidth = 1.25
    valuecolnum = 1+ dfInput.columns.get_loc('IMP_TH')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoAccChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': '#F79646',
                                            'dash_type': 'round_dot'},
                            'name': ['ROUTE', 0, valuecolnum],
                            })
    
    # Add odometry accuracy serie (underestimation) UNDER_ACCURACY
    lineWidth = 1.0
    valuecolnum = 1+ dfInput.columns.get_loc('UNDER_ACCURACY')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoAccChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': '#416FA6',},
                            'name': ['ROUTE', 0, valuecolnum],
                            })
    
    # Add odometry accuracy serie (underestimation) OVER_ACCURACY
    lineWidth = 1.0
    valuecolnum = 1+ dfInput.columns.get_loc('OVER_ACCURACY')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoAccChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            },
                            'name': ['ROUTE', 0, valuecolnum],
                            })
    

    return odoAccChart

# R_DrawOverUnderErrorChart(wb, dfSingleMission)
# Draw estimation error chart
def R_DrawOverUnderErrorChart(wb, dfSingleMission):

    dfInput = pandas.DataFrame(dfSingleMission)

    # check if OVEREST_ERROR data
    #dfTempOver = dfInput[~dfInput['OVEREST_ERROR'].isna()]
    #dfTempUnder = dfInput[~dfInput['UNDEREST_ERROR'].isna()]
    #if dfTempOver.empty & dfTempUnder.empty: 
    #    return False
    
    lastrow = len(dfInput.axes[0]) # Last excel row for chart
            
    # Create speed chart
    odoestLineChart = wb.add_chart({'type': 'line'})
    
            
            
    # Format chart (# 1 cm= 37,79527559055 px)
    # default with 480 px (12,7 cm)
    chartWidth = float(lastrow/10*37.79)
    if chartWidth < 480: chartWidth = 480
    odoestLineChart.set_title({'name': 'OVER UNDER ESTIMATION ERROR (est/dist %)',
                                  'name_font': {'size': 14, 'bold': False}})
    odoestLineChart.set_size({'width': chartWidth,
                         'y_scale': 1.5})
    odoestLineChart.set_legend({'position': 'bottom'})

    df = dfInput[~dfInput['L_DOUBTOVER'].isna()]
    maxover = abs(df['L_DOUBTOVER'].min())
    df = dfInput[~dfInput['L_DOUBTUNDER'].isna()]
    maxunder = df['L_DOUBTUNDER'].max()

    maxaxis = max([maxover, maxunder])
    strAux = str(maxaxis)
    if strAux == 'nan': 
        maxaxis = 0.0
    maxaxis = 10 + round(maxaxis)

    odoestLineChart.set_y_axis({
                        'major_gridlines': {
                        'visible': True,
                        'line': {'width': 0.75, 'color': '#D9D9D9'}
                        },
                        'line': {'none': True},
                        'num_font': {'size': 10},
                        'name': 'CI Confident interval (m)',
                        'name_font': {'size': 10, 'bold': False},
                        'max': maxaxis,
                        'min': -maxaxis
                        })
    odoestLineChart.set_x_axis({'major_tick_mark': 'none',
                            'line': {'none': True},
                            'num_font': {'size': 8, 'rotation': -70},
                            'name': 'Balises',
                            'name_font': {'size': 10, 'bold': False},
                            'label_position': 'low'
                            })
    df = dfInput[~dfInput['OVEREST_ERROR'].isna()]
    maxover = abs(df['OVEREST_ERROR'].min())
    df = dfInput[~dfInput['UNDEREST_ERROR'].isna()]
    maxunder = df['UNDEREST_ERROR'].max()

    maxaxis = max([maxover, maxunder])
    strAux = str(maxaxis)
    if strAux == 'nan': 
        maxaxis = 0.0
    maxaxis = 0.1 + round(maxaxis, 1)

    odoestLineChart.set_y2_axis({'visible': True,
                                'line': {'none': True},
                                'num_font': {'size': 10},
                                'num_format': '0.00%',
                                'name': 'estimation/distance (%)',
                                'name_font': {'size': 10, 'bold': False},
                                'max': maxaxis,
                                'min': -maxaxis
                                })
    
           
    # Add L_DOUBTOVER values serie
    lineWidth = 0.75
    valuecolnum = 1+ dfInput.columns.get_loc('L_DOUBTOVER')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoestLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth},
                            'name': ['ROUTE', 0, valuecolnum]
    
    })
    # Add L_DOUBTUNDER values serie
    lineWidth = 0.75
    valuecolnum = 1+ dfInput.columns.get_loc('L_DOUBTUNDER')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoestLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth},
                            'name':       ['ROUTE', 0, valuecolnum]
                            })
    

    # Add overestimation error serie
    lineWidth = 1.5
    valuecolnum = 1+ dfInput.columns.get_loc('OVEREST_ERROR')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoestLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': '#4A7EBB',
                                            'dash_type': 'round_dot'},
                            'name': ['ROUTE', 0, valuecolnum],
                            'y2_axis': True,
                            })
    
    # Add underestimation error serie
    lineWidth = 1.5
    valuecolnum = 1+ dfInput.columns.get_loc('UNDEREST_ERROR')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    odoestLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': '#BE4B48',
                                            'dash_type': 'round_dot'},
                            'name': ['ROUTE', 0, valuecolnum],
                            'y2_axis': True,
                            })
    
    return odoestLineChart

# R_DrawOverUnderSlopeChart(wb, dfSingleMission)
# Draw odo_err chart
def R_DrawOverUnderSlopeChart(wb, dfSingleMission):

    global glSlopeSS41Th

    dfInput = pandas.DataFrame(dfSingleMission)
    
            
    # Create speed chart
    slopeLineChart = wb.add_chart({'type': 'line'})
    
            
            
    # Format chart (# 1 cm= 37,79527559055 px)
    # default with 480 px (12,7 cm)
    lastrow = len(dfInput.axes[0]) # Last excel row for chart
    chartWidth = float(lastrow/10*37.79)
    if chartWidth < 480: chartWidth = 480
    slopeLineChart.set_title({'name': 'OVER UNDER ESTIMATION SLOPE (%)',
                                  'name_font': {'size': 14, 'bold': False}})
    slopeLineChart.set_size({'width': chartWidth,
                         'y_scale': 1.5})
    slopeLineChart.set_legend({'position': 'bottom'})

    df = dfInput[~dfInput['L_DOUBTOVER'].isna()]
    maxover = abs(df['L_DOUBTOVER'].min())
    df = dfInput[~dfInput['L_DOUBTUNDER'].isna()]
    maxunder = df['L_DOUBTUNDER'].max()

    maxaxis = max([maxover, maxunder])
    strAux = str(maxaxis)
    if strAux == 'nan': 
        maxaxis = 0.0
    maxaxis = 10 + round(maxaxis)

    slopeLineChart.set_y_axis({
                        'major_gridlines': {
                        'visible': True,
                        'line': {'width': 0.75, 'color': '#D9D9D9'}
                        },
                        'line': {'none': True},
                        'num_font': {'size': 10},
                        'name': 'Confident interval (m)',
                        'name_font': {'size': 10, 'bold': False},
                        'max': maxaxis,
                        'min': -maxaxis
                        })
    slopeLineChart.set_x_axis({'major_tick_mark': 'none',
                            'line': {'none': True},
                            'num_font': {'size': 8, 'rotation': -70},
                            'name': 'Balises',
                            'name_font': {'size': 10, 'bold': False},
                            'label_position': 'low'
                            })
    df = dfInput[~dfInput['OVERSLOPE'].isna()]
    maxover = abs(df['OVERSLOPE'].min())
    df = dfInput[~dfInput['UNDERSLOPE'].isna()]
    maxunder = df['UNDERSLOPE'].max()

    maxaxis = max([maxover, maxunder])
    strAux = str(maxaxis)
    if strAux == 'nan': 
        maxaxis = 0.0
    maxaxis = 0.01 + round(maxaxis, 1)
    if maxaxis < glSlopeSS41Th:
        maxaxis = 0.01 + glSlopeSS41Th

    slopeLineChart.set_y2_axis({'visible': True,
                                 'line': {'none': True},
                                'num_font': {'size': 10},
                                'num_format': '0.00%',
                                'name': 'over/under estimation slope (%)',
                                'name_font': {'size': 10, 'bold': False},
                                'max': maxaxis,
                                'min': -maxaxis
                                })
    
           
    # Add L_DOUBTOVER values serie
    lineWidth = 0.5
    valuecolnum = 1+ dfInput.columns.get_loc('L_DOUBTOVER')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    slopeLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth},
                            'name': ['ROUTE', 0, valuecolnum]
                            })
    # Add L_DOUBTUNDER values serie
    lineWidth = 0.5
    valuecolnum = 1+ dfInput.columns.get_loc('L_DOUBTUNDER')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    slopeLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth},
                            'name':       ['ROUTE', 0, valuecolnum]
                            })
    

    # Add overestimation slope serie
    lineWidth = 2.0
    valuecolnum = 1+ dfInput.columns.get_loc('OVERSLOPE')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    slopeLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': '#4A7EBB',
                                            'dash_type': 'round_dot'},
                            'name': ['ROUTE', 0, valuecolnum],
                            'y2_axis': True,
                            })
    
    # Add underestimation slope serie
    lineWidth = 2.0
    valuecolnum = 1+ dfInput.columns.get_loc('UNDERSLOPE')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    slopeLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': '#BE4B48',
                                            'dash_type': 'round_dot'},
                            'name': ['ROUTE', 0, valuecolnum],
                            'y2_axis': True,
                            })
    
    # Add overestimation slope threshold serie
    lineWidth = 1.0
    valuecolnum = 1+ dfInput.columns.get_loc('SS41_OVERSLOPE')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    slopeLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': 'red',
                                            'dash_type': 'round_dot'},
                            'name': ['ROUTE', 0, valuecolnum],
                            'y2_axis': True,
                            })
    
    # Add overestimation slope threshold serie
    lineWidth = 1.0
    valuecolnum = 1+ dfInput.columns.get_loc('SS41_UNDERSLOPE')
    catcolnum = 1 + dfInput.columns.get_loc('NID_BG') 
    slopeLineChart.add_series({
                            'categories': ['ROUTE', 1, catcolnum, lastrow, catcolnum],
                            'values':     ['ROUTE', 1, valuecolnum, lastrow, valuecolnum],
                            'line':       {'width': lineWidth, 
                                            'color': 'orange',
                                            'dash_type': 'round_dot'},
                            'name': ['ROUTE', 0, valuecolnum],
                            'y2_axis': True,
                            })
    

    
    
    return slopeLineChart

# -----------------------------------  R_Ato_Functions    -------------------------------------------
# Ato functions. Input file R_File (normalized routes files)
# R_Ato_GetRoutesToCompare(strFile)
# Get routes info to compare 
# Get summary of missions from file
# Get list of balises in mission
def R_Ato_GetRoutesToCompare(strFile):

    outDfMissionSummary = pandas.DataFrame()
    outDfBGsInRoute = pandas.DataFrame()
    #serieBGs = pandas.Series()
    outArrOfDataframes = []

    dfMissions = pandas.DataFrame()
    dfRoutes = pandas.DataFrame()
    dfBook = pandas.DataFrame()

    try:
        dfBook = pandas.read_excel(strFile, 
                    sheet_name = None,
                    dtype={'NID_BG': str})
    except Exception as ex:
        print(color.Fore.RED + "Can´t read file: " + str(ex) + color.Style.RESET_ALL) 
        return outArrOfDataframes   

    dfMissions = dfBook['MISSIONS']
    # Removing unnamed columns using drop function
    dfMissions.drop(dfMissions.columns[dfMissions.columns.str.contains(
                    'unnamed', case=False)], axis=1, inplace=True)
                    
    dfRoutes = dfBook['ROUTES']
    # Removing unnamed columns using drop function
    dfRoutes.drop(dfRoutes.columns[dfRoutes.columns.str.contains(
                'unnamed', case=False)], axis=1, inplace=True)
    # file name format: 'C:\\Users\\manuel.caceres\\Documents\\Desarrollo\\src\\inputs\\test\\R_2024-03-04_1_17244.xlsx'
    strAux = strFile
    ix = 0
    n = 0
    while True:
        try:   
            iAux = strAux.index('\\')
            strAux = strAux[iAux +1:len(strAux)]
            if n == 0:
                ix += iAux
            else:
                ix += (iAux + 1)
            n += 1
        except:
            break
    fileName = strAux
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
    fileNameNoext = fileName[0:ix].lstrip()

    listSerBGs = []
    outDfMissionSummary = pandas.DataFrame(dfMissions)
    emptyList = []
    for x in range (len(outDfMissionSummary.axes[0])):
        emptyList.insert(x, "")

    outDfMissionSummary.insert(0, 'MISSION_ID', emptyList)
    outDfMissionSummary.insert(0, 'FILENAME', emptyList)
    

    for x in range(len(dfMissions.axes[0])):
        ixSoM = dfMissions['IxSoM'][x]
        ixEoM = dfMissions['IxEoM'][x]
        
        # Filter by each mission
        dfSingleMission = dfRoutes.iloc[ixSoM:ixEoM+1]
        missionid = fileNameNoext + '_' + str(x)
        columnName = fileNameNoext + '_' + str(x)
        serieBGs = dfSingleMission[dfSingleMission['NID_BG'].str.contains('-0', na = False)]['NID_BG'].apply(utl_removeBaliseExension)
        serieBGs.rename(columnName, inplace=True) 
        serieBGs.reset_index(drop=True, inplace=True)
        listSerBGs.insert(x, serieBGs)
        outDfMissionSummary.iloc[x, outDfMissionSummary.columns.get_loc('MISSION_ID')] = missionid
        outDfMissionSummary.iloc[x, outDfMissionSummary.columns.get_loc('FILENAME')] = fileName
        
    if len(listSerBGs) > 0: # File has at least one mission
        outDfBGsInRoute = pandas.concat(listSerBGs, axis=1)

        outArrOfDataframes.insert(0, outDfMissionSummary)
        outArrOfDataframes.insert(1, outDfBGsInRoute)


    return outArrOfDataframes

# R_Ato_CompareRoutes(arrInputRoutes)
# Returns an array of matching routes (array of cRoutePath objects)
# arrInputRoutes: array of routes to compare (class cRoutePath)
# percMatch: min % of matched balises to consider matching routes
# numMatch: min number of matches balises to consider relevant matching routes
#
def R_Ato_CompareRoutes(arrInputRoutes, minMatchPercentage, minMatchBalises):

    # Initialization
    if (len(arrInputRoutes)) > 0:
        routeLevel = arrInputRoutes[0].getLevel() # Level to compare
    else:
        routeLevel = 0 
    
    dfAlreadyCompared = pandas.DataFrame()
    listAlreadyCompared = []
    boolAlreadyCompared = False

    # Auxiliar function to compare list of balises within the route
    def utl_Equal(list):
        # list[0]: left
        # list[1]: shifted
        if(list[0] == list[2]):
            return 1
        else:
            return 0

    arrMatchingdInputRoutes = []

    #if((minMatchPercentage == 0) | (minMatchBalises == 0)):
    #    return arrMatchingdInputRoutes
    
    # Loops to compare all routes in arrInputRoutes
    numSeq = 0 # Initialization
    for x in range(len(arrInputRoutes)):
        boolAlreadyCompared = False
        # left: first route
        left = arrInputRoutes[x].getListBG()
        if (len(left) < minMatchBalises):
            break # Less number of balises than minimum required
        serleft = pandas.Series(left)
        for y in range(x+1,len(arrInputRoutes)):
            # right second route
            right = arrInputRoutes[y].getListBG()
            if (len(right) < minMatchBalises):
                break # Less number of balises than minimum required

            # If level >= 1 check if both routes have been already processed
            if routeLevel > 0:
                listParents = []
                listParents += arrInputRoutes[x].getLevel0ParentsList()
                listParents += arrInputRoutes[y].getLevel0ParentsList()
                listParents = list(set(listParents)) # unique values
                listParents.sort()

                
                for z in range(len(dfAlreadyCompared.axes[0])):
                    listAlreadyCompared = dfAlreadyCompared.iloc[z].to_list()
                    listAlreadyCompared = [value for value in listAlreadyCompared if str(value) != 'nan']
                    listAlreadyCompared.sort()
                    if len(listParents) < len(listAlreadyCompared):
                        listA = listParents
                        listB = listAlreadyCompared
                    else:
                        listB = listParents
                        listA = listAlreadyCompared
                    
                    if utl_is_sublist(listA, listB):
                        boolAlreadyCompared = True
                        break
                
            if boolAlreadyCompared:
                continue # Level 0 routes have been previously   

            
            serright = pandas.Series (right)  
            # shift: auxiliar serie to shift second route and compare with first one        
            sershift = pandas.Series(serright)
            # compare: auxiliar serie to identify balise macthing in both routes
            # 1 when matching balises; 0 when not
            compare = pandas.Series ()
            # Data frame with routes to compare and aux series for comparison process
            dfMatchingRoutes = pandas.DataFrame({'left': serleft, 'right': serright, 'shift': sershift, 'compare': compare})

            # Compare routes
            nMatchingBalises = 0
            # dfMatchingBlocks: auxiliar dataframe with matching blocks within both routes
            # contains: 
            # 'leftindex': left index starting index in left route
            # 'nshift': shifted index in right route to match with left route
            # 'nmatch': number of matching balises in the block
            dfMatchingBlocks = pandas.DataFrame() 
            listLeftindex = []
            listNshift = []
            listSum = []            
            nmatchIndex = 0 # variable to 
            #init = -(len(df['right'].axes[0]) - minMatchBalises)
            #length = len(df['right'].axes[0]) - minMatchBalises + 1
            init = -(len(dfMatchingRoutes['right'].axes[0]))
            length = len(dfMatchingRoutes['right'].axes[0]) + 1
            for n in range(init, length):

                dfMatchingRoutes['shift'] = dfMatchingRoutes['right'].shift(n, fill_value=numpy.nan)
                dfMatchingRoutes['compare'] = dfMatchingRoutes.apply(utl_Equal, axis=1)
                match = dfMatchingRoutes['compare'].sum()
                if match > 2: # 2 is minimun number of matching balise to take into account the matching
                    listLeftindex.insert(nmatchIndex, dfMatchingRoutes[dfMatchingRoutes['compare'] == 1].index[0]) # first matching balise index)
                    listNshift.insert(nmatchIndex, n)
                    listSum.insert(nmatchIndex, match)
                    nmatchIndex += 1

            # df matching to contain different matching blocks with the shifted index and nimber of matching balises in the block
            if nmatchIndex > 0:
                dfMatchingBlocks['leftindex'] = numpy.array(listLeftindex)
                dfMatchingBlocks['nshift'] = numpy.array(listNshift)
                dfMatchingBlocks['nmatch'] = numpy.array(listSum)  

                # Case of duplicated balises in some of compared missions; delete blocks wit less number of matching balises
                serList = pandas.Series([True]*len(dfMatchingBlocks.axes[0]))
                for ix in range(len(dfMatchingBlocks.axes[0])):
                    leftindex = dfMatchingBlocks['leftindex'][ix]
                    if len(dfMatchingBlocks[dfMatchingBlocks['leftindex'] == leftindex]) > 1: # Duplicated balise in mission
                        print('Info. Balise found more than once in mission: '
                            +  str(dfMatchingRoutes['left'][leftindex]) + '.')
                        print('Removing duplicated path.')
                        ixmax = dfMatchingBlocks[dfMatchingBlocks['leftindex'] == leftindex]['nmatch'].idxmax() # Search block with max number of matching balises
                        serList &= dfMatchingBlocks['leftindex'].apply(lambda x: x != leftindex)
                        serList |= ~pandas.Series(dfMatchingBlocks.index.values.tolist()).apply(lambda x: x != ixmax)
                        
                        
                dfMatchingBlocks = dfMatchingBlocks[serList]
                dfMatchingBlocks.reset_index(drop=True, inplace=True)
                        
                nMatchingBalises = dfMatchingBlocks['nmatch'].sum()
            
            # order dfMatching by left index
            if len(dfMatchingBlocks.axes[0]) > 1:
                dfMatchingBlocks.sort_values(by = ['leftindex'], inplace = True, ignore_index=True)
            
            if ((nMatchingBalises > 0) & (nMatchingBalises >= minMatchBalises)): # matching number higher than min expected
                
                # Check matching percentage
                boolCheckPercentage = True
                while boolCheckPercentage == True:

                    numBlocks = len(dfMatchingBlocks.axes[0])                
                    ixFirst = dfMatchingBlocks['leftindex'][0]
                    ixLast = dfMatchingBlocks['leftindex'][numBlocks-1] + dfMatchingBlocks['nmatch'][numBlocks-1] - 1
                    dfAux = dfMatchingRoutes.iloc[ixFirst:ixLast+1]
                    nmatch = dfMatchingBlocks['nmatch'].sum()   

                    
                    matchPercentage = nmatch/dfAux['left'].count()
                    
                    if numBlocks <= 1: # Case one matching block; no way of improving matching percentage
                        break 
                    if matchPercentage >= minMatchPercentage: # Case: required matching percentage achieved
                        break
                    
                    # Case: requiered percentage not achieved and more than one matching block
                    # In this case we can remove blocks to increase percentage
                    nmatchFor = dfMatchingBlocks['nmatch'][0]
                    nmatchBack = dfMatchingBlocks['nmatch'][len(dfMatchingBlocks['nmatch']) - 1]

                    if(nmatchFor < nmatchBack):
                        dfMatchingBlocks.drop([0], inplace=True)
                    else:
                        dfMatchingBlocks.drop([len(dfMatchingBlocks['nmatch']) - 1], inplace=True)
                    
                    dfMatchingBlocks.reset_index(drop=True, inplace=True)
                    nMatchingBalises = dfMatchingBlocks['nmatch'].sum()

            # Check again as number of matching balises could be reduced above (when trying to increase matching percentage)
            if ((nMatchingBalises > 0) & (nMatchingBalises >= minMatchBalises)): # matching NEW number higher than min expected
              
                # if matching percentage greater than expected, set matched route
                if matchPercentage >= minMatchPercentage:
                    listBGs = [] # Initialization
                    for p in range(len(dfMatchingBlocks.axes[0])):
                        # shift list to matching block
                        dfMatchingRoutes['shift'] = dfMatchingRoutes['right'].shift(dfMatchingBlocks['nshift'][p], fill_value=numpy.nan)
                        # set matching list
                        dfMatchingRoutes['compare'] = dfMatchingRoutes.apply(utl_Equal, axis=1)
                        # find first and last index for matching balises
                        ixFirst = dfMatchingBlocks['leftindex'][p]
                        ixLast = dfMatchingBlocks['leftindex'][p] + dfMatchingBlocks['nmatch'][p] - 1
                        print('left: ' + str(x) + '-' + 'right:' + str(y) + '-' + str(dfMatchingRoutes['left'][ixFirst]) +
                               '-' + str(dfMatchingRoutes['left'][ixLast]) + '-' + str(dfMatchingBlocks['nmatch'][p]) + '-' + str(round(matchPercentage, 4) * 100) + "%")
                        dfAux = dfMatchingRoutes.iloc[ixFirst:ixLast+1]
                        dfAux.reset_index(drop=True, inplace=True)  # !!!!!!!
                        dfNotmatch = dfAux[dfAux['compare'] == 0]
                        cindex = 0
                        listIndex = dfNotmatch.index.values.tolist()
                        prevIndex = 0
                        for i in listIndex:   
                            if (prevIndex > 0) & (prevIndex != (i-1)):
                                cindex = 0 # reset                     
                            if cindex == 0:
                                dfAux.iloc[i, dfAux.columns.get_loc('left')] = '*' # delete balise
                            else:
                                dfAux.iloc[i, dfAux.columns.get_loc('left')] = 'delete' # delete balise
                            cindex += 1
                            prevIndex = i
                        dfAux = dfAux[dfAux['left'] != 'delete']
                        dfAux.reset_index(drop=True, inplace=True)
                        # listBGs = append listBGs from each block
                        # separate each block with an '*'
                        if len(listBGs) == 0:
                            listBGs = dfAux['left'].tolist()
                        else:
                            listBGs.insert(len(listBGs), '*') # insert '*' between blocks of matching balises
                            listBGs.extend(dfAux['left'].tolist())
                        
                    # Next: create matching routes objects
                    # parent routes: x and y
                    # Case 2: *args = id, parentList, listBG
                    prefix = 'PATH'
                    parentList = [arrInputRoutes[x], arrInputRoutes[y]]
                    level = arrInputRoutes[x].getLevel() + 1
                    objectid = prefix + '_' + str(level) + '_' + f"{numSeq:03d}"
                    numSeq += 1 # Next one
                    objRPath = cRoutePath(objectid, parentList, listBGs)
                    objRPath.setMatchPerc(matchPercentage)
                    arrMatchingdInputRoutes.insert(len(arrMatchingdInputRoutes), objRPath)

                    # If level >= 1 set routes that have already been processed
                    if routeLevel > 0:
                        listParents = []
                        listParents += arrInputRoutes[x].getLevel0ParentsList()
                        listParents += arrInputRoutes[y].getLevel0ParentsList()
                        listParents = list(set(listParents)) # unique values
                        listParents.sort()
                        dfTemp = pandas.DataFrame(listParents)
                        dfTemp = dfTemp.transpose()

                        dfAlreadyCompared = pandas.concat([dfAlreadyCompared, dfTemp], ignore_index=True)
                    
    # Remove paths included in next level
    listdel = []
    xdel = 0
    newArrRoutePaths = []
    for n in range(len(arrInputRoutes)):
        newArrRoutePaths.insert(n, arrInputRoutes[n]) # copy

    for xhigh in range(len(arrMatchingdInputRoutes)):
        objhigh = arrMatchingdInputRoutes[xhigh]
        listhigh = objhigh.getListBG()
        for xlow in range(len(newArrRoutePaths)):
            objlow = newArrRoutePaths[xlow]
            listlow = objlow.getListBG()
            if len(listhigh) == len(listlow):
                if listlow == listhigh:
                    listdel.insert(xdel, xlow)
    
    listdel = list(set(listdel)) # unique values  
           
    for xlow in listdel:
        newArrRoutePaths[xlow] = None
   
    newArrRoutePaths = [item for item in newArrRoutePaths if item is not None] 

    arrOutput = [arrMatchingdInputRoutes, newArrRoutePaths]

    return arrOutput

# classes
class cRoutePath:
  
    def __init__(self, *args):

        # Case 1: *args = id, listBG, nidengine, missionid, pathFile, routeFile
        # Case 2: *args = id, parentList, listBG

        if len(args) == 6:
            id = args[0]
            listBG = args[1]
            nidengine = args[2]
            missionid = args[3]
            pathFile = args[4]
            routeFile = args[5]

            self.id = id
            self.parentList = []
            self.level = 0
            self.missionid = missionid
            self.pathFile = pathFile
            self.routeFile = routeFile
            self.listBG = listBG
            self.nidengine = nidengine
            self.matchperc = 0.0
            
        elif len(args) == 3:

            id = args[0]
            parentList = args[1]
            listBG = args[2]

            if len(parentList) != 2:
                return None # Not allowed
            elif parentList[0].getLevel() != parentList[1].getLevel():
                return None # Not allowed
            
            self.id = id
            self.parentList = [None, None]    
            self.parentList[0] = parentList[0].shallow_copy()
            self.parentList[1] = parentList[1].shallow_copy()
            self.level = parentList[0].getLevel() + 1 # Both parents have the same level
            self.missionid = ''
            self.nidengine = ''
            self.routeFile = ''
            self.pathFile = ''
            self.listBG = listBG
            self.matchperc = 0.0
            
        else:
            return None

    def deep_copy(self):
        return copy.deepcopy(self)

    def shallow_copy(self):
        return copy.copy(self)
    
    def getId(self):
        return self.id
    
    #def setParent(self, parent):
    #    if parent == None:
    #        self.parent = None
    #    else:
    #        self.parent = parent.shallow_copy()
    
    def getParentList(self):
        return self.parentList
    
    def getLevel0ParentsList(self):

        if self.level == 0:  
            outputlist = [self.id]
        else:
            if len(self.parentList) < 2: return 
            outputlist = self.parentList[0].getLevel0ParentsList()
            outputlist += self.parentList[1].getLevel0ParentsList()
        return outputlist
            

    def getMissionid(self):
        return self.missionid
    
    def getPathFile(self):
        return self.pathFile
    
    def getNidengine(self):
        return self.nidengine
 
    def setNidengine(self, nidengine):
        self.nidengine = nidengine
    
    def getRouteFile(self):
        return self.routeFile

    #def setLevel(self, level):
    #    self.level = level

    def getLevel(self):
        return int(self.level)

    def getListBG(self):
        return self.listBG
    
    def getNumBG(self):

        list = self.listBG
        list = [item for item in list if str(item) != '*']
        list = [item for item in list if str(item) != 'nan']

        return len(list)
    
    def setMatchPerc(self, perc):
        self.matchperc = perc
        return True
    
    def getMatchPerc(self):

        return self.matchperc


# Utilities
def utl_changeSign(x):
    if x != "": return -x
    else: return x

def utl_delUnknownValue(x):
    if x == 32767: return ''
    elif x == '32767': return ''
    else: return x

def utl_removeBaliseExension(extendedBalise):
    strBG = str(extendedBalise)
    try:
        ix = strBG.index('-0')
        return(strBG[0:ix])
    except:
        return strBG
    

def utl_ChangeTypeStrFromInt (var):
    try:
        var = str(int(var))
    except:
        var = ''
    return var

def utl_ChangeTypeFloatFromStr (var):
    try:
        var = float(str(var))
    except:
        var = ''
    return var

def utl_ChangeDateFormat (date):
    try:
        strDate = str(date)
        i = strDate.index('/')
        dd = strDate[0:i].rjust(2, '0')
        strDate = strDate[i+1:len(strDate)]
        i = strDate.index('/')
        mm = strDate[0:i].rjust(2, '0')
        yy = strDate[i+1:len(strDate)]
        strDate = dd + '/' + mm + '/' + yy
        
    except:
        strDate = ''

    return strDate


def utl_ChangeTimeFormat (date):
    try:
        strTime = str(date)
        i = strTime.index(':')
        hour = strTime[0:i].rjust(2, '0')
        strTime = strTime[i+1:len(strTime)]
        i = strTime.index(':')
        minutes = strTime[0:i].rjust(2, '0')
        strTime = strTime[i+1:len(strTime)]
        try:
            i = strTime.index('.') # case seconds with decimals
            seconds = strTime[0:i].rjust(2, '0')
        except:
            seconds = strTime[0:len(strTime)].rjust(2, '0')
        strTime = hour + ':' + minutes + ':' + seconds # H:m:s
        
    except:
        strTime = ''

    return strTime
        
def utl_GetSxy(list, ymean, xmean):
    # 
    # Sxy = (x-xmean)*(y-ymean)
    # list[0] = y
    # list[1] = x
    # To access a value by position, use `ser.iloc[pos]
    
    sxy = (list.iloc[1]-xmean)*(list.iloc[0] - ymean)

    return sxy

def utl_GetSxx(list, xmean):
    # 
    # Sxx = (x-xmean)^2
    # list[0] = y
    # list[1] = x
    # To access a value by position, use `ser.iloc[pos]
    
    sxx = (list.iloc[1]-xmean)**2

    return sxx


def utl_GetMxOdo(x, slope):
    y = x*slope
    return y

def utl_is_sublist(A, B):
    if not A:
        return True
    if not B:
        return False
    if A[0] == B[0]:
        return utl_is_sublist(A[1:], B[1:])
    return utl_is_sublist(A, B[1:])