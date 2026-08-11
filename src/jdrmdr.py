import pandas as pd
import numpy as np
import re
import sys
import colorama as color
import xlsxwriter as xls
import warnings

import config as cfg
import jdrmdrcfg as jdrmdrcfg
import ss027

# Version control
# Author: Manuel Cáceres Marzal
# V1.0 Date 30/10/2025
#   First version pre-compiling
#
# V1.1 Date 09/12/2025
#   Adding compiling phase
#
#   Pending:
#       - Take into account Q_DIR for packets reseived from TS
#
# V 2.1 Date 06/02/2026
#   - Update version number to make it consistent across all modules
#
# V 3.0 Date 06/02/2026
#   - Include BGERROR
#
# V 3.1 Date 03/03/2026
#   - Included in charts: EBC, SBC, START, DMIACK, DRIVERACK
#
# V3.1.1  Date 20/07/2026
#
# Fixed bug at D_REF get value:
#   Before: regular expression '[+-]?\d*[.,]\d+|\d+'   removes '-' when integer
#   After: change for regular expression '[+-]?(?:\d*[.,]\d+|\d+)'

jdrmdrVersion = '3.1.1'

# Global variables

# Set of global variables to track progress
glNumberofCompiledRegisters = 0
glPh1Progress = 0.0
glPh1Functions = 0 
glPh1NumberRows = 0
glFunctionNum = 0

# global variable glLocale for content format
# Values
# ENG for english number formats and texts
#   , for thousands
#   . for decimals
# SPA for spanish number format and texts
#   . for thousands
#   , for decimals

glLocale = cfg.glLocale

# Functions over original Alstom JRU file
# naming: AlFunctionName


# AlPrecompile(inputFile)
# Alstom decoder
# Translate Alstom text file with JRU messages to an excel file very similar to Hasler JRU files
# Input: text file decoded with alstom JRMDR (test with JRMDRversion 7.6.0)
# Output file columns:
# Record Id
# Date
# Time
# Distance (km)
# V_TRAIN (km/h)
# MESSAGE ID
# MESSAGE NAME
# LENGTH
# HEADER
# DATA
# PACKET RBC
# PACKET

def AlPrecompile(inputfile):
    
    
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum
    global glNumberofCompiledRegisters

    glPh1Progress = 0.0
    glPh1Functions = 4 
    glPh1NumberRows = 0
    glFunctionNum = 0

    # DataFrames with input data
    dfJruMessages = pd.DataFrame ()
    # Data Frame for pre compiled data
    dfPrecompiledAlstom = pd.DataFrame()

    # Open file and put JRU messages in dataframe
    listJruMessages = []
    msgId = None
    msgIdPrefix = jdrmdrcfg.msgIdPrefix
    block = []
    level = 0
    reading = False
    print("Loading Alstom File. " + inputfile)
    
    with open(inputfile, encoding="utf-8", errors='ignore') as f: 
        # !Important: errors='ignore' used to avoid errors when non-utf-8 chars into the file
        for line in f:
            line = line.rstrip("\n")
            if line.startswith(msgIdPrefix):
                msgId = line.strip()
                msgId = msgId[len(msgIdPrefix):len(msgId)]
            elif "JRU (" in line:
                reading = True
                level = 1
                block = []
            elif reading:

                if ":" not in line:
                    level += line.count("(")
                    level -= line.count(")")
                    #level += line.count("(")
                    #level -= line.count(")")
                        
                if level == 0:
                    listJruMessages.append((msgId, "\n".join(block).strip() + '\n'))
                    reading = False
                    block = []
                else:
                    block.append(line)
            
        
    
    # list JRUMessages is a list with each list element an array of two fields:
    # [0] msgId
    # [1] JruBlock
    dfJruMessages = pd.DataFrame(np.array(listJruMessages),
                    columns=['Record Id', 'JRU'])
    
    print("Pre-compiling File")
    glNumberofCompiledRegisters = len(dfJruMessages.axes[0])

    glPh1NumberRows = len(dfJruMessages.axes[0])
    
    sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
    sys.stdout.flush()

    # DataFrames with partial data
    dfHeader = pd.DataFrame()

    # Extracting data from Alstom JRU
    glFunctionNum = 0
    listDfReturn = AlGetHeaderFromJru(dfJruMessages)

    # listDfReturn = [dfHeader, dfDate, dfTime, dfDistance, dfSpeed, 
    #              dfJruNidMessage, dfJruMessageName, dfJruMessageLength]

    dfHeader = listDfReturn[0]
    dfDate = listDfReturn[1]
    dfTime = listDfReturn[2]
    dfDistance = listDfReturn[3]
    dfSpeed = listDfReturn[4]
    dfJruNidMessage = listDfReturn[5]
    dfJruMessageName = listDfReturn[6]
    dfJruMessageLength = listDfReturn[7]

    # Get common data from JRU message
    glFunctionNum += 1
    dfData = AlgetDataFromJru (dfJruMessages)

    # Get RBC packets from JRU message
    glFunctionNum += 1
    dfRbcPacket = AlgetRbcPacketsFromJru (dfJruMessages)

    # Get balise packets from JRU message
    glFunctionNum += 1
    dfBalisePacket = AlgetBalisePacketsFromJru (dfJruMessages)

    sys.stdout.write("\n")
    sys.stdout.flush()
        

    # Aggregate all Data Frames
    listDf = []
    x = 0
    if len(dfJruMessages['Record Id']) >0:
        listDf.insert(x, dfJruMessages['Record Id'])
        x+=1
    if len(dfDate) >0:
        listDf.insert(x, dfDate)
        x+=1
    if len(dfTime) >0:
        listDf.insert(x, dfTime)
        x+=1
    if len(dfDistance) >0:
        listDf.insert(x, dfDistance)
        x+=1
    if len(dfSpeed) >0:
        listDf.insert(x, dfSpeed)
        x+=1
    if len(dfJruNidMessage) >0:
        listDf.insert(x, dfJruNidMessage)
        x+=1
    if len(dfJruMessageName) >0:
        listDf.insert(x, dfJruMessageName)
        x+=1
    if len(dfJruMessageLength) >0:
        listDf.insert(x, dfJruMessageLength)
        x+=1
    if len(dfHeader) >0:
        listDf.insert(x, dfHeader)
        x+=1
    if len(dfData) >0:
        listDf.insert(x, dfData)
        x+=1
    if len(dfRbcPacket) >0:
        listDf.insert(x, dfRbcPacket)
        x+=1
    if len(dfBalisePacket) >0:
        listDf.insert(x, dfBalisePacket)
        x+=1
    
    #if len(dfCommondata) >0:
    #    listDf.insert(x, dfCommondata)
    #    x+=1
    
    try:
        dfPrecompiledAlstom = pd.concat(listDf, axis=1)
    except Exception as ex:
        print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)

    # Change type of Record Id as integer to sort by it
    def loc_ChangeTypeIntFromStr (var):
        try:
            var = int(str(var))
        except:
            var = None
        return var
    serGrabarId = pd.Series(dfPrecompiledAlstom['Record Id'].apply(loc_ChangeTypeIntFromStr))
    dfPrecompiledAlstom['Record Id'] = serGrabarId  
    # Sort dataframe by 'Record Id' descendant to fix data from start
    dfPrecompiledAlstom.sort_values(by='Record Id', ascending=False, inplace=True)
    
    # drop column 0
    # dfPrecompiledAlstom.set_index('Record Id', inplace=True)

    return dfPrecompiledAlstom


# AlCompile(hasDf, listRbcFilter, listJruFilter)
# Alstom (JDRMDR) decoder
# Compile Alstom precompilde file
#
# Input dataframe columns:
#
# # Record Id
# Date
# Time
# Distance (km)
# V_TRAIN (km/h)
# MESSAGE ID
# MESSAGE NAME
# LENGTH
# HEADER
# DATA
# PACKET RBC
# PACKET
#
# Output file columns:
# NID_MESSAGE
# Record Id / Record Id
# 	Date
# 	Time
# 	NID_ENGINE
# 	Q_SCALE	NID_LRBG
# 	D_LRBG	Q_DIRLRBG
# 	Q_DLRBG
# 	L_DOUBTOVER
# 	L_DOUBTUNDER
# 	V_TRAIN
# 	M_LEVEL
# 	M_MODE
# 	PACKET_RBC
# 	D_REF
# 	Rbc Packets
# 	NID_BG
# 	Q_LINK
# 	Balise Packets
# 	BG_LINKS
# 	D_LINKS


def AlCompile(dfPrecompiled, listRbcFilter, listJruFilter, listJruDiscard):
    
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum
    global glNumberofCompiledRegisters

    glPh1Progress = 0.0
    glPh1Functions = 16
    glPh1NumberRows = 0
    glFunctionNum = 0

    # Make a copy of input dataframe
    inputdf = pd.DataFrame(dfPrecompiled)

    print("Compiling File")
    glNumberofCompiledRegisters = len(inputdf.axes[0])

    inputdf = AlApplyRowFilter(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glPh1NumberRows = len(inputdf.axes[0])
    
    sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
    sys.stdout.flush()

    # DataFrames with partial data
    dfNidmessage = pd.DataFrame()
    dfCommondata = pd.DataFrame()
    dfTrainEngine = pd.DataFrame()
    dfTrainpos = pd.DataFrame()
    dfVeltrain = pd.DataFrame()
    dfEtcslevel = pd.DataFrame()
    dfEtcsmode = pd.DataFrame()
    dfPacketrbc = pd.DataFrame()
    dfDatabalise = pd.DataFrame()
    dfPacketBalise = pd.DataFrame()
    dfVperm = pd.DataFrame()
    dfNidErrorBg = pd.DataFrame()
    dfMessageErrorBg = pd.DataFrame()
    dfServiceBrakeCommandState = pd.DataFrame()
    dfEmergencyBrakeCommandState = pd.DataFrame()
    dfMoSymbolAck = pd.DataFrame()
    dfStSymbolRadioUp = pd.DataFrame()
    dfStSymbolRadioDown = pd.DataFrame()
    dfDriverAction = pd.DataFrame()
    dfTcRadioHole = pd.DataFrame()
    

    # Data Frame for compiled data
    dfCompiledAlstom = pd.DataFrame()

    # Extract data from type-hasler JRU
    glFunctionNum = 0
    dfNidmessage = AlGetNidmesageFromMessage(inputdf)
    glFunctionNum += 1
    dfCommondata = AlGetCommonData(inputdf)
    glFunctionNum += 1
    dfTrainEngine = AlGetTrainEngineFromHeader(inputdf)
    glFunctionNum += 1
    dfTrainpos = AlGetTrainPosFromHeader(inputdf)
    glFunctionNum += 1
    dfVeltrain = AlGetVelocityFromHeader(inputdf)
    glFunctionNum += 1
    dfEtcslevel = AlGetetcsLevelFromHeader(inputdf)
    glFunctionNum += 1
    dfEtcsmode = AlGetetcsModeFromHeader(inputdf)
    glFunctionNum += 1
    dfPacketrbc = AlGetPacketRBC(inputdf)
    glFunctionNum += 1
    dfDatabalise = AlGetDataFromBalise(inputdf)
    glFunctionNum += 1
    dfPacketBalise = AlGetPacketFromBalise(inputdf, dfTrainpos, dfDatabalise)
    glFunctionNum += 1
    dfVperm = AlGetVPerm(inputdf)
    # Get balise groupo error
    glFunctionNum += 1
    listDfReturn = AlgetBaliseGroupErrorFromData (inputdf)
    # listReturn = [dfNidErrorBg, dfMessageErrorBg] 
    dfNidErrorBg = listDfReturn[0]
    dfMessageErrorBg = listDfReturn[1]
    glFunctionNum += 1
    dfServiceBrakeCommandState = AlgetServiceBrakeCommandedFromData(inputdf)
    glFunctionNum += 1
    dfEmergencyBrakeCommandState = AlgetEmergencyBrakeCommandedFromData(inputdf)
    glFunctionNum += 1
    listDfReturn = AlgetDmiSymbolStatusFromData(inputdf)
    # listReturn = [dfMoSymbolAck, dfStSymbolRadioUp, dfStSymbolRadioDown, dfTcRadioHole]
    dfMoSymbolAck = listDfReturn[0]
    dfStSymbolRadioUp = listDfReturn[1]
    dfStSymbolRadioDown = listDfReturn[2]
    dfTcRadioHole = listDfReturn[3]
    glFunctionNum += 1
    dfDriverAction = AlgetDriverActionFromData(inputdf)
    
    sys.stdout.write("\n")
    sys.stdout.flush()
        

    # Aggregate all Data Frames
    listDf = []
    x = 0
    if len(dfNidmessage) >0:
        listDf.insert(x, dfNidmessage)
        x+=1
    if len(dfCommondata) >0:
        listDf.insert(x, dfCommondata)
        x+=1
    if len(dfTrainEngine) >0:
        listDf.insert(x, dfTrainEngine)
        x+=1
    if len(dfTrainpos) >0:
        listDf.insert(x, dfTrainpos)
        x+=1
    if len(dfVeltrain) >0:
        listDf.insert(x, dfVeltrain)
        x+=1
    if len(dfEtcslevel) >0:
        listDf.insert(x, dfEtcslevel)
        x+=1
    if len(dfEtcsmode) >0:
        listDf.insert(x, dfEtcsmode)
        x+=1   
    if len(dfPacketrbc) >0:
        listDf.insert(x, dfPacketrbc)
        x+=1 
    if len(dfDatabalise) >0:
        listDf.insert(x, dfDatabalise)
        x+=1
    if len(dfPacketBalise) >0:
        listDf.insert(x, dfPacketBalise)
        x+=1
    if len(dfVperm) >0:
        listDf.insert(x, dfVperm)
        x+=1
    if len(dfNidErrorBg) >0:
        listDf.insert(x, dfNidErrorBg)
        x+=1
    if len(dfMessageErrorBg) >0:
        listDf.insert(x, dfMessageErrorBg)
        x+=1
    if len(dfServiceBrakeCommandState) >0:
        listDf.insert(x, dfServiceBrakeCommandState)
        x+=1
    if len(dfEmergencyBrakeCommandState) >0:
        listDf.insert(x, dfEmergencyBrakeCommandState)
        x+=1
    if len(dfMoSymbolAck) >0:
        listDf.insert(x, dfMoSymbolAck)
        x+=1
    if len(dfStSymbolRadioUp) >0:
        listDf.insert(x, dfStSymbolRadioUp)
        x+=1
    if len(dfStSymbolRadioDown) >0:
        listDf.insert(x, dfStSymbolRadioDown)
        x+=1
    if len(dfDriverAction) >0:
        listDf.insert(x, dfDriverAction)
        x+=1
    if len(dfTcRadioHole) >0:
        listDf.insert(x, dfTcRadioHole)
        x+=1

    # Merge columns from balises and RBC

    def loc_MergeRBCBaliseData(list):
        # 
        # list[0] = bg_links from rbc
        # list[1] = bg_links from balises
        
        #if list[1] != '':
        #    serBglinks = list[1]
        #elif list[0] != '':
        #    serBglinks = list[0]
        #else:
        #    serBglinks = ''

        # In a future version, integer keys will always be treated as labels (consistent with 
        # DataFrame behavior). To access a value by position, use `ser.iloc[pos]`
        # ser.iloc[pos]
        if list.iloc[1] != '':
            serBglinks = list.iloc[1]
        elif list.iloc[0] != '':
            serBglinks = list.iloc[0]
        else:
            serBglinks = ''
      


        return serBglinks
    
    # Merge BG_LINKS
    dfBglinks = pd.DataFrame({'BGLINK_RBC' : dfPacketrbc['BG_LINKS'], 'BGLINK_BALISE' : dfPacketBalise['BG_LINKS'] })
    serBglinks = dfBglinks.apply(loc_MergeRBCBaliseData, axis=1)

    dfPacketrbc['BG_LINKS'] = serBglinks
    dfPacketBalise.drop(['BG_LINKS'], axis=1, inplace=True)

    # Merge SSP
    dfSSP = pd.DataFrame({'SSP_RBC' : dfPacketrbc['SSP'], 'SSP_BALISE' : dfPacketBalise['SSP'] })
    serSSP = dfSSP.apply(loc_MergeRBCBaliseData, axis=1)

    dfPacketrbc['SSP'] = serSSP
    dfPacketBalise.drop(['SSP'], axis=1, inplace=True)

    # Merge RBC_PACKET (simulated from balise)
    dfMA = pd.DataFrame({'PACKET_RBC' : dfPacketrbc['PACKET_RBC'], 'MA_BALISE' : dfPacketBalise['PACKET_RBC_SIMUL'] })
    serMA = dfMA.apply(loc_MergeRBCBaliseData, axis=1)

    dfPacketrbc['PACKET_RBC'] = serMA
    dfPacketBalise.drop(['PACKET_RBC_SIMUL'], axis=1, inplace=True)

    # Merge LX
    dfLX = pd.DataFrame({'LX_RBC' : dfPacketrbc['LX'], 'LX_BALISE' : dfPacketBalise['LX'] })
    serLX = dfLX.apply(loc_MergeRBCBaliseData, axis=1)

    dfPacketrbc['LX'] = serLX
    dfPacketBalise.drop(['LX'], axis=1, inplace=True)

    try:
        dfCompiledAlstom = pd.concat(listDf, axis=1)
    except Exception as ex:
        print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)
    
    # shift column 'BG_LINKS' to last position 
    #last_column = dfCompiledAlstom.pop('BG_LINKS') 
    #dfCompiledAlstom.insert(len(dfCompiledAlstom.columns), 'BG_LINKS', last_column) 


    return dfCompiledAlstom



# AlGetHeaderFromJru (dfInputJruMessages)
# Alstom decoder
# Input: dataframe with all JRU messages
# Input format: columns=['Record Id', 'JRU']
# Input Format: row = message id, JRU message in text file (one row by JRU register)
# Function: get header from JRU message and extract from here: date, time, distance (null),
# v_train, message id, message name and length
# Creates a data frame columns = ['Record Id', 'Date', 'Distance km', V_TRAIN (km/h),
#           'V_TRAIN (km/h)', 'MESSAGE ID', 'MESSAGE NAME', 'LENGTH', 'HEADER'
# listReturn = [dfHeader, dfDate, dfTime, dfDistance, dfSpeed, 
#                  dfJruNidMessage, dfJruMessageName, dfJruMessageLength]
def AlGetHeaderFromJru(dfInputJruMessages):

    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    listHeader = []
    dfHeader = pd.DataFrame()
    listDate = []
    dfDate = pd.DataFrame()
    listTime = []
    dfTime = pd.DataFrame()
    listDistance = []
    dfDistance = pd.DataFrame()
    listSpeed = []
    dfSpeed = pd.DataFrame()
    listJruNidMessage = []
    dfJruNidMessage = pd.DataFrame()
    listJruMessageName = []
    dfJruMessageName = pd.DataFrame()
    listJruMessageLength = []
    dfJruMessageLength = pd.DataFrame()

    length = dfInputJruMessages["JRU"].size
    for x in range(length):
        strJRU = dfInputJruMessages["JRU"][x]
        strAux = strJRU
        iAux = 0
        ixLine = 0
        strHeader = ""
        ixStartBlock = 0
        lengthBlock = 0
        startreading = False
        stopreading = False

        # ---------------- Get Header
        while True:
            try:   
                iAux = strAux.index('\n')
                line = strAux[0:iAux]
                strAux = strAux[iAux+1:len(strAux)]

                if "NID_MESSAGE(SUBSET) :" in line:
                    startreading = True
                    ixStartBlock = ixLine
                elif startreading:
                    if "M_MODE :" in line:
                        stopreading = True 
                        lengthBlock = ixLine + iAux # remove \n
                        break
                ixLine += (iAux + 1)
                        
            except:
                line = strAux[0:iAux]
                strAux = strAux[iAux+1:len(strAux)]
                if "M_MODE :" in line:
                        stopreading = True 
                        lengthBlock = ixLine + iAux + 1
                print(color.Fore.YELLOW + 
                  "AlGetHeaderFromJru-WARNING. Found JRU message with bad format in HEADER block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
                break

        if startreading and stopreading:
            strHeader = strJRU[ixStartBlock:ixStartBlock + lengthBlock].strip()
        else:
            print(color.Fore.RED + 
                  "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE(SUBSET) block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)

        # Insertamos bloque en la lista
        listHeader.insert(x, strHeader)

        # ------------------- Get Date
        strDate = ''
        strYear = ''
        strMonth = ''
        strDay = ''
        # Get year
        try:
            match = re.search(r'DATE.YEAR\s:', strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strYear = strHeader[start:end]
            match = re.search('\\d+', strYear)
            strYear = strYear[match.start():match.end()]
            if len(strYear) == 4:
                strYear = strYear[2:4]

        except:
            print(color.Fore.RED + 
                  "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in DATE.YEAR : block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)

        # Get month
        try:
            match = re.search("DATE.MONTH\\s:", strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strMonth = strHeader[start:end]
            match = re.search("\d+", strMonth)
            strMonth = strMonth[match.start():match.end()]
            
        except:
            print(color.Fore.RED + 
                  "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in DATE.MONTH : block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)  

        # Get day
        try:
            match = re.search("DATE.DAY\s:", strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strDay = strHeader[start:end]
            match = re.search("\d+", strDay)
            strDay = strDay[match.start():match.end()]
            
        except:
            print(color.Fore.RED + 
                  "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in DATE.DAY : block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)         

        # Insertamos valor en la lista
        strDate = strDay + '/' + strMonth + '/' + strYear
        listDate.insert(x, strDate)


        # ------------------- Get Time
        strHour = ''
        strMinutes = ''
        strSeconds = ''
        strMiliseconds = ''
        strTime = ''
        # Get hour
        try:
            match = re.search("TIME.HOUR\s:", strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strHour = strHeader[start:end]
            match = re.search("\d+", strHour)
            strHour = strHour[match.start():match.end()]

        except:
            print(color.Fore.RED + 
                  "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in TIME.HOUR : block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)

        # Get minutes
        try:
            match = re.search("TIME.MINUTES\s:", strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strMinutes = strHeader[start:end]
            match = re.search("\d+", strMinutes)
            strMinutes = strMinutes[match.start():match.end()]

        except:
            print(color.Fore.RED + 
                  "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in TIME.MINUTES : block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)

        # Get seconds
        try:
            match = re.search("TIME.SECONDS\s:", strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strSeconds = strHeader[start:end]
            match = re.search("\d+", strSeconds)
            strSeconds = strSeconds[match.start():match.end()]

        except:
            print(color.Fore.RED + 
                  "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in TIME.MINUTES : block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)

        # Get miliseconds
        try:
            match = re.search("TIME.MILLISECONDS\s:", strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strMiliseconds = strHeader[start:end]
            match = re.search("\d+", strMiliseconds)
            strMiliseconds = strMiliseconds[match.start():match.end()]

        except:
            print(color.Fore.RED + 
                  "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in TIME.MINUTES : block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)

        # Insertamos valor en la lista
        strTime = strHour + ':' + strMinutes + ':' + strSeconds + '.' + strMiliseconds
        listTime.insert(x, strTime)

        # ------------------- Get Distance
        # pending (not in JRU file)
        strDistance = ''
        
        # Insertamos valor en la lista
        listDistance.insert(x, strDistance)

        # ------------------- Get Speed
        strSpeed = ''
        
        # Get v_train
        try:
            match = re.search("V_TRAIN\s:", strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strSpeed = strHeader[start:end]
            match = re.search("\d+", strSpeed)
            strSpeed = strSpeed[match.start():match.end()]
            
    
        except:
            if (strSpeed != '') and "Unknown" in strSpeed:

                strSpeed = ''

            else:

                print(color.Fore.RED + 
                    "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in V_TRAIN : block. Message number: " + 
                    dfInputJruMessages["Record Id"][x] + 
                    color.Style.RESET_ALL)
                exit(0)

        # Insert value into list
        listSpeed.insert(x, strSpeed)
        
        # ------------------- Get JRU NID MESSAGE
        strJruNidMessage = ''
        # Get nid_message
        try:
            match = re.search("NID_MESSAGE\(SUBSET\)\s:", strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strJruNidMessage = strHeader[start:end]
            match = re.search("\d+", strJruNidMessage)
            strJruNidMessage = strJruNidMessage[match.start():match.end()]
    
        except:
            print(color.Fore.RED + 
                "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE : block. Message number: " + 
                dfInputJruMessages["Record Id"][x] + 
                color.Style.RESET_ALL)
            exit(0)

        # Insert value into list
        listJruNidMessage.insert(x, strJruNidMessage)

        # ------------------- Get JRU MESSAGE NAME
        strJruMessageName = ''
        intNidMessage = None
        try:
            intNidMessage = int(strJruNidMessage)
            if intNidMessage >= 50 and intNidMessage <= 254:
                strJruMessageName = 'SPARE'
            else:
                strJruMessageName = ss027.dfSS027['MESSAGE'][int(strJruNidMessage)]

        except:

            print(color.Fore.RED + 
                "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE : block. Message number: " + 
                dfInputJruMessages["Record Id"][x] + 
                color.Style.RESET_ALL)
            exit(0)

        # Insert value into list
        listJruMessageName.insert(x, strJruMessageName)

        # ------------------- Get JRU MESSAGE LENGTH
        strJruMessageLength = ''
        intMessageLength = None
        # Get L_MESSAGE
        try:
            match = re.search("L_MESSAGE\s:", strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strJruMessageLength = strHeader[start:end]
            match = re.search("\d+", strJruMessageLength)
            strJruMessageLength = strJruMessageLength[match.start():match.end()]
            intMessageLength = int(strJruMessageLength)
    
        except:
            print(color.Fore.RED + 
                "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in L_MESSAGE : block. Message number: " + 
                dfInputJruMessages["Record Id"][x] + 
                color.Style.RESET_ALL)
            exit(0)

        # Insert value into list
        listJruMessageLength.insert(x, intMessageLength)

        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfHeader = pd.DataFrame(np.array(listHeader), columns=['HEADER'])
    dfDate = pd.DataFrame(np.array(listDate), columns=['Date'])
    dfTime = pd.DataFrame(np.array(listTime), columns=['Time'])
    dfDistance = pd.DataFrame(np.array(listDistance), columns=['Distance (km)'])
    dfSpeed = pd.DataFrame(np.array(listSpeed), columns=['V_TRAIN (km/h)'])
    dfJruNidMessage = pd.DataFrame(np.array(listJruNidMessage), columns=['MESSAGE ID'])
    dfJruMessageName = pd.DataFrame(np.array(listJruMessageName), columns=['MESSAGE NAME'])
    dfJruMessageLength = pd.DataFrame(np.array(listJruMessageLength), columns=['LENGTH'])
    
    listReturn = [dfHeader, dfDate, dfTime, dfDistance, dfSpeed, 
                  dfJruNidMessage, dfJruMessageName, dfJruMessageLength]

    return listReturn


# AlgetDataFromJru (dfInputJruMessages)
# Alstom decoder
# Input: dataframe with all JRU messages
# Input format: columns=['Record Id', 'JRU']
# Input Format: row = message id, JRU message in text file (one row by JRU register)
# Function: get additional data (neither header nor packets) from JRU message
# return dfData = ['DATA', data value]


def AlgetDataFromJru (dfInputJruMessages):

    global glFunctionNum

    listData = []
    dfData = pd.DataFrame()

    length = dfInputJruMessages["JRU"].size
    for x in range(length):
        strJRU = dfInputJruMessages["JRU"][x]
        strData = ''
        strHeader = ''
        strNoHeader = ''
        strAux = strJRU
        iAux = 0
        ixLine = 0
        
        ixStartBlock = 0
        lengthBlock = 0
        startreading = False
        stopreading = False

        # ---------------- Split Header 
        while True:
            try:   
                iAux = strAux.index('\n')
                line = strAux[0:iAux]
                strAux = strAux[iAux+1:len(strAux)]

                if "NID_MESSAGE(SUBSET) :" in line:
                    startreading = True
                    ixStartBlock = ixLine
                elif startreading:
                    if "M_MODE :" in line:
                        stopreading = True 
                        lengthBlock = ixLine + iAux # remove \n
                        break
                ixLine += (iAux + 1)
                        
            except:
                line = strAux[0:iAux]
                strAux = strAux[iAux+1:len(strAux)]
                if "M_MODE :" in line:
                        stopreading = True 
                        lengthBlock = ixLine + iAux + 1
                print(color.Fore.YELLOW + 
                  "AlGetDataFromJru-WARNING. Found JRU message with bad format in HEADER block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
                break

        if startreading and stopreading:
            strHeader = strJRU[ixStartBlock:ixStartBlock + lengthBlock].strip() 
            strNoHeader = strJRU[ixStartBlock + lengthBlock:len(strJRU)].strip() + '\n' # !importat to add '\n' at the end
        else:
            print(color.Fore.RED + 
                  "AlGetHeaderFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE(SUBSET) block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)
            
        # ------------------- Get JRU NID MESSAGE
        strJruNidMessage = ''
        intNidMessage = None
        # Get nid_message
        try:
            match = re.search("NID_MESSAGE\(SUBSET\)\s:", strHeader)
            start = match.end()
            match = re.search("\n", strHeader[start:len(strHeader)])
            end = start + match.start()
            strJruNidMessage = strHeader[start:end]
            match = re.search("\d+", strJruNidMessage)
            strJruNidMessage = strJruNidMessage[match.start():match.end()]
    
        except:
            print(color.Fore.RED + 
                "AlGetDataFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE : block. Message number: " + 
                dfInputJruMessages["Record Id"][x] + 
                color.Style.RESET_ALL)
            exit(0)

        
        try:

            intNidMessage = int(strJruNidMessage)

        except:

            print(color.Fore.RED + 
                "AlGetDataFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE : block. Message number: " + 
                dfInputJruMessages["Record Id"][x] + 
                color.Style.RESET_ALL)
            exit(0)

        # Get Data depending of JRU NID_MESSAGE

        strAux = strNoHeader
        iAux = 0
        ixLine = 0
        ixStartBlock = 0
        lengthBlock = 0
        reading = False

        # All messages different of train, rbc or balise messages
        if intNidMessage == 6: # Telegram from balise

            # ---------------- Get data content
            strAux = strNoHeader
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'EUROBALISEHEADER (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strData = "\n".join(block)
                            reading = False
                            break
                        block.append(line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    strData = "\n".join(block)
                    print(color.Fore.YELLOW + 
                        "AlGetDataFromJru-WARNING. Found JRU message with bad format in EUROBALISE HEADER block. Message number: " + 
                        dfInputJruMessages["Record Id"][x] + 
                        color.Style.RESET_ALL)
                    break


        elif intNidMessage == 9: # Message from RBC

            # ---------------- Get data content
            strAux = strNoHeader
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'RBC_MESSAGETK2T (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strData = "\n".join(block)
                            reading = False
                            break
                        block.append(line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    strData = "\n".join(block)
                    print(color.Fore.YELLOW + 
                        "AlGetDataFromJru-WARNING. Found JRU message with bad format in RBC_MESSAGETK2T block. Message number: " + 
                        dfInputJruMessages["Record Id"][x] + 
                        color.Style.RESET_ALL)
                    break


        elif intNidMessage == 10: # Message to RBC

            # ---------------- Get data content
            strAux = strNoHeader
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'RBC_MESSAGET2TK (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strData = "\n".join(block)
                            reading = False
                            break
                        block.append(line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    strData = "\n".join(block)
                    print(color.Fore.YELLOW + 
                        "AlGetDataFromJru-WARNING. Found JRU message with bad format in RBC_MESSAGET2TK block. Message number: " + 
                        dfInputJruMessages["Record Id"][x] + 
                        color.Style.RESET_ALL)
                    break




        else: 

            # ---------------- Get data content
            strAux = strNoHeader
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            iAux = strAux.index('\n')
            firstLine = strAux[0:iAux].strip()
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    strAux = strAux[iAux+1:len(strAux)]

                    if firstLine in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strData = firstLine + '\n' + "\n".join(block) + '\n)'
                            reading = False
                            break
                        block.append(line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    strData = firstLine + '\n' + "\n".join(block) + '\n)'
                    print(color.Fore.YELLOW + 
                        "AlGetDataFromJru-WARNING. Found JRU message with bad format in DATA block. Message number: " + 
                        dfInputJruMessages["Record Id"][x] + 
                        color.Style.RESET_ALL)
                    break


        # Insertamos valor en la lista
        listData.insert(x, strData)
            
        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfData = pd.DataFrame(np.array(listData), columns=['DATA']) 
    return dfData

# AlgetDataFromJru (dfInputJruMessages)
# Alstom decoder
# Input: dataframe with all JRU messages
# Input format: columns=['Record Id', 'JRU']
# Input Format: row = message id, JRU message in text file (one row by JRU register)
# Function: get RBC packets (FROM RBC & TO RBC) from JRU messages
# return dfPacketRBC = ['RBC PACKET', packets

def AlgetRbcPacketsFromJru (dfInputJruMessages):

    global glFunctionNum

    listRbcPackets = []
    dfRbcPacket = pd.DataFrame()

    length = dfInputJruMessages["JRU"].size
    for x in range(length):
        strJRU = dfInputJruMessages["JRU"][x]
        strJruHeader = ''
        strJruNoheader = ''
        strMessageFromRbc = ''
        strMessageFromRbcHeader = ''
        strMessageToRbc = ''
        strMessageToRbcHeader = ''
        strAux = strJRU
        iAux = 0
        ixLine = 0
        
        ixStartBlock = 0
        lengthBlock = 0
        startreading = False
        stopreading = False

        # ---------------- Split JRU Header
        while True:
            try:   
                iAux = strAux.index('\n')
                line = strAux[0:iAux]
                strAux = strAux[iAux+1:len(strAux)]

                if "NID_MESSAGE(SUBSET) :" in line:
                    startreading = True
                    ixStartBlock = ixLine
                elif startreading:
                    if "M_MODE :" in line:
                        stopreading = True 
                        lengthBlock = ixLine + iAux # remove \n
                        break
                ixLine += (iAux + 1)
                        
            except:
                line = strAux[0:iAux]
                strAux = strAux[iAux+1:len(strAux)]
                if "M_MODE :" in line:
                        stopreading = True 
                        lengthBlock = ixLine + iAux + 1
                print(color.Fore.YELLOW + 
                  "AlgetRbcPacketsFromJru-WARNING. Found JRU message with bad format in HEADER block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
                break

        if startreading and stopreading:
            strJruHeader = strJRU[ixStartBlock:ixStartBlock + lengthBlock].strip() 
            strJruNoheader = strJRU[ixStartBlock + lengthBlock:len(strJRU)].strip() + '\n' # !importat to add '\n' at the end
        else:
            print(color.Fore.RED + 
                  "AlgetRbcPacketsFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE(SUBSET) block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)
            
        # ------------------- Get JRU NID MESSAGE
        strJruNidMessage = ''
        intNidMessage = None
        # Get nid_message
        try:
            match = re.search("NID_MESSAGE\(SUBSET\)\s:", strJruHeader)
            start = match.end()
            match = re.search("\n", strJruHeader[start:len(strJruHeader)])
            end = start + match.start()
            strJruNidMessage = strJruHeader[start:end]
            match = re.search("\d+", strJruNidMessage)
            strJruNidMessage = strJruNidMessage[match.start():match.end()]
    
        except:
            print(color.Fore.RED + 
                "AlgetRbcPacketsFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE : block. Message number: " + 
                dfInputJruMessages["Record Id"][x] + 
                color.Style.RESET_ALL)
            exit(0)

        
        try:

            intNidMessage = int(strJruNidMessage)

        except:

            print(color.Fore.RED + 
                "AlgetRbcPacketsFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE : block. Message number: " + 
                dfInputJruMessages["Record Id"][x] + 
                color.Style.RESET_ALL)
            exit(0)

        # Get Data depending of JRU NID_MESSAGE
        if intNidMessage == 9: # Message from RBC

            strMessageFromRbc = strJruNoheader

            # Format:
            # MESSAGE FROM RBC (RBC BaseLine 3 Maintenance Release 1) (
            #    RBC_MESSAGETK2T (
            #        NID_C : 74-->74
            #        NID_RBC : 9001-->9001
            #        NID_MESSAGE : 3-->Movement Authority
            #        L_MESSAGE : 142
            #        T_TRAIN : 135370
            #        M_ACK : 1-->Acknowledgement required
            #        NID_LRBG1 : 74
            #        NID_LRBG2 : 9369
            #    )
            #    MSG_TK2T-3 (
            #    NID_PACKET : 15 -->Level 2/3 Movement Authority
            #    PKT_TK2T_15 (
            #        Q_DIR : 1 -->Nominal
            #        L_PACKET : 135
            #        Q_SCALE : 1 -->1 m scale
            #        V_EMA : 0 -->0km/h
            #        T_EMA : 1023 -->infinite
            #        niter : 0
            #        L_ENDSECTION : 1650 -->1650 m
            #        Q_SECTIONTIMER : 0 -->No section Timer Information
            #        Q_ENDTIMER : 0 -->No End section timer information
            #        Q_DANGERPOINT : 1 -->Danger point information to follow
            #        D_DP : 22 -->22 m
            #        V_RELEASEDP : 0 -->0 km/h
            #        Q_OVERLAP : 1 -->Overlap information to follow
            #        D_STARTOL : 700 -->700 m
            #        T_OL : 1023 -->infinite
            #        D_OL : 61 -->61 m
            #        V_RELEASEOL : 3 -->15 km/h
            #    )
            #    OPTIONAL_PKT_TK2T (
            #        NID_PACKET : 21 -->Gradient Profile
            #        PKT_TK2T_21 (
            #            Q_DIR : 1 -->Nominal
            #            L_PACKET : 126
            #                   .....
            #        ) packet 21
            #        NID_PACKET : 27 -->International Static Speed Profile
            #        PKT_TK2T_27 (
            #            Q_DIR : 1 -->Nominal
            #            L_PACKET : 114
            #               ....
            #        ) packet 27
            #           ... other packets  ..
            #    ) optional packets
            #    ) message 3
            # ) message from rbc
            #

            # Split header from message from Rbc
            strAux = strMessageFromRbc
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'RBC_MESSAGETK2T (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strMessageFromRbcHeader = "\n".join(block)
                            reading = False
                            break
                        block.append('    ' + line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    ixLine += len(strAux)
                    strMessageFromRbcHeader = "\n".join(block)
                    print(color.Fore.YELLOW + 
                        "AlGetPacketsFromJru-WARNING. Found JRU message with bad format in RBC_MESSAGETK2T HEADER block. Message number: " + 
                        dfInputJruMessages["Record Id"][x] + 
                        color.Style.RESET_ALL)
                    break

            # Get RBC NID_MESSAGE from rbc header
            intRbcNidMessage = None
            strRbcNidMessage = ''
            # Get nid_message
            try:
                match = re.search("NID_MESSAGE\s:", strMessageFromRbcHeader)
                start = match.end()
                match = re.search("\n", strMessageFromRbcHeader[start:len(strMessageFromRbcHeader)])
                end = start + match.start()
                strRbcNidMessage = strMessageFromRbcHeader[start:end]
                match = re.search("\d+", strRbcNidMessage)
                strRbcNidMessage = strRbcNidMessage[match.start():match.end()]
        
            except:
                print(color.Fore.RED + 
                    "AlgetRbcPacketsFromJru-ERROR. Found message from RBC with bad format in NID_MESSAGE : block. Message number: " + 
                    dfInputJruMessages["Record Id"][x] + 
                    color.Style.RESET_ALL)
                exit(0)

            
            try:

                intRbcNidMessage = int(strRbcNidMessage)

            except:

                print(color.Fore.RED + 
                    "AlgetRbcPacketsFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE : block. Message number: " + 
                    dfInputJruMessages["Record Id"][x] + 
                    color.Style.RESET_ALL)
                exit(0)


            # Get MSG_TK2T-xx ( from message from RBC
            strAux = strMessageFromRbc
            strMSG_TK2T = ''
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'MSG_TK2T-' + strRbcNidMessage + ' (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strMSG_TK2T = "\n".join(block)
                            reading = False
                            break
                        block.append(line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    ixLine += len(strAux)
                    strMSG_TK2T = "\n".join(block)
                    print(color.Fore.YELLOW + 
                        "AlGetPacketsFromJru-WARNING. Found RBC message with bad format in MSG_TK2T-xx ( block. Message number: " + 
                        dfInputJruMessages["Record Id"][x] + 
                        color.Style.RESET_ALL)
                    break
            
            # Get rbcMessageHeader strMSG_Tk2THeader
            strAux = strMSG_TK2T
            strMSG_Tk2THeader = ''
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'NID_PACKET :' in line or 'OPTIONAL_PKT_TK2T (' in line:
                        strMSG_Tk2THeader = "\n".join(block)
                        break
                    
                    block.append(line)
                    ixLine += (iAux + 1)
                    

                except:

                    strMSG_Tk2THeader = ''
                    break
            

            # ---------------- Get packets from message from RBC
            listPacketsInRbcMessage = []
            numPackets = 0   
            strPacketRbc = ''            
            strData = strMSG_TK2T
            boolMorePackets = True
            while boolMorePackets:

                intNidRbcPacket = None
                ixNextData = 0
                strNidRbcPacketLine = ''
                strNidRbcPacket = ''
                

                # Get RbcNidPacket
                try:
                    match = re.search("NID_PACKET\s:", strData)
                    start = match.end()
                    match = re.search("\n", strData[start:len(strData)])
                    end = start + match.start()
                    ixNextData = end
                    strNidRbcPacketLine = strData[start:end]
                    match = re.search("\d+", strNidRbcPacketLine)
                    strNidRbcPacket = strNidRbcPacketLine[match.start():match.end()]
                    strData = strData[ixNextData:len(strData)]

                    try:

                        intNidRbcPacket = int(strNidRbcPacket)

                    except:

                        print(color.Fore.RED + 
                            "AlgetRbcPacketsFromJru-ERROR. Found RBC message with bad format in NID_PACKET : block. Message number: " + 
                            dfInputJruMessages["Record Id"][x] + 
                            color.Style.RESET_ALL)
                        exit(0)
                    
                    # Get packet content : PKT_TK2T_xx
                    strAux = strData
                    strPKT_TK2T = ''
                    iAux = 0
                    ixLine = 0
                    ixStartBlock = 0
                    lengthBlock = 0
                    reading = False
                
                    block = []
                    while True:

                        try:   
                            iAux = strAux.index('\n')
                            line = strAux[0:iAux]
                            line = line.rstrip("\n")
                            line = line.lstrip()
                            strAux = strAux[iAux+1:len(strAux)]

                            if 'PKT_TK2T_' + strNidRbcPacket + ' (' in line:
                                reading = True
                                level = 1
                                
                            elif reading:

                                if ":" not in line:
                                    level += line.count("(")
                                    level -= line.count(")")
                                #level += line.count("(")
                                #level -= line.count(")")
                                if level == 0:
                                    strPKT_TK2T = "\n".join(block)
                                    reading = False
                                    break
                                block.append('    ' + line)
                                
                            ixLine += (iAux + 1)
                            

                        except:
                            line = strAux
                            ixLine += len(strAux)
                            strPKT_TK2T = "\n".join(block)
                            print(color.Fore.YELLOW + 
                                "AlGetPacketsFromJru-WARNING. Found RBC message with bad format in PKT_TK2Txx ( block. Message number: " + 
                                dfInputJruMessages["Record Id"][x] + 
                                color.Style.RESET_ALL)
                            break

                    strPacketRbc = 'PACKET: ' + strNidRbcPacket 
                    strPacketRbc += '\n' + strPKT_TK2T

                    listPacketsInRbcMessage.insert(numPackets, strPacketRbc)
                    numPackets += 1
                
                except:
                    boolMorePackets = False
        
            strAux = 'MESSAGE RBC: ' + strRbcNidMessage +'\n'
            if strMSG_Tk2THeader != '':
                strAux += strMSG_Tk2THeader + '\n'
            strAux += "\n".join(listPacketsInRbcMessage)
            listRbcPackets.append(strAux)
        
        elif intNidMessage == 10: # Message to RBC
            
            strMessageToRbc = strJruNoheader

            # Format:
            # MESSAGS TO RBC (RBC BaseLine 3 Maintenance Release 1) (
            # RBC_MESSAGET2TK (
            #    NID_C : 74-->74
            #    NID_RBC : 9001-->9001
            #    NID_MESSAGE : 136-->Train Position Report
            #    L_MESSAGE : 25
            #    T_TRAIN : 174993
            #    NID_ENGINE : 33586
            #)
            # MSG_T2TK-136 (
            #    PKT_T2TK_0OR1 (
            #        NID_PACKET : 0 -->Position Report
            #        PKT_T2TK_0 (
            #            L_PACKET : 122
            #            Q_SCALE : 0 -->10 cm scale
            #            NID_LRBG : 1221807 -->1221807
            #            D_LRBG : 1462 -->146,2 m
            #            Q_DIRLRBG : 1 -->Nominal
            #            Q_DLRBG : 0 -->Reverse
            #            L_DOUBTOVER : 139 -->13,9 m
            #            L_DOUBTUNDER : 151 -->15,1 m
            #            Q_LENGTH : 0 -->No train integrity information available
            #            V_TRAIN : 0 -->0 km/h
            #            Q_DIRTRAIN : 2 -->Unknown
            #            M_MODE : 6 -->Stand By
            #            M_LEVEL : 1 -->Level NTC specified by NID_NTC
            #            NID_NTC : 6 -->6
            #          )
            #       )
            #    OPTIONAL_PKT_T2TK (
            #    )
            #   )
            # )

            # Split header from message to Rbc
            strAux = strMessageToRbc
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'RBC_MESSAGET2TK (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strMessageToRbcHeader = "\n".join(block)
                            reading = False
                            break
                        block.append('    ' + line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    ixLine += len(strAux)
                    strMessageToRbcHeader = "\n".join(block)
                    print(color.Fore.YELLOW + 
                        "AlGetPacketsFromJru-WARNING. Found JRU message with bad format in RBC_MESSAGET2TK HEADER block. Message number: " + 
                        dfInputJruMessages["Record Id"][x] + 
                        color.Style.RESET_ALL)
                    break

            # Get RBC NID_MESSAGE from rbc header
            intRbcNidMessage = None
            strRbcNidMessage = ''
            # Get nid_message
            try:
                match = re.search("NID_MESSAGE\s:", strMessageToRbcHeader)
                start = match.end()
                match = re.search("\n", strMessageToRbcHeader[start:len(strMessageToRbcHeader)])
                end = start + match.start()
                strRbcNidMessage = strMessageToRbcHeader[start:end]
                match = re.search("\d+", strRbcNidMessage)
                strRbcNidMessage = strRbcNidMessage[match.start():match.end()]
        
            except:
                print(color.Fore.RED + 
                    "AlgetRbcPacketsFromJru-ERROR. Found message from RBC with bad format in NID_MESSAGE : block. Message number: " + 
                    dfInputJruMessages["Record Id"][x] + 
                    color.Style.RESET_ALL)
                exit(0)

            
            try:

                intRbcNidMessage = int(strRbcNidMessage)

            except:

                print(color.Fore.RED + 
                    "AlgetRbcPacketsFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE : block. Message number: " + 
                    dfInputJruMessages["Record Id"][x] + 
                    color.Style.RESET_ALL)
                exit(0)


            # Get MSG_T2TK-xx ( from message from RBC
            strAux = strMessageToRbc
            strMSG_T2TK = ''
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'MSG_T2TK-' + strRbcNidMessage + ' (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strMSG_T2TK = "\n".join(block) + '\n'
                            reading = False
                            break
                        block.append(line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    ixLine += len(strAux)
                    strMSG_T2TK = "\n".join(block) + '\n'
                    print(color.Fore.YELLOW + 
                        "AlGetPacketsFromJru-WARNING. Found RBC message with bad format in MSG_T2TK-xx ( block. Message number: " + 
                        dfInputJruMessages["Record Id"][x] + 
                        color.Style.RESET_ALL)
                    break

            # In case block PKT_T2TK_0OR1 () exists, then remove this label
            strAux = strMSG_T2TK
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []

            match = re.search("PKT_T2TK_0OR1\s\(", strAux)
            if match != None:
            
                while True:

                    try:   
                        iAux = strAux.index('\n')
                        line = strAux[0:iAux]
                        line = line.rstrip("\n")
                        line = line.lstrip()
                        strAux = strAux[iAux+1:len(strAux)]

                        if 'PKT_T2TK_0OR1 (' in line:
                            reading = True
                            level = 1
                            
                        elif reading:

                            if ":" not in line:
                                level += line.count("(")
                                level -= line.count(")")
                            #level += line.count("(")
                            #level -= line.count(")")
                            if level == 0:
                                strMSG_T2TK = "\n".join(block) + '\n'
                                reading = False
                                break
                            block.append(line)
                            
                        ixLine += (iAux + 1)
                        

                    except:
                        line = strAux
                        ixLine += len(strAux)
                        strMSG_T2TK = "\n".join(block) + '\n'
                        print(color.Fore.YELLOW + 
                            "AlGetPacketsFromJru-WARNING. Found RBC message with bad format in PKT_T2TK_0OR1 ( block. Message number: " + 
                            dfInputJruMessages["Record Id"][x] + 
                            color.Style.RESET_ALL)
                        break
            
            # Get rbcMessageHeader strMSG_T2TKHeader
            strAux = strMSG_T2TK
            strMSG_T2TKHeader = ''
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'NID_PACKET :' in line or 'OPTIONAL_PKT_T2TK (' in line:
                        strMSG_T2TKHeader = "\n".join(block)
                        break
                    
                    block.append(line)
                    ixLine += (iAux + 1)
                    

                except:

                    strMSG_T2TKHeader = ''
                    break
            

            # ---------------- Get packets from message to RBC
            listPacketsInRbcMessage = []
            numPackets = 0   
            strPacketRbc = ''            
            strData = strMSG_T2TK
            boolMorePackets = True
            while boolMorePackets:

                intNidRbcPacket = None
                ixNextData = 0
                strNidRbcPacketLine = ''
                strNidRbcPacket = ''
                

                # Get RbcNidPacket
                try:
                    match = re.search("NID_PACKET\s:", strData)
                    start = match.end()
                    match = re.search("\n", strData[start:len(strData)])
                    end = start + match.start()
                    ixNextData = end
                    strNidRbcPacketLine = strData[start:end]
                    match = re.search("\d+", strNidRbcPacketLine)
                    strNidRbcPacket = strNidRbcPacketLine[match.start():match.end()]
                    strData = strData[ixNextData:len(strData)]

                    try:

                        intNidRbcPacket = int(strNidRbcPacket)

                    except:

                        print(color.Fore.RED + 
                            "AlgetRbcPacketsFromJru-ERROR. Found RBC message with bad format in NID_PACKET : block. Message number: " + 
                            dfInputJruMessages["Record Id"][x] + 
                            color.Style.RESET_ALL)
                        exit(0)
                    
                    # Get packet content : PKT_T2TK_xx
                    strAux = strData
                    strPKT_T2TK = ''
                    iAux = 0
                    ixLine = 0
                    ixStartBlock = 0
                    lengthBlock = 0
                    reading = False
                
                    block = []
                    while True:

                        try:   
                            iAux = strAux.index('\n')
                            line = strAux[0:iAux]
                            line = line.rstrip("\n")
                            line = line.lstrip()
                            strAux = strAux[iAux+1:len(strAux)]

                            if dfInputJruMessages["Record Id"][x] == '148':
                                pepe = 0

                            if ('PKT_T2TK-' + strNidRbcPacket + ' (' in line or 'PKT_T2TK_' + strNidRbcPacket + ' (' in line):
                                
                                reading = True
                                level = 1
                                
                            elif reading:

                                if ":" not in line:
                                    level += line.count("(")
                                    level -= line.count(")")
                                #level += line.count("(")
                                #level -= line.count(")")
                                if level == 0:
                                    strPKT_T2TK = "\n".join(block)
                                    reading = False
                                    break
                                block.append('    ' + line)
                                
                            ixLine += (iAux + 1)
                            

                        except:
                            line = strAux
                            ixLine += len(strAux)
                            strPKT_T2TK = "\n".join(block)
                            print(color.Fore.YELLOW + 
                                "AlGetPacketsFromJru-WARNING. Found RBC message with bad format in PKT_T2TKxx ( block. Message number: " + 
                                dfInputJruMessages["Record Id"][x] + 
                                color.Style.RESET_ALL)
                            break

                    strPacketRbc = 'PACKET: ' + strNidRbcPacket 
                    strPacketRbc += '\n' + strPKT_T2TK

                    listPacketsInRbcMessage.insert(numPackets, strPacketRbc)
                    numPackets += 1
                
                except:
                    boolMorePackets = False
        
            strAux = 'MESSAGE RBC: ' + strRbcNidMessage +'\n'
            if strMSG_T2TKHeader != '':
                strAux += strMSG_T2TKHeader + '\n'
            strAux += "\n".join(listPacketsInRbcMessage)
            listRbcPackets.append(strAux)


        else:
            listRbcPackets.append('')

        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

            
    dfRbcPacket = pd.DataFrame(np.array(listRbcPackets), columns=['PACKET RBC']) 

    return dfRbcPacket



# AlgetBalisePacketsFromJru (dfInputJruMessages)
# Alstom decoder
# Input: dataframe with all JRU messages
# Input format: columns=['Record Id', 'JRU']
# Input Format: row = message id, JRU message in text file (one row by JRU register)
# Function: get RBC packets (FROM RBC & TO RBC) from JRU messages
# return dfPacketRBC = ['RBC PACKET', packets

def AlgetBalisePacketsFromJru (dfInputJruMessages):

    global glFunctionNum

    listBalisePackets = []
    dfBalisePacket = pd.DataFrame()

    length = dfInputJruMessages["JRU"].size
    for x in range(length):
        strJRU = dfInputJruMessages["JRU"][x]
        strJruHeader = ''
        strJruNoheader = ''
        strTelegramFromBalise = ''
        strTelegramFromBaliseHeader = ''
        strAux = strJRU
        iAux = 0
        ixLine = 0
        
        ixStartBlock = 0
        lengthBlock = 0
        startreading = False
        stopreading = False

        # ---------------- Split JRU Header
        while True:
            try:   
                iAux = strAux.index('\n')
                line = strAux[0:iAux]
                strAux = strAux[iAux+1:len(strAux)]

                if "NID_MESSAGE(SUBSET) :" in line:
                    startreading = True
                    ixStartBlock = ixLine
                elif startreading:
                    if "M_MODE :" in line:
                        stopreading = True 
                        lengthBlock = ixLine + iAux # remove \n
                        break
                ixLine += (iAux + 1)
                        
            except:
                line = strAux[0:iAux]
                strAux = strAux[iAux+1:len(strAux)]
                if "M_MODE :" in line:
                        stopreading = True 
                        lengthBlock = ixLine + iAux + 1
                print(color.Fore.YELLOW + 
                  "AlgetBalisePacketsFromJru-WARNING. Found JRU message with bad format in HEADER block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
                break

        if startreading and stopreading:
            strJruHeader = strJRU[ixStartBlock:ixStartBlock + lengthBlock].strip() 
            strJruNoheader = strJRU[ixStartBlock + lengthBlock:len(strJRU)].strip() + '\n' # !importat to add '\n' at the end
        else:
            print(color.Fore.RED + 
                  "AlgetBalisePacketsFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE(SUBSET) block. Message number: " + 
                  dfInputJruMessages["Record Id"][x] + 
                  color.Style.RESET_ALL)
            exit(0)
            
        # ------------------- Get JRU NID MESSAGE
        strJruNidMessage = ''
        intNidMessage = None
        # Get nid_message
        try:
            match = re.search("NID_MESSAGE\(SUBSET\)\s:", strJruHeader)
            start = match.end()
            match = re.search("\n", strJruHeader[start:len(strJruHeader)])
            end = start + match.start()
            strJruNidMessage = strJruHeader[start:end]
            match = re.search("\d+", strJruNidMessage)
            strJruNidMessage = strJruNidMessage[match.start():match.end()]
    
        except:
            print(color.Fore.RED + 
                "AlgetBalisePacketsFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE : block. Message number: " + 
                dfInputJruMessages["Record Id"][x] + 
                color.Style.RESET_ALL)
            exit(0)

        
        try:

            intNidMessage = int(strJruNidMessage)

        except:

            print(color.Fore.RED + 
                "AlgetBalisePacketsFromJru-ERROR. Found JRU message with bad format in NID_MESSAGE : block. Message number: " + 
                dfInputJruMessages["Record Id"][x] + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 6: # Telegram from balise

            strTelegramFromBalise = strJruNoheader

            # Format:
            #TELEGRAM FROM BALISE (
            #    EUROBALISEHEADER (
            #        Q_UPDOWN : 1-->Up link telegram
            #        M_VERSION : 32-->Version 2.0, introduced in SRS 3.3.0
            #        Q_MEDIA : 0-->Balise
            #        N_PIG : 0-->I am the 1st
            #        N_TOTAL : 0-->1 balise in the group
            #        M_DUP : 0-->No duplicates
            #        M_MCOUNT : 255-->The telegram fits with all telegrams of the same balise group)
            #        NID_C : 74-->74
            #        NID_BG : 9391-->9391
            #        Q_LINK : 1-->Linked
            #    )
            #    NID_PACKET : 66 -->Temporary Speed Restriction Revocation
            #    PKT_TK2T_66 (
            #        Q_DIR : 1 -->Nominal
            #        L_PACKET : 31
            #        NID_TSR : 126 -->Reserved for non RBC transmission (balise, loop or radio infill)
            #    )
            #    NID_PACKET : 66 -->Temporary Speed Restriction Revocation
            #    PKT_TK2T_66 (
            #        Q_DIR : 0 -->Reverse
            #        L_PACKET : 31
            #        NID_TSR : 126 -->Reserved for non RBC transmission (balise, loop or radio infill)
            #    )
            #    NID_PACKET : 255 -->End of information
            #    PKT_TK2T_255 (
            #    )
            #)

            # Split header from telegram from balise
            strAux = strTelegramFromBalise
            iAux = 0
            ixLine = 0
            ixStartBlock = 0
            lengthBlock = 0
            reading = False
        
            block = []
            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'TELEGRAM FROM BALISE (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strTelegramFromBaliseHeader = "\n".join(block) + '\n'
                            reading = False
                            break
                        block.append('    ' + line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    ixLine += len(strAux)
                    strTelegramFromBaliseHeader = "\n".join(block) + '\n'
                    print(color.Fore.YELLOW + 
                        "AlgetBalisePacketsFromJru-WARNING. Found JRU message with bad format in TELEGRAM FROM BALISE HEADER block. Message number: " + 
                        dfInputJruMessages["Record Id"][x] + 
                        color.Style.RESET_ALL)
                    break
            
            # ---------------- Get packets from message from RBC
            listPacketsInTelegram = []
            numPackets = 0   
            strPacketRbc = ''            
            strData = strTelegramFromBaliseHeader
            boolMorePackets = True
            while boolMorePackets:

                ixNextData = 0
                intNidTelegramPacket = None
                strNidTelegramPacketLine = ''
                strNidTelegramPacket = ''
                

                # Get RbcNidPacket
                try:
                    match = re.search("NID_PACKET\s:", strData)
                    start = match.end()
                    match = re.search("\n", strData[start:len(strData)])
                    end = start + match.start()
                    ixNextData = end
                    strNidTelegramPacketLine = strData[start:end]
                    match = re.search("\d+", strNidTelegramPacketLine)
                    strNidTelegramPacket = strNidTelegramPacketLine[match.start():match.end()]
                    strData = strData[ixNextData:len(strData)]

                    try:

                        intNidTelegramPacket = int(strNidTelegramPacket)
                        # Only used to detect exceptions in case not integer

                    except:

                        print(color.Fore.RED + 
                            "AlgetBalisePacketsFromJru-ERROR. Found RBC message with bad format in NID_PACKET : block. Message number: " + 
                            dfInputJruMessages["Record Id"][x] + 
                            color.Style.RESET_ALL)
                        exit(0)
                    
                    # Get packet content : PKT_TK2T_xx
                    strAux = strData
                    strPKT_TK2T = ''
                    iAux = 0
                    ixLine = 0
                    ixStartBlock = 0
                    lengthBlock = 0
                    reading = False
                
                    block = []
                    while True:

                        try:   
                            iAux = strAux.index('\n')
                            line = strAux[0:iAux]
                            line = line.rstrip("\n")
                            line = line.lstrip()
                            strAux = strAux[iAux+1:len(strAux)]

                            if 'PKT_TK2T_' + strNidTelegramPacket + ' (' in line:
                                reading = True
                                level = 1
                                
                            elif reading:

                                if ":" not in line:
                                    level += line.count("(")
                                    level -= line.count(")")
                                #level += line.count("(")
                                #level -= line.count(")")
                                if level == 0:
                                    strPKT_TK2T = "\n".join(block)
                                    reading = False
                                    break
                                block.append('    ' + line)
                                
                            ixLine += (iAux + 1)
                            

                        except:
                            line = strAux
                            ixLine += len(strAux)
                            strPKT_TK2T = "\n".join(block)
                            print(color.Fore.YELLOW + 
                                "AlgetBalisePacketsFromJru-WARNING. Found telegram with bad format in PKT_TK2Txx ( block. Message number: " + 
                                dfInputJruMessages["Record Id"][x] + 
                                color.Style.RESET_ALL)
                            print("NID_PACKET: " + strNidTelegramPacketLine) 
                            break

                    strPacketRbc = 'PACKET: ' + strNidTelegramPacket 
                    strPacketRbc += '\n' + strPKT_TK2T

                    listPacketsInTelegram.insert(numPackets, strPacketRbc)
                    numPackets += 1
                
                except:
                    boolMorePackets = False
        
            strAux += "\n".join(listPacketsInTelegram)
            listBalisePackets.append(strAux)


        else:
            listBalisePackets.append('')

        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

            
    dfBalisePacket = pd.DataFrame(np.array(listBalisePackets), columns=['PACKET']) 

    return dfBalisePacket




# AlApplyRowFilter(alDf, listRbcFilter, listJruFilter)
# Alstom decoder
# Apply filter for JRU messages and RBC packets returns a re-indexed df with the filtered file
# Accept two filters (list): PACKET_RBC ID & JRUMESSAGE ID

def AlApplyRowFilter(alDf, listRbcFilter, listJruFilter, listJruDiscard):

    # -----------  WARNING --------------------------------------------------
    # To review; not working properly when listRbCFilter is not an empty array
    # Use empty array fort Rbc Filter in main.py until correction
    # -----------------------------------------------------------------------

    inputdf = pd.DataFrame(alDf) # copia local del objeto del argumento
    listRbcPackets = np.array(listRbcFilter) # copia local del filtro de mensajes RBC
    listJruMessages = np.array(listJruFilter) # copia local del filtro de mensajes JRU
    listJruDiscardMessages = np.array(listJruDiscard)
    
    # filtramos df por tipos de mensjaes solicitados
    boolList = []
    serFilterJruMessages = pd.Series() # Filter for messages in list listJruFilter
    serFilterJruDiscardedMess = pd.Series()
    serFilterRbcMessages = pd.Series() # Filter for rbc only-Rbc messages contained in ListRbcFilter
    serFilterNonRbcMessages = pd.Series() # Filter for messages in listJruFilter but non-RBC messages
    serFilterRbcPackets = pd.Series() # Fil
    serFinalFilter = pd.Series() # Final filter compound from above ones

    # boolList
    for x in range (len(inputdf.axes[0])):
        boolList.insert(x, False)
    
    

    # Filter for required JRU Messages
    if len(listJruMessages) > 0:
        # Filtramos mensajes JRU
        for x in range(len(listJruMessages)):
            serFilterJruMessages[x] = inputdf["MESSAGE NAME"].str.contains(listJruMessages[x], na=False, regex=False)
            if x > 0:
                # Sum of filters
                serFilterJruMessages[0] |= serFilterJruMessages[x]

        # Filtering
        inputdf = inputdf[serFilterJruMessages[0]]
        inputdf.index = np.arange(0, len(inputdf.axes[0]))

    # Filter for discarded JRU Messages
    if len(listJruDiscardMessages) > 0:
        # Filtramos mensajes JRU
        for x in range(len(listJruDiscardMessages)):
            serFilterJruDiscardedMess[x] = inputdf["MESSAGE NAME"].str.contains(listJruDiscardMessages[x], na=False, regex=False)
            if x > 0:
                # Sum of filters
                serFilterJruDiscardedMess[0] |= serFilterJruDiscardedMess[x]
        serFilterJruDiscardedMess[0] = ~serFilterJruDiscardedMess[0]        

        # Filtering
        inputdf = inputdf[serFilterJruDiscardedMess[0]]
        inputdf.index = np.arange(0, len(inputdf.axes[0]))

    # Filter for only-RBC required JRU Messages
    listJruRbcMessages = np.array(['MESSAGE TO RBC', 
                            'MESSAGE FROM RBC'] )
    for x in range(len(listJruRbcMessages)):
        serFilterRbcMessages[x] = inputdf["MESSAGE NAME"].str.contains(listJruRbcMessages[x], na=False, regex=False)
        if x > 0:
            serFilterRbcMessages[0] |= serFilterRbcMessages[x]

    # Filter for non-RBC messages 
    serFilterNonRbcMessages = ~serFilterRbcMessages[0]
     
    # Filter for required RBC packets
    if len(listRbcPackets) > 0:
        for x in range(len(listRbcPackets)):
            strRbcPacket = 'PACKET RBC: '+ str(listRbcPackets[x]) +'.'
            serFilterRbcPackets[x] = inputdf["PACKET RBC"].str.contains(strRbcPacket, na=False, regex=False)
            if x > 0:
                serFilterRbcPackets[0] |= serFilterRbcPackets[x]

    # Now we have differente filters
    # serFilterNonRbcMessages for messages not from/to RBC (balises)
    # serFilterRbcMessages[0] for messages from/to RBC
    # serFilterRbcPackets[0] for required (as in listRbcFilter ) packets from/to RBC 
    if serFilterRbcPackets.size > 0:
        serFinalFilter = serFilterNonRbcMessages | serFilterRbcPackets[0]

    if len(serFinalFilter) > 0:
        inputdf = inputdf[serFinalFilter]

    # reindexamos nuevo dataframe
    inputdf.index = np.arange(0, len(inputdf.axes[0]))
    
    return inputdf

# AlGetNidmesageFromHeader(AlDf, listRbcFilter)
# Alstom(JDRMDR) decoder
# Get message id from header, returns a re-indexed df
# NID_MESSAGE

def AlGetNidmesageFromMessage(alDf):
    
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pd.DataFrame(alDf) # copia local del objeto del argumento

    dfNidmessage = pd.DataFrame()
    listNidmessage = []
    
    length = inputdf["HEADER"].size
    for x in range(length):

        # get NID_MESSAGE
        # i.e. MESSAGE FROM BALISE [6]
        strMessageName = inputdf["MESSAGE NAME"][x]
        strMessageId = inputdf["MESSAGE ID"][x]

        
        nidMessage = strMessageName + ' [' + strMessageId + ']'
        # Insertamos valor en la lista de tipos de mensajes
        listNidmessage.insert(x, nidMessage)
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()
        
    # Creamos dataframe resultado  
    dfNidmessage["NID_MESSAGE"] = np.array(listNidmessage)

    return dfNidmessage

# AlGetCommondata(hasDf, listRbcFilter)
# Alstom (JDRMDR) decoder
# Get some message data, returns a re-indexed df
# Date
# Time

def AlGetCommonData(hasDf):
    
    global glLocale
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pd.DataFrame(hasDf) # copia local del objeto del argumento

    dfCommondata = pd.DataFrame()
    
    # Ajustamos formato de datos
    # inputdf['Date'] = pandas.to_datetime(inputdf['Date'], format="%d/%m/%y").dt.date
    # inputdf['Time'] = pandas.to_datetime(inputdf['Time'], format="%H:%M:%S").dt.time
    # Creamos dataframe resultado  
    if glLocale == 'SPA':
        df = inputdf[["Grabar Id", "Fecha", "Hora"]] # only used to check file is in spanish format
        inputdf.rename(columns={"Fecha": "Date", "Hora": "Time", "Grabar Id": "Record Id"}, inplace = True)
        dfCommondata = inputdf[["Record Id", "Date", "Time"]]
    elif glLocale == 'ENG':
        dfCommondata = inputdf[["Record Id", "Date", "Time"]]
    else:
        print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
        exit(1)
        
    

    # Adjus string to dd/mm/aa
    for x in range(len(inputdf.axes[0])):
        strDate = str(inputdf['Date'][x])
        i = strDate.index('/')
        dd = strDate[0:i].rjust(2, '0')
        strDate = strDate[i+1:len(strDate)]
        i = strDate.index('/')
        mm = strDate[0:i].rjust(2, '0')
        yy = strDate[i+1:len(strDate)]
        strDate = dd + '/' + mm + '/' + yy

        inputdf.iloc[x, inputdf.columns.get_loc('Date')] = strDate

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()
    
    return dfCommondata

# AlGetTrainEngineHeader(hasDf)
# Alstom JDRMDR decoder
# Get train engine id from header, returns a re-indexed df
# NID_ENGINE

def AlGetTrainEngineFromHeader(hasDf):

    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum
    
    inputdf = pd.DataFrame(hasDf) # copia local del objeto del argumento
    
    dfTrainengine = pd.DataFrame()
    listTrainengine = []
    
    length = inputdf["HEADER"].size
    for x in range(length):
        strHead = inputdf["HEADER"][x]

        #get NID ENGINE
        
        try:
            #NID_ENGINE : 17232 
            match = re.search("NID_ENGINE\s:", strHead)
            start = match.start()
            match = re.search("\n", strHead[start:len(strHead)])
            end = start+match.end()
            strTrainengine = strHead[start:end]
            # get nid_engine value
            match = re.search("NID_ENGINE\s:", strTrainengine)
            start = match.end()
            match = re.search("\d+", strTrainengine[start:len(strTrainengine)])
            idEngine = strTrainengine[start+match.start():start+match.end()]
        except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Unknown", strTrainengine)
                    start = match.end()
                    idEngine = 'Unknown'
                except Exception as ex:
                    print(color.Fore.RED + "ERROR-HEADER. NID_ENGINE not found. " + str(ex) + color.Style.RESET_ALL)
                    exit()
        
        
        # Insertamos valor en la lista de idEngine
        listTrainengine.insert(x, idEngine)
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

        
    # Creamos df train engine
    dfTrainengine["NID_ENGINE"] = listTrainengine
    
    return dfTrainengine

# AlGetTrainPosHeader(alDf)
# Alstom (JDRMDR) decoder
# Get train postion data from header, returns a re-indexed df
#    TRAIN_POSITION.Q_SCALE : 2-->10 m scale
#    TRAIN_POSITION.NID_C : 383-->Identity number of the country or region = 383
#    TRAIN_POSITION.NID_BG : 717-->717
#    TRAIN_POSITION.D_LRBG : 5198-->Value = 51980.00  m
#    TRAIN_POSITION.Q_DIRLRBG : 0-->Reverse
#    TRAIN_POSITION.Q_DLRBG : 1-->Nominal
#    TRAIN_POSITION.L_DOUBTOVER : 75-->Over-reading error = 750.00 m
#    TRAIN_POSITION.L_DOUBTUNDER : 89-->Under-reading error = 890.00 m
def AlGetTrainPosFromHeader(alDf):
    
    global glLocale
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pd.DataFrame(alDf) # copia local del objeto del argumento
    
    dfTrainpos = pd.DataFrame()
    listQscale = []
    listLrbg = []
    listDistlrbg = []
    listQdirlrbg = []
    listQdlrbg = []
    listLDoubtover = []
    listLDoubtunder = []

    length = inputdf["HEADER"].size
    try:
        for x in range(length):
            strHead = inputdf["HEADER"][x]

            #get TRAIN POSITION
            #    TRAIN_POSITION.Q_SCALE : 2-->10 m scale
            #    TRAIN_POSITION.NID_C : 383-->Identity number of the country or region = 383
            #    TRAIN_POSITION.NID_BG : 717-->717
            #    TRAIN_POSITION.D_LRBG : 5198-->Value = 51980.00  m
            #    TRAIN_POSITION.Q_DIRLRBG : 0-->Reverse
            #    TRAIN_POSITION.Q_DLRBG : 1-->Nominal
            #    TRAIN_POSITION.L_DOUBTOVER : 75-->Over-reading error = 750.00 m
            #    TRAIN_POSITION.L_DOUBTUNDER : 89-->Under-reading error = 890.00 m
            match = re.search("TRAIN_POSITION.Q_SCALE", strHead)
            start = match.start()
            match = re.search("TRAIN_POSITION.L_DOUBTUNDER", strHead[start:len(strHead)])
            end = start + match.start()
            match = re.search("\n", strHead[start+match.start():len(strHead)])
            end += match.end()
            strTrainpos = strHead[start:end]
            
            # get q_scale
            # TRAIN_POSITION.Q_SCALE : 2-->10 m scale
            # Qualifier for the distance/length scale
            # 0 - 10 cm scale
            # 1 - 1 m scale
            # 2 - 10 m scale
            # 3 - Spare

            try:

                match = re.search("Q_SCALE\s:", strTrainpos)
                start = match.start()
                match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
                end = start + match.start()
                strQscale = strTrainpos[start:end]
                # get q_scale value
                match = re.search("Q_SCALE\s:", strQscale)
                start = match.end()
                match = re.search("\d+", strQscale[start:len(strQscale)])
                qscale = strQscale[start+match.start():start+match.end()]
            
            except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Unknown", strQscale)
                    start = match.end()
                    qscale = '3' # Spare SS026-7
                    print("INFO-HEADER. Unknown Q_SCALE. Record Id: " + str(alDf['Record Id'][x]))
                except Exception as ex:
                    print(color.Fore.RED + "ERROR-HEADER. Found error in Q_SCALE. Record Id: " + str(alDf['Record Id'][x]) + str(ex) + color.Style.RESET_ALL)
                    exit()
            
            # Insertamos valor en la lista de Qscale
            listQscale.insert(x, qscale)

            # get LRBG string
            # TRAIN_POSITION.NID_BG : 717-->717

            try:

                match = re.search("NID_BG\s:", strTrainpos)
                start = match.start()
                match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
                end = start + match.start()
                strLrbg = strTrainpos[start:end]
                # get LRBG ID
                match = re.search("NID_BG\s:", strLrbg)
                start = match.end()
                match = re.search("\d+", strLrbg[start:len(strLrbg)])
                idLrbg = strLrbg[start+match.start():start+match.end()]
                
            except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Unknown", strLrbg)
                    start = match.end()
                    idLrbg = '16383' # Unknown balise SS 026-7
                    print("INFO-HEADER. Unknown NID_BG. Record Id: " + str(alDf['Record Id'][x]))
                except Exception as ex:
                    print(color.Fore.RED + "ERROR-HEADER. Found error in NID_BG. Record Id: " + str(alDf['Record Id'][x]) + str(ex) + color.Style.RESET_ALL)
                    exit()
            
            # Insertamos valor en la lista de Lrbg
            listLrbg.insert(x, idLrbg)

            # get LRBG distance
            # TRAIN_POSITION.D_LRBG : 5198-->Value = 51980.00  m
            # Distance between the last relevant balise group and the estimated front end 
            #   of the train (the side of the active cab).

            try:
                # get D_LRBG
                match = re.search("D_LRBG\s:", strTrainpos)
                start = match.start()
                match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
                end = start + match.start()
                strDistlrbg = strTrainpos[start:end]
                
                match = re.search("Value\s=", strDistlrbg)
                start = match.end()
                match = re.search("\d*[.,]\d+|\d+", strDistlrbg[start:len(strDistlrbg)])
                D_lrbg = strDistlrbg[start+match.start():start+match.end()]
                # Insertamos valor en la lista de D_LRBG
                # Previamente lo convertimos a float

                if glLocale == 'SPA': 
                    D_lrbg = D_lrbg.replace('.','')
                    D_lrbg = float(D_lrbg.replace(",", "."))
                elif glLocale == 'ENG': D_lrbg = float(D_lrbg.replace(",", ""))
                else:
                    print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
                    exit(1)

            except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Unknown", strDistlrbg)
                    start = match.end()
                    D_lrbg = '32767' # Unknown SS026-7
                    print("INFO-HEADER. Unknown D_LRBG. Record Id: " + str(alDf['Record Id'][x]))
                except Exception as ex:
                    print(color.Fore.RED + "ERROR-HEADER. Found error in D_LRBG. Record Id: " + str(alDf['Record Id'][x]) + str(ex) + color.Style.RESET_ALL)
                    exit()


            listDistlrbg.insert(x, D_lrbg)

            # get Q_DIRLRBG
            # TRAIN_POSITION.Q_DIRLRBG : 0-->Reverse
            # Orientation of the train in relation to the direction of the LRBG
            # 0 - Reverse
            # 1 - Nominal
            # 2 - Unknown
            # 3 - Spare

            try:

                match = re.search("Q_DIRLRBG\s:", strTrainpos)
                start = match.start()
                match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
                end = start + match.start()
                strQdirlrbg = strTrainpos[start:end]
                # get orientation
                match = re.search("Q_DIRLRBG\s:", strQdirlrbg)
                start = match.end()
                match = re.search("\d+", strQdirlrbg[start:len(strQdirlrbg)])
                Q_dirlrbg = strQdirlrbg[start+match.start():start+match.end()]
            
            except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Unknown", strQdirlrbg)
                    start = match.end()
                    Q_dirlrbg = '2' # Unknown SS026-7
                    print("INFO-HEADER. Unknown Q_DIRLRBG. Record Id: " + str(alDf['Record Id'][x]))
                except Exception as ex:
                    print(color.Fore.RED + "ERROR-HEADER. Found error in Q_DIRLRBG. Record Id: " + str(alDf['Record Id'][x]) + str(ex) + color.Style.RESET_ALL)
                    exit()
            
            
            # Insertamos valor en la lista de Q_DIRLRBG
            listQdirlrbg.insert(x, Q_dirlrbg)
            
            # get Q_DLRBG
            #  TRAIN_POSITION.Q_DLRBG : 1-->Nominal
            # Qualifier telling on which side of the LRBG the estimated front end is
            # 0 - Reverse
            # 1 - Nominal
            # 2 - Unknown
            # 3 - Spare

            try:
                
                match = re.search("Q_DLRBG\s:", strTrainpos)
                start = match.start()
                match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
                end = start + match.start()
                strQdlrbg = strTrainpos[start:end]
                # get orientation
                match = re.search("Q_DLRBG\s:", strQdlrbg)
                start = match.end()
                match = re.search("\d+", strQdlrbg[start:len(strQdlrbg)])
                Q_dlrbg = strQdlrbg[start+match.start():start+match.end()]
            
            except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Unknown", strQdlrbg)
                    start = match.end()
                    Q_dlrbg = '2'    # Unknown SS026-7
                    print("INFO-HEADER. Unknown Q_DLRBG. Record Id: " + str(alDf['Record Id'][x]))
                except Exception as ex:
                    print(color.Fore.RED + "ERROR-HEADER. Found error in Q_DLRBGG. Record Id: " + str(alDf['Record Id'][x]) + str(ex) + color.Style.RESET_ALL)
                    exit()
            
            # Insertamos valor en la lista de Q_DIRLRBG
            listQdlrbg.insert(x, Q_dlrbg)


            # get L_DOUBTOVER
            # TRAIN_POSITION.L_DOUBTOVER : 75-->Over-reading error = 750.00 m
            # The over-reading amount plus the Q_LOCACC of the LRBG

            try:

                match = re.search("L_DOUBTOVER\s:", strTrainpos)
                start = match.start()
                match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
                end = start + match.start()
                strDoubtover = strTrainpos[start:end]
                # get D_LRBG
                match = re.search("Over-reading\serror\s=", strDoubtover)
                start = match.end()
                match = re.search("\d*[.,]\d+|\d+", strDoubtover[start:len(strDoubtover)])
                L_doubtover = strDoubtover[start+match.start():start+match.end()]
                # Insertamos valor en la lista de D_LRBG
                # Previamente lo convertimos a float
                if glLocale == 'SPA': 
                    L_doubtover = float(L_doubtover.replace(",", "."))
                elif glLocale == 'ENG': L_doubtover = float(L_doubtover.replace(",", ""))
                else:
                    print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
                    exit(1)
            
            except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Unknown", strDoubtover)
                    start = match.end()
                    L_doubtover = '32767' # Unknown SS026-7
                    print("INFO-HEADER. Unknown L_DOUBTOVER. Record Id: " + str(alDf['Record Id'][x]))
                except Exception as ex:
                    print(color.Fore.RED + "ERROR-HEADER. Found error in L_DOUBTOVER. Record Id: " + str(alDf['Record Id'][x]) + str(ex) + color.Style.RESET_ALL)
                    exit()
            
            
            listLDoubtover.insert(x, L_doubtover)

            # get L_DOUBTUNDER
            # TRAIN_POSITION.L_DOUBTUNDER : 89-->Under-reading error = 890.00 m
            # The under-reading amount plus the Q_LOCACC of the LRBG

            try:

                match = re.search("L_DOUBTUNDER\s:", strTrainpos)
                start = match.start()
                match = re.search("\n", strTrainpos[start:len(strTrainpos)]) 
                end = start + match.start()
                strDoubtunder = strTrainpos[start:end]
                # get D_LRBG
                match = re.search("Under-reading\serror\s=", strDoubtunder)
                start = match.end()
                match = re.search("\d*[.,]\d+|\d+", strDoubtunder[start:len(strDoubtunder)])
                L_doubtunder = strDoubtunder[start+match.start():start+match.end()]
                # Insertamos valor en la lista de D_LRBG
                # Previamente lo convertimos a float
                if glLocale == 'SPA': 
                    L_doubtunder = float(L_doubtunder.replace(",", "."))
                elif glLocale == 'ENG': L_doubtunder = float(L_doubtunder.replace(",", ""))
                else:
                    print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
                    exit(1)

            except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Unknown", strDoubtunder)
                    start = match.end()
                    L_doubtunder = '32767' # Unknown SS026-7
                    print("INFO-HEADER. Unknown L_DOUBTUNDER. Record Id: " + str(alDf['Record Id'][x]))
                except Exception as ex:
                    print(color.Fore.RED + "ERROR-HEADER. Found error in L_DOUBTUNDER. Record Id: " + str(alDf['Record Id'][x]) + str(ex) + color.Style.RESET_ALL)
                    exit()        
            
            
            listLDoubtunder.insert(x, L_doubtunder)

            glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
            sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
            sys.stdout.flush()

    except Exception as ex:
        print(strTrainpos)
        print(ex)
        exit(0)

    # Creamos variable train pos
    # Se ajustan tipos de los datos
    dfTrainpos["Q_SCALE"] = np.array(listQscale, dtype=int)
    dfTrainpos["NID_LRBG"] = listLrbg # se mantiene como string
    dfTrainpos["D_LRBG"] = listDistlrbg # se ajustó a float al insertar en la lista 
    dfTrainpos["Q_DIRLRBG"] = np.array(listQdirlrbg, dtype = int)
    dfTrainpos["Q_DLRBG"] = np.array(listQdlrbg, dtype = int)
    dfTrainpos["L_DOUBTOVER"] = listLDoubtover # se ajustó a float al insertar en la lista
    dfTrainpos["L_DOUBTUNDER"] = listLDoubtunder # se ajustó a float al insertar en la lista

    return dfTrainpos


# AlGetVelocityFromHeader(hasDf)
# Alstom (JDRMDR) decoder
# Get train velocity from header, returns a re-indexed df
# V_TRAIN : 1023-->Standstill
# V_TRAIN : 6-->6 km/h

def AlGetVelocityFromHeader(alDf):
    
    global glLocale
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pd.DataFrame(alDf) # copia local del objeto del argumento
    
    dfTrainvel = pd.DataFrame()
    listVeltrain = []
    
    
    length = inputdf["HEADER"].size
    for x in range(length):
        strHead = inputdf["HEADER"][x]
        # get V_TRAIN
        # V_TRAIN : 1023-->Standstill
        # V_TRAIN : 6-->6 km/h

        try:

            match = re.search("V_TRAIN\s:", strHead)
            start = match.start()
            match = re.search("\n", strHead[start:len(strHead)])
            end = start+match.end()
            strVeltrain = strHead[start:end]
            # get v_train value
            match = re.search("\-\-\>", strVeltrain)
            start = match.end()
            match = re.search("\d*[.,]\d+|\d+", strVeltrain[start:len(strVeltrain)])
            velTrain = strVeltrain[start+match.start():start+match.end()]
            # Case 1023 --> standstill
            if velTrain == '1023': velTrain = '0'
            # Insertamos valor en la lista de velocidades
            # Previamente lo convertimos a float
            if glLocale == 'SPA': 
                velTrain = float(velTrain.replace(",", "."))
            elif glLocale == 'ENG': velTrain = float(velTrain.replace(",", ""))
            else:
                print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
                exit(1)
        
        
        except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Standstill", strVeltrain)
                    start = match.end()
                    velTrain = 0.0 # Standstill
                    
                except Exception as ex:

                    try:
                        match = re.search("Unknown", strVeltrain)
                        start = match.end()
                        velTrain = None # Unknown 
                        print("INFO-HEADER. Unknown V_TRAIN. Record Id: " + str(alDf['Record Id'][x]))

                    except Exception as err:
                        print(color.Fore.RED + "ERROR-HEADER. Found error in V_TRAIN. Record Id: " + str(alDf['Record Id'][x]) + str(ex) + color.Style.RESET_ALL)
                        exit() 
        

        listVeltrain.insert(x, velTrain)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()



    # Creamos dataframe resultado  
    dfTrainvel["V_TRAIN"] = listVeltrain
    
    return dfTrainvel


# AlGetetcsLevelFromHeader(hasDf)
# Alstom decoder (JDRMDR) decoder
# Get etcs level from header, returns a re-indexed df
# M_LEVEL

def AlGetetcsLevelFromHeader(alDf):
    
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pd.DataFrame(alDf) # copia local del objeto del argumento
    
    dfEtcsLevel = pd.DataFrame()
    listEtcsLevel = []
    
    
    length = inputdf["HEADER"].size
    for x in range(length):
        strHead = inputdf["HEADER"][x]
        # get M_LEVEL
        # 0 - Level 0
        # 1 - Level NTC specified by NID_NTC
        # 2 - Level 1
        # 3 - Level 2
        # 4-7 - Spare
        # M_LEVEL : 2-->Level 1
            
        try:
        
            match = re.search("M_LEVEL\s:", strHead)
            start = match.start()
            match = re.search("\n", strHead[start:len(strHead)])
            end = start+match.end()
            strMlevel = strHead[start:end]
            # get m_level value (number)
            match = re.search("M_LEVEL\s:", strMlevel)
            start = match.end()
            match = re.search("\d+", strMlevel[start:len(strMlevel)])
            mLevel = strMlevel[start+match.start():start+match.end()]
        
        
        except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Unknown", strMlevel)
                    start = match.end()
                    mLevel = 'Unknown' # Unknown
                    print("INFO-HEADER. Unknown M_LEVEL. Record Id: " + str(alDf['Record Id'][x]))
                except Exception as ex:
                    print(color.Fore.RED + "ERROR-HEADER. Found error in M_LEVEL. Record Id: " + str(alDf['Record Id'][x]) + str(ex) + color.Style.RESET_ALL)
                    exit() 
        
        
        # Insertamos valor en la lista de velocidades
        listEtcsLevel.insert(x, mLevel)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()


    # Creamos dataframe resultado  
    dfEtcsLevel["M_LEVEL"] = np.array(listEtcsLevel) # no dtype for 'Unknown'; if other case, use: dtype=int)
    
    return dfEtcsLevel

# AlGetetcsModeFromHeader(hasDf)
# Alstom decoder (JDRMDR) decoder
# Get etcs mode from header, returns a re-indexed df
# M_MODE

def AlGetetcsModeFromHeader(alDf):
    
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pd.DataFrame(alDf) # copia local del objeto del argumento
    
    dfEtcsMode = pd.DataFrame()
    listEtcsMode = []
    
    
    length = inputdf["HEADER"].size
    for x in range(length):
        strHead = inputdf["HEADER"][x]
        # get M_MODE
        # Onboard operating mode
        # 0 - Full Supervision
        # 1 - On Sight
        # 2 - Staff Responsible
        # 3 - Shunting
        # 4 - Unfitted
        # 5 - Sleeping
        # 6 - Stand By
        # 7 - Trip
        # 8 - Post Trip
        # 9 - System Failure
        # 10 - Isolation
        # 11 - Non Leading
        # 12 - Limited Supervision
        # 13 - National System
        # 14 - Reversing
        # 15 - Passive Shunting
        # 16 - Automatic Driving
        # 17 - Supervised Manoeuvre
        # 18-31 - Spare
        # M_MODE : 6-->Stand By
            
        try:
            
            match = re.search("M_MODE\s:", strHead)
            start = match.start()
            end = len(strHead) # Último campo de la celda
            strMmode = strHead[start:end]
            # get m_level value (number)
            match = re.search("M_MODE\s:", strMmode)
            start = match.end()
            match = re.search("\d+", strMmode[start:len(strMmode)])
            mMode = strMmode[start+match.start():start+match.end()]
        
        except Exception as ex:
                # check if Unknown value
                try:
                    match = re.search("Unknown", strMmode)
                    start = match.end()
                    mMode = 'Unknown' # Unknown
                    print("INFO-HEADER. Unknown M_MODE. Record Id: " + str(alDf['Record Id'][x]))
                except Exception as ex:
                    print(color.Fore.RED + "ERROR-HEADER. Found error in M_MODE. Record Id: " + str(alDf['Record Id'][x]) + str(ex) + color.Style.RESET_ALL)
                    exit() 
        
        # Insertamos valor en la lista de modos
        listEtcsMode.insert(x, mMode)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()


    # Creamos dataframe resultado  
    dfEtcsMode["M_MODE"] = np.array(listEtcsMode) # , dtype=int)
    
    return dfEtcsMode

# HasGetPacketRBC(hasDf, listRbcFilter, listJruFilter)
# Hasler decoder
# Get data from PACKET_RBC content
# (under development)
# PACKET_RBC (packet id)
# PACKET: 15. Level 2/3 Movement Authority
# 1 .. n packets con su diferente contenido
# Apply filter for JRU messages and RBC packets returns a re-indexed df with the filtered file
def AlGetPacketRBC(alDf):
    
    global glLocale
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pd.DataFrame(alDf) # copia local del objeto del argumento
    

    dfPacketrbc = pd.DataFrame()
    listPacketrbc = []
    listDref = []
    listSS026Packets = []
    # For packet 5
    listLinksBG = [] 
    # For packet 27
    listSSP = []
    # For packet 88
    listLX = []
    
    length = inputdf["PACKET RBC"].size
    try: 
        for x in range(length):
            strPacketrbc = inputdf["PACKET RBC"][x]
            strDatarbc = inputdf["DATA"][x]

            # Hasler
            #PACKET RBC: 33. MA with Shifted Location Reference
            #NID_MESSAGE (8): 33
            #L_MESSAGE (10): 267 byte(s)
            #T_TRAIN (32): 7031.87 s [703187]
            #M_ACK (1): Acknowledgement required
            #NID_LRBG: 
            #    NID_C (10): Reserved [391]
            #    NID_BG (14): 1656
            #Q_SCALE (2): 1 m
            #D_REF (16): -92,00 m [-92]
            #PACKET: 15. Level 2/3 Movement Authority
            # 1 .. n packets con su diferente contenido


            # JDRMDR structure
            # MESSAGE RBC: 33
            #   Q_SCALE : 0 -->10 cm scale
            #   D_REF : 64779 -->-75,7 m
            # PACKET: 15
            #    Q_DIR : 0 -->Reverse
            #    L_PACKET : 88
            #    Q_SCALE : 1 -->1 m scale
            #    V_EMA : 0 -->0km/h
            #    T_EMA : 1023 -->infinite
            #    niter : 0
            #    L_ENDSECTION : 4669 -->4669 m
            #    Q_SECTIONTIMER : 0 -->No section Timer Information
            #    Q_ENDTIMER : 0 -->No End section timer information
            #    Q_DANGERPOINT : 1 -->Danger point information to follow
            #    D_DP : 1 -->1 m
            #    V_RELEASEDP : 0 -->0 km/h
            #    Q_OVERLAP : 0 -->No overlap information
            # PACKET: 27
            #    Q_DIR : 0 -->Reverse
            #    L_PACKET : 86
            #    Q_SCALE : 1 -->1 m scale
            #    D_STATIC : 57 -->57 m
            #    V_STATIC : 10 -->50km/h
            #    Q_FRONT : 0 -->Train length delay on validity end point of profile element
            #    niter : 0
            #    niter : 1
            #    D_STATIC[0] : 4623 -->4623 m
            #    V_STATIC[0] : 127 -->No numericalvalue telling the static speed profile description ends at D_STATIC(n)
            #    Q_FRONT[0] : 0 -->Train length delay on validity end point of profile element
            #    niter[0] : 0
            # PACKET: 21
            #    Q_DIR : 0 -->Reverse
            #    L_PACKET : 78
            #    Q_SCALE : 1 -->1 m scale
            #    D_GRADIENT : 57 -->57 m
            #    Q_GDIR : 0 -->downhill
            #    G_A : 0 -->0/oo
            #    niter : 1
            #    D_GRADIENT[0] : 4623 -->4623 m
            #    Q_GDIR[0] : 0 -->downhill
            #    G_A[0] : 255 -->Non numerical value telling that the current gradien description ends at D_GRADIENT(n)
            # PACKET: 5
            #    Q_DIR : 0 -->Reverse
            #    L_PACKET : 69
            #    Q_SCALE : 0 -->10 cm scale
            #    D_LINK : 1600 -->160 m
            #    Q_NEWCOUNTRY : 0 -->Same country / railway administration, no NID_C follows
            #    NID_BG : 501 -->501
            #    Q_LINKORIENTATION : 1 -->The balise group is seen by the train in nominal direction
            #    Q_LINKREACTION : 2 -->No reaction
            #    Q_LOCACC : 1 -->1m
            #    niter : 0

            #get PACKET_RBC header 

            # if pd.isnull(strPacketrbc):
            if strPacketrbc == '' or strDatarbc == '':
                listPacketrbc.insert(x, "")
                listDref.insert(x, "")
                listSS026Packets.insert(x, "")
                listLinksBG.insert(x, "")
                listSSP.insert(x,"")
                listLX.insert(x,"")
            else:
                match = re.search("MESSAGE\sRBC:", strPacketrbc)
                start = match.start()
                try:
                    match = re.search("PACKET:", strPacketrbc[start:len(strPacketrbc)])
                    end = start + match.start()
                except:
                    # Case: no ss026 packets
                    end = len(strPacketrbc)
                strPacketheader = strPacketrbc[start:end].rstrip() # Clean trailing special chars
                strPacketheader = strPacketheader.lstrip() # Clean leading special chars
                # get PACKET_RBC ID
                match = re.search("MESSAGE\sRBC:", strPacketheader)
                start = match.end()
                match = re.search("\d+", strPacketheader[start:len(strPacketheader)])
                idPacketrbc = strPacketheader[start+match.start():start+match.end()]
                # Insertamos valor en la lista
                listPacketrbc.insert(x, idPacketrbc)

                # get D_REF (case packet 33 MA shifted)
                try:
                    match = re.search("D_REF\s:", strPacketheader)
                    start = match.end()
                    #   D_REF : 64779 -->-75,7 m
                    match = re.search("\-\-\>", strPacketheader[start:len(strPacketheader)])
                    start = start + match.end()
                    # match = re.search("[+-]?\d*[.,]\d+|\d+", strPacketheader[start:len(strPacketheader)]) # Note: it doesn´t work with negative integer
                    match = re.search("[+-]?(?:\d*[.,]\d+|\d+)", strPacketheader[start:len(strPacketheader)])
                    dref = strPacketheader[start+match.start():start+match.end()]

                    # WARNING: JDRMDR uses spanish format in RBC messages for D_REF
                    #if glLocale == 'SPA': dref = float(dref.replace(",", "."))
                    #elif glLocale == 'ENG': dref = float(dref.replace(",", ""))
                    #else:
                    #    print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
                    #    exit(1)
                    dref = float(dref.replace(",", ".")) # Warning (exception for JDRMDR)
                    
                except:
                    dref = ""
                
                # Insertamos valor en la lista
                listDref.insert(x, dref)
                
                # get list of SS26 packets in PACKET_RBC
                # PACKET: 15. Level 2/3 Movement Authority
                # PACKET: 5. Linking
                # PACKET: 27. SSP
                # PACKET: ...
                boolWhile = True
                boolPacket5 = False
                boolPacket27 = False
                boolPacket88 = False
                
                boolAtleast1SS026Packet = False
                istart = 0
                strSS026packet = None
                idSS026packet = None
                while boolWhile == True: 
                    
                    match = re.search("PACKET:", strPacketrbc[istart:len(strPacketrbc)])
                    if match == None:
                        if strSS026packet is None: strSS026packet = ""
                        boolWhile = False
                        continue

                    boolAtleast1SS026Packet = True
                    start = istart + match.end()
                    match = re.search("\d+", strPacketrbc[start:len(strPacketrbc)])
                    idSS026packet = strPacketrbc[start+match.start():start+match.end()]
                    if idSS026packet == '5': 
                        boolPacket5 = True # identificamos que en este PACKET RBC viene un paquete 5 (linking)
                    if idSS026packet == '27': 
                        boolPacket27 = True # identificamos que en este PACKET RBC viene un paquete 27 (SSP)
                    if idSS026packet == '88': 
                        boolPacket88 = True # identificamos que en este PACKET RBC viene un paquete 27 (SSP)
                    
                    
                    istart = start
                    if boolAtleast1SS026Packet == False: idSS026packet = '-'

                    if strSS026packet is None:
                        strSS026packet = str(idSS026packet)
                    else: 
                        strSS026packet += ", "+str(idSS026packet)
                # Insertamos valor en la lista
                listSS026Packets.insert(x, strSS026packet)

                # Get reference LRBG from data
                # NID_C : 383-->383
                # NID_RBC : 1-->1
                # NID_MESSAGE : 33-->MA with Shifted Location Reference
                # L_MESSAGE : 94
                # T_TRAIN : 153665
                # M_ACK : 1-->Acknowledgement required
                # NID_LRBG1 : 383
                # NID_LRBG2 : 517

                # Reference LRBG is NID_LRBG2
                match = re.search("NID_LRBG2\s:", strDatarbc)
                if match != None:                
                    start = match.end()
                    strRefNidLrbg = str(strDatarbc[start:len(strDatarbc)])
                    strRefNidLrbg = strRefNidLrbg.lstrip()
                    strRefNidLrbg = strRefNidLrbg.rstrip()
                    
                else:
                    strRefNidLrbg = ''



                # ------------------   Get linking from Packet 5  (starts) --------------

                if boolPacket5 == False:
                    listLinksBG.insert(x, "")
                else:
                          
                    # get complete Packet 5 string
                    match = re.search("PACKET:\s5\n", strPacketrbc) # !important \n
                    startP5 = match.start()
                    nextStart = match.end()
                    match = re.search("PACKET:", strPacketrbc[nextStart:len(strPacketrbc)])
                    if match != None:
                        endP5 = nextStart + match.start()
                    else:
                        endP5 = len(strPacketrbc)

                    strPacket5 = strPacketrbc[startP5:endP5]
                    strPacket5 = strPacket5.rstrip() # Clean trailing special chars
                    strPacket5 = strPacket5.lstrip() # Clean leading special chars
                

                    strListLink = AlGetPacket5FromPaketRBC(strPacket5, strRefNidLrbg, dref)
                    listLinksBG.insert(x, strListLink)

                # ------------------   Get SSP from Packet 27  (starts) --------------

                if boolPacket27 == False:
                    listSSP.insert(x, "")
                else:

                    # get complete Packet 27 string
                    match = re.search("PACKET:\s27\n", strPacketrbc) # !important \n
                    startP27 = match.start()
                    nextStart = match.end()
                    match = match = re.search("PACKET:", strPacketrbc[nextStart:len(strPacketrbc)])
                    if match != None:
                        endP27 = nextStart + match.start()
                    else:
                        endP27 = len(strPacketrbc)

                    strPacket27 = strPacketrbc[startP27:endP27]
                    strPacket27 = strPacket27.rstrip() # Clean trailing special chars
                    strPacket27 = strPacket27.lstrip() # Clean leading special chars
                

                    strListSSP = AlGetPacket27FromPaketRBC(strPacket27, strRefNidLrbg)
                    listSSP.insert(x, strListSSP)

                # ------------------   Get LX info from Packet 88 (level crossing) --------------

                if boolPacket88 == False:
                        
                    listLX.insert(x, "")

                else:

                    # get complete Packet 88 string
                    
                    match = re.search("PACKET:\s88\n", strPacketrbc) # !important \n
                    startP88 = match.start()
                    nextStart = match.end()
                    match = match = re.search("PACKET:", strPacketrbc[nextStart:len(strPacketrbc)])
                    if match != None:
                        endP88= nextStart + match.start()
                    else:
                        endP88 = len(strPacketrbc)

                    strPacket88 = strPacketrbc[startP88:endP88]
                    strPacket88 = strPacket88.rstrip() # Clean trailing special chars
                    strPacket88 = strPacket88.lstrip() # Clean leading special chars

                    lxstatus = AlGetPacket88FromPaketRBC(strPacket88, strRefNidLrbg)

                    # Set LX STATUS
                    listLX.insert(x, lxstatus)           

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    except: 
        sys.stdout.write("\n")
        sys.stdout.flush()

        print(color.Fore.YELLOW + 
              "AlGetPacketRBC-Error in Packet RBC. Record ID = "+ str(inputdf['Record Id'][x]) +
              color.Style.RESET_ALL)

        

    # Creamos variable train pos
    # Se ajustan tipos de los datos
    dfPacketrbc["PACKET_RBC"] = np.array(listPacketrbc)
    dfPacketrbc["D_REF"] = np.array(listDref) # Previously converted to float with "." as decimal separator
    dfPacketrbc["Rbc Packets"] = np.array(listSS026Packets)
    dfPacketrbc["BG_LINKS"] = np.array(listLinksBG, dtype=str)
    dfPacketrbc["SSP"] = np.array(listSSP, dtype=str)
    dfPacketrbc["LX"] = np.array(listLX, dtype=str)

    return dfPacketrbc


# AlGetPacket5FromPaketRBC(strPacket5, refNidlrbg, dref)
# Alstom(JDRMDR) decoder
# Get data from MESSAGE RBC/PACKET 5 content
# Apply filter for JRU messages and RBC packets
# Returns a strings with a list of NID_BGs and D_LINKS in the format:
# nidbg0:dlink0, nidbg1:dlink1, ...
def AlGetPacket5FromPaketRBC(strPacket5, refNidlrbg, dref):
   
    global glLocale 
    
    # Get list of links in packet 5
    # Hasler format
    # PACKET: 5. Linking
    #         NID_PACKET (8): 5. Linking
    #         Q_DIR (2): Both directions [2]
    #         L_PACKET (13): 810 bit(s)
    #         Q_SCALE (2): 1 m
    #         D_LINK (15): 305,00 m [305]
    #         Q_NEWCOUNTRY (1): Same country / railway administration, no NID_C follows [0]
    #         NID_BG (14): 1338
    #         Q_LINKORIENTATION (1): The balise group is seen by the train in reverse direction [0]
    #         Q_LINKREACTION (2): No reaction [2]
    #         Q_LOCACC (6): 10 m
    #         N_ITER (5): 19
    #             D_LINK (15): 249,00 m [249]
    #             Q_NEWCOUNTRY (1): Same country / railway administration, no NID_C follows [0]
    #             NID_BG (14): 1336
    #             Q_LINKORIENTATION (1): The balise group is seen by the train in reverse direction [0]
    #             Q_LINKREACTION (2): No reaction [2]
    #             Q_LOCACC (6): 10 m / Q_LINKACC
    #               .... (iter)

    # ALstom /JDRMDR) frmat
    # PACKET: 5
    # Q_DIR : 1 -->Nominal
    # L_PACKET : 108
    # Q_SCALE : 0 -->10 cm scale
    # D_LINK : 768 -->76,8 m
    # Q_NEWCOUNTRY : 0 -->Same country / railway administration, no NID_C follows
    # NID_BG : 517 -->517
    # Q_LINKORIENTATION : 1 -->The balise group is seen by the train in nominal direction
    # Q_LINKREACTION : 1 -->Apply service brake
    # Q_LOCACC : 1 -->1m
    # niter : 1
    # D_LINK[0] : 1310 -->131 m
    # Q_NEWCOUNTRY[0] : 0 -->Same country / railway administration, no NID_C follows
    # NID_BG[0] : 521 -->521
    # Q_LINKORIENTATION[0] : 1 -->The balise group is seen by the train in nominal direction
    # Q_LINKREACTION[0] : 1 -->Apply service brake
    # Q_LOCACC[0] : 1 -->1m
                    

    boolWhilePck5 = True
    ilink = 0
    istart = 0
    startLinkBlock = 0
    startLinkHeader = 0
    strListLink = "" 
    # LRBG is set as first balise in listLinks
    strListLink += str(refNidlrbg) + ':0'  


    # Get link header from NID_PACKET to Q_LOCACC (first block)
    match = re.search("Q_DIR\s:", strPacket5[istart:len(strPacket5)])
                            
    startLinkHeader = istart + match.start() # Start of linking header                        
    start = istart + match.end()
    
    try:
        match = re.search("niter\s:", strPacket5[start:len(strPacket5)])
        end = start + match.start()
    except Exception as ex:
        # Case: no packet 5
        print(color.Fore.YELLOW + 'AlGetPacket5FromBalise-Exception: ' + str(ex) + color.Style.RESET_ALL)
        

    strLinkHeader = strPacket5[startLinkHeader:end].rstrip() # Clean trailing special chars
    strLinkHeader = strLinkHeader.lstrip() # Clean leading special chars

    
    # get Q_LOCACC from link header
    try:
        match = re.search("Q_LOCACC", strLinkHeader)
        start = match.end()
        match = re.search("\d+", strLinkHeader[start:len(strLinkHeader)])
    except: 
        try:
            # get Q_LINKACC from link header
            match = re.search("Q_LINKACC", strLinkHeader)
            start = match.end()
            match = re.search("\d+", strLinkHeader[start:len(strLinkHeader)])
        except Exception as ex:
            print(color.Fore.YELLOW + "Packet 5. Q_LOCACC/Q_LINKACC not found in header. " + str(ex) + color.Style.RESET_ALL)
            exit()

    qlocacc = 0 # To indicate no referenced qlocacc; strLinkHeader[start+match.start():start+match.end()]
    strListLink += ":" +str(qlocacc)

    istart = 0 # initialization to start with list of blocks
    while boolWhilePck5 == True: 
                                    
        # Obtenemos todo la estructura del enlace
        # desde D_LINK hasta Q_LOCACC
        match = re.search("D_LINK", strPacket5[istart:len(strPacket5)])
        if match == None:
            # There is no more link information
            boolWhilePck5 = False
            continue
                        
        startLinkBlock = istart + match.start() # Start of linking block
        ilink += 1 # To identify the first block which ends in N_ITER
                        
        start = istart + match.end()
        # If first link block
        if ilink == 1:
            match = re.search("niter\s:", strPacket5[start:len(strPacket5)])
            end = start + match.start()
        else: 
            try:
                match = re.search("D_LINK", strPacket5[start:len(strPacket5)])
                end = start + match.start()
            except:
                # Case: no more ss026 packets
                end = len(strPacket5)
        strLinkblock = strPacket5[startLinkBlock:end].rstrip() # Clean trailing special chars
        strLinkblock = strLinkblock.lstrip() # Clean leading special chars
                
        istart = start # start next link block
        # strlinkblock structure
        # First block:
        # D_LINK : 43 -->43 m
        #   Q_NEWCOUNTRY : 0 -->Same country / railway administration, no NID_C follows
        #   NID_BG : 715 -->715
        #   Q_LINKORIENTATION : 1 -->The balise group is seen by the train in nominal direction
        #   Q_LINKREACTION : 1 -->Apply service brake
        #   Q_LOCACC : 1 -->1m
        #
        # Iteractions:
        #       D_LINK[0] : 667 -->667 m
        #       Q_NEWCOUNTRY[0] : 0 -->Same country / railway administration, no NID_C follows
        #       NID_BG[0] : 709 -->709
        #       Q_LINKORIENTATION[0] : 0 -->The balise group is seen by the train in reverse direction
        #       Q_LINKREACTION[0] : 2 -->No reaction
        #       Q_LOCACC[0] : 2 -->2m

        # get NID_BG from link block
        match = re.search("NID_BG", strLinkblock)
        start = match.end()
        match = re.search(":", strLinkblock[start:len(strLinkblock)])
        start = start + match.end()
        match = re.search("\d+", strLinkblock[start:len(strLinkblock)])
        nidbg = strLinkblock[start+match.start():start+match.end()]
        # check nid_bg to be repeated (case Level 2)
        if nidbg == refNidlrbg:
            continue
        strListLink += "-"+str(nidbg)
        # set last nidbg as reference
        refNidlrbg = nidbg

        # get D_LINK
        # D_LINK : 43 -->43 m
        match = re.search("D_LINK", strLinkblock)
        start = match.end()
        match = re.search("\-\-\>", strLinkblock[start:len(strLinkblock)])
        start = start + match.end()
        match = re.search("\d*[.,]\d+|\d+", strLinkblock[start:len(strLinkblock)])
        d_link = strLinkblock[start+match.start():start+match.end()]

        # WARNING: JDRMDR uses ',' as decimals separator 
        # Depending on locale variable we adapt string and separators to standard float xywz.00
        #if glLocale == 'SPA': 
        #    d_link = d_link.replace(".", "")
        #    d_link = float(d_link.replace(",", "."))
        #elif glLocale == 'ENG': d_link = float(d_link.replace(",", ""))
        #else:
        #    print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
        #    exit(1)
        d_link = d_link.replace(".", "")
        d_link = float(d_link.replace(",", "."))  # WARNING 

        # If first link block
        if ilink == 1 and dref != "":
            # increase shifted distance
            d_link += dref

        # Insertamos valor en la lista de D_LINK
        strListLink += ":"+str(d_link)

        # get Q_LOCACC from link block
        #   Q_LOCACC : 1 -->1m
        try:
            match = re.search("Q_LOCACC", strLinkblock)
            start = match.end()
            match = re.search("\-\-\>", strLinkblock[start:len(strLinkblock)])
            start = start + match.end()
            match = re.search("\d+", strLinkblock[start:len(strLinkblock)])
        except:
            # get Q_LINKACC from link header
            match = re.search("Q_LINKACC", strLinkblock)
            start = match.end()
            match = re.search("\-\-\>", strLinkblock[start:len(strLinkblock)])
            start = start + match.end()
            match = re.search("\d+", strLinkblock[start:len(strLinkblock)])

        qlocacc = strLinkblock[start+match.start():start+match.end()]
        strListLink += ":"+str(qlocacc)
        
                    
    return strListLink

# AlGetPacket27FromPaketRBC(strPacket27, refNidlrbg)
# Alstom(JDRMDR) decoder
# Get data from MESSAGE RBC/PACKET 27 content
# Apply filter for JRU messages and RBC packets
# Returns a strings with a list of SSP in the format:
# nidbg0:dlink0, nidbg1:dlink1, ...
def AlGetPacket27FromPaketRBC(strPacket27, refNidlrbg):
   
    
                                  
    # Hasler struct
    # PACKET: 27. International Static Speed Profile
    #         NID_PACKET (8): 27. International Static Speed Profile
    #         Q_DIR (2): Both directions [2]
    #         L_PACKET (13): 114 bit(s)
    #         Q_SCALE (2): 1 m
    #         D_STATIC (15): 0,00 m [0]
    #         V_STATIC (7): 80 km/h [16]
    #         Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #         N_ITER (5): 0
    #         N_ITER (5): 2
    #             D_STATIC (15): 2.554,00 m [2554]
    #             V_STATIC (7): 40 km/h [8]
    #             Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #             N_ITER (5): 0
    #             D_STATIC (15): 624,00 m [624]
    #             V_STATIC (7): Non numerical value telling that the static speed profile description ends at D_STATIC(n) [127]
    #             Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #             N_ITER (5): 0

    # Alstom JDRMDR struct:
    # NID_PACKET : 27 -->International Static Speed Profile
    #    PKT_TK2T_27 (
    #        Q_DIR : 1 -->Nominal
    #        L_PACKET : 226
    #        Q_SCALE : 1 -->1 m scale
    #        D_STATIC : 0 -->0 m
    #        V_STATIC : 14 -->70km/h
    #        Q_FRONT : 0 -->Train length delay on validity end point of profile element
    #        niter : 0
    #        niter : 6
    #       D_STATIC[0] : 158 -->158 m
    #       V_STATIC[0] : 8 -->40km/h
    #       Q_FRONT[0] : 0 -->Train length delay on validity end point of profile element
    #       niter[0] : 0
    #       D_STATIC[1] : 500 -->500 m
    #       V_STATIC[1] : 12 -->60km/h
    #       Q_FRONT[1] : 0 -->Train length delay on validity end point of profile element
    #      niter[1] : 0
    #       D_STATIC[2] : 1100 -->1100 m
    #       V_STATIC[2] : 10 -->50km/h
    #       Q_FRONT[2] : 0 -->Train length delay on validity end point of profile element
    #       niter[2] : 0
    #       D_STATIC[3] : 400 -->400 m
    #       V_STATIC[3] : 12 -->60km/h
    #       Q_FRONT[3] : 0 -->Train length delay on validity end point of profile element
    #       niter[3] : 0
    #      D_STATIC[4] : 900 -->900 m
    #       V_STATIC[4] : 9 -->45km/h
    #       Q_FRONT[4] : 0 -->Train length delay on validity end point of profile element
    #       niter[4] : 0
    #       D_STATIC[5] : 100 -->100 m
    #       V_STATIC[5] : 10 -->50km/h
    #       Q_FRONT[5] : 0 -->Train length delay on validity end point of profile element
    #       niter[5] : 0
    #   )             

    # -----------------  Get list of distances and velocities  ----------------------
    boolWhilePck27 = True
    iSSP = 0
    istart = 0
    startSSPBlock = 0
    # LRBG is set as reference in list for dstatic
    strListSSP = str(refNidlrbg)
    while boolWhilePck27 == True: 


        # Get structure
        # from D_STATIC to last N_ITER (now we dont take into account V_DIFF & Q_DIFF)
        match = re.search("D_STATIC", strPacket27[istart:len(strPacket27)])
        if match == None:
            # There is no more link information
            boolWhilePck27 = False
            continue
        
        startSSPBlock = istart + match.start() # Start of SSP block
        
        start = istart + match.end() # After D_STATIC :
        
                        
        iSSP += 1 # To identify the first block which ends in N_ITER

        # If first SSP block
        if iSSP == 1:
            # D_STATIC : 0 -->0 m
            match = re.search("\s:", strPacket27[start:len(strPacket27)])
            start = start + match.end()
            match = re.search("niter\s:", strPacket27[start:len(strPacket27)])
            # jumping of first N_ITER used for Q_DIFF
            start = start + match.end()
            match = re.search("niter\s:", strPacket27[start:len(strPacket27)])
            start = start + match.end()
            end = start

                        
        else: 
            try:
                # D_STATIC[0] : 158 -->158 m
                match = re.search("\[", strPacket27[start:len(strPacket27)])
                start = start + match.end()
                match = re.search("niter\[", strPacket27[start:len(strPacket27)])
                start = start + match.start()
                end = start
            except:
                # Case: no ss026 packets
                start = len(strPacket27)
                end = len(strPacket27)
                       
        strSSPblock = strPacket27[startSSPBlock:end].rstrip() # Clean trailing special chars
        strSSPblock = strSSPblock.lstrip() # Clean leading special chars
                
        istart = start # start next ssp block

                        
        # strSSPblock structure
        #   D_STATIC : 0 -->0 m
        #   V_STATIC : 14 -->70km/h
        #   Q_FRONT : 0 -->Train length delay on validity end point of profile element

        # get D_STATIC from ssp block
        match = re.search("D_STATIC", strSSPblock)
        start = match.end()
        match = re.search("\-\-\>", strSSPblock[start:len(strSSPblock)])
        start = start + match.end()
        match = re.search("\d*[.,]\d+|\d+", strSSPblock[start:len(strSSPblock)])
        dstatic = strSSPblock[start+match.start():start+match.end()]
        strListSSP += "-"+str(dstatic)

        # get V_STATIC from ssp block
        match = re.search("V_STATIC", strSSPblock)
        start = match.end()
        match = re.search("\-\-\>", strSSPblock[start:len(strSSPblock)])
        start = start + match.end()
        match = re.search("Q_FRONT", strSSPblock[start:len(strSSPblock)])
        end = start + match.start()
        match = re.search("\d*[.,]\d+|\d+", strSSPblock[start:end])
        if match == None:
            vstatic = 0
        else:
            vstatic = strSSPblock[start+match.start():start+match.end()]
        if vstatic == '1':
            pepe = 0
        strListSSP += ":"+str(vstatic)                     
    
    return strListSSP

# AlGetPacket88FromPaketRBC(strPacket27, refNidlrbg)
# Alstom(JDRMDR) decoder
# Get data from MESSAGE RBC/PACKET 27 content
# Apply filter for JRU messages and RBC packets
# Returns a strings with a list of SSP in the format:
# nidbg0:dlink0, nidbg1:dlink1, ...
def AlGetPacket88FromPaketRBC(strPacket88, refNidlrbg):
   
    
                                  
    # Hasler struct
    # PACKET: 27. International Static Speed Profile
    #         NID_PACKET (8): 27. International Static Speed Profile
    #         Q_DIR (2): Both directions [2]
    #         L_PACKET (13): 114 bit(s)
    #         Q_SCALE (2): 1 m
    #         D_STATIC (15): 0,00 m [0]
    #         V_STATIC (7): 80 km/h [16]
    #         Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #         N_ITER (5): 0
    #         N_ITER (5): 2
    #             D_STATIC (15): 2.554,00 m [2554]
    #             V_STATIC (7): 40 km/h [8]
    #             Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #             N_ITER (5): 0
    #             D_STATIC (15): 624,00 m [624]
    #             V_STATIC (7): Non numerical value telling that the static speed profile description ends at D_STATIC(n) [127]
    #             Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #             N_ITER (5): 0

    # Alstom JDRMDR struct:
    # NID_PACKET : 88 -->Level crossing information
    #    PKT_TK2T_88 (
    #    Q_DIR : 0 -->Reverse
    #    L_PACKET : 72
    #    Q_SCALE : 1 -->1 m scale
    #    NID_LX : 138 -->Reserved for RBC transmission
    #    D_LX : 8223 -->8223 m
    #    L_LX : 170 -->170 m
    #    Q_LXSTATUS : 1 -->LX is not protected
    #    V_LX : 6 -->30km/h
    #    Q_STOPLX : 0 -->No stop required
    # )        

    # -----------------  Get list of distances and velocities  ----------------------

    # get Q_LXSTATUS from ssp block
    try:
        match = re.search("Q_LXSTATUS\s:", strPacket88)
        start = match.end()
        match = re.search("\d*[.,]\d+|\d+", strPacket88[start:len(strPacket88)])
        lxstatus = strPacket88[start+match.start():start+match.end()]
    except:
        print(color.Fore.RED +
              "Error in packet 88 from RBC - Q_LXSTATUS. LRBG = "+ refNidlrbg +
              color.Style.RESET_ALL)
        exit()
                           
    
    return lxstatus



# AlGetDataFromBalise(hasDf, listRbcFilter, listJruFilter)
# Alstom (JDRMDR) decoder
# Get DATA from Balise message
# (under development)
# NID_BG
# N_PIG
# Q_LINK

def AlGetDataFromBalise(alDf):

    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pd.DataFrame(alDf) # copia local del objeto del argumento

    dfDataBalise = pd.DataFrame()
    listNidBG = []
    listQlink = []
    
    length = inputdf["DATA"].size
    try: 
        for x in range(length):
            strDataBalise = inputdf["DATA"][x]

            # Hasler struct
            # Q_UPDOWN (1): Up link telegram [1]
            # M_VERSION (7): Class 1 [16]
            # Q_MEDIA (1): Balise [0]
            # N_PIG (3): First position [0]
            # N_TOTAL (3): 1 balise(s) in the group [0]
            # M_DUP (2): No duplicates [0]
            # M_MCOUNT (8): The telegram fits with all telegrams of the same balise group. [255]
            # NID_C (10): Reserved [391]
            # NID_BG (14): 1363
            # Q_LINK (1): Linked [1]

            # Alstom JDRMDR struct
            # Q_UPDOWN : 1-->Up link telegram
            # M_VERSION : 33-->Version 2.1, introduced in SRS 3.5.0
            # Q_MEDIA : 0-->Balise
            # N_PIG : 0-->I am the 1st
            # N_TOTAL : 1-->2 balises in the group
            # M_DUP : 0-->No duplicates
            # M_MCOUNT : 255-->The telegram fits with all telegrams of the same balise group)
            # NID_C : 383-->383
            # NID_BG : 103-->103
            # Q_LINK : 1-->Linked


            #get DATA string 
            if pd.isnull(strDataBalise):
                listNidBG.insert(x, "")
                listQlink.insert(x, "")
            else:

                # get N_PIG
                try:
                    match = re.search("N_PIG\s:", strDataBalise)
                    start = match.end()
                    match = re.search("\d+", strDataBalise[start:len(strDataBalise)])
                    npig = strDataBalise[start+match.start():start+match.end()]
                except:
                    npig = ""

                # get NID_BG
                try:
                    match = re.search("NID_BG\s:", strDataBalise)
                    start = match.end()
                    match = re.search("\d+", strDataBalise[start:len(strDataBalise)])
                    nidbg = strDataBalise[start+match.start():start+match.end()]
                except:
                    nidbg = ""
                if npig != "": nidbg = nidbg+"-"+npig
                # Insertamos valor en la lista
                listNidBG.insert(x, nidbg)

                # get Q_LINK
                try:
                    match = re.search("Q_LINK\s:", strDataBalise)
                    start = match.end()
                    match = re.search("\d+", strDataBalise[start:len(strDataBalise)])
                    qlink = strDataBalise[start+match.start():start+match.end()]
                except:
                    qlink = ""
                # Insertamos valor en la lista
                listQlink.insert(x, qlink)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    except:
        sys.stdout.write("\n")
        sys.stdout.flush() 

        print(color.Fore.YELLOW +
              "Error en Data from Balise. Record ID = "+ str(inputdf['Record Id'][x]) +
              color.Style.RESET_ALL)
        

        

    # Creamos variable train pos
    # Se ajustan tipos de los datos
    dfDataBalise["NID_BG"] = np.array(listNidBG)
    dfDataBalise["Q_LINK"] = np.array(listQlink)
    
    return dfDataBalise

# AlGetPacketFromBalise(hasDf)
# Alstom (JDRMDR) decoder
# Get Packets from Balise telegram
# (under development)
# PACKET(packet id)

def AlGetPacketFromBalise(alDf, dfTrainPos, dfDataBalise):

    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pd.DataFrame(alDf) # copia local del objeto del argumento


    dfPacketBalise = pd.DataFrame()
    listSS026Packets = []
    listDiscardedSS026Packets = []
    # For packet 5
    listLinksBG = [] 
    # For packet 27
    listSSP = []
    # For packet 12
    listMA = []

    # For packet 88
    listLX = []
    
    
    length = inputdf["PACKET"].size
    try: 
        for x in range(length):
            # Get message header
            # strTelegramHeader = inputdf["HEADER"][x]
            # Get message data
            strTelegramData = inputdf["DATA"][x]
            # Get packets struct
            strBalisePacket = inputdf["PACKET"][x]

            # Hasler struct
            # PACKET: 3. National Values
            # NID_PACKET (8): 3. National Values
            # Q_DIR (2): Reverse [0]
            # L_PACKET (13): 186 bit(s)
            # Q_SCALE (2): 1 m
            # ...
            # PACKET: 42. Session Management
            # NID_PACKET (8): 42. Session Management
            # Q_DIR (2): Reverse [0]
            # L_PACKET (13): 113 bit(s)
            # Q_RBC (1): Establish communication session [1]
            # NID_C (10): Reserved [391]
            # NID_RBC (14): 4
            # ...

            # Alstom JDRMDR struct
            # PACKET: 3
            # Q_DIR : 1 -->Nominal
            # L_PACKET : 240
            # Q_SCALE : 1 -->1 m scale
            # D_VALIDNV : 0 -->0 m
            # NID_C : 383 -->383
            # niter : 1
            # NID_C[0] : 383 -->383
            # V_NVSHUNT : 6 -->30km/h
            # V_NVSTFF : 20 -->100km/h
            # V_NVONSIGHT : 6 -->30km/h
            # V_NVLIMSUPERV : 100 -->500km/h
            # V_NVUNFIT : 28 -->140km/h
            # V_NVREL : 3 -->15km/h
            # D_NVROLL : 5 -->5 m
            # Q_NVSBTSMPERM : 1 -->Yes
            # Q_NVEMRRLS : 0 -->Revoke emergency brake command at standstill
            # Q_NVGUIPERM : 0 -->No
            # Q_NVSBFBPERM : 0 -->No
            # Q_NVINHSMICPERM : 0 -->No
            # V_NVALLOWOVTRP : 0 -->0km/h
            # V_NVSUPOVTRP : 6 -->30km/h
            # D_NVOVTRP : 80 -->80 m
            # T_NVOVTRP : 40 -->40s
            # D_NVPOTRP : 50 -->50 m
            # M_NVCONTACT : 1 -->Apply service brake
            # T_NVCONTACT : 20 -->20s
            # M_NVDERUN : 1 -->Yes
            # D_NVSTFF : 32767 -->infinite
            # Q_NVDRIVER_ADHES : 0 -->Not Allowed
            # A_NVMAXREDADH1 : 20 -->1 m/s2
            # A_NVMAXREDADH2 : 14 -->0,7 m/s2
            # A_NVMAXREDADH3 : 14 -->0,7 m/s2
            # Q_NVLOCACC : 12 -->12m
            # M_NVAVADH : 0 -->0
            # M_NVEBCL : 9 -->Confidence level = 99.9999999%
            # Q_NVKINT : 0 -->No integrated correction factors follow

             
            if pd.isnull(strBalisePacket) or (strBalisePacket == ''):
                listSS026Packets.insert(x, "")
                listDiscardedSS026Packets.insert(x, "")
                listLinksBG.insert(x, "")
                listSSP.insert(x,"")
                listMA.insert(x,"")
                listLX.insert(x,"")
            else:

                # get list of packets in PACKET
                # PACKET: 3. National Values
                # PACKET: 42. Session Management
                # PACKET: ...
                boolWhile = True
                boolPacket5 = False # Linking
                boolPacket27 = False # SSP
                boolPacket12 = False # MA
                boolPacket88 = False # LX
                boolAtleast1SS026Packet = False
                strCurrentNidLrbg = ''
                strNextNidLrbg = ''
                strCurrNidBg = ''
                qdirNidLrbg = ''

                # Después buscar en dfTrainPos (que hemos pasado por variable)
                # la siguiente nidlrbg y comparar con la baliza leída
                # Si la siguiente nidlrbg no coincide con la baliza leída, descartamos
                # Si coincide pero Q_DIRLRBG no coincide con Q_DIR y Q_DIR no es igual 2 (ambas direcciones) descartamos paquete


                # Check if packets in telegram are valids for train orientation in relation to direction of balise
                # We get different variables
                #   - Balise that sends telegram: strCurrNidBg
                #   - Q_DIR from each packet: qdirSS026packet
                #   - Q_DIRLRBG from train position related to strCurrNidBG. For that:
                #       - Search first JRU message where NID_LRBG is equal to strCurrNidBg
                #       - In that message we get  TRAIN_POSITION.Q_DIRLRBG, that is train orientation with respect the strCurrNidBg
                #       - In case that next LRBG (that is not unknown) is differet from strCurrNidBG, we discard packets
                #           as we don´t know the train orientation with respect the balise
                #       - Exception to previous point: when Q_DIR from packet is equal to 2 that means "both directions"

                # strCurrNidBg is the balise over which the train is currently positioned and whose packets are going to be analysed
                strCurrNidBg =  dfDataBalise["NID_BG"][x] # nidbg-x need to remove -x
                # remove -x            
                try:
                    iAux = strCurrNidBg.index('-')
                    strCurrNidBg = strCurrNidBg[0:iAux]
                except:
                    strCurrNidBg = ''

                
                strCurrentNidLrbg = dfTrainPos["NID_LRBG"][x]
                if '829' in strCurrentNidLrbg:
                    pepe = 0   
                
                for iNextBg in range(x, dfTrainPos["NID_LRBG"].size):
                    strTempNidlrg = dfTrainPos["NID_LRBG"][iNextBg]
                    if strTempNidlrg != strCurrentNidLrbg and strTempNidlrg != '16383': # 16383 is unknown balise
                        strNextNidLrbg = strTempNidlrg
                        qdirNidLrbg = str(int(dfTrainPos["Q_DIRLRBG"][iNextBg])) 
                        break
                    else:

                        # Case strCurrNidBg == strCurrentNidLrbg (Train has change its direction)
                        if strCurrNidBg == strCurrentNidLrbg:
                            strNextNidLrbg = strTempNidlrg
                            qdirNidLrbg = str(int(dfTrainPos["Q_DIRLRBG"][x])) 
                            break

                        if str(int(dfTrainPos["Q_DIRLRBG"][iNextBg])) != str(int(dfTrainPos["Q_DIRLRBG"][x])):
                            # Train has changed its direction (same lrbg - different orientation)
                            # Possible movement in SB (not detected balise)
                            strNextNidLrbg = strTempNidlrg
                            qdirNidLrbg = '2' # unknown (not used later)
                            break # We dont search any more

                
                # At this moment 
                # strCurrNidBg is the balise over which the train is currently positioned and whose packets are going to be analysed
                # strNextNidLrbg is next valid LRBG being established for the train inn next mesages 
                strSS026packet = ''
                strDiscardedSS026packet = ''    
                istart = 0
                while boolWhile == True: 
                    
                    idSS026packet = None
                    qdirSS026packet = ''
                    boolDiscardPacket = False
                
                    match = re.search("PACKET:", strBalisePacket[istart:len(strBalisePacket)])
                    if match == None:
                        #if strSS026packet is None: strSS026packet = ""
                        boolWhile = False
                        continue

                    boolAtleast1SS026Packet = True
                    start = istart + match.end()
                    match = re.search("\d+", strBalisePacket[start:len(strBalisePacket)])
                    idSS026packet = strBalisePacket[start+match.start():start+match.end()]
                    istart2 = start # for q_dir
                    istart = start # for next packet
                    
                    # Get Q_DIR packet to be compared to train orientation Q_DIRLRBG 
                    try:
                        match = re.search("Q_DIR\s: ", strBalisePacket[istart2:len(strBalisePacket)])
                        start2 = istart2 + match.end()
                        match = re.search("\d+", strBalisePacket[start2:len(strBalisePacket)])
                        qdirSS026packet = strBalisePacket[start2+match.start():start2+match.end()]    

                    except Exception as ex:

                        qdirSS026packet = ''
                        if idSS026packet != '255':
                            print(color.Fore.YELLOW +
                                "AlGetPacketFromBalise-Q_DIR not found in packet from balise. Record ID = "+ str(inputdf['Record Id'][x]) +
                                color.Style.RESET_ALL)
                                       
                    if boolAtleast1SS026Packet == False: idSS026packet = '-'

                    # Check if packet is discarded for train direction
                    if strCurrNidBg != strNextNidLrbg: # Discarded packet for train orientation

                        boolDiscardPacket = True

                    else:

                        if (qdirSS026packet != '2') and (qdirNidLrbg != qdirSS026packet):
                            
                            boolDiscardPacket = True

                    if boolDiscardPacket == False:
                    
                        if strSS026packet == '':
                            strSS026packet = str(idSS026packet)
                        else: 
                            strSS026packet += ", "+str(idSS026packet)

                        if idSS026packet == '5': 
                            boolPacket5 = True # identificamos que en este PACKET RBC viene un paquete 5 (linking)
                            
                        if idSS026packet == '27': 
                            boolPacket27 = True # identificamos que en este PACKET RBC viene un paquete 27 (SSP)
                            
                        if idSS026packet == '12': 
                            boolPacket12 = True # identificamos que en este PACKET RBC viene un paquete 12 (MA)

                        if idSS026packet == '88': 
                            boolPacket88 = True # identificamos que en este PACKET RBC viene un paquete 8 (LX)
                              
                    
                    else:
                        if strDiscardedSS026packet == '':
                            strDiscardedSS026packet = str(idSS026packet)
                        else: 
                            strDiscardedSS026packet += ", "+str(idSS026packet)

                
                # Insert values in lists
                listSS026Packets.insert(x, strSS026packet)
                listDiscardedSS026Packets.insert(x, strDiscardedSS026packet)
                              

                # ------------------   Get linking from Packet 5  (starts) --------------

                if boolPacket5 == False:
                    listLinksBG.insert(x, "")
                else:
                    # get NID_LRBG
                    match = re.search("NID_BG\s:", strTelegramData)
                    start = match.end()
                    match = re.search("\d+", strTelegramData[start:len(strTelegramData)])
                    nidlrbg = strTelegramData[start+match.start():start+match.end()]
                    
                    # get complete Packet 5 string (in the right orientation)
                    
                    try:
                        istart = 0
                        while True:
                            match = re.search("PACKET:\s5\n", strBalisePacket[istart:len(strBalisePacket)])
                            startP5 = istart + match.start()
                            if match == None:
                                print(color.Fore.RED +
                                        "AlGetPacketFromBalise-PACKET: 5 not found when expected in packet from balise. Record ID = "+ str(inputdf['Record Id'][x]) +
                                        color.Style.RESET_ALL)
                                exit(0)

                            start = istart + match.end()
                            istart2 = start # for q_dir
                            istart = start # for next packet
                            
                            # Get Q_DIR packet to be compared to train orientation Q_DIRLRBG 
                            try:
                                match = re.search("Q_DIR\s: ", strBalisePacket[istart2:len(strBalisePacket)])
                                start2 = istart2 + match.end()
                                match = re.search("\d+", strBalisePacket[start2:len(strBalisePacket)])
                                qdirPacket5 = strBalisePacket[start2+match.start():start2+match.end()]    

                            except Exception as ex:

                                print(color.Fore.RED +
                                        "AlGetPacketFromBalise-Q_DIR not found in packet 5 from balise. Record ID = "+ str(inputdf['Record Id'][x]) +
                                        color.Style.RESET_ALL)
                                exit(0)

                            if qdirPacket5 == qdirNidLrbg:
                                
                                nextStart = istart # Find end of packet 5
                                match = re.search("PACKET:", strBalisePacket[nextStart:len(strBalisePacket)])
                                if match != None:
                                    endP5 = nextStart + match.start()
                                else:
                                    endP5 = len(strBalisePacket)

                                break # found packet in right train orientation
                            
                    except:
                        print(color.Fore.RED +
                                    "AlGetPacketFromBalise-Packet 5 expected for train orientation not found. Record ID = "+ str(inputdf['Record Id'][x]) +
                                    color.Style.RESET_ALL)
                        exit(0)


                    strPacket5 = strBalisePacket[startP5:endP5]
                    strPacket5 = strPacket5.rstrip() # Clean trailing special chars
                    strPacket5 = strPacket5.lstrip() # Clean leading special chars
                

                    strListLink = AlGetPacket5FromBalise(strPacket5, nidlrbg)
                    listLinksBG.insert(x, strListLink)


                # ------------------   Get SSP from Packet 27  (starts) --------------

                if boolPacket27 == False:
                    listSSP.insert(x, "")
                else:
                    
                    
                    # get NID_LRBG
                    match = re.search("NID_BG\s:", strTelegramData)
                    start = match.end()
                    match = re.search("\d+", strTelegramData[start:len(strTelegramData)])
                    nidlrbg = strTelegramData[start+match.start():start+match.end()]
                    
                    # get complete Packet 27 string (in the right orientation)
                    
                    try:
                        istart = 0
                        while True:
                            match = re.search("PACKET:\s27\n", strBalisePacket[istart:len(strBalisePacket)])
                            startP27 = istart + match.start()
                            if match == None:
                                print(color.Fore.RED +
                                        "AlGetPacketFromBalise-PACKET: 27 not found when expected in packet from balise. Record ID = "+ str(inputdf['Record Id'][x]) +
                                        color.Style.RESET_ALL)
                                exit(0)

                            start = istart + match.end()
                            istart2 = start # for q_dir
                            istart = start # for next packet
                            
                            # Get Q_DIR packet to be compared to train orientation Q_DIRLRBG 
                            try:
                                match = re.search("Q_DIR\s: ", strBalisePacket[istart2:len(strBalisePacket)])
                                start2 = istart2 + match.end()
                                match = re.search("\d+", strBalisePacket[start2:len(strBalisePacket)])
                                qdirPacket27 = strBalisePacket[start2+match.start():start2+match.end()]    

                            except Exception as ex:

                                print(color.Fore.RED +
                                        "AlGetPacketFromBalise-Q_DIR not found in packet 27 from balise. Record ID = "+ str(inputdf['Record Id'][x]) +
                                        color.Style.RESET_ALL)
                                exit(0)

                            if qdirPacket27 == qdirNidLrbg:
                                
                                nextStart = istart # Find end of packet 27
                                match = re.search("PACKET:", strBalisePacket[nextStart:len(strBalisePacket)])
                                if match != None:
                                    endP27 = nextStart + match.start()
                                else:
                                    endP27 = len(strBalisePacket)

                                break # found packet in right train orientation
                            
                    except:
                        print(color.Fore.RED +
                                    "AlGetPacketFromBalise-Packet 27 expected for train orientation not found. Record ID = "+ str(inputdf['Record Id'][x]) +
                                    color.Style.RESET_ALL)
                        exit(0)
                    
                    
                    
                    
                    
                    strPacket27 = strBalisePacket[startP27:endP27]
                    strPacket27 = strPacket27.rstrip() # Clean trailing special chars
                    strPacket27 = strPacket27.lstrip() # Clean leading special chars

                    strListSSP = AlGetPacket27FromBalise(strPacket27, nidlrbg)
                    listSSP.insert(x, strListSSP)
                
                # ------------------   Get MA from Packet 12 (starts) --------------

                if boolPacket12 == False:
                        
                    listMA.insert(x, "")

                else:
                    # Set PACKET RBC = 3 to simulate an MA and set SoM in the first MA received
                    listMA.insert(x, '3')


                # ------------------   Get LX info from Packet 88 (level crossing) --------------

                if boolPacket88 == False:
                        
                    listLX.insert(x, "")

                else:

                    # get NID_LRBG
                    match = re.search("NID_BG\s:", strTelegramData)
                    start = match.end()
                    match = re.search("\d+", strTelegramData[start:len(strTelegramData)])
                    nidlrbg = strTelegramData[start+match.start():start+match.end()]
                    
                    # get complete Packet 88 string
                    
                    try:
                        istart = 0
                        while True:
                            match = re.search("PACKET:\s88\n", strBalisePacket[istart:len(strBalisePacket)])
                            startP88 = istart + match.start()
                            if match == None:
                                print(color.Fore.RED +
                                        "AlGetPacketFromBalise-PACKET: 88 not found when expected in packet from balise. Record ID = "+ str(inputdf['Record Id'][x]) +
                                        color.Style.RESET_ALL)
                                exit(0)

                            start = istart + match.end()
                            istart2 = start # for q_dir
                            istart = start # for next packet
                            
                            # Get Q_DIR packet to be compared to train orientation Q_DIRLRBG 
                            try:
                                match = re.search("Q_DIR\s: ", strBalisePacket[istart2:len(strBalisePacket)])
                                start2 = istart2 + match.end()
                                match = re.search("\d+", strBalisePacket[start2:len(strBalisePacket)])
                                qdirPacket88 = strBalisePacket[start2+match.start():start2+match.end()]    

                            except Exception as ex:

                                print(color.Fore.RED +
                                        "AlGetPacketFromBalise-Q_DIR not found in packet 88 from balise. Record ID = "+ str(inputdf['Record Id'][x]) +
                                        color.Style.RESET_ALL)
                                exit(0)

                            if qdirPacket88 == qdirNidLrbg:
                                
                                nextStart = istart # Find end of packet 27
                                match = re.search("PACKET:", strBalisePacket[nextStart:len(strBalisePacket)])
                                if match != None:
                                    endP88 = nextStart + match.start()
                                else:
                                    endP88 = len(strBalisePacket)

                                break # found packet in right train orientation
                            
                    except:
                        print(color.Fore.RED +
                                    "AlGetPacketFromBalise-Packet 27 expected for train orientation not found. Record ID = "+ str(inputdf['Record Id'][x]) +
                                    color.Style.RESET_ALL)
                        exit(0)


                    strPacket88 = strBalisePacket[startP88:endP88]
                    strPacket88 = strPacket88.rstrip() # Clean trailing special chars
                    strPacket88 = strPacket88.lstrip() # Clean leading special chars

                    lxstatus = AlGetPacket88FromBalise(strPacket88, strCurrNidBg)

                    # Set LX STATUS
                    listLX.insert(x, lxstatus)

                    
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()    
                
    except: 
        sys.stdout.write("\n")
        sys.stdout.flush()

        print(color.Fore.YELLOW +
              "AlGetPacketFromBalise-Error en Packet from Balise. Record ID = "+ str(inputdf['Record Id'][x]) +
              color.Style.RESET_ALL)
        

       

    # Creamos variables de paquetes RBC
    # Se ajustan tipos de los datos
    dfPacketBalise["Balise Packets"] = np.array(listSS026Packets)
    dfPacketBalise["Discarded Packets"] = np.array(listDiscardedSS026Packets)
    dfPacketBalise["BG_LINKS"] = np.array(listLinksBG, dtype=str)
    dfPacketBalise["SSP"] = np.array(listSSP, dtype=str)
    dfPacketBalise["PACKET_RBC_SIMUL"] = np.array(listMA, dtype=str)
    dfPacketBalise["LX"] = np.array(listLX, dtype=str)
    
    return dfPacketBalise


# AlGetPacket5FromBalise(strPacket5, refNidlrbg)
# Alstom (JDRMDR) decoder
# Get data from PACKET_RBC/PACKET 5 content
# Apply filter for JRU messages and RBC packets
# Returns a strings with a list of NID_BGs and D_LINKS inthe format:
# nidbg0:dlink0, nidbg1:dlink1, ...
def AlGetPacket5FromBalise(strPacket5, refNidlrbg):
   
    global glLocale 
    
    # Hasler struct
    # Get list of links in packet 5
    # PACKET: 5. Linking
    #         NID_PACKET (8): 5. Linking
    #         Q_DIR (2): Both directions [2]
    #         L_PACKET (13): 810 bit(s)
    #         Q_SCALE (2): 1 m
    #         D_LINK (15): 305,00 m [305]
    #         Q_NEWCOUNTRY (1): Same country / railway administration, no NID_C follows [0]
    #         NID_BG (14): 1338
    #         Q_LINKORIENTATION (1): The balise group is seen by the train in reverse direction [0]
    #         Q_LINKREACTION (2): No reaction [2]
    #         Q_LOCACC (6): 10 m
    #         N_ITER (5): 19
    #             D_LINK (15): 249,00 m [249]
    #             Q_NEWCOUNTRY (1): Same country / railway administration, no NID_C follows [0]
    #             NID_BG (14): 1336
    #             Q_LINKORIENTATION (1): The balise group is seen by the train in reverse direction [0]
    #             Q_LINKREACTION (2): No reaction [2]
    #             Q_LOCACC (6): 10 m / Q_LINKACC
    #               .... (iter)


    # Alstom JDRMDR Struct
    # PACKET: 5
    #   Q_DIR : 1 -->Nominal
    #   L_PACKET : 186
    #   Q_SCALE : 1 -->1 m scale
    #   D_LINK : 43 -->43 m
    #   Q_NEWCOUNTRY : 0 -->Same country / railway administration, no NID_C follows
    #   NID_BG : 715 -->715
    #   Q_LINKORIENTATION : 1 -->The balise group is seen by the train in nominal direction
    #   Q_LINKREACTION : 1 -->Apply service brake
    #   Q_LOCACC : 1 -->1m
    #   niter : 3
    #       D_LINK[0] : 667 -->667 m
    #       Q_NEWCOUNTRY[0] : 0 -->Same country / railway administration, no NID_C follows
    #       NID_BG[0] : 709 -->709
    #       Q_LINKORIENTATION[0] : 0 -->The balise group is seen by the train in reverse direction
    #       Q_LINKREACTION[0] : 2 -->No reaction
    #       Q_LOCACC[0] : 2 -->2m
    #       D_LINK[1] : 1880 -->1880 m
    #           ... 
                       

    boolWhilePck5 = True
    ilink = 0
    istart = 0
    startLinkBlock = 0
    startLinkHeader = 0
    strListLink = "" 
    # LRBG is set as first balise in listLinks
    strListLink += str(refNidlrbg) + ':0'  


    # Get link header from NID_PACKET to Q_LOCACC (first block)
    match = re.search("Q_DIR\s:", strPacket5[istart:len(strPacket5)])
                            
    startLinkHeader = istart + match.start() # Start of linking header                        
    start = istart + match.end()
    
    try:
        match = re.search("niter\s:", strPacket5[start:len(strPacket5)])
        end = start + match.start()
    except Exception as ex:
        # Case: no packet 5
        print(color.Fore.YELLOW + 'AlGetPacket5FromBalise-Exception: ' + str(ex) + color.Style.RESET_ALL)
        

    strLinkHeader = strPacket5[startLinkHeader:end].rstrip() # Clean trailing special chars
    strLinkHeader = strLinkHeader.lstrip() # Clean leading special chars

    
    # get Q_LOCACC from link header
    try:
        match = re.search("Q_LOCACC", strLinkHeader)
        start = match.end()
        match = re.search("\d+", strLinkHeader[start:len(strLinkHeader)])
    except: 
        try:
            # get Q_LINKACC from link header
            match = re.search("Q_LINKACC", strLinkHeader)
            start = match.end()
            match = re.search("\d+", strLinkHeader[start:len(strLinkHeader)])
        except Exception as ex:
            print(color.Fore.YELLOW + "Packet 5. Q_LOCACC/Q_LINKACC not found in header. " + str(ex) + color.Style.RESET_ALL)
            exit()

    qlocacc = 0 # To indicate no referenced qlocacc; strLinkHeader[start+match.start():start+match.end()]
    strListLink += ":" +str(qlocacc)

    istart = 0 # initialization to start with list of blocks
    while boolWhilePck5 == True: 
                                    
        # Obtenemos todo la estructura del enlace
        # desde D_LINK hasta Q_LOCACC
        match = re.search("D_LINK", strPacket5[istart:len(strPacket5)])
        if match == None:
            # There is no more link information
            boolWhilePck5 = False
            continue
                        
        startLinkBlock = istart + match.start() # Start of linking block
        ilink += 1 # To identify the first block which ends in N_ITER
                        
        start = istart + match.end()
        # If first link block
        if ilink == 1:
            match = re.search("niter\s:", strPacket5[start:len(strPacket5)])
            end = start + match.start()
        else: 
            try:
                match = re.search("D_LINK", strPacket5[start:len(strPacket5)])
                end = start + match.start()
            except:
                # Case: no more ss026 packets
                end = len(strPacket5)
        strLinkblock = strPacket5[startLinkBlock:end].rstrip() # Clean trailing special chars
        strLinkblock = strLinkblock.lstrip() # Clean leading special chars
                
        istart = start # start next link block
        # strlinkblock structure
        # First block:
        # D_LINK : 43 -->43 m
        #   Q_NEWCOUNTRY : 0 -->Same country / railway administration, no NID_C follows
        #   NID_BG : 715 -->715
        #   Q_LINKORIENTATION : 1 -->The balise group is seen by the train in nominal direction
        #   Q_LINKREACTION : 1 -->Apply service brake
        #   Q_LOCACC : 1 -->1m
        #
        # Iteractions:
        #       D_LINK[0] : 667 -->667 m
        #       Q_NEWCOUNTRY[0] : 0 -->Same country / railway administration, no NID_C follows
        #       NID_BG[0] : 709 -->709
        #       Q_LINKORIENTATION[0] : 0 -->The balise group is seen by the train in reverse direction
        #       Q_LINKREACTION[0] : 2 -->No reaction
        #       Q_LOCACC[0] : 2 -->2m

        # get NID_BG from link block
        match = re.search("NID_BG", strLinkblock)
        start = match.end()
        match = re.search(":", strLinkblock[start:len(strLinkblock)])
        start = start + match.end()
        match = re.search("\d+", strLinkblock[start:len(strLinkblock)])
        nidbg = strLinkblock[start+match.start():start+match.end()]
        strListLink += "-"+str(nidbg)

        # get D_LINK
        # D_LINK : 43 -->43 m
        match = re.search("D_LINK", strLinkblock)
        start = match.end()
        match = re.search("\-\-\>", strLinkblock[start:len(strLinkblock)])
        start = start + match.end()
        match = re.search("\d*[.,]\d+|\d+", strLinkblock[start:len(strLinkblock)])
        d_link = strLinkblock[start+match.start():start+match.end()]

        # Depending on locale variable we adapt string and separators to standard float xywz.00
        if glLocale == 'SPA': 
            d_link = d_link.replace(".", "")
            d_link = float(d_link.replace(",", "."))
        elif glLocale == 'ENG': d_link = float(d_link.replace(",", ""))
        else:
            print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
            exit(1)

        # Insertamos valor en la lista de D_LINK
        strListLink += ":"+str(d_link)

        # get Q_LOCACC from link block
        #   Q_LOCACC : 1 -->1m
        try:
            match = re.search("Q_LOCACC", strLinkblock)
            start = match.end()
            match = re.search("\-\-\>", strLinkblock[start:len(strLinkblock)])
            start = start + match.end()
            match = re.search("\d+", strLinkblock[start:len(strLinkblock)])
        except:
            # get Q_LINKACC from link header
            match = re.search("Q_LINKACC", strLinkblock)
            start = match.end()
            match = re.search("\-\-\>", strLinkblock[start:len(strLinkblock)])
            start = start + match.end()
            match = re.search("\d+", strLinkblock[start:len(strLinkblock)])

        qlocacc = strLinkblock[start+match.start():start+match.end()]
        strListLink += ":"+str(qlocacc)
        
                    
    return strListLink


# AlGetPacket27FromBalise(strPacket27, refNidlrbg)
# Alstom (JDRMDR) decoder
# Get data from PACKET 27 content
# Returns a strings with a list of SSP in the format:
# nidbg0:dlink0, nidbg1:dlink1, ...
def AlGetPacket27FromBalise(strPacket27, refNidlrbg):
   
    
                                  
    # Hasler struct
    # PACKET: 27. International Static Speed Profile
    #         NID_PACKET (8): 27. International Static Speed Profile
    #         Q_DIR (2): Both directions [2]
    #         L_PACKET (13): 114 bit(s)
    #         Q_SCALE (2): 1 m
    #         D_STATIC (15): 0,00 m [0]
    #         V_STATIC (7): 80 km/h [16]
    #         Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #         N_ITER (5): 0
    #         N_ITER (5): 2
    #             D_STATIC (15): 2.554,00 m [2554]
    #             V_STATIC (7): 40 km/h [8]
    #             Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #             N_ITER (5): 0
    #             D_STATIC (15): 624,00 m [624]
    #             V_STATIC (7): Non numerical value telling that the static speed profile description ends at D_STATIC(n) [127]
    #             Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #             N_ITER (5): 0

    # Alstom JDRMDR struct:
    # NID_PACKET : 27 -->International Static Speed Profile
    #    PKT_TK2T_27 (
    #        Q_DIR : 1 -->Nominal
    #        L_PACKET : 226
    #        Q_SCALE : 1 -->1 m scale
    #        D_STATIC : 0 -->0 m
    #        V_STATIC : 14 -->70km/h
    #        Q_FRONT : 0 -->Train length delay on validity end point of profile element
    #        niter : 0
    #        niter : 6
    #       D_STATIC[0] : 158 -->158 m
    #       V_STATIC[0] : 8 -->40km/h
    #       Q_FRONT[0] : 0 -->Train length delay on validity end point of profile element
    #       niter[0] : 0
    #       D_STATIC[1] : 500 -->500 m
    #       V_STATIC[1] : 12 -->60km/h
    #       Q_FRONT[1] : 0 -->Train length delay on validity end point of profile element
    #      niter[1] : 0
    #       D_STATIC[2] : 1100 -->1100 m
    #       V_STATIC[2] : 10 -->50km/h
    #       Q_FRONT[2] : 0 -->Train length delay on validity end point of profile element
    #       niter[2] : 0
    #       D_STATIC[3] : 400 -->400 m
    #       V_STATIC[3] : 12 -->60km/h
    #       Q_FRONT[3] : 0 -->Train length delay on validity end point of profile element
    #       niter[3] : 0
    #      D_STATIC[4] : 900 -->900 m
    #       V_STATIC[4] : 9 -->45km/h
    #       Q_FRONT[4] : 0 -->Train length delay on validity end point of profile element
    #       niter[4] : 0
    #       D_STATIC[5] : 100 -->100 m
    #       V_STATIC[5] : 10 -->50km/h
    #       Q_FRONT[5] : 0 -->Train length delay on validity end point of profile element
    #       niter[5] : 0
    #   ) 
                    

    # -----------------  Get list of distances and velocities  ----------------------
    boolWhilePck27 = True
    iSSP = 0
    istart = 0
    startSSPBlock = 0
    # LRBG is set as reference in list for dstatic
    strListSSP = str(refNidlrbg)
    while boolWhilePck27 == True: 
                                    
        # Get structure
        # from D_STATIC to last N_ITER (now we dont take into account V_DIFF & Q_DIFF)
        match = re.search("D_STATIC", strPacket27[istart:len(strPacket27)])
        if match == None:
            # There is no more link information
            boolWhilePck27 = False
            continue
        
        startSSPBlock = istart + match.start() # Start of SSP block
        
        start = istart + match.end() # After D_STATIC :
        
                        
        iSSP += 1 # To identify the first block which ends in N_ITER
                        
        # If first SSP block
        if iSSP == 1:
            # D_STATIC : 0 -->0 m
            match = re.search("\s:", strPacket27[start:len(strPacket27)])
            start = start + match.end()
            match = re.search("niter\s:", strPacket27[start:len(strPacket27)])
            # jumping of first N_ITER used for Q_DIFF
            start = start + match.end()
            match = re.search("niter\s:", strPacket27[start:len(strPacket27)])
            start = start + match.end()
            end = start

                        
        else: 
            try:
                # D_STATIC[0] : 158 -->158 m
                match = re.search("\[", strPacket27[start:len(strPacket27)])
                start = start + match.end()
                match = re.search("niter\[", strPacket27[start:len(strPacket27)])
                start = start + match.start()
                end = start
            except:
                # Case: no ss026 packets
                start = len(strPacket27)
                end = len(strPacket27)
                       
        strSSPblock = strPacket27[startSSPBlock:end].rstrip() # Clean trailing special chars
        strSSPblock = strSSPblock.lstrip() # Clean leading special chars
                
        istart = start # start next ssp block

                        
        # strSSPblock structure
        #   D_STATIC : 0 -->0 m
        #   V_STATIC : 14 -->70km/h
        #   Q_FRONT : 0 -->Train length delay on validity end point of profile element

        # get D_STATIC from ssp block
        match = re.search("D_STATIC", strSSPblock)
        start = match.end()
        match = re.search("\-\-\>", strSSPblock[start:len(strSSPblock)])
        start = start + match.end()
        match = re.search("\d*[.,]\d+|\d+", strSSPblock[start:len(strSSPblock)])
        dstatic = strSSPblock[start+match.start():start+match.end()]
        strListSSP += "-"+str(dstatic)

        # get V_STATIC from ssp block
        match = re.search("V_STATIC", strSSPblock)
        start = match.end()
        match = re.search("\-\-\>", strSSPblock[start:len(strSSPblock)])
        start = start + match.end()
        match = re.search("\d*[.,]\d+|\d+", strSSPblock[start:len(strSSPblock)])
        vstatic = strSSPblock[start+match.start():start+match.end()]
        strListSSP += ":"+str(vstatic)                   
    
    return strListSSP

# AlGetPacket27FromBalise(strPacket27, refNidlrbg)
# Alstom (JDRMDR) decoder
# Get data from PACKET 27 content
# Returns a strings with a list of SSP in the format:
# nidbg0:dlink0, nidbg1:dlink1, ...
def AlGetPacket88FromBalise(strPacket88, strCurrNidBg):
       
                                  
    # Hasler struct
    # PACKET: 27. International Static Speed Profile
    #         NID_PACKET (8): 27. International Static Speed Profile
    #         Q_DIR (2): Both directions [2]
    #         L_PACKET (13): 114 bit(s)
    #         Q_SCALE (2): 1 m
    #         D_STATIC (15): 0,00 m [0]
    #         V_STATIC (7): 80 km/h [16]
    #         Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #         N_ITER (5): 0
    #         N_ITER (5): 2
    #             D_STATIC (15): 2.554,00 m [2554]
    #             V_STATIC (7): 40 km/h [8]
    #             Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #             N_ITER (5): 0
    #             D_STATIC (15): 624,00 m [624]
    #             V_STATIC (7): Non numerical value telling that the static speed profile description ends at D_STATIC(n) [127]
    #             Q_FRONT (1): Train length delay on validity end point of profile element [0]
    #             N_ITER (5): 0

    # Alstom JDRMDR struct:
    # NID_PACKET : 88 -->Level crossing information
        # PKT_TK2T_88 (
        #    Q_DIR : 1 -->Nominal
        #    L_PACKET : 64
        #    Q_SCALE : 1 -->1 m scale
        #    NID_LX : 1 -->Reserved for non RBC transmission (balise, loop or radio infill)
        #    D_LX : 400 -->400 m
        #    L_LX : 40 -->40 m
        #    Q_LXSTATUS : 0 -->LX is protected
        #)

                    
    # get Q_LXSTATUS from ssp block
    try:
        match = re.search("Q_LXSTATUS\s:", strPacket88)
        start = match.end()
        match = re.search("\d*[.,]\d+|\d+", strPacket88[start:len(strPacket88)])
        lxstatus = strPacket88[start+match.start():start+match.end()]
    except:
        print(color.Fore.RED +
              "Error in packet 88 from Balise - Q_LXSTATUS. Balise = "+ strCurrNidBg +
              color.Style.RESET_ALL)
        exit()
                   
    
    return lxstatus



def AlGetVPermOld(alDf):

    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pd.DataFrame(alDf) # local copy
    
    dfMrsp = pd.DataFrame() # output df
    listVperm = [] # List of Vperm values
    
    length = len(inputdf.axes[0])
    try: 
        for x in range(length):
            strDataVperm = ''
            vmrsp = 0
            try:
                indexDlink = inputdf["DATA"][x].index('V_PERM :')
            except:
                listVperm.insert(x, None)
                continue
            
            strDataVperm = inputdf["DATA"][x]
            # Format:  V_PERM : 40-->40-->Speed = 40 km/h


            #get DATA string 
            if pd.isnull(strDataVperm):
                listVperm.insert(x, None)
            else:

                # get V_PERM
                try:
                    match = re.search("V_PERM\s:", strDataVperm)
                    start = match.end()
                    match = re.search("\n", strDataVperm[start:len(strDataVperm)])
                    end = start+match.end()
                    strDataVperm = strDataVperm[start:end]
                    start = 0
                    match = re.search("Speed\s=", strDataVperm[start:len(strDataVperm)])
                    start = start + match.end()
                    match = re.search("\d+", strDataVperm[start:len(strDataVperm)])
                    vmrsp = int(strDataVperm[start+match.start():start+match.end()])
                except Exception as ex:

                    match = re.search("1023", strDataVperm)
                    start = match.end()
                    vmrsp = '1023' # Unknown

                
                # Insert value into list
                listVperm.insert(x, vmrsp)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    except:
        sys.stdout.write("\n")
        sys.stdout.flush() 

        print(color.Fore.RED +
              "Error in Data from Balise-V_PERM. Record ID = "+ str(inputdf['Record Id'][x]) +
              color.Style.RESET_ALL)
        exit()
        

        

    dfMrsp["V_PERM"] = np.array(listVperm) # For compatibility with Hasler old versions
    
    return dfMrsp


# AlGetVPerm (alDf)
# Alstom (JDRMDR) decoder
# Input: dataframe with type-hasler decoder
# Input format: columns=['Record Id', 'JRU']
# return Service Brake Command State from  DATA column
# !!! Fix with baseline/release

def AlGetVPerm (alDf):

    global glFunctionNum

    dfVperm = pd.DataFrame() # output df
    listVperm = [] # List of Vperm values

    inputdf = pd.DataFrame(alDf) # copia local del objeto del argumento
    
    length = inputdf["HEADER"].size
    
    for x in range(length):
        
        strSpeedDistanceMonitorBlock = ''
        strSpeedDistanceMonitorInfo = ''
           
        # ------------------- Get NID MESSAGE
        
        intNidMessage = None
        
        try:

            intNidMessage = int(inputdf["MESSAGE ID"][x])

        except:

            print(color.Fore.RED + 
                "AlGetVPerm-ERROR. Found message with bad format in NID_MESSAGE : block. Message number: " + 
                str(inputdf["Record Id"][x]) + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 20: # SPEED AND DISTANCE MONITORING INFORMATION message

            strSpeedDistanceMonitorBlock = inputdf["DATA"][x]

            # SPEED AND DISTANCE MONITORING INFORMATION (
            #        M_SDMTYPE : 0-->Ceiling speed monitoring (CSM)
            #        M_SDMSUPSTAT : 0-->Normal Status
            #        V_PERM : 1023-->None
            #        V_SBI : 1023-->None
            #        V_TARGET : 1023-->None
            #        D_TARGET : 32767-->None
            #        V_RELEASE : 1023-->None
            #        M_TTI : 0-->None
            # )

            # Split content from BALISE GROUP ERROR
            strAux = strSpeedDistanceMonitorBlock
            iAux = 0
            ixLine = 0
            reading = False
            block = []

            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'SPEED AND DISTANCE MONITORING INFORMATION (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strSpeedDistanceMonitorInfo = "\n".join(block) + '\n'
                            reading = False
                            break
                        block.append('    ' + line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    if ":" not in line:
                        level += line.count("(")
                        level -= line.count(")")

                    if level == 0:
                        strSpeedDistanceMonitorInfo = "\n".join(block) + '\n'
                    else:
                        print(color.Fore.YELLOW + 
                            "AlGetVPerm-WARNING. Found message with bad format in SPEED AND DISTANCE MONITORING INFORMATION block. Message number: " + 
                            str(inputdf["Record Id"][x]) + 
                            color.Style.RESET_ALL)
                    break
            
            # ---------------- Get error from message   
                    
            strSDMData = strSpeedDistanceMonitorInfo

            try:

                # Format:  V_PERM : 40-->40-->Speed = 40 km/h
                # or   V_PERM : 1023-->None


                #get DATA string 
                if pd.isnull(strSDMData):
                    listVperm.insert(x, None)
                else:

                    # get V_PERM
                    try:
                        match = re.search("V_PERM\s:", strSDMData)
                        start = match.end()
                        match = re.search("\n", strSDMData[start:len(strSDMData)])
                        end = start+match.end()
                        strSDMData = strSDMData[start:end]
                        start = 0
                        match = re.search("Speed\s=", strSDMData[start:len(strSDMData)])
                        start = start + match.end()
                        match = re.search("\d+", strSDMData[start:len(strSDMData)])
                        vmrsp = int(strSDMData[start+match.start():start+match.end()])
                    except Exception as ex:

                        match = re.search("1023", strSDMData)
                        start = match.end()
                        vmrsp = 1023 # Unknown

                    # Insert value into list
                    listVperm.insert(x, vmrsp)
                
            except:

                print(color.Fore.YELLOW + 
                        "AlGetVPerm-WARNING. Found JRU message with bad format in V_PERM block. Message number: " + 
                        str(inputdf["Record Id"][x]) + 
                        color.Style.RESET_ALL)
                listVperm.insert(x, None)
                


        else:
            listVperm.append('')
            
        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfVperm = pd.DataFrame(np.array(listVperm), columns=['V_PERM'])
            
    return dfVperm



# AlgetBaliseGroupErrorFromData (hasDf)
# Alstom (JDRMDR) decoder
# Input: dataframe with type-hasler decoder
# Input format: columns=['Record Id', 'JRU']
# return BG Error from  DATA column

def AlgetBaliseGroupErrorFromData (hasDf):

    global glFunctionNum

    listNidErrorBg = []
    dfNidErrorBg = pd.DataFrame()
    listMessageErrorBg = []
    dfMessageErrorBg = pd.DataFrame()
    listReturn = []


    inputdf = pd.DataFrame(hasDf) # copia local del objeto del argumento
    
    length = inputdf["DATA"].size
    
    for x in range(length):
        
        strBaliseGroupError = ''
        strBaliseGroupErrorData = ''        
            
        # ------------------- Get NID MESSAGE
        
        intNidMessage = None
        
        try:

            intNidMessage = int(inputdf["MESSAGE ID"][x])

        except:

            print(color.Fore.RED + 
                "AlgetBaliseGroupErrorFromData-ERROR. Found message with bad format in NID_MESSAGE : block. Message number: " + 
                str(inputdf["Record Id"][x]) + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 12: # Balise group error message

            strBaliseGroupError = inputdf["DATA"][x]
            strNidErrorBg = ''
            strMessageError = ''

            # BALISE GROUP ERROR (
            # NID_C : 383-->383
            # NID_ERRORBG : 639-->639
            # M_ERROR : 0-->Balise group: linking consistency error
            # )

            # Split content from BALISE GROUP ERROR
            strAux = strBaliseGroupError
            iAux = 0
            ixLine = 0
            reading = False
            block = []

            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'BALISE GROUP ERROR (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strBaliseGroupErrorData = "\n".join(block) + '\n'
                            reading = False
                            break
                        block.append('    ' + line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    if ":" not in line:
                        level += line.count("(")
                        level -= line.count(")")

                    if level == 0:
                        strBaliseGroupErrorData = "\n".join(block) + '\n'
                    else:
                        print(color.Fore.YELLOW + 
                            "AlgetBaliseGroupErrorFromData-WARNING. Found message with bad format in BALISE GROUP ERROR block. Message number: " + 
                            str(inputdf["Record Id"][x]) + 
                            color.Style.RESET_ALL)
                    break
            
            # ---------------- Get error from message   
                    
            strBGEData = strBaliseGroupErrorData

            try:
                # get NID_ERRORBG
                # NID_ERRORBG : 639-->639
                match = re.search("NID_ERRORBG\s:", strBGEData)
                start = match.end()
                match = re.search("\-\-\>", strBGEData[start:len(strBGEData)])
                #start = start + match.end()
                end = start + match.start()
                #match = re.search("\d*[.,]\d+|\d+", strBGEData[start:len(strBGEData)])
                match = re.search("\d*[.,]\d+|\d+", strBGEData[start:end])
                strNidErrorBg = strBGEData[start+match.start():start+match.end()]

                # get M_ERROR
                # M_ERROR : 0-->Balise group: linking consistency error
                match = re.search("M_ERROR\s:", strBGEData)
                start = match.end()
                #match = re.search("\-\-\>", strBGEData[start:len(strBGEData)])
                #start = start + match.end()
                match = re.search("\n", strBGEData[start:len(strBGEData)])
                end = start+match.end() -1 # -1 to remove \n
                strMessageError = strBGEData[start:end] 

                listNidErrorBg.append(strNidErrorBg)
                listMessageErrorBg.append(strMessageError)
            except:

                print(color.Fore.YELLOW + 
                        "AlgetBaliseGroupErrorFromData-WARNING. Found JRU message with bad format in BALISE GROUP ERROR block. Message number: " + 
                        str(inputdf["Record Id"][x]) + 
                        color.Style.RESET_ALL)
                listNidErrorBg.append('')
                listMessageErrorBg.append('')


        else:
            listNidErrorBg.append('')
            listMessageErrorBg.append('')

        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfNidErrorBg = pd.DataFrame(np.array(listNidErrorBg), columns=['NID_ERRORBG'])
    dfMessageErrorBg = pd.DataFrame(np.array(listMessageErrorBg), columns=['M_ERRORBG'])
        
    listReturn = [dfNidErrorBg, dfMessageErrorBg]
            
    return listReturn

# AlgetServiceBrakeCommandedFromData (hasDf)
# Alstom (JDRMDR) decoder
# Input: dataframe with type-hasler decoder
# Input format: columns=['Record Id', 'JRU']
# return Service Brake Command State from  DATA column

def AlgetServiceBrakeCommandedFromData (hasDf):

    global glFunctionNum

    listServiceBrakeCommandState = []
    dfServiceBrakeCommandState = pd.DataFrame()


    inputdf = pd.DataFrame(hasDf) # copia local del objeto del argumento
    
    length = inputdf["DATA"].size
    
    for x in range(length):
        
        strServiceBrakeCommandState = ''
        strServiceBrakeCommandStateData = ''        
            
        # ------------------- Get NID MESSAGE
        
        intNidMessage = None
        
        try:

            intNidMessage = int(inputdf["MESSAGE ID"][x])

        except:

            print(color.Fore.RED + 
                "AlgetServiceBrakeCommandedFromData-ERROR. Found message with bad format in NID_MESSAGE : block. Message number: " + 
                str(inputdf["Record Id"][x]) + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 4: # Service Brake Command State message

            strServiceBrakeCommandState = inputdf["DATA"][x]

            # SERVICE BRAKE COMMAND STATE (
            #        M_BRAKE_COMMAND_STATE : 1-->Commanded
            # )

            # Split content from BALISE GROUP ERROR
            strAux = strServiceBrakeCommandState
            iAux = 0
            ixLine = 0
            reading = False
            block = []

            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'SERVICE BRAKE COMMAND STATE (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strServiceBrakeCommandStateData = "\n".join(block) + '\n'
                            reading = False
                            break
                        block.append('    ' + line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    if ":" not in line:
                        level += line.count("(")
                        level -= line.count(")")

                    if level == 0:
                        strServiceBrakeCommandStateData = "\n".join(block) + '\n'
                    else:
                        print(color.Fore.YELLOW + 
                            "AlgetServiceBrakeCommandedFromData-WARNING. Found message with bad format in SERVICE BRAKE COMMAND STATE block. Message number: " + 
                            str(inputdf["Record Id"][x]) + 
                            color.Style.RESET_ALL)
                    break
            
            # ---------------- Get error from message   
                    
            strSBCData = strServiceBrakeCommandStateData

            try:
                # get M_BRAKE_COMMAND_STATE
                # M_BRAKE_COMMAND_STATE : 1-->Commanded
                match = re.search("M_BRAKE_COMMAND_STATE\s:", strSBCData)
                start = match.end()
                #match = re.search("\-\-\>", strBGEData[start:len(strBGEData)])
                #start = start + match.end()
                #end = start + match.start()
                match = re.search("\d*[.,]\d+|\d+", strSBCData[start:len(strSBCData)])
                #match = re.search("\d*[.,]\d+|\d+", strBGEData[start:end])
                strBrakeCommandState = strSBCData[start+match.start():start+match.end()]

                listServiceBrakeCommandState.append(strBrakeCommandState)
                
            except:

                print(color.Fore.YELLOW + 
                        "AlgetServiceBrakeCommandedFromData-WARNING. Found JRU message with bad format in SERVICE BRAKE COMMAND STATE block. Message number: " + 
                        str(inputdf["Record Id"][x]) + 
                        color.Style.RESET_ALL)
                listServiceBrakeCommandState.append('')
                


        else:
            listServiceBrakeCommandState.append('')
            
        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfServiceBrakeCommandState = pd.DataFrame(np.array(listServiceBrakeCommandState), columns=['SBC_STATE'])
            
    return dfServiceBrakeCommandState


# AlgetEmergencyBrakeCommandedFromData (hasDf)
# Alstom (JDRMDR) decoder
# Input: dataframe with type-hasler decoder
# Input format: columns=['Record Id', 'JRU']
# return Emergency Brake Command State from  DATA column

def AlgetEmergencyBrakeCommandedFromData (hasDf):

    global glFunctionNum

    listEmergencyBrakeCommandState = []
    dfEmergencyBrakeCommandState = pd.DataFrame()


    inputdf = pd.DataFrame(hasDf) # copia local del objeto del argumento
    
    length = inputdf["DATA"].size
    
    for x in range(length):
        
        strEmergencyBrakeCommandState = ''
        strEmergencyBrakeCommandStateData = ''        
            
        # ------------------- Get NID MESSAGE
        
        intNidMessage = None
        
        try:

            intNidMessage = int(inputdf["MESSAGE ID"][x])

        except:

            print(color.Fore.RED + 
                "AlgetEmergencyBrakeCommandedFromData-ERROR. Found message with bad format in NID_MESSAGE : block. Message number: " + 
                str(inputdf["Record Id"][x]) + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 3: # Emergency Brake Command State message

            strEmergencyBrakeCommandState = inputdf["DATA"][x]

            # EMERGENCY BRAKE COMMAND STATE (
            #         M_BRAKE_COMMAND_STATE : 1-->Commanded
            # )

            # Split content from BALISE GROUP ERROR
            strAux = strEmergencyBrakeCommandState
            iAux = 0
            ixLine = 0
            reading = False
            block = []

            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'EMERGENCY BRAKE COMMAND STATE (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        if ":" not in line:
                            level += line.count("(")
                            level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strEmergencyBrakeCommandStateData = "\n".join(block) + '\n'
                            reading = False
                            break
                        block.append('    ' + line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    if ":" not in line:
                        level += line.count("(")
                        level -= line.count(")")

                    if level == 0:
                        strEmergencyBrakeCommandStateData = "\n".join(block) + '\n'
                    else:
                        print(color.Fore.YELLOW + 
                            "AlgetEmergencyBrakeCommandedFromData-WARNING. Found message with bad format in EMERGENCY BRAKE COMMAND STATE block. Message number: " + 
                            str(inputdf["Record Id"][x]) + 
                            color.Style.RESET_ALL)
                    break
            
            # ---------------- Get error from message   
                    
            strEBCData = strEmergencyBrakeCommandStateData

            try:
                # get M_BRAKE_COMMAND_STATE
                # M_BRAKE_COMMAND_STATE : 1-->Commanded
                match = re.search("M_BRAKE_COMMAND_STATE\s:", strEBCData)
                start = match.end()
                #match = re.search("\-\-\>", strBGEData[start:len(strBGEData)])
                #start = start + match.end()
                #end = start + match.start()
                match = re.search("\d*[.,]\d+|\d+", strEBCData[start:len(strEBCData)])
                #match = re.search("\d*[.,]\d+|\d+", strBGEData[start:end])
                strBrakeCommandState = strEBCData[start+match.start():start+match.end()]

                listEmergencyBrakeCommandState.append(strBrakeCommandState)
                
            except:

                print(color.Fore.YELLOW + 
                        "AlgetEmergencyBrakeCommandedFromData-WARNING. Found JRU message with bad format in EMERGENCY BRAKE COMMAND STATE block. Message number: " + 
                        str(inputdf["Record Id"][x]) + 
                        color.Style.RESET_ALL)
                listEmergencyBrakeCommandState.append('')
                


        else:
            listEmergencyBrakeCommandState.append('')
            
        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfEmergencyBrakeCommandState = pd.DataFrame(np.array(listEmergencyBrakeCommandState), columns=['EBC_STATE'])
            
    return dfEmergencyBrakeCommandState


# AlgetBaliseGroupErrorFromData (hasDf)
# Alstom (JDRMDR) decoder
# Input: dataframe with type-hasler decoder
# Input format: columns=['Record Id', 'JRU']
# return BG Error from  DATA column

def AlgetDmiSymbolStatusFromData (hasDf):

    global glFunctionNum

    # Symbols
    # Acknoledgment for Limited Supervision
    strMoSymbolAckLs = 'MO22'
    strMoSymbolAckSR = 'MO10'
    strMoSymbolAckTrip = 'MO05'
    listMoSymbolAck = [] 
    dfMoSymbolAck = pd.DataFrame()

    # Safe radio connection “Connection Up”
    strStSymRadioUp = 'ST03'
    listStSymbolRadioUp = [] 
    dfStSymbolRadioUp = pd.DataFrame()
    
    # Safe radio connection “Connection Lost/Set-Up failed”
    strStSymRadioDown = 'ST04'
    listStSymbolRadioDown = []
    dfStSymbolRadioDown = pd.DataFrame()

    # Radio Hole  (Track Condition in DMI)
    strTcRadioHole = 'TC12'
    listTcRadioHole = []
    dfTcRadioHole = pd.DataFrame()
    
    listReturn = []


    inputdf = pd.DataFrame(hasDf) # copia local del objeto del argumento
    
    length = inputdf["DATA"].size
    
    for x in range(length):
        
        strDmiSymbolStatus = ''
        strDmiSymbolStatusContent = ''        
            
        # ------------------- Get NID MESSAGE
        
        intNidMessage = None
        
        try:

            intNidMessage = int(inputdf["MESSAGE ID"][x])

        except:

            print(color.Fore.RED + 
                "AlgetDmiSymbolStatusFromData-ERROR. Found message with bad format in NID_MESSAGE : block. Message number: " + 
                str(inputdf["Record Id"][x]) + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 21: # DMI symbol status

            strDmiSymbolStatus = inputdf["DATA"][x]

            # DMI SYMBOL STATUS (
            #         DMI_SYMB_STATUS : 0|0|0|8|32|0|0|0|0|0|0|0|0|0-->MO13 | ST01
            # )

            # Check vaid Data
            strAux = strDmiSymbolStatus
            iAux = 0
            ixLine = 0
            reading = False
            block = []

            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'DMI SYMBOL STATUS (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        level += line.count("(")
                        level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strDmiSymbolStatusContent = "\n".join(block) + '\n'
                            reading = False
                            break
                        block.append('    ' + line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    level += line.count("(")
                    level -= line.count(")")

                    if level == 0:
                        strDmiSymbolStatusContent = "\n".join(block) + '\n'
                    else:
                        print(color.Fore.YELLOW + 
                            "AlgetDmiSymbolStatusFromData-WARNING. Found message with bad format in DMI SYMBOL STATUS block. Message number: " + 
                            str(inputdf["Record Id"][x]) + 
                            color.Style.RESET_ALL)
                    break
            
            # ---------------- # Search for symbol status   
                    
            strDSTData = strDmiSymbolStatusContent
            strDSTData.rstrip() # trim data

            # remove newline in case it exists in the middel
            ch = ['\n']
            for i in ch:
                strDSTData = strDSTData.replace(i, '')

            # DMI Ack Limited Supervision
            if strMoSymbolAckLs in strDSTData:
                listMoSymbolAck.append(strMoSymbolAckLs)
            # DMI Ack Staff Responsible
            elif strMoSymbolAckSR in strDSTData:
                listMoSymbolAck.append(strMoSymbolAckSR)
            # DMI Ack Trip
            elif strMoSymbolAckTrip in strDSTData:
                listMoSymbolAck.append(strMoSymbolAckTrip)
            else:
                listMoSymbolAck.append('')

            # Radio UP
            if strStSymRadioUp in strDSTData:
                listStSymbolRadioUp.append(strStSymRadioUp)
            else:
                listStSymbolRadioUp.append('')
            
            # Radio DOWN
            if strStSymRadioDown in strDSTData:
                listStSymbolRadioDown.append(strStSymRadioDown)
            else:
                listStSymbolRadioDown.append('')

            # Track Condiiton Radio Hole
            if strTcRadioHole in strDSTData:
                listTcRadioHole.append(strTcRadioHole)
            else:
                listTcRadioHole.append('')


        else:
            listMoSymbolAck.append('')
            listStSymbolRadioUp.append('')
            listStSymbolRadioDown.append('')
            listTcRadioHole.append('')


        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfMoSymbolAck = pd.DataFrame(np.array(listMoSymbolAck), columns=['MOSYMB'])
    dfStSymbolRadioUp = pd.DataFrame(np.array(listStSymbolRadioUp), columns=[strStSymRadioUp])
    dfStSymbolRadioDown = pd.DataFrame(np.array(listStSymbolRadioDown), columns=[strStSymRadioDown])
    dfTcRadioHole = pd.DataFrame(np.array(listTcRadioHole), columns=[strTcRadioHole])
        
    listReturn = [dfMoSymbolAck, dfStSymbolRadioUp, dfStSymbolRadioDown, dfTcRadioHole]
            
    return listReturn


# AlgetDriverActionFromData (hasDf)
# Alstom (JDRMDR) decoder
# Input: dataframe with type-hasler decoder
# Input format: columns=['Record Id', 'JRU']
# return BG Error from  DATA column

def AlgetDriverActionFromData (hasDf):

    global glFunctionNum

    # Symbols
    # Acknoledgment for Limited Supervision
    driverAckLs = '13'
    driverAckSR = '3'
    driverAckTrip = '2'
    driverStart = '19'
    driverOveride = '14'
    listDriverAction = [] 
    dfDriverAction = pd.DataFrame()

    
    listReturn = []


    inputdf = pd.DataFrame(hasDf) # copia local del objeto del argumento
    
    length = inputdf["DATA"].size
    
    for x in range(length):
        
        strDriverAtion = ''
        strDriverAtionContent = ''        
            
        # ------------------- Get NID MESSAGE
        
        intNidMessage = None
        
        try:

            intNidMessage = int(inputdf["MESSAGE ID"][x])

        except:

            print(color.Fore.RED + 
                "AlgetDriverActionFromData-ERROR. Found message with bad format in NID_MESSAGE : block. Message number: " + 
                str(inputdf["Record Id"][x]) + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 11: # Drivers action

            strDriverAtion = inputdf["DATA"][x]

            # DRIVER'S ACTION (
            #         M_DRIVERACTIONS : 13-->Ack of Limited supervision mode
            # )

            # Check vaid Data
            strAux = strDriverAtion
            iAux = 0
            ixLine = 0
            reading = False
            block = []

            while True:

                try:   
                    iAux = strAux.index('\n')
                    line = strAux[0:iAux]
                    line = line.rstrip("\n")
                    line = line.lstrip()
                    strAux = strAux[iAux+1:len(strAux)]

                    if 'DRIVER\'S ACTION (' in line:
                        reading = True
                        level = 1
                        
                    elif reading:

                        level += line.count("(")
                        level -= line.count(")")
                        #level += line.count("(")
                        #level -= line.count(")")
                        if level == 0:
                            strDriverAtionContent = "\n".join(block) + '\n'
                            reading = False
                            break
                        block.append('    ' + line)
                        
                    ixLine += (iAux + 1)
                    

                except:
                    line = strAux
                    level += line.count("(")
                    level -= line.count(")")

                    if level == 0:
                        strDriverAtionContent = "\n".join(block) + '\n'
                    else:
                        print(color.Fore.YELLOW + 
                            "AlgetDriverActionFromData-WARNING. Found message with bad format in DRIVER´s ACTION block. Message number: " + 
                            str(inputdf["Record Id"][x]) + 
                            color.Style.RESET_ALL)
                    break
            
            # ---------------- # Search for symbol status   
                    
            strDRACData = strDriverAtionContent
            strDRACData.rstrip() # trim data

            # remove newline in case it exists in the middel
            ch = ['\n']
            for i in ch:
                strDRACData = strDRACData.replace(i, '')

            try:
                # get M_DRIVERACTIONS
                # M_DRIVERACTIONS : 13-->Ack of Limited supervision mode
                match = re.search("M_DRIVERACTIONS\s:", strDRACData)
                start = match.end()
                #match = re.search("\-\-\>", strBGEData[start:len(strBGEData)])
                #start = start + match.end()
                #end = start + match.start()
                match = re.search("\d*[.,]\d+|\d+", strDRACData[start:len(strDRACData)])
                #match = re.search("\d*[.,]\d+|\d+", strBGEData[start:end])
                mDriverAction = strDRACData[start+match.start():start+match.end()]

                # Driver acknowledges LS mode
                if mDriverAction == driverAckLs:
                    listDriverAction.append(driverAckLs)
                # Driver acknowledges Staff Responsibke
                elif mDriverAction == driverAckSR:
                    listDriverAction.append(driverAckSR)
                # Driver acknowledges Trip
                elif mDriverAction == driverAckTrip:
                    listDriverAction.append(driverAckTrip)
                # Driver pushes start
                elif mDriverAction == driverStart:
                    listDriverAction.append(driverStart)
                # Driver override
                elif mDriverAction == driverOveride:
                    listDriverAction.append(driverOveride)                    
                else:
                    listDriverAction.append('')
            
                
            except:

                print(color.Fore.YELLOW + 
                        "AlgetDriverActionFromData-WARNING. Found JRU message with bad format in DRIVER'S ACTION block. Message number: " + 
                        str(inputdf["Record Id"][x]) + 
                        color.Style.RESET_ALL)
                listDriverAction.append('')




        else:
            listDriverAction.append('')

        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfDriverAction = pd.DataFrame(np.array(listDriverAction), columns=['DRACT']) # Driver Action
        
    return dfDriverAction
