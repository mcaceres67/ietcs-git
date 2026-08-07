import pandas
import numpy
import re
import sys
import colorama as color
import xlsxwriter as xls
import warnings

import src.config as cfg
import src.hascfg as hascfg

# Version control
# Author: Manuel Cáceres Marzal
# V1.0 Date 10/09/2024
#   First version
# V1.1 Date 22/09/2024
#   - Includes accuracy of distances measured on-board (SS41-5.3) 
#   - Takes into accounts different versions of hasler files with:
#       - Takes into account a locale configuration parameter to compile files in different languages and different number formats as per locales
#       - Either NID_ENGINE or NID_OPERATIONAL in hasler header column
#       - Either Q_LOCACC or Q_LINKACC in balise accuracy variable
#       - Includes data time homogeneization
#   - Includes new calculations as odometry accuracy evry 5.000 meters, correlation if odo_error
#   - Includes new charts related with new calculated values
#   - Includes a first set of statistics to be implemented
#   - Fix some bugs of previous version
#   - Improves using of record filters previous to analytics
#   - Open points: to fix bug with distance calculation when not linked MAs
# V1.2 Date: 20/10/2024
#   - Includes new error calculations: usafeerror and osafeerror
#   - Fix some bugs: miscalculation of calibrated distances because of error searching of lrbg_extended (use of fullmatch instead of contains)
#   - Fixed: to fix bug with distance calculation when not linked MAs
#   - Optimization of time procession for CHasComChangeDtypes function
#   - Fix chart errors when maxaxis is nan value
#   - Fixed: calculate distance when jumping balises in linkig (case boolLinkedError = missing balise)
#   - Include processing of missions with balises found more than once in same missions (mainly testing cases without SoM message)
#   - Resetting of distance when two MAs not linked (mainly in transitiosn from NS, pr SL)
#   - Due to previous point, calculation of total distance taking into account resetting in the middle of mission 
# V1.3 Date: 19/11/24
#   - To include Level 1 MAs, and packet 5 (linking) and 27 (SSP)
#   - Pending: review some cases for warning3: qlocacc not found (in case a unique register with this balise as LRBG) 
#   - Includes calibration process to estimate 
#       - Atenna distance to front-end
#       - Position error due to distance travelled (odometry or onboard error)
#       - Position error due to track components (dlink error, adherence problems, ...)
#       - Includes functions to search matching routes to compare runs
# V1.4 Date 12/12/24
#   - Fixed error for accumulated odo_err (last 5.000 meters) when reset of distance in the middle of the misssion (case of transition to NS and back to ERTMS)
#   - Fixed error: bad configuration of qlocacc value for balise in MA message header; using value of 0 in this case
#   - Fixed error: corrected formula for over/under reading slope
# 
# V 2.0 Date 28/12/25
#   - New version. Splitting code from compilation to allow multiple formats compilation
#   - Functions for JRU compilation are removed from this module (ietcs) and added to specific ones for each jru format
#
# V 2.1 Date: 17/11/25
#   - New formula for odometry error vs speed correlation. We add a new value for odometry error slope at each position. Name for these
#       new variables: SpotOrSlope (for over-reading) and SpotUrSlope (for under-reading)
#
# V 3.0 Date 06/02/2026
#   - New functions in other modules (Alstom)
#
# # V 3.1 Date 03/03/2026
#   - Pending:
#       1) Process packets 5, 27, 12 from balises in HasGetPacketFromBalise function
#       2) Merge BG_LINKS, SSP and MA from balises and RBC in HasCompile function
#       3) Get EB Command State and SB Command state
#       4) DMIACK and DRACK
#       5) Included in charts: Safe connection established/interrupted
#       6) Include discarded packets in reverse
#       7) Insert LX column (up to now not necessary as there are no LXs in lines with hasler decoder)
#
hasVersion = '3.1.1' # Temporary

# Global variables

# Compile global variables
glNumberofCompiledRegisters = 0

# Set of global variables to figure out progress
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

# Functions over original Hasler file
# naming: HasFunctionName


# HasCompile(hasDf, listRbcFilter, listJruFilter)
# Hasler decoder
# Compile Hasler original file
# Output file columns:
# NID_MESSAGE
# Grabar Id / Record Id
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


def HasCompile(file, listRbcFilter, listJruFilter, listJruDiscard):

    print("Loading Hasler File. " + file)
    try:
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")
            inputdf = pandas.read_excel(file,
            sheet_name = 1,
            header = 1,)
            # engine = 'openpyxl')
    except Exception as ex:
        print(color.Fore.YELLOW + "Open file: " + str(ex) + color.Style.RESET_ALL)

    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum
    global glNumberofCompiledRegisters

    glPh1Progress = 0.0
    glPh1Functions = 14
    glPh1NumberRows = 0
    glFunctionNum = 0

    print("Compiling File")
    glNumberofCompiledRegisters = len(inputdf.axes[0])

    inputdf = HasApplyRowFilter(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glPh1NumberRows = len(inputdf.axes[0])
    
    sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
    sys.stdout.flush()

    # DataFrames with partial data
    dfNidmessage = pandas.DataFrame()
    dfCommondata = pandas.DataFrame()
    dfTrainEngine = pandas.DataFrame()
    dfTrainpos = pandas.DataFrame()
    dfVeltrain = pandas.DataFrame()
    dfEtcslevel = pandas.DataFrame()
    dfEtcsmode = pandas.DataFrame()
    dfPacketrbc = pandas.DataFrame()
    dfDatabalise = pandas.DataFrame()
    dfPacketBalise = pandas.DataFrame()
    dfVperm = pandas.DataFrame()
    dfNidErrorBg = pandas.DataFrame()
    dfMessageErrorBg = pandas.DataFrame()
    dfEmergencyBrakeCommandState = pandas.DataFrame()

    

    # Data Frame for compiled data
    dfCompiledHasler = pandas.DataFrame()

    # Extract data from hasler JRU
    glFunctionNum = 0
    dfNidmessage = HasGetNidmesageFromHeader(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glFunctionNum += 1
    dfCommondata = HasGetCommonData(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glFunctionNum += 1
    dfTrainEngine = HasGetTrainEngineFromHeader(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glFunctionNum += 1
    dfTrainpos = HasGetTrainPosFromHeader(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glFunctionNum += 1
    dfVeltrain = HasGetVelocityFromHeader(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glFunctionNum += 1
    dfEtcslevel = HasGetetcsLevelFromHeader(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glFunctionNum += 1
    dfEtcsmode = HasGetetcsModeFromHeader(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glFunctionNum += 1
    dfPacketrbc = HasGetPacketRBC(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glFunctionNum += 1
    dfDatabalise = HasGetDataFromBalise(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glFunctionNum += 1
    dfPacketBalise = HasGetPacketFromBalise(inputdf, listRbcFilter, listJruFilter, listJruDiscard)
    glFunctionNum += 1
    dfVperm = HasGetVPerm(inputdf)
    sys.stdout.write("\n")
    sys.stdout.flush()

    # Get balise groupo error
    glFunctionNum += 1
    listDfReturn = HasGetBaliseGroupErrorFromData (inputdf)

    # listReturn = [dfNidErrorBg, dfMessageErrorBg] 
    #
    dfNidErrorBg = listDfReturn[0]
    dfMessageErrorBg = listDfReturn[1]

    glFunctionNum += 1
    dfServiceBrakeCommandState = HasGetServiceBrakeCommandedFromData(inputdf)
    glFunctionNum += 1
    dfEmergencyBrakeCommandState = HasGetEmergencyBrakeCommandedFromData(inputdf)

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

    try:
        dfCompiledHasler = pandas.concat(listDf, axis=1)
    except Exception as ex:
        print(color.Fore.YELLOW + str(ex) + color.Style.RESET_ALL)

    # Column order
    #cols = list(dfHasExtract.columns)
    #a, b = cols.index('BG_LINKS'), cols.index('Q_LINK')
    #cols[b], cols[a] = cols[a], cols[b]
    #dfHasExtract = dfHasExtract[cols]


    # shift column 'BG_LINKS' to last position 
    #last_column = dfCompiledHasler.pop('BG_LINKS') 
    #dfCompiledHasler.insert(len(dfCompiledHasler.columns), 'BG_LINKS', last_column) 

   
    return dfCompiledHasler


# HasGetTrainPosHeader(hasDf, listRbcFilter, listJruFilter)
# Hasler decoder
# Get train postion data from header, returns a re-indexed df
# Q_SCALE
# NID_LRBG
# D_LRBG 
# Q_DIRLRBG
# Q_DLRBG  
# L_DOUBTOVER 
# L_DOUBTUNDER  
# Accept two filters (list): PACKET_RBC ID & JRUMESSAGE ID
def HasGetTrainPosFromHeader(hasDf, listRbcFilter, listJruFilter, listJruDiscard):
    
    global glLocale
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro de mensajes RBC
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes balizas
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfTrainpos = pandas.DataFrame()
    listQscale = []
    listLrbg = []
    listDistlrbg = []
    listQdirlrbg = []
    listQdlrbg = []
    listLDoubtover = []
    listLDoubtunder = []

    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

    length = inputdf["HEADER"].size
    for x in range(length):
        strHead = inputdf["HEADER"][x]

        #get TRAIN POSITION
        #TRAIN_POSITION: 
        #Q_SCALE (2): 1 m
        #NID_LRBG: 
        #    NID_C (10): Reserved [391]
        #    NID_BG (14): 1361
        #D_LRBG (15): 52,00 m [52]
        #Q_DIRLRBG (2): Nominal [1]
        #Q_DLRBG (2): Nominal [1]
        #L_DOUBTOVER (15): 15,00 m [15]
        #L_DOUBTUNDER (15): 15,00 m [15]
        try:
            match = re.search("TRAIN_POSITION:", strHead)
            start = match.start() # type: ignore
            match = re.search("L_DOUBTUNDER\s\(15\):", strHead[start:len(strHead)])
            end = start + match.start() # type: ignore
            match = re.search("\n", strHead[start+match.start():len(strHead)])
            end += match.end() # type: ignore
            strTrainpos = strHead[start:end]
        except:
            strTrainpos = ''
        
        # get q_scale
        # Q_SCALE (2): 1 m
        # Qualifier for the distance/length scale
        # 0 - 10 cm scale
        # 1 - 1 m scale
        # 2 - 10 m scale
        # 3 - Spare
        try:
            match = re.search("Q_SCALE\s\(2\):", strTrainpos)
            start = match.start()  # type: ignore
            match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
            end = start + match.start() # type: ignore
            strQscale = strTrainpos[start:end]
            # get q_scale value
            match = re.search("Q_SCALE\s\(2\):", strQscale)
            start = match.end() # type: ignore
            match = re.search("\d+", strQscale[start:len(strQscale)])
            qscale = strQscale[start+match.start():start+match.end()] # type: ignore
        except:
            qscale = ''
        # Insertamos valor en la lista de Qscale
        listQscale.insert(x, qscale)

        # get LRBG string
        # NID_BG (14): 1361
        try:
            match = re.search("NID_BG\s\(14\):", strTrainpos)
            start = match.start() # type: ignore
            match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
            end = start + match.start() # type: ignore
            strLrbg = strTrainpos[start:end]
            # get LRBG ID
            match = re.search("NID_BG\s\(14\):", strLrbg)
            start = match.end() # type: ignore
            match = re.search("\d+", strLrbg[start:len(strLrbg)])
            idLrbg = strLrbg[start+match.start():start+match.end()] # type: ignore
        except:
            idLrbg = ''
        # Insertamos valor en la lista de Lrbg
        listLrbg.insert(x, idLrbg)

        # get LRBG distance
        # D_LRBG (15): 52,00 m [52]
        # Distance between the last relevant balise group and the estimated front end 
        #   of the train (the side of the active cab).

        try:
            match = re.search("D_LRBG\s\(15\):", strTrainpos)
            start = match.start() # type: ignore
            match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
            end = start + match.start() # type: ignore
            strDistlrbg = strTrainpos[start:end]
            # get D_LRBG
            match = re.search("D_LRBG\s\(15\):", strDistlrbg)
            start = match.end() # type: ignore
            match = re.search("\d*[.,]\d+|\d+", strDistlrbg[start:len(strDistlrbg)])
            D_lrbg = strDistlrbg[start+match.start():start+match.end()] # type: ignore
        except:
            D_lrbg = ''
        # Insertamos valor en la lista de D_LRBG
        # Previamente lo convertimos a float
        if glLocale == 'SPA': 
            D_lrbg = D_lrbg.replace('.','')
            D_lrbg = float(D_lrbg.replace(",", "."))
        elif glLocale == 'ENG': D_lrbg = float(D_lrbg.replace(",", ""))
        else:
            print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
            exit(1)
        listDistlrbg.insert(x, D_lrbg)

        # get Q_DIRLRBG
        # Q_DIRLRBG (2): Nominal [1]
        # Orientation of the train in relation to the direction of the LRBG
        # 0 - Reverse
        # 1 - Nominal
        # 2 - Unknown
        # 3 - Spare
        try:
            match = re.search("Q_DIRLRBG\s\(2\):", strTrainpos)
            start = match.start() # type: ignore
            match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
            end = start + match.start() # type: ignore
            strQdirlrbg = strTrainpos[start:end]
            # get orientation
            match = re.search("Q_DIRLRBG\s\(2\):", strQdirlrbg)
            start = match.end() # type: ignore
            match = re.search("\d+", strQdirlrbg[start:len(strQdirlrbg)])
            Q_dirlrbg = strQdirlrbg[start+match.start():start+match.end()] # type: ignore
        except:
            Q_dirlrbg = ''
        # Insertamos valor en la lista de Q_DIRLRBG
        listQdirlrbg.insert(x, Q_dirlrbg)
        
        # get Q_DLRBG
        # Q_DLRBG (2): Nominal [1]
        # Qualifier telling on which side of the LRBG the estimated front end is
        # 0 - Reverse
        # 1 - Nominal
        # 2 - Unknown
        # 3 - Spare

        try:
            match = re.search("Q_DLRBG\s\(2\):", strTrainpos)
            start = match.start() # type: ignore
            match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
            end = start + match.start() # type: ignore
            strQdlrbg = strTrainpos[start:end]
            # get orientation
            match = re.search("Q_DLRBG\s\(2\):", strQdlrbg)
            start = match.end() # type: ignore
            match = re.search("\d+", strQdlrbg[start:len(strQdlrbg)])
            Q_dlrbg = strQdlrbg[start+match.start():start+match.end()] # type: ignore
        except:
            Q_dlrbg = ''
        # Insertamos valor en la lista de Q_DIRLRBG
        listQdlrbg.insert(x, Q_dlrbg)


        # get L_DOUBTOVER
        #L_DOUBTOVER (15): 15,00 m [15]
        # The over-reading amount plus the Q_LOCACC of the LRBG
        try:
            match = re.search("L_DOUBTOVER\s\(15\):", strTrainpos)
            start = match.start() # type: ignore
            match = re.search("\n", strTrainpos[start:len(strTrainpos)])  
            end = start + match.start() # type: ignore
            strDoubtover = strTrainpos[start:end]
            # get D_LRBG
            match = re.search("L_DOUBTOVER\s\(15\):", strDoubtover)
            start = match.end() # type: ignore
            match = re.search("\d*[.,]\d+|\d+", strDoubtover[start:len(strDoubtover)])
            L_doubtover = strDoubtover[start+match.start():start+match.end()] # type: ignore
        except:
            L_doubtover = ''
        # Insertamos valor en la lista de D_LRBG
        # Previamente lo convertimos a float
        if glLocale == 'SPA': 
            L_doubtover = float(L_doubtover.replace(",", "."))
        elif glLocale == 'ENG': L_doubtover = float(L_doubtover.replace(",", ""))
        else:
            print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
            exit(1)
        listLDoubtover.insert(x, L_doubtover)

        # get L_DOUBTUNDER
        # L_DOUBTUNDER (15): 15,00 m [15]
        # The under-reading amount plus the Q_LOCACC of the LRBG

        try:
            match = re.search("L_DOUBTUNDER\s\(15\):", strTrainpos)
            start = match.start() # type: ignore
            match = re.search("\n", strTrainpos[start:len(strTrainpos)]) 
            end = start + match.start() # type: ignore
            strDoubtunder = strTrainpos[start:end]
            # get D_LRBG
            match = re.search("L_DOUBTUNDER\s\(15\):", strDoubtunder)
            start = match.end() # type: ignore
            match = re.search("\d*[.,]\d+|\d+", strDoubtunder[start:len(strDoubtunder)])
            L_doubtunder = strDoubtunder[start+match.start():start+match.end()] # type: ignore
        except:
            L_doubtunder = ''
        # Insertamos valor en la lista de D_LRBG
        # Previamente lo convertimos a float
        if glLocale == 'SPA': 
            L_doubtunder = float(L_doubtunder.replace(",", "."))
        elif glLocale == 'ENG': L_doubtunder = float(L_doubtunder.replace(",", ""))
        else:
            print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
            exit(1)
                
        listLDoubtunder.insert(x, L_doubtunder)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()



    # Creamos variable train pos
    # Se ajustan tipos de los datos
    dfTrainpos["Q_SCALE"] = numpy.array(listQscale, dtype=int)
    dfTrainpos["NID_LRBG"] = listLrbg # se mantiene como string
    dfTrainpos["D_LRBG"] = listDistlrbg # se ajustó a float al insertar en la lista 
    dfTrainpos["Q_DIRLRBG"] = numpy.array(listQdirlrbg, dtype = int)
    dfTrainpos["Q_DLRBG"] = numpy.array(listQdlrbg, dtype = int)
    dfTrainpos["L_DOUBTOVER"] = listLDoubtover # se ajustó a float al insertar en la lista
    dfTrainpos["L_DOUBTUNDER"] = listLDoubtunder # se ajustó a float al insertar en la lista

    return dfTrainpos


# HasGetTrainEngineHeader(hasDf, listRbcFilter, listJruFilter)
# Hasler decoder
# Get train engine id from header, returns a re-indexed df
# NID_ENGINE
# Accept two filters (list): PACKET_RBC ID & JRUMESSAGE ID
def HasGetTrainEngineFromHeader(hasDf, listRbcFilter, listJruFilter, listJruDiscard):

    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum
    
    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes balizas
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfTrainengine = pandas.DataFrame()
    listTrainengine = []
    
    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

    length = inputdf["HEADER"].size
    for x in range(length):
        strHead = inputdf["HEADER"][x]

        #get NID ENGINE
        
        try:
            #NID_ENGINE (24): 17232 
            match = re.search("NID_ENGINE\s\(24\):", strHead)
            start = match.start()
            match = re.search("\n", strHead[start:len(strHead)])
            end = start+match.end() 
            strTrainengine = strHead[start:end]
            # get nid_engine value
            match = re.search("NID_ENGINE\s\(24\):", strTrainengine)
            start = match.end()
            match = re.search("\d+", strTrainengine[start:len(strTrainengine)])
            idEngine = strTrainengine[start+match.start():start+match.end()]
        except:
            try: 
                # NID_OPERATIONAL (32): 00005142
                match = re.search("NID_OPERATIONAL\s\(32\):", strHead)
                start = match.start()
                match = re.search("\n", strHead[start:len(strHead)])
                end = start+match.end()
                strTrainengine = strHead[start:end]
                # get nid_engine value
                match = re.search("NID_OPERATIONAL\s\(32\):", strTrainengine)
                start = match.end()
                match = re.search("\d+", strTrainengine[start:len(strTrainengine)])
                idEngine = strTrainengine[start+match.start():start+match.end()]
            except Exception as ex:
                print(color.Fore.YELLOW + "HEADER. NID_ENGINE/NID_OPERATIONAL not found. " + str(ex) + color.Style.RESET_ALL)
                exit()
        
        
        # Insertamos valor en la lista de idEngine
        listTrainengine.insert(x, idEngine)
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

        
    # Creamos df train engine
    dfTrainengine["NID_ENGINE"] = listTrainengine
    
    return dfTrainengine


# HasGetVelocityFromHeader(hasDf, listRbcFilter)
# Hasler decoder
# Get train velocity from header, returns a re-indexed df
# V_TRAIN
# Accept two filters (list): PACKET_RBC ID & JRUMESSAGE ID
def HasGetVelocityFromHeader(hasDf, listRbcFilter, listJruFilter, listJruDiscard):
    
    global glLocale
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes balizas
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfTrainvel = pandas.DataFrame()
    listVeltrain = []
    
    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

    length = inputdf["HEADER"].size
    for x in range(length):
        strHead = inputdf["HEADER"][x]
        # get V_TRAIN
        #V_TRAIN (7): 0 km/h [0] 
            
        match = re.search("V_TRAIN\s\(7\):", strHead)
        start = match.start()
        match = re.search("\n", strHead[start:len(strHead)])
        end = start+match.end()
        strVeltrain = strHead[start:end]
        # get v_train value
        match = re.search("V_TRAIN\s\(7\):", strVeltrain)
        start = match.end()
        match = re.search("\d*[.,]\d+|\d+", strVeltrain[start:len(strVeltrain)])
        velTrain = strVeltrain[start+match.start():start+match.end()]
        # Insertamos valor en la lista de velocidades
        # Previamente lo convertimos a float
        if glLocale == 'SPA': 
            velTrain = float(velTrain.replace(",", "."))
        elif glLocale == 'ENG': velTrain = float(velTrain.replace(",", ""))
        else:
            print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
            exit(1)
        listVeltrain.insert(x, velTrain)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()



    # Creamos dataframe resultado  
    dfTrainvel["V_TRAIN"] = listVeltrain
    
    return dfTrainvel


# HasGetetcsLevelFromHeader(hasDf, listRbcFilter)
# Hasler decoder
# Get etcs level from header, returns a re-indexed df
# M_LEVEL
# Accept two filters (list): PACKET_RBC ID & JRUMESSAGE ID
def HasGetetcsLevelFromHeader(hasDf, listRbcFilter, listJruFilter, listJruDiscard):
    
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes balizas
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfEtcsLevel = pandas.DataFrame()
    listEtcsLevel = []
    
    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

    length = inputdf["HEADER"].size
    for x in range(length):
        strHead = inputdf["HEADER"][x]
        # get M_LEVEL
        # 0 - Level 0
        # 1 - Level NTC specified by NID_NTC
        # 2 - Level 1
        # 3 - Level 2
        # 4-7 - Spare
        # M_LEVEL (3): Level 2 [3]
            
        match = re.search("M_LEVEL\s\(3\):", strHead)
        start = match.start()
        match = re.search("\n", strHead[start:len(strHead)])
        end = start+match.end()
        strMlevel = strHead[start:end]
        # get m_level value (number)
        match = re.search("\[", strMlevel)
        start = match.end()
        match = re.search("\d+", strMlevel[start:len(strMlevel)])
        mLevel = strMlevel[start+match.start():start+match.end()]
        # Insertamos valor en la lista de velocidades
        listEtcsLevel.insert(x, mLevel)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()


    # Creamos dataframe resultado  
    dfEtcsLevel["M_LEVEL"] = numpy.array(listEtcsLevel, dtype=int)
    
    return dfEtcsLevel


# HasGetNidmesageFromHeader(hasDf, listRbcFilter)
# Hasler decoder
# Get message id from header, returns a re-indexed df
# NID_MESSAGE
# Accept two filters (list): PACKET_RBC ID & JRUMESSAGE ID
def HasGetNidmesageFromHeader(hasDf, listRbcFilter, listJruFilter, listJruDiscard):
    
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes balizas
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfNidmessage = pandas.DataFrame()
    listNidmessage = []
    
    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

    length = inputdf["HEADER"].size
    for x in range(length):
        strHead = inputdf["HEADER"][x]
        # get NID_MESSAGE
        # NID_MESSAGE (8): MESSAGE FROM BALISE [6]    
        match = re.search("NID_MESSAGE\s\(", strHead)
        start = match.start()
        match = re.search("\n", strHead[start:len(strHead)])
        end = start+match.end()
        strNidmessage = strHead[start:end]
        # get nid_messagel value (number)
        match = re.search(":\s", strNidmessage)
        start = match.end()
        match = re.search("\n", strNidmessage[start:len(strNidmessage)])
        nidMessage = strNidmessage[start:start+match.end()-1]
        # Insertamos valor en la lista de tipos de mensajes
        listNidmessage.insert(x, nidMessage)
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()
        
    # Creamos dataframe resultado  
    dfNidmessage["NID_MESSAGE"] = numpy.array(listNidmessage)

    return dfNidmessage


# HasGetetcsModeFromHeader(hasDf, listRbcFilter)
# Hasler decoder
# Get etcs mode from header, returns a re-indexed df
# M_MODE
# Accept two filters (list): PACKET_RBC ID & JRUMESSAGE ID
def HasGetetcsModeFromHeader(hasDf, listRbcFilter,listJruFilter, listJruDiscard):
    
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes balizas
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfEtcsMode = pandas.DataFrame()
    listEtcsMode = []
    
    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

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
        # M_MODE (4): Staff Responsible [2]
            
        match = re.search("M_MODE\s\(4\):", strHead)
        start = match.start()
        end = len(strHead) # Último campo de la celda
        strMmode = strHead[start:end]
        # get m_level value (number)
        match = re.search("\[", strMmode)
        start = match.end()
        match = re.search("\d+", strMmode[start:len(strMmode)])
        mMode = strMmode[start+match.start():start+match.end()]
        # Insertamos valor en la lista de modos
        listEtcsMode.insert(x, mMode)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()


    # Creamos dataframe resultado  
    dfEtcsMode["M_MODE"] = numpy.array(listEtcsMode, dtype=int)
    
    return dfEtcsMode


# HasGetCommondata(hasDf, listRbcFilter)
# Hasler decoder
# Get some message data, returns a re-indexed df
# Date
# Time
# Accept two filters (list): PACKET_RBC ID & JRUMESSAGE ID
def HasGetCommonData(hasDf, listRbcFilter, listJruFilter, listJruDiscard):
    
    global glLocale
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes balizas
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfCommondata = pandas.DataFrame()
    
    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

    # Ajustamos formato de datos
    # inputdf['Fecha'] = pandas.to_datetime(inputdf['Fecha'], format="%d/%m/%y").dt.date
    # inputdf['Hora'] = pandas.to_datetime(inputdf['Hora'], format="%H:%M:%S").dt.time
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



    glPh1Progress += 1/glPh1Functions*100
    sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
    sys.stdout.flush()
    
    return dfCommondata



# HasGetPacketRBC(hasDf, listRbcFilter, listJruFilter)
# Hasler decoder
# Get data from PACKET_RBC content
# (under development)
# PACKET_RBC (packet id)
# PACKET: 15. Level 2/3 Movement Authority
# 1 .. n packets con su diferente contenido
# Apply filter for JRU messages and RBC packets returns a re-indexed df with the filtered file
def HasGetPacketRBC(hasDf, listRbcFilter, listJruFilter, listJruDiscard):
    
    global glLocale
    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro de mensajes RBC
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes balizas
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfPacketrbc = pandas.DataFrame()
    listPacketrbc = []
    listDref = []
    listSS026Packets = []
    # For packet 5
    listLinksBG = [] 
    # For packet 27
    listSSP = []
    
    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

    length = inputdf["PACKET RBC"].size
    try: 
        for x in range(length):
            strPacketrbc = inputdf["PACKET RBC"][x]
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

            #get PACKET_RBC header 
            if pandas.isnull(strPacketrbc):
                listPacketrbc.insert(x, "")
                listDref.insert(x, "")
                listSS026Packets.insert(x, "")
                listLinksBG.insert(x, "")
                listSSP.insert(x,"")
                
            else:
                match = re.search("PACKET\sRBC:", strPacketrbc)
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
                match = re.search("PACKET\sRBC:", strPacketheader)
                start = match.end()
                match = re.search("\d+", strPacketheader[start:len(strPacketheader)])
                idPacketrbc = strPacketheader[start+match.start():start+match.end()]
                # Insertamos valor en la lista
                listPacketrbc.insert(x, idPacketrbc)

                # get D_REF (case packet 33 MA shifted)
                try:
                    match = re.search("D_REF\s\(16\):", strPacketheader)
                    start = match.end()
                    match = re.search("[+-]?\d*[.,]\d+|\d+", strPacketheader[start:len(strPacketheader)])
                    dref = strPacketheader[start+match.start():start+match.end()]
                    if glLocale == 'SPA': dref = float(dref.replace(",", "."))
                    elif glLocale == 'ENG': dref = float(dref.replace(",", ""))
                    else:
                        print(color.Fore.RED + 'Please define input file language' + color.Style.RESET_ALL)
                        exit(1)
                    
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
                
                boolAtleast1SS026Packet = False
                istart = 0
                startPckBlock = 0
                strSS026packet = None
                idSS026packet = None
                while boolWhile == True: 
                    
                    match = re.search("PACKET:", strPacketrbc[istart:len(strPacketrbc)])
                    if match == None:
                        if strSS026packet is None: strSS026packet = ""
                        boolWhile = False
                        continue
                    startPckBlock = match.start() # Start of packet bloc
                    boolAtleast1SS026Packet = True
                    start = istart + match.end()
                    match = re.search("\d+", strPacketrbc[start:len(strPacketrbc)])
                    idSS026packet = strPacketrbc[start+match.start():start+match.end()]
                    if idSS026packet == '5': 
                        boolPacket5 = True # identificamos que en este PACKET RBC viene un paquete 5 (linking)
                    if idSS026packet == '27': 
                        boolPacket27 = True # identificamos que en este PACKET RBC viene un paquete 27 (SSP)
                    istart = start
                    if boolAtleast1SS026Packet == False: idSS026packet = '-'

                    if strSS026packet is None:
                        strSS026packet = str(idSS026packet)
                    else: 
                        strSS026packet += ", "+str(idSS026packet)
                # Insertamos valor en la lista
                listSS026Packets.insert(x, strSS026packet)

                # ------------------   Get linking from Packet 5  (starts) --------------

                if boolPacket5 == False:
                    listLinksBG.insert(x, "")
                else:
                    # get NID_LRBG
                    match = re.search("NID_BG\s\(14\):", strPacketheader)
                    start = match.end()
                    match = re.search("\d+", strPacketheader[start:len(strPacketheader)])
                    nidlrbg = strPacketheader[start+match.start():start+match.end()]
                    
                    # get complete Packet 5 string
                    match = re.search("PACKET:\s5\.", strPacketrbc)
                    startP5 = match.start()
                    nextStart = match.end()
                    match = match = re.search("PACKET:", strPacketrbc[nextStart:len(strPacketrbc)])
                    if match != None:
                        endP5 = nextStart + match.start()
                    else:
                        endP5 = len(strPacketrbc)

                    strPacket5 = strPacketrbc[startP5:endP5]
                    strPacket5 = strPacket5.rstrip() # Clean trailing special chars
                    strPacket5 = strPacket5.lstrip() # Clean leading special chars
                

                    strListLink = HasGetPacket5FromPaketRBC(strPacket5, nidlrbg, dref)
                    listLinksBG.insert(x, strListLink)

                # ------------------   Get SSP from Packet 27  (starts) --------------

                if boolPacket27 == False:
                    listSSP.insert(x, "")
                else:
                    # get NID_LRBG
                    match = re.search("NID_BG\s\(14\):", strPacketheader)
                    start = match.end()
                    match = re.search("\d+", strPacketheader[start:len(strPacketheader)])
                    nidlrbg = strPacketheader[start+match.start():start+match.end()]
                    
                    # get complete Packet 27 string
                    match = re.search("PACKET:\s27\.", strPacketrbc)
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
                

                    strListSSP = HasGetPacket27FromPaketRBC(strPacket27, nidlrbg)
                    listSSP.insert(x, strListSSP)


        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    except: 
        sys.stdout.write("\n")
        sys.stdout.flush()

        print(color.Fore.YELLOW + 
              "Error en Packet RBC. Record ID = "+ str(inputdf['Record Id'][x]) +
              color.Style.RESET_ALL)

        

    # Creamos variables de paquetes RBC
    # Se ajustan tipos de los datos
    dfPacketrbc["PACKET_RBC"] = numpy.array(listPacketrbc)
    dfPacketrbc["D_REF"] = numpy.array(listDref) # Previously converted to float with "." as decimal separator
    dfPacketrbc["Rbc Packets"] = numpy.array(listSS026Packets)
    dfPacketrbc["BG_LINKS"] = numpy.array(listLinksBG, dtype=str)
    dfPacketrbc["SSP"] = numpy.array(listSSP, dtype=str)

    return dfPacketrbc

# HasGetPacket5FromPaketRBC(strPacket5, refNidlrbg, dref)
# Hasler decoder
# Get data from PACKET_RBC/PACKET 5 content
# Apply filter for JRU messages and RBC packets
# Returns a strings with a list of NID_BGs and D_LINKS inthe format:
# nidbg0:dlink0, nidbg1:dlink1, ...
def HasGetPacket5FromPaketRBC(strPacket5, refNidlrbg, dref):
   
    global glLocale 
    
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
                    

    boolWhilePck5 = True
    ilink = 0
    istart = 0
    startLinkBlock = 0
    startLinkHeader = 0
    strListLink = "" 
    # LRBG is set as first balise in listLinks
    strListLink += str(refNidlrbg) + ':0'  


    # Get link header from NID_PACKET to Q_LOCACC (first block)
    match = re.search("NID_PACKET\s\(8\)", strPacket5[istart:len(strPacket5)])
                            
    startLinkHeader = istart + match.start() # Start of linking header                        
    start = istart + match.end()
    
    try:
        match = re.search("N_ITER\s\(5\):", strPacket5[start:len(strPacket5)])
        end = start + match.start()
    except Exception as ex:
        # Case: no packet 5
        print(color.Fore.YELLOW + 'HasGetPacket5FromPaketRBC-Exception: ' + str(ex) + color.Style.RESET_ALL)
        

    strLinkHeader = strPacket5[startLinkHeader:end].rstrip() # Clean trailing special chars
    strLinkHeader = strLinkHeader.lstrip() # Clean leading special chars

    
    # get Q_LOCACC from link header
    try:
        match = re.search("Q_LOCACC\s\(6\):", strLinkHeader)
        start = match.end()
        match = re.search("\d+", strLinkHeader[start:len(strLinkHeader)])
    except: 
        try:
            # get Q_LINKACC from link header
            match = re.search("Q_LINKACC\s\(6\):", strLinkHeader)
            start = match.end()
            match = re.search("\d+", strLinkHeader[start:len(strLinkHeader)])
        except Exception as ex:
            print(color.Fore.YELLOW + "Packet 5. Q_LOCACC/Q_LINKACC not found. " + str(ex) + color.Style.RESET_ALL)
            exit()

    qlocacc = 0 # To indicate no referenced qlocacc; strLinkHeader[start+match.start():start+match.end()]
    strListLink += ":" +str(qlocacc)

    istart = 0 # initialization to start with list of blocks
    while boolWhilePck5 == True: 
                                    
        # Obtenemos todo la estructura del enlace
        # desde D_LINK hasta Q_LOCACC
        match = re.search("D_LINK\s\(15\):", strPacket5[istart:len(strPacket5)])
        if match == None:
            # There is no more link information
            boolWhilePck5 = False
            continue
                        
        startLinkBlock = istart + match.start() # Start of linking block
        ilink += 1 # To identify the first block which ends in N_ITER
                        
        start = istart + match.end()
        # If first link block
        if ilink == 1:
            match = re.search("N_ITER\s\(5\):", strPacket5[start:len(strPacket5)])
            end = start + match.start()
        else: 
            try:
                match = re.search("D_LINK\s\(15\):", strPacket5[start:len(strPacket5)])
                end = start + match.start()
            except:
                # Case: no more ss026 packets
                end = len(strPacket5)
        strLinkblock = strPacket5[startLinkBlock:end].rstrip() # Clean trailing special chars
        strLinkblock = strLinkblock.lstrip() # Clean leading special chars
                
        istart = start # start next link block
        # strlinkblock structure
        # D_LINK (15): 305,00 m [305]
        # Q_NEWCOUNTRY (1): Same country / railway administration, no NID_C follows [0]
        # NID_BG (14): 1338
        # Q_LINKORIENTATION (1): The balise group is seen by the train in reverse direction [0]
        # Q_LINKREACTION (2): No reaction [2]
        # Q_LOCACC (6): 10 m

        # get NID_BG from link block
        match = re.search("NID_BG\s\(14\):", strLinkblock)
        start = match.end()
        match = re.search("\d+", strLinkblock[start:len(strLinkblock)])
        nidbg = strLinkblock[start+match.start():start+match.end()]
        strListLink += "-"+str(nidbg)


        # get D_LINK
        match = re.search("D_LINK\s\(15\):", strLinkblock)
        start = match.end()
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

        # If first link block
        if ilink == 1 and dref != "":
            # increase shifted distance
            d_link += dref

        # Insertamos valor en la lista de D_LINK
        strListLink += ":"+str(d_link)

        # get Q_LOCACC from link block
        try:
            match = re.search("Q_LOCACC\s\(6\):", strLinkblock)
            start = match.end()
            match = re.search("\d+", strLinkblock[start:len(strLinkblock)])
        except:
            # get Q_LINKACC from link header
            match = re.search("Q_LINKACC\s\(6\):", strLinkblock)
            start = match.end()
            match = re.search("\d+", strLinkblock[start:len(strLinkblock)])

        qlocacc = strLinkblock[start+match.start():start+match.end()]
        strListLink += ":"+str(qlocacc)
        
                    
    return strListLink

# HasGetPacket27FromPaketRBC(strPacket27, refNidlrbg)
# Hasler decoder
# Get data from PACKET_RBC/PACKET 27 content
# Apply filter for JRU messages and RBC packets
# Returns a strings with a list of SSP in the format:
# nidbg0:dlink0, nidbg1:dlink1, ...
def HasGetPacket27FromPaketRBC(strPacket27, refNidlrbg):
   
    
                                  
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
        match = re.search("D_STATIC\s\(15\):", strPacket27[istart:len(strPacket27)])
        if match == None:
            # There is no more link information
            boolWhilePck27 = False
            continue
                        
        startSSPBlock = istart + match.start() # Start of SSP block
        iSSP += 1 # To identify the first block which ends in N_ITER
                        
        start = istart + match.end()
        # If first SSP block
        if iSSP == 1:
            match = re.search("N_ITER\s\(5\):", strPacket27[start:len(strPacket27)])
            end = start + match.start()
            # jumping of first N_ITER used for Q_DIFF
            start = end
            match = re.search("N_ITER\s\(5\):", strPacket27[start:len(strPacket27)])
                        
        else: 
            try:
                match = re.search("D_STATIC\s\(15\):", strPacket27[start:len(strPacket27)])
                end = start + match.start()
            except:
                # Case: no ss026 packets
                end = len(strPacket27)
                       
        strSSPblock = strPacket27[startSSPBlock:end].rstrip() # Clean trailing special chars
        strSSPblock = strSSPblock.lstrip() # Clean leading special chars
                
        istart = start # start next ssp block

                        
        # strSSPblock structure
        #     D_STATIC (15): 0,00 m [0]
        #     V_STATIC (7): 80 km/h [16]
        #     Q_FRONT (1): Train length delay on validity end point of profile element [0]

        # get D_STATIC from ssp block
        match = re.search("D_STATIC\s\(15\):", strSSPblock)
        start = match.end()
        match = re.search("\d*[.,]\d+|\d+", strSSPblock[start:len(strSSPblock)])
        dstatic = strSSPblock[start+match.start():start+match.end()]
        strListSSP += "-"+str(dstatic)

        # get V_STATIC from ssp block
        match = re.search("V_STATIC\s\(7\):", strSSPblock)
        start = match.end()
        match = re.search("\d*[.,]\d+|\d+", strSSPblock[start:len(strSSPblock)])
        vstatic = strSSPblock[start+match.start():start+match.end()]
        strListSSP += ":"+str(vstatic)                   
    
    return strListSSP


# HasGetDataFromBalise(hasDf, listRbcFilter, listJruFilter)
# Hasler decoder
# Get DATA from Balise message
# (under development)
# NID_BG
# N_PIG
# Q_LINK
# Apply filter for JRU messages and RBC packets returns a re-indexed df with the filtered file
def HasGetDataFromBalise(hasDf, listRbcFilter, listJruFilter, listJruDiscard):

    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro de mensajes RBC
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes balizas
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfDataBalise = pandas.DataFrame()
    listNidBG = []
    listQlink = []
    
    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

    length = inputdf["DATA"].size
    try: 
        for x in range(length):
            strDataBalise = inputdf["DATA"][x]
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

            #get DATA string 
            if pandas.isnull(strDataBalise):
                listNidBG.insert(x, "")
                listQlink.insert(x, "")
            else:

                # get N_PIG
                try:
                    match = re.search("N_PIG\s\(3\):", strDataBalise)
                    start = match.end()
                    match = re.search("\d+", strDataBalise[start:len(strDataBalise)])
                    npig = strDataBalise[start+match.start():start+match.end()]
                except:
                    npig = ""

                # get NID_BG
                try:
                    match = re.search("NID_BG\s\(14\):", strDataBalise)
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
                    match = re.search("Q_LINK\s\(1\):", strDataBalise)
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
    dfDataBalise["NID_BG"] = numpy.array(listNidBG)
    dfDataBalise["Q_LINK"] = numpy.array(listQlink)
    
    return dfDataBalise


# HasGetPacketFromBalise(hasDf, listRbcFilter, listJruFilter)
# Hasler decoder
# Get DATA from Balise message
# (under development)
# PACKET_RBC (packet id)
# Apply filter for JRU messages and RBC packets returns a re-indexed df with the filtered file
def HasGetPacketFromBalise(hasDf, listRbcFilter, listJruFilter, listJruDiscard):

    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro de mensajes RBC
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes balizas
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfPacketBalise = pandas.DataFrame()
    listBalisePacket = []
    
    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

    length = inputdf["PACKET"].size
    try: 
        for x in range(length):
            strBalisePacket = inputdf["PACKET"][x]
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
             
            if pandas.isnull(strBalisePacket):
                listBalisePacket.insert(x, "")
            else:
                # get list of packets in PACKET
                # PACKET: 3. National Values
                # PACKET: 42. Session Management
                # PACKET: ...
                boolWhile = True
                boolAtleast1SS026Packet = False
                istart = 0
                strSS026packet = None
                idSS026packet = None
                while boolWhile == True: 
                    
                    match = re.search("PACKET:", strBalisePacket[istart:len(strBalisePacket)])
                    if match == None:
                        if strSS026packet is None: strSS026packet = ""
                        boolWhile = False
                        continue

                    boolAtleast1SS026Packet = True
                    start = istart + match.end()
                    match = re.search("\d+", strBalisePacket[start:len(strBalisePacket)])
                    idSS026packet = strBalisePacket[start+match.start():start+match.end()]
                    istart = start
                    if boolAtleast1SS026Packet == False: idSS026packet = '-'

                    if strSS026packet is None:
                        strSS026packet = str(idSS026packet)
                    else: 
                        strSS026packet += ", "+str(idSS026packet)
                # Insertamos valor en la lista
                listBalisePacket.insert(x, strSS026packet)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()    
                
    except: 
        sys.stdout.write("\n")
        sys.stdout.flush()

        print(color.Fore.YELLOW +
              "Error en Packet from Balise. Record ID = "+ str(inputdf['Record Id'][x]) +
              color.Style.RESET_ALL)
        

       

    # Creamos variable train pos
    # Se ajustan tipos de los datos
    dfPacketBalise["Balise Packets"] = numpy.array(listBalisePacket)
    
    return dfPacketBalise


def HasGetVMRSP_old(hasDf, listRbcFilter, listJruFilter, listJruDiscard):

    global glPh1Progress
    global glPh1Functions
    global glPh1NumberRows
    global glFunctionNum

    inputdf = pandas.DataFrame(hasDf) # local copy
    listRbcPackets = numpy.array(listRbcFilter) # local copy
    listJruMessages = numpy.array(listJruFilter) # local copy
    listJruDiscardMessages = numpy.array(listJruDiscard)

    dfMrsp = pandas.DataFrame() # output df
    listMrsp = [] # List of Mrsp values
    
    # filtramos df por tipos de paquetes solicitados
    # inputdf = HasApplyRowFilter(inputdf, listRbcPackets, listJruMessages, listJruDiscardMessages)

    length = len(inputdf.axes[0])
    try: 
        for x in range(length):
            strDataMrsp = ''
            vmrsp = 0
            try:
                indexDlink = inputdf["HEADER"][x].index('MOST RESTRICTIVE SPEED PROFILE [20]')
            except:
                listMrsp.insert(x, None)
                continue
            
            strDataMrsp = inputdf["DATA"][x]
            # Format: V_MRSP (7): 30 km/h [6]


            #get DATA string 
            if pandas.isnull(strDataMrsp):
                listMrsp.insert(x, None)
            else:

                # get V_MRSP
                try:
                    match = re.search("V_MRSP\s\(7\):", strDataMrsp)
                    start = match.end()
                    match = re.search("\d+", strDataMrsp[start:len(strDataMrsp)])
                    vmrsp = int(strDataMrsp[start+match.start():start+match.end()])
                except:
                    vmrsp = 0

                
                # Insert value into list
                listMrsp.insert(x, vmrsp)

        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    except:
        sys.stdout.write("\n")
        sys.stdout.flush() 

        print(color.Fore.YELLOW +
              "Error en Data from Balise. Record ID = "+ str(inputdf['Record Id'][x]) +
              color.Style.RESET_ALL)
        

        

    dfMrsp["V_MRSP"] = numpy.array(listMrsp)
    
    return dfMrsp
      

# HasGetVPerm (HasDf)
# Hasler decoder
# Input: dataframe with type-hasler decoder
# Input format: columns=['Record Id', 'JRU']
# return Service Brake Command State from  DATA column
# !!! Fix with baseline/release

def HasGetVPerm (HasDf):

    global glFunctionNum

    dfVperm = pandas.DataFrame() # output df
    listVperm = [] # List of Vperm values

    inputdf = pandas.DataFrame(HasDf) # copia local del objeto del argumento
    
    length = inputdf["HEADER"].size
    
    for x in range(length):
        
        strPermitedSpeedBlock = ''
        strPermitedSpeedInfo = ''
           
        # ------------------- Get NID MESSAGE
        
        intNidMessage = None
        
        try:

            intNidMessage = int(inputdf["MESSAGE ID"][x])

        except:

            if cfg.glLocale == 'ENG': strRecord = str(inputdf["Record Id"][x])
            elif cfg.glLocale == 'SPA': strRecord = str(inputdf["Grabar Id"][x])
            else: strRecord = ''
            print(color.Fore.RED + 
                "HasGetVPerm-ERROR. Found message with bad format in NID_MESSAGE : block. Message number: " + 
                strRecord + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 27: # PERMITTED SPEED message

            strPermitedSpeedBlock = inputdf["DATA"][x]

            

            strPermitedSpeedInfo = strPermitedSpeedBlock
            
            # ---------------- Get error from message   
                    
            strPSData = strPermitedSpeedInfo

            try:

                # !!! Fix with baseline/release  !!!!!!!!!!!
                # Format:  V_PERSPEED (7): 5 km/h [1]
                # Format 2: V_PERMITTED (7): 150 km/h [30]
                boolV_PERSPEED = False
                boolV_PERMITTED = False


                #get DATA string 
                if pandas.isnull(strPSData):
                    listVperm.insert(x, None)
                else:


                    # get V_PERMITTED
                    try:
                        match = re.search("V_PERMITTED\s\(7\):", strPSData)
                        start = match.end()
                        boolV_PERMITTED = True
                    except:
                        boolV_PERMITTED = False

                    # get V_PERSPEED
                    try:
                        match = re.search("V_PERSPEED\s\(7\):", strPSData)
                        start = match.end()
                        boolV_PERSPEED = True
                    except:
                        boolV_PERSPEED = False

                    if boolV_PERSPEED == True:

                        try:

                            match = re.search("V_PERSPEED\s\(7\):", strPSData)
                            start = match.end()
                            match = re.search("\d+", strPSData[start:len(strPSData)])
                            vmrsp = int(strPSData[start+match.start():start+match.end()])

                        except Exception as ex:

                            vmrsp = None # Unknown

                    elif boolV_PERMITTED == True:

                        try:

                            match = re.search("V_PERMITTED\s\(7\):", strPSData)
                            start = match.end()
                            match = re.search("\d+", strPSData[start:len(strPSData)])
                            vmrsp = int(strPSData[start+match.start():start+match.end()])

                        except Exception as ex:

                            vmrsp = None # Unknown

                    else:

                        vmrsp = None

                    # Insert value into list
                    listVperm.insert(x, vmrsp)
                
            except:
                if cfg.glLocale == 'ENG': strRecord = str(inputdf["Record Id"][x])
                elif cfg.glLocale == 'SPA': strRecord = str(inputdf["Grabar Id"][x])
                else: strRecord = ''
                print(color.Fore.YELLOW + 
                        "AlGetVPerm-WARNING. Found JRU message with bad format in V_PERM block. Message number: " + 
                        strRecord + 
                        color.Style.RESET_ALL)
                listVperm.insert(x, None)
                


        else:
            listVperm.append('')
            
        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfVperm = pandas.DataFrame(numpy.array(listVperm), columns=['V_PERM'])
            
    return dfVperm


# HasGetBaliseGroupErrorFromData (hasDf)
# Alstom (JDRMDR) decoder
# Input: dataframe with type-hasler decoder
# Input format: columns=['Record Id', 'JRU']
# return BG Error from  DATA column

def HasGetBaliseGroupErrorFromData (hasDf):

    global glFunctionNum

    listNidErrorBg = []
    dfNidErrorBg = pandas.DataFrame()
    listMessageErrorBg = []
    dfMessageErrorBg = pandas.DataFrame()
    listReturn = []


    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    
    length = inputdf["DATA"].size
    
    for x in range(length):
        
        strBaliseGroupError = ''
        strBaliseGroupErrorData = ''        
            
        # ------------------- Get NID MESSAGE
        
        intNidMessage = None
        
        try:

            intNidMessage = int(inputdf["MESSAGE ID"][x])

        except:

            if cfg.glLocale == 'ENG': strRecord = str(inputdf["Record Id"][x])
            elif cfg.glLocale == 'SPA': strRecord = str(inputdf["Grabar Id"][x])
            else: strRecord = ''
            print(color.Fore.RED + 
                "HasGetBaliseGroupErrorFromData-ERROR. Found message with bad format in NID_MESSAGE : block. Message number: " + 
                strRecord + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 12: # Balise group error message

            strBaliseGroupError = inputdf["DATA"][x]
            strNidErrorBg = ''
            strMessageError = ''

            # NID_C (10): Spain, Córdoba - Málaga [355]
            # NID_BG (14): 2716
            # M_ERROR (8): Balise consistency: linking [0]

            
            # ---------------- Get error from message   
                    
            strBGEData = strBaliseGroupError

            try:
                # get NID_ERRORBG
                # NID_ERRORBG : 639-->639
                match = re.search("NID_BG\s\(14\):", strBGEData)
                start = match.end()
                match = re.search("\d*[.,]\d+|\d+", strBGEData[start:len(strBGEData)])
                strNidErrorBg = strBGEData[start+match.start():start+match.end()]

                # get M_ERROR
                # M_ERROR : 0-->Balise group: linking consistency error
                match = re.search("M_ERROR\s\(8\):", strBGEData)
                start = match.end()
                strMessageError = strBGEData[start+match.start():start+match.end()] 

                listNidErrorBg.append(strNidErrorBg)
                listMessageErrorBg.append(strMessageError)
            except:

                if cfg.glLocale == 'ENG': strRecord = str(inputdf["Record Id"][x])
                elif cfg.glLocale == 'SPA': strRecord = str(inputdf["Grabar Id"][x])
                else: strRecord = ''
                print(color.Fore.YELLOW + 
                        "HasGetBaliseGroupErrorFromData-WARNING. Found JRU message with bad format in BALISE GROUP ERROR block. Message number: " + 
                        strRecord + 
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

    dfNidErrorBg = pandas.DataFrame(numpy.array(listNidErrorBg), columns=['NID_ERRORBG'])
    dfMessageErrorBg = pandas.DataFrame(numpy.array(listMessageErrorBg), columns=['M_ERRORBG'])
        
    listReturn = [dfNidErrorBg, dfMessageErrorBg]
            
    return listReturn


# HasgetServiceBrakeCommandedFromData (hasDf)
# Hasler decoder
# Input: dataframe with type-hasler decoder
# Input format: columns=['Record Id', 'JRU']
# return Service Brake Command State from  DATA column

def HasGetServiceBrakeCommandedFromData (hasDf):

    global glFunctionNum

    listServiceBrakeCommandState = []
    dfServiceBrakeCommandState = pandas.DataFrame()


    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    
    length = inputdf["DATA"].size
    
    for x in range(length):
        
        strServiceBrakeCommandState = ''
        strServiceBrakeCommandStateData = ''        
            
        # ------------------- Get NID MESSAGE
        
        intNidMessage = None
        
        try:

            intNidMessage = int(inputdf["MESSAGE ID"][x])

        except:

            if cfg.glLocale == 'ENG': strRecord = str(inputdf["Record Id"][x])
            elif cfg.glLocale == 'SPA': strRecord = str(inputdf["Grabar Id"][x])
            else: strRecord = ''
            print(color.Fore.RED +            
                "HasgetServiceBrakeCommandedFromData. Found message with bad format in NID_MESSAGE : block. Message number: " + 
                strRecord + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 4: # Service Brake Command State message

            strServiceBrakeCommandState = inputdf["DATA"][x]

            # M_BRAKE_ORDER (1): Application [1]
            # M_BRAKE_ORDER (1): Revocation [0]

            
            # ---------------- Get error from message   
                    
            strSBCData = strServiceBrakeCommandState

            try:
                # To be completed with control of versions/releases
                # M_BRAKE_ORDER / M_BRAKE_STATE 
                # M_BRAKE_ORDER (1): Application [1]
                # M_BRAKE_ORDER (1): Revocation [0]
                # M_BRAKE_STATE (1): Application [1]
                # M_BRAKE_STATE (1): Revocation [0]
                try:
                    match = re.search("M_BRAKE_ORDER\s\(1\):", strSBCData)
                    start = match.end()
                except:
                    match = re.search("M_BRAKE_STATE\s\(1\):", strSBCData)
                    start = match.end()
                match = re.search("\[", strSBCData[start:len(strSBCData)])
                start = start + match.end()
                #end = start + match.start()
                match = re.search("\d*[.,]\d+|\d+", strSBCData[start:len(strSBCData)])
                #match = re.search("\d*[.,]\d+|\d+", strBGEData[start:end])
                strBrakeCommandState = strSBCData[start+match.start():start+match.end()]

                listServiceBrakeCommandState.append(strBrakeCommandState)
                
            except:

                
                if cfg.glLocale == 'ENG': strRecord = str(inputdf["Record Id"][x])
                elif cfg.glLocale == 'SPA': strRecord = str(inputdf["Grabar Id"][x])
                else: strRecord = ''
                print(color.Fore.YELLOW + 
                        "HasgetServiceBrakeCommandedFromData-WARNING. Found JRU message with bad format in SERVICE BRAKE COMMAND STATE block. Message number: " + 
                        strRecord + 
                        color.Style.RESET_ALL)
                listServiceBrakeCommandState.append('')
                


        else:
            listServiceBrakeCommandState.append('')
            
        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfServiceBrakeCommandState = pandas.DataFrame(numpy.array(listServiceBrakeCommandState), columns=['SBC_STATE'])
            
    return dfServiceBrakeCommandState


# HasGetEmergencyBrakeCommandedFromData (hasDf)
# Hasler decoder decoder
# Input: dataframe with type-hasler decoder
# Input format: columns=['Record Id', 'JRU']
# return Emergency Brake Command State from  DATA column

def HasGetEmergencyBrakeCommandedFromData (hasDf):

    global glFunctionNum

    listEmergencyBrakeCommandState = []
    dfEmergencyBrakeCommandState = pandas.DataFrame()


    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    
    length = inputdf["DATA"].size
    
    for x in range(length):
        
        strEmergencyBrakeCommandState = ''
        strEmergencyBrakeCommandStateData = ''        
            
        # ------------------- Get NID MESSAGE
        
        intNidMessage = None
        
        try:

            intNidMessage = int(inputdf["MESSAGE ID"][x])

        except:

            if cfg.glLocale == 'ENG': strRecord = str(inputdf["Record Id"][x])
            elif cfg.glLocale == 'SPA': strRecord = str(inputdf["Grabar Id"][x])
            else: strRecord = ''
            print(color.Fore.RED + 
                "HasGetEmergencyBrakeCommandedFromData-ERROR. Found message with bad format in NID_MESSAGE : block. Message number: " + 
                strRecord + 
                color.Style.RESET_ALL)
            exit(0)

        if intNidMessage == 3: # Emergency Brake Command State message

            strEmergencyBrakeCommandState = inputdf["DATA"][x]
            # M_BRAKE_ORDER (1): Application [1]
            # M_BRAKE_ORDER (1): Revocation [0]

                
            # ---------------- Get error from message   
                        
            strSBCData = strEmergencyBrakeCommandState

            try:
                # To be completed with control of versions/releases
                # M_BRAKE_ORDER / M_BRAKE_STATE 
                # M_BRAKE_ORDER (1): Application [1]
                # M_BRAKE_ORDER (1): Revocation [0]
                # M_BRAKE_STATE (1): Application [1]
                # M_BRAKE_STATE (1): Revocation [0]
                try:
                    match = re.search("M_BRAKE_ORDER\s\(1\):", strSBCData)
                    start = match.end()
                except:
                    match = re.search("M_BRAKE_STATE\s\(1\):", strSBCData)
                    start = match.end()
                match = re.search("\[", strSBCData[start:len(strSBCData)])
                start = start + match.end()
                #end = start + match.start()
                match = re.search("\d*[.,]\d+|\d+", strSBCData[start:len(strSBCData)])
                #match = re.search("\d*[.,]\d+|\d+", strBGEData[start:end])
                strBrakeCommandState = strSBCData[start+match.start():start+match.end()]

                listEmergencyBrakeCommandState.append(strBrakeCommandState)
                    
            except:

                    
                if cfg.glLocale == 'ENG': strRecord = str(inputdf["Record Id"][x])
                elif cfg.glLocale == 'SPA': strRecord = str(inputdf["Grabar Id"][x])
                else: strRecord = ''
                print(color.Fore.YELLOW + 
                        "HasgetServiceBrakeCommandedFromData-WARNING. Found JRU message with bad format in SERVICE BRAKE COMMAND STATE block. Message number: " + 
                        strRecord + 
                        color.Style.RESET_ALL)
                listEmergencyBrakeCommandState.append('')
                


        else:
            listEmergencyBrakeCommandState.append('')
            
        # Calculate progress
        glPh1Progress = 1/glPh1Functions*(glFunctionNum*glPh1NumberRows+x+1)/glPh1NumberRows*100
        sys.stdout.write("Progress: %d%%   \r" % (glPh1Progress) )
        sys.stdout.flush()

    dfEmergencyBrakeCommandState = pandas.DataFrame(numpy.array(listEmergencyBrakeCommandState), columns=['EBC_STATE'])
            
    return dfEmergencyBrakeCommandState




# HasApplyRowFilter(hasDf, listRbcFilter, listJruFilter)
# Hasler decoder
# Apply filter for JRU messages and RBC packets returns a re-indexed df with the filtered file
# Accept two filters (list): PACKET_RBC ID & JRUMESSAGE ID

def HasApplyRowFilter(hasDf, listRbcFilter, listJruFilter, listJruDiscard):

    # -----------  WARNING --------------------------------------------------
    # To review; not working properly when listRbCFilter is not an empty array
    # Use empty array fort Rbc Filter in main.py until correction
    # -----------------------------------------------------------------------

    inputdf = pandas.DataFrame(hasDf) # copia local del objeto del argumento
    listRbcPackets = numpy.array(listRbcFilter) # copia local del filtro de mensajes RBC
    listJruMessages = numpy.array(listJruFilter) # copia local del filtro de mensajes JRU
    listJruDiscardMessages = numpy.array(listJruDiscard)
    
    # filtramos df por tipos de mensjaes solicitados
    boolList = []
    serFilterJruMessages = pandas.Series() # Filter for messages in list listJruFilter
    serFilterJruDiscardedMess = pandas.Series()
    serFilterRbcMessages = pandas.Series() # Filter for rbc only-Rbc messages contained in ListRbcFilter
    serFilterNonRbcMessages = pandas.Series() # Filter for messages in listJruFilter but non-RBC messages
    serFilterRbcPackets = pandas.Series() # Fil
    serFinalFilter = pandas.Series() # Final filter compound from above ones

    # boolList
    for x in range (len(inputdf.axes[0])):
        boolList.insert(x, False)
    
    

    # Filter for required JRU Messages
    if len(listJruMessages) > 0:
        # Filtramos mensajes JRU
        for x in range(len(listJruMessages)):
            serFilterJruMessages[x] = inputdf["HEADER"].str.contains(listJruMessages[x], na=False, regex=False)
            if x > 0:
                # Sum of filters
                serFilterJruMessages[0] |= serFilterJruMessages[x]

        # Filtering
        inputdf = inputdf[serFilterJruMessages[0]]
        inputdf.index = numpy.arange(0, len(inputdf.axes[0]))

    # Filter for discarded JRU Messages
    if len(listJruDiscardMessages) > 0:
        # Filtramos mensajes JRU
        for x in range(len(listJruDiscardMessages)):
            serFilterJruDiscardedMess[x] = inputdf["HEADER"].str.contains(listJruDiscardMessages[x], na=False, regex=False)
            if x > 0:
                # Sum of filters
                serFilterJruDiscardedMess[0] |= serFilterJruDiscardedMess[x]
        serFilterJruDiscardedMess[0] = ~serFilterJruDiscardedMess[0]        

        # Filtering
        inputdf = inputdf[serFilterJruDiscardedMess[0]]
        inputdf.index = numpy.arange(0, len(inputdf.axes[0]))

    # Filter for only-RBC required JRU Messages
    listJruRbcMessages = numpy.array(['MESSAGE TO RBC [10]', 
                            'MESSAGE FROM RBC [9]'] )
    for x in range(len(listJruRbcMessages)):
        serFilterRbcMessages[x] = inputdf["HEADER"].str.contains(listJruRbcMessages[x], na=False, regex=False)
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
    inputdf.index = numpy.arange(0, len(inputdf["HEADER"]))
    
    return inputdf
