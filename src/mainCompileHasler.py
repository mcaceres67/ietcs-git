
import os
import pandas as pd
import numpy as np
import colorama
import datetime as date
import pathlib

import src.config as cfg
import src.hascfg as hascfg
import src.hasler as has



startTotalTime = date.datetime.now()
numberofProcessedRegisters = 0
# ----------------------   CONFIG BLOCK - STARTS ----------------------------------------------------

# Filter for packets to compile
listRbcFilter = cfg.listRbcFilter
listJruFilter =  cfg.listJruFilter
listJruDiscard = cfg.listJruDiscard

# Input and Output paths
inputPath = hascfg.inputPath
outputPath = hascfg.outputPath

# Variables for multiple files processing
boolMultipleFile = False    # To indicate if process multiple files in directory or a single file
# When boolMultipleFile False; the file to compile
#inputSinglefile = '[Inno] Ejemplo Excel JRU (01-03-21)-repaired.xls'
inputSinglefile = 'JruC_ESC S-103 CoMa 22-24 jun 2024 C1 (5)_splitted.xlsx'

# ----------------------   CONFIG BLOCK - ENDS ----------------------------------------------------

# Print configuration info
print('Hasler JRU. Version: ' + has.hasVersion)
print('Using locale: ' + has.glLocale)
print('Processing multiple files: ' + str(boolMultipleFile))

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

print('Files to compile: ' + str(listInputFiles))

print(colorama.Fore.GREEN + 'Hasler compilation process.' + colorama.Style.RESET_ALL)


userContinue = input("Continue? [y/n] Enter to continue. ")
if not((userContinue == 'y') | (userContinue == 'Y') | (userContinue == '')):
    print('Bye.')
    exit()

dfAggregatedAntennaDistances = pd.DataFrame() # Initialization
arrDfAntennaDistances = [] # Initialization
for numseq in range(len(listInputFiles)):

    #inputfile = inputpath + listInputFiles[0]
    inputfilePath = pathlib.Path.joinpath(inputPath, listInputFiles[numseq])
    strinputfile = str(inputfilePath)

    # Get inputfilename without extension
    strAux = listInputFiles[numseq]
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
    inputFileNameNoextension = listInputFiles[numseq][0:ix].lstrip()  
    compiledFilePath = pathlib.Path.joinpath(outputPath,'JruC_' + inputFileNameNoextension + '.xlsx')
    stroutputCompiledFile = str(compiledFilePath)

    # Processing analytics
    arrEngine = []
    startTime = date.datetime.now()

    #  ---------------------- PHASE 1 ---------------------------
    t0 = date.datetime.now()
    print(colorama.Fore.GREEN + "Compilation phase starts" + colorama.Style.RESET_ALL)
        
    # Phase 1
    # Compile original Hasler file
    try:
        dfCompiledHasler = has.HasCompile(strinputfile, listRbcFilter, listJruFilter, listJruDiscard)
        numberofProcessedRegisters += has.glNumberofCompiledRegisters
        #dfCompiledHasler = has.HasCompile('./inputs/2024-03-07_1.xls', listRbcFilter, listJruFilter)
    except Exception as ex:
        print(colorama.Fore.YELLOW + "Can´t compile Hasler file. Exception: " + str(ex) + colorama.Style.RESET_ALL)
        exit()

        
    print("Saving compiled file. " + stroutputCompiledFile)
    # Create output directory if not exists
    pathlib.Path(outputPath).mkdir(parents=True, exist_ok=True)

    # Excel write (end of Phase 1)
        
    try:
        with pd.ExcelWriter(stroutputCompiledFile) as writer:
            dfCompiledHasler.to_excel(writer)
    except Exception as ex:
        print(colorama.Fore.YELLOW + "Can´t save compiled file. Error: " + str(ex) + colorama.Style.RESET_ALL)

    t1 = date.datetime.now()
    delta = t1-t0
    delta = delta.total_seconds()
    print(colorama.Fore.GREEN + "Compilation ends" + ". Partial time: " + str(delta) + ' secs' + colorama.Style.RESET_ALL)
        
    
endTotalTime = date.datetime.now()
delta = endTotalTime - startTotalTime
delta = delta.total_seconds()
print("Total process time: " + str(delta) + ' secs')
print("Total of registers compiled: " + str(has.glNumberofCompiledRegisters))
