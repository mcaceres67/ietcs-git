
import os
import pandas as pd
import numpy as np
import colorama
import datetime as date
import pathlib
import colorama as color

import config as cfg
import jdrmdrcfg
import jdrmdr



startTotalTime = date.datetime.now()
numberofProcessedRegisters = 0
# ----------------------   CONFIG BLOCK - STARTS ----------------------------------------------------

# Filter for packets to compile
listRbcFilter = cfg.listRbcFilter
listJruFilter =  cfg.listJruFilter
listJruDiscard = cfg.listJruDiscard

# Input and Output paths
inputPath = jdrmdrcfg.inputPath
outputPath = jdrmdrcfg.outputPath

# Variables for multiple files processing
boolMultipleFile = True    # To indicate if process multiple files in directory or a single file
# When boolMultipleFile False; the file to compile
inputSinglefile = 'JDRMDR Export 2026-02-05 16h 52m 25s.txt'
#inputSinglefile = 'test.txt'


# ----------------------   CONFIG BLOCK - ENDS ----------------------------------------------------

# Print configuration info
print('ALSTOM JRU. Version: ' + jdrmdr.jdrmdrVersion)
print(color.Fore.YELLOW + 'Using locale: ' + cfg.glLocale + color.Style.RESET_ALL)
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

print(colorama.Fore.GREEN + 'Alstonm compilation process.' + colorama.Style.RESET_ALL)


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
    PREcompiledFilePath = pathlib.Path.joinpath(outputPath,'AlstomJRU_' + inputFileNameNoextension + '.xlsx')
    stroutputPREcompiledFile = str(PREcompiledFilePath)

    # Processing analytics
    arrEngine = []
    startTime = date.datetime.now()

    #  ---------------------- PHASE 1 ---------------------------
    t0 = date.datetime.now()
    print(colorama.Fore.GREEN + "Compilation starts" + colorama.Style.RESET_ALL)
    # Phase 1
    # Compile original Alstom file
    try:
        print(colorama.Fore.GREEN + "Pre-compiling phase" + colorama.Style.RESET_ALL)    
        dfPreCompiledAlstom = jdrmdr.AlPrecompile(strinputfile)
        numberofProcessedRegisters += jdrmdr.glNumberofCompiledRegisters
    except Exception as ex:
        print(colorama.Fore.YELLOW + "Can´t compile Alstom file. Pre-compiling phase. Exception: " + str(ex) + colorama.Style.RESET_ALL)
        exit()


    print("Saving pre-compiled file. " + stroutputPREcompiledFile)
    # Create output directory if not exists
    pathlib.Path(outputPath).mkdir(parents=True, exist_ok=True)

    # Excel write 
    try:
        with pd.ExcelWriter(stroutputPREcompiledFile) as writer:
            writer.book.formats[0].set_text_wrap()  # update global format            
            writer.book.formats[0].set_align('top')  # update global format 
            dfPreCompiledAlstom.to_excel(writer)
            sheet = writer.book.get_worksheet_by_name('Sheet1')
            sheet.autofit() 
            #cell_format =  writer.book.add_format({'border':1})
            #sheet.set_column(0, 2, None, cell_format)

    except Exception as ex:
        print(colorama.Fore.YELLOW + "Can´t save compiled file. Error: " + str(ex) + colorama.Style.RESET_ALL)


    try:
        print(colorama.Fore.GREEN + "Compiling phase" + colorama.Style.RESET_ALL)    
        dfCompiledAlstom = jdrmdr.AlCompile(dfPreCompiledAlstom, listRbcFilter, listJruFilter, listJruDiscard)
    except Exception as ex:
        print(colorama.Fore.YELLOW + "Can´t compile Alstom file. Compiling phase. Exception: " + str(ex) + colorama.Style.RESET_ALL)
        exit()

        
    print("Saving compiled file. " + stroutputCompiledFile)
    # Create output directory if not exists
    pathlib.Path(outputPath).mkdir(parents=True, exist_ok=True)

    # Excel write (end of Phase 1)
        
    try:
        with pd.ExcelWriter(stroutputCompiledFile) as writer:
            writer.book.formats[0].set_text_wrap()  # update global format            
            writer.book.formats[0].set_align('top')  # update global format 
            dfCompiledAlstom.to_excel(writer)
            sheet = writer.book.get_worksheet_by_name('Sheet1')
            sheet.autofit() 
            #cell_format =  writer.book.add_format({'border':1})
            #sheet.set_column(0, 2, None, cell_format)

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
print("Total of registers compiled: " + str(jdrmdr.glNumberofCompiledRegisters))
