import numpy as np
import pathlib
mod_path = pathlib.Path(__file__).parent
# Important: pathfile format:
# dir1/dir2

# ----------------------   CONFIG BLOCK - STARTS ----------------------------------------------------

# To change process behaviour use next variables

# Filter for packets to compile
# Track to train 3: SR Authorisation
# Track to train 33 MA with Shifted Location Reference
# Track to train 3 Movement Authority
# Train to Trcak 136: Train Position Report
# Train to track 150: End of Mission
# Train to track 157: SoM Position Report
# Train to track 132: MA Request
listRbcFilter = [] # np.array(['33', '3', '136', '150', '157'])


listJruFilter =  [] #np.array(['MESSAGE FROM BALISE', 'MESSAGE TO RBC',
                    #    'MESSAGE FROM RBC', 'PERMITTED SPEED',
                    #    'MOST RESTRICTIVE SPEED PROFILE'] )

listJruDiscard =  np.array(['STM INFORMATION', 'ETCS ON-BOARD PROPRIETARY JURIDICAL DATA'])
                   #'GENERAL MESSAGE']) 
                    #       'STOP DISPLAYING PLAIN TEXT MESSAGES', 
                    #       'START DISPLAYING PLAIN TEXT MESSAGES'])

# To filter/not filter registers with LRBG unknown (16383). Only applies for RBC messages; balise messages with unknown LRBG are never filtered
glboolFilterUnknownBalises = False # ! Important; let value = True

# global variable glLocale for content format
# Values
# ENG for english number formats and texts
#   , for thousands
#   . for decimals
# SPA for spanish number format and texts
#   . for thousands
#   , for decimals

glLocale = 'ENG' # english format
#glLocale = 'SPA' # spanish format

compiledFilePrefix = 'JruC_'

# ----------------------   CONFIG BLOCK - ENDS ----------------------------------------------------

# ----------------------   DIGITAL MAP CONFIGURATION - STARTS ----------------------------------------------------

dmFilePath = '../inputs/L-G/DM'
dmFilePath = (mod_path / dmFilePath).resolve()
dmFile = 'DM_LeonGuardo.xlsx'  
#dmFile = '' 

# ----------------------   DIGITAL MAP CONFIGURATION - ENDS ----------------------------------------------------