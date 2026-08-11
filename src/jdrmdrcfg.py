import config as cfg
import numpy as np
import pathlib

# ----------------------   CONFIG BLOCK - STARTS ----------------------------------------------------
msgIdPrefix = "Msg "

# Input and Output paths
mod_path = pathlib.Path(__file__).parent
# Important: inputPath and outputPath format:
# dir1/dir2

#inputPath = '../inputs/L-G/20251126-27/IERTMS'
#outputPath = '../outputs/L-G/220251126-27'
#inputPath = '../inputs/L-G/20260121-22/IERTMS'
#outputPath = '../outputs/L-G/20260121-22'
#inputPath = '../inputs/L-G/20260202-03/IERTMS'
#outputPath = '../outputs/L-G/20260202-03'
#inputPath = '../inputs/L-G/20260216-17/IERTMS'
#outputPath = '../outputs/L-G/20260216-17'
#inputPath = '../inputs/L-G/20260217-18/IERTMS'
#outputPath = '../outputs/L-G/20260217-18'
#inputPath = '../inputs/L-G/20260218-19/IERTMS'
#outputPath = '../outputs/L-G/20260218-19'
#inputPath = '../inputs/L-G/20260219-20/IERTMS'
#outputPath = '../outputs/L-G/20260219-20'
#inputPath = '../inputs/L-G/20260316-17/IERTMS'
#outputPath = '../outputs/L-G/20260316-17'
#inputPath = '../inputs/L-G/20260317-18/IERTMS'
#outputPath = '../outputs/L-G/20260317-18'
#inputPath = '../inputs/L-G/20260318-19/IERTMS'
#outputPath = '../outputs/L-G/20260318-19'
#inputPath = '../inputs/L-G/20260319-20/IERTMS'
#outputPath = '../outputs/L-G/20260319-20'
#inputPath = '../inputs/L-G/20260325-26/IERTMS'
#outputPath = '../outputs/L-G/20260325-26'
#inputPath = '../inputs/L-G/20260326-27/IERTMS'
#outputPath = '../outputs/L-G/20260326-27'
#inputPath = '../inputs/L-G/20260327-28/IERTMS'
#outputPath = '../outputs/L-G/20260327-28'
#inputPath = '../inputs/L-G/20260409-10/IERTMS'
#outputPath = '../outputs/L-G/20260409-10'
#inputPath = '../inputs/L-G/20260411-12/IERTMS'
#outputPath = '../outputs/L-G/20260411-12'
#inputPath = '../inputs/L-G/20260417-18/IERTMS'
#outputPath = '../outputs/L-G/20260417-18'
#inputPath = '../inputs/L-G/20260418-19/IERTMS'
#outputPath = '../outputs/L-G/20260418-19'
#inputPath = '../inputs/L-G/20260419-20/IERTMS'
#outputPath = '../outputs/L-G/20260419-20'
#inputPath = '../inputs/L-G/20260531-0601/IERTMS'
#outputPath = '../outputs/L-G/20260531-0601'
inputPath = '../inputs/L-G/20260602-03/IERTMS'
outputPath = '../outputs/L-G/20260602-03'
#inputPath = '../inputs/L-G LIF/20260701/IERTMS'
#outputPath = '../outputs/L-G LIF/20260701'
inputPath = (mod_path / inputPath).resolve()
outputPath = (mod_path / outputPath).resolve()

# ----------------------   CONFIG BLOCK - ENDS ----------------------------------------------------