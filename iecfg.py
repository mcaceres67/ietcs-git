# import config as cfg

import numpy as np
import pathlib


# Input and Output paths
mod_path = pathlib.Path(__file__).parent
# Important: inputPath and outputPath format:
# dir1/dir2

#inputPath = '../outputs/Israel 21'
#outputPath = '../outputs/Israel 21'
inputPath = '../outputs/S-103/C1'
outputPath = '../outputs/S-103/C1'
#inputPath = '../outputs/Isla Valencia S-112.012/Cab 23'
#outputPath = '../outputs/Isla Valencia S-112.012/Cab 23'
#inputPath = '../outputs/L-G/20251126-27'
#outputPath = '../outputs/L-G/20251126-27'
#inputPath = '../outputs/L-G/20260121-22'
#outputPath = '../outputs/L-G/20260121-22'
#inputPath = '../outputs/L-G/20260202-03'
#outputPath = '../outputs/L-G/20260202-03'
#inputPath = '../outputs/L-G/20260216-17'
#outputPath = '../outputs/L-G/20260216-17'
#inputPath = '../outputs/L-G/20260217-18'
#outputPath = '../outputs/L-G/20260217-18'
#inputPath = '../outputs/L-G/20260218-19'
#outputPath = '../outputs/L-G/20260218-19'
#inputPath = '../outputs/L-G/20260219-20'
#outputPath = '../outputs/L-G/20260219-20'
#inputPath = '../outputs/L-G/20260316-17'
#outputPath = '../outputs/L-G/20260316-17'
#inputPath = '../outputs/L-G/20260317-18'
#outputPath = '../outputs/L-G/20260317-18'
#inputPath = '../outputs/L-G/20260318-19'
#outputPath = '../outputs/L-G/20260318-19'
#inputPath = '../outputs/L-G/20260319-20'
#outputPath = '../outputs/L-G/20260319-20'
#inputPath = '../outputs/L-G/20260325-26'
#outputPath = '../outputs/L-G/20260325-26'
#inputPath = '../outputs/L-G/20260326-27'
#outputPath = '../outputs/L-G/20260326-27'
#inputPath = '../outputs/L-G/20260327-28'
#outputPath = '../outputs/L-G/20260327-28'
#inputPath = '../outputs/L-G/20260409-10'
#outputPath = '../outputs/L-G/20260409-10'
#inputPath = '../outputs/L-G/20260411-12'
#outputPath = '../outputs/L-G/20260411-12'
#inputPath = '../outputs/L-G/20260417-18'
#outputPath = '../outputs/L-G/20260417-18'
#inputPath = '../outputs/L-G/20260418-19'
#outputPath = '../outputs/L-G/20260418-19'
#inputPath = '../outputs/L-G/20260419-20'
#outputPath = '../outputs/L-G/20260419-20'
#inputPath = '../outputs/L-G/20260531-0601'
#outputPath = '../outputs/L-G/20260531-0601'
#inputPath = '../outputs/L-G/20260602-03'
#outputPath = '../outputs/L-G/20260602-03'
#inputPath = '../outputs/L-G LIF/20260701'
#outputPath = '../outputs/L-G LIF/20260701'
calibrationPath = outputPath + '/calibration'
statisticsPath = outputPath + '/statistics'


inputPath = (mod_path / inputPath).resolve()
outputPath = (mod_path / outputPath).resolve()
calibrationPath = (mod_path / calibrationPath).resolve()
statisticsPath = (mod_path / statisticsPath).resolve()

# Calibration process config
nCalibrationIteractions = 10
nameCalibrationFile = 'CalibrationFile.xlsx'
minCalibrationAccuracy = 1E-4

# ----------------------   CONFIG BLOCK - ENDS ----------------------------------------------------