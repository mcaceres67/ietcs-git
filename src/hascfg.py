import src.config as cfg
import numpy as np
import pathlib

# ----------------------   CONFIG BLOCK - STARTS ----------------------------------------------------

# Input and Output paths
mod_path = pathlib.Path(__file__).parent
# Important: inputPath and outputPath format:
# dir1/dir2

#inputPath = '../inputs/Israel 21'
#outputPath = '../outputs/Israel 21'
inputPath = '../inputs/S-103/C1'
outputPath = '../outputs/S-103/C1'
#inputPath = '../inputs/Isla Valencia S-112.012/Cab 23'
#outputPath = '../outputs/Isla Valencia S-112.012/Cab 23'

inputPath = (mod_path / inputPath).resolve()
outputPath = (mod_path / outputPath).resolve()

# ----------------------   CONFIG BLOCK - ENDS ----------------------------------------------------