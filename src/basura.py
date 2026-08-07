import numpy as np
import pandas as pd
import re

strPacketheader = 'MESSAGE RBC: 33\nQ_SCALE : 0 -->10 cm scale\nD_REF : 64866 -->-67 m'

match = re.search("D_REF\s:", strPacketheader)
start = match.end()
#   D_REF : 64779 -->-75,7 m
match = re.search("\-\-\>", strPacketheader[start:len(strPacketheader)])
start = start + match.end()
#match = re.search("[+-]?\d*[.,]\d+|\d+", strPacketheader[start:len(strPacketheader)])
match = re.search("[+-]?(?:\d*[.,]\d+|\d+)", strPacketheader[start:len(strPacketheader)])
dref = strPacketheader[start+match.start():start+match.end()]

pepe = 0



