import os
import glob
from datetime import datetime

pattern = "CBD_*.csv"
files_names = glob.glob(pattern)
files_names.sort()
file_path = files_names[1]
filename = os.path.basename(file_path)
parts = filename.split("_")
num = parts[2].split("-")
x = ( num[0] + "-" + str(int(num[0])+6))
header = []
header.append(x + ' Target')
header.append(x + ' Achieved')
print(header)