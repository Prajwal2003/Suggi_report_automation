import os

folder_path = "/Users/ashwyn/hello"
file_names = os.listdir(folder_path)
    
for file in file_names:
    print(file)

