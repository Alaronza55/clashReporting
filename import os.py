import os

Pythonfile=os.getcwd()
print(Pythonfile)

users = os.path.expanduser('~')
print(users)

os.chdir(users)
listLogin=os.listdir(".")
print(users)

os.chdir("Downloads")
listDownloads=os.listdir(".")
print(listDownloads)

Downloads=os.getcwd()
print(Downloads)

# for file in os.listdir(Downloads):
#     if file.startswith("XXX"): 
#         old_name = str(file)
#         old_name_path = os.path.join(Downloads,old_name)
#         new_name_path = os.path.join(Downloads,new_name)
#         os.rename(old_name_path,new_name_path)