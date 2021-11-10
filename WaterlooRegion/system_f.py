from os import listdir
from os.path import isfile, join

def get_excelname(path):
    list = [f for f in listdir(path) if isfile(join(path,f)) and '.xlsx' in join(path,f)]
    return list
