from os import listdir
from os.path import isfile, join

def get_excelname(path):
    list = [f for f in listdir(path) if isfile(join(path,f)) and '.xlsx' in join(path,f)]
    return list


def get_codeconcern(sentence):
    if 'c:' in sentence:
        Boolean = 'Y'
        sentence_list = sentence.split('\n')
        for item in sentence_list:
            # this is added because only 1 item can be code concern
            if 'c:' in item:
                item = item.replace('c:', '').lstrip().rstrip()
                comm = 'The asset had a code concern, ' + item[0].lower() + item[1:]
                if comm[len(comm) - 1] != '.':
                    comm = comm + '.'
    else:
        Boolean = 'N'
        comm = 'The asset did not have any code concerns.'
    return (Boolean, comm)


sentence = 'jasndkjsabfjkasbfjkashfb jaaskj njkdabf jhabfjh a\nc: The asset is not good for the exisitng code \n dasd'
package = get_codeconcern(sentence)
print(package[0], package[1])
