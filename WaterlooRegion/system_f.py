from os import listdir
from os.path import isfile, join
import CAFunctions.CAFunctions as ca


def get_excelname(path):
    list = [f for f in listdir(path) if isfile(join(path,f)) and '.xlsx' in join(path,f)]
    return list


def get_codeconcern(sentence):
    # return tuple:
    # [0]: Code concern Yes or No.
    # [1]: Code concern comment
    if 'c:' in sentence:
        Boolean = 'Y'
        sentence_list = sentence.split('\n')
        for item in sentence_list:
            # this is added because only 1 item can be code concern
            if 'c:' in item:
                item = item.replace('c:', '').lstrip().rstrip()
                Comment = 'The asset had a code concern, ' + item[0].lower() + item[1:]
                if Comment[len(Comment) - 1] != '.':
                    Comment = Comment + '.'
    else:
        Boolean = 'N'
        Comment = 'The asset did not have any code concerns.'
    return (Boolean, Comment)


def get_conditionR(condition_str):
    condition_str = condition_str.lower()
    condition_rating = ca.Converter_UniAssetConditionConversion(condition_str)
    return condition_rating


def get_performance(condition, inspector_comment):
    # return tuple:
    # [0]: performance rating
    # [1]: performance rating comments
    if condition == 1 or 2 or 3:
        performance_rating = condition
        performance_sentence = 'The asset did not have any performance issues.'
    if condition == 4 or 5 and 'performance' not in inspector_comment:
        performance_rating = 3
        performance_sentence = 'The asset did not have any performance issues.'
    else:
        performance_rating = 4
        performance_sentence = 'The asset had performance issue.'
    return (performance_rating, performance_sentence)


def get_inspectioncom(cond)




sentence = 'jasndkjsabfjkasbfjkashfb jaaskj njkdabf jhabfjh a\nc: The asset is not good for the exisitng code \n dasd'
package = get_codeconcern(sentence)
print(package[0], package[1])
