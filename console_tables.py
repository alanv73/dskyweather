import operator

def printTable(myDict, colList=None):
    """ Pretty print a list of dictionaries (myDict) as a dynamically sized table.
    If column names (colList) aren't specified, they will show in random order.
    Author: Thierry Husson - Use it as you want but don't blame me.
##    Mod 7/26/2018: Alan Van Art - check all dicts for max column count. darksky
##    data may return data with different column counts - we'll see what happens
##    when I see the condition again.
##    line 10: myDict[0] -> myDict[max([(k,len(m)) for k, m in enumerate(myDict)], key=operator.itemgetter(1))[0]]
    """
##    if not colList: colList = list(myDict[0].keys() if myDict else [])

    if not colList:
        colList = list(
            myDict[
                max(
                    [(k,len(m)) for k, m in enumerate(myDict)],
                    key=operator.itemgetter(1))[0]
            ].keys() if myDict else [])


    myList = [colList] # 1st row = header

##    for item in myDict: myList.append([str(item[col] or '') for col in colList])
    for item in myDict: myList.append([str(item.get(col,'')) for col in colList])

    colSize = [max(map(len,col)) for col in zip(*myList)]
    formatStr = ' | '.join(["{{:<{}}}".format(i) for i in colSize])
    myList.insert(1, ['-' * i for i in colSize]) # Seperating line

    txt=''
    for item in myList:
##       print(formatStr.format(*item))
       txt = txt + formatStr.format(*item) + '\n'
    return txt
