import datetime
import os
import sys
from typing import Tuple, Any, Dict, List

import pandas as pd


def fileRead(filePath, sheet, rowLimit):
    frame = pd.read_excel(filePath, sheet).values.tolist()
    date = None
    for i in frame[0]:
        if type(i) is datetime.datetime:
            date = i
    account = []
    for i in frame[1]:
        if type(i) is str:
            account.append(i)
    records = {"account": []}
    for i in frame[2]:
        if i not in records.keys() and type(i) is str:
            records[i] = []
    titles = list(records.keys())
    length = len(titles) - 1

    for line in frame[3:]:
        if not rowLimit:
            break
        for i in range(length * (len(line) // length)):
            if i % length == 0:
                if type(line[i]) is float:
                    rowLimit -= 1
                else:
                    rowLimit = 100
                    records["account"].append(account[i // length])
            if rowLimit != 100:
                continue
            if type(line[i % length]) is not float or line[i % length] > 0:
                records[titles[i % length + 1]].append(line[i % length])
            else:
                records[titles[i % length + 1]].append(None)
    return date, records


def writeFile(outPath, sheetName, records):
    writer = pd.ExcelWriter(outPath)
    output = pd.DataFrame(records)
    output.to_excel(writer, sheet_name=sheetName)
    writer.close()


def formatFile(filePath, outPath, startSheet, endSheet):
    for i in range(startSheet - 1, endSheet):
        if os.path.exists(filePath):
            sheet_name = list(pd.read_excel(filePath, sheet_name=None).keys())[i]
            date, records = fileRead(filePath, i, 100)
        else:
            print("No such file or directory.")
            sys.exit()
        if os.access(outPath, os.W_OK):
            writeFile(outPath, sheet_name + date.strftime("%Y-%m"), records)
        else:
            print("Can't write file.")
            sys.exit()


if __name__ == '__main__':
    filePath = "J:/ExcelProcess/data.xls"
    outPath = "J:/ExcelProcess/out.xlsx"
    outData = []
    if os.path.exists(filePath):
        for i in range(6, 26):
            sheet_name = list(pd.read_excel(filePath, sheet_name=None).keys())[i]
            date, records = fileRead(filePath, i, 100)
            outData.append([sheet_name + date.strftime("%Y-%m"), records])
    else:
        print("No such file or directory.")
        sys.exit()
    if os.access(outPath, os.W_OK):
        with pd.ExcelWriter(outPath) as writer:
            for sheet_name, records in outData:
                output = pd.DataFrame(records)
                output.to_excel(writer, sheet_name)
    else:
        print("Can't write file.")
        sys.exit()
