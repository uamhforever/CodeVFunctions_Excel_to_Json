import json
import pandas as pd
import re


def getFuncArguments(arg_type="", arg_name="", arg_description=""):
    argn = {"type": "", "name": "", "description": ""}
    argn["type"] = arg_type
    argn["name"] = arg_name
    argn["description"] = arg_description
    return argn


savejsonfilepath = r"test.json"
excelfilepath = r"PredefinedFunctions.xlsx"
excelSheetName = 'Sheet1'
data = {"predefined": []}
with pd.ExcelFile(excelfilepath) as xls:
    df1 = pd.read_excel(xls, excelSheetName)
    nline = len(df1.index)
    nkey_true = df1.iloc[:, 1].count() - df1.iloc[:, 0].count()

    df_builtinsType = df1[pd.isnull(df1.iloc[:, 0]) == False]
    builtinsTypeCount = sum(pd.isnull(df1.iloc[:, 0]) == False)
    for idx in range(builtinsTypeCount):
        builtins_container = {"title": "", "description": "", "funcs": {}}
        title = str(df_builtinsType.iloc[idx, 1])
        description = "predefined " + title.lower() + \
            ". These can be called in any Macro-Plus sequence file."
        builtins_container["title"] = title
        builtins_container["description"] = description
        data["predefined"].append(builtins_container)

    i = 0  # row index
    j = -1  # bulitinsType index pointer FIXME(-1 may be cause the program bug)
    regex = re.compile(r'\b\s+\b')
    while i < nline:
        if not (pd.isnull(df1.iloc[i, 0])):
            j += 1
        else:

            if not (pd.isnull(df1.iloc[i, 1])):
                funcName = df1.iloc[i, 1]
                currentFuncNameLine = i
                if not (pd.isnull(df1.iloc[i, 2])):
                    returnType = df1.iloc[i, 2]
                else:
                    returnType = ""
                if not (pd.isnull(df1.iloc[i, 3])):
                    nArgs = df1.iloc[i, 3]
                else:
                    nArgs = 0
                if not (pd.isnull(df1.iloc[i, 4])):
                    argReturnType = df1.iloc[i, 4]
                else:
                    argReturnType = ""
                if not (pd.isnull(df1.iloc[i, 5])):
                    argName = df1.iloc[i, 5]
                else:
                    argName = ""
                if not (pd.isnull(df1.iloc[i, 6])):
                    argDescription = df1.iloc[i, 6]
                else:
                    argDescription = ""
                if not (pd.isnull(df1.iloc[i, 7])):
                    description = df1.iloc[i, 7]
                else:
                    description = ""
                func = {"returnType": "", "description": "",
                        "parameters": [], "signature": ""}
                func["returnType"] = regex.sub('_', returnType)
                func["description"] = description
                parameters = []
                parameters.append(getFuncArguments(
                    arg_type=argReturnType, arg_name=argName, arg_description=argDescription))
                func["parameters"] = parameters
                signature = returnType + " " + funcName + "("
                if nArgs > 0:
                    signature += regex.sub('_', argReturnType) + " " + argName

                data["predefined"][j]["funcs"][funcName] = func
            else:
                if not (pd.isnull(df1.iloc[i, 4])):
                    argReturnType = df1.iloc[i, 4]
                if not (pd.isnull(df1.iloc[i, 5])):
                    argName = df1.iloc[i, 5]
                if not (pd.isnull(df1.iloc[i, 6])):
                    argDescription = df1.iloc[i, 6]
                if nArgs > 1:
                    signature += ", " + \
                        regex.sub('_', argReturnType) + " " + argName
                    data["predefined"][j]["funcs"][funcName]["parameters"].append(getFuncArguments(
                        arg_type=argReturnType, arg_name=argName, arg_description=argDescription))

            if i == currentFuncNameLine + nArgs - 1 or nArgs == 0:
                signature += ")"
                data["predefined"][j]["funcs"][funcName]["signature"] = signature
        i += 1

# print(data)
with open(savejsonfilepath, mode='w') as f:
    json.dump(data, f, ensure_ascii=False, indent=4, separators=(',', ': '))

# with open(filepath, mode='r') as f:
#     x = json.load(f)
