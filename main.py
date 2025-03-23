import os
import orjson as oj
import pandas as pd


semanticList = []
reportList = []

semanticDF = pd.DataFrame()
reportDF =pd.DataFrame()

def jsonLoad(jFile):
    try:
        with open(jFile,'rb') as file:
            content = file.read().decode("utf-8")
            data = oj.loads(content)
            return data
    except FileNotFoundError:
        print(f"Error: File Not Found {jFile}")
        return None
    except oj.JSONDecodeError:
        print(f"Error: Invalid JSON format {jFile}")
        return None
    
def reportPages(rData):
    pageDict = {}
    for page in rData:
        name = page['displayName']
        pageDict = {"Page": name}
        
        print(page.keys())

        for key in page.keys():
            #columnDict[f"{key}"] = column[f"{key}"]
            pageDict[f"{key}"] = page[f"{key}"]
            #print(type(key))
            if "visualContainers" in key:
                visualContainers = page['visualContainers']
                for visuals in visualContainers:
                    config = oj.loads(visuals['config'])
                    filters = visuals['filters']
                    #print(type(config))
                    for configItem in config:
                        print((config[f"{configItem}"]))


                #visualContainers = page['visualContainers']
                #print(name)
                #for visuals in visualContainers:
                #    config = oj.loads(visuals['config'])
                #    filters = visuals['filters']
                    #print(filters)
                    #print(type(config))
                    #config = oj.loads(visuals['config'])
                    #configName = config['name']
                    #print(configName)
                    #visual = config['singleVisual']
                    #print(visual)
                    #visualType = visual['visualType']
                    #print(f"report {visualType}")
        reportList.append(pageDict)
    df = pd.DataFrame(reportList)
    print(df)
            


def semanticModel(mData):
    for table in mData:
        #print(f"Name: {table['name']} {table.keys()}")
        name = table['name']
        #semanticDic = {"Model": name}
        #semanticList.append(semanticDic)
        #semanticDic.clear()
        #semanticDF.append(semanticList, ignore_index=True)
        for column in table['columns']:
           # print(f"{column['name']}: {column.keys()}")
            columnDict = {"Table Name": name}
            for key in column.keys():
                columnDict[f"{key}"] = column[f"{key}"]
                if isinstance(column[f"{key}"], list):
                    for subList in column[f"{key}"]:
                        #print(subList)
                        for subKey in subList.keys():
                            #print(subKey)
                            columnDict[f"{key}"] = subList[f"{subKey}"]

           # if column['type'] is not None:
           #     columnType = column['type']
           # else:
           #     columnType = ""
           # print(columnType)
            
            #columnDict = {"Model": name, "Column Name":column['name'], "Data Type": column['dataType'], "Summarize By": column['summarizeBy'], "Type": columnType }
            semanticList.append(columnDict)
    #print(f"semantic Model {semanticList}")
    df = pd.DataFrame(semanticList)
    #print(df)
    df.to_excel('test.xlsx', sheet_name="test", index=False)


def main():

    

    reportFile = "c:\\Users\\Derek\\Documents\\code\\python\\PowerBI\\powerbicheckpy\\Report\\test2.Report\\report.json"
    modelFile = "c:\\Users\\Derek\\Documents\\code\\python\\PowerBI\\powerbicheckpy\\Report\\test2.SemanticModel\\model.bim"



    report = jsonLoad(reportFile)['sections']
    model = jsonLoad(modelFile)['model']['tables']

    
    #print(report)
    #print(len(model))
    semanticModel(model)
    reportPages(report)
    #print(semanticDF)



if __name__ == '__main__':main()