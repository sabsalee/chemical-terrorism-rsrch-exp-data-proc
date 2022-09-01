import os, csv
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import NamedStyle
from openpyxl.utils.cell import get_column_letter
from module.timeinfo import TimeInfo



class CsvDataSheet:
    def __init__(self, type, name, path) -> None:
        self.type:str = type
        self.name:str = name
        self.path:str = path
        self.dirPath:str = path[:-(len(name)+1)]


class TimeInfoNotMatched(Exception):
    def __str__(self) -> str:
        return "실험시작 시간이 서로 다른 파일이 한 폴더 안에 있는 것 같습니다. 이 프로그램은 동시에 시작한 실험 파일만을 처리할 수 있습니다. 다시 시도하세요."



def filter_csv_from_dir(dirPath) -> list:
    processed_file_objects = []
    # 경로안에 있는 유효한 파일들 걸러내기
    fileList = os.listdir(dirPath)
    csvList = [e for e in fileList if ".csv" in e[-4:]]
    # 유효한 파일의 이름으로 이 데이터의 종류 결정하기
    for file in csvList:
        if "ID" in file:
            processed_file_objects.append(CsvDataSheet("dummy", file, f"{dirPath}/{file}"))
        elif "Total" in file:
            processed_file_objects.append(CsvDataSheet("total", file, f"{dirPath}/{file}"))
        elif 'CO2' in file:
            processed_file_objects.append(CsvDataSheet("CO2", file, f"{dirPath}/{file}"))
        elif 'He' in file:
            processed_file_objects.append(CsvDataSheet("He", file, f"{dirPath}/{file}"))
    return processed_file_objects



def preprocess_csv_to_df(csvDataSheet:CsvDataSheet, time:TimeInfo) -> pd.DataFrame: # CSV파일 가공하는 함수
    if csvDataSheet.type == "dummy":
        df = pd.read_csv(csvDataSheet.path, encoding="CP949")
        df = df.drop([df.columns[5], df.columns[6], df.columns[7]], axis='columns') # 온도, 습도, 대기압 삭제
        df[df.columns[0]] = pd.to_datetime(df[df.columns[0]], format="%Y-%m-%d-%H-%M-%S") # 측정시간 Datetime으로 변환

        # 시작 시간 추출
        if time.isExist() == False:
            time.set_time("start", df.iloc[0,0])
        

        indexCount = 0 # 범위에 포함되는 인덱스 잘라내기 위한 변수
        for i in range(df.shape[0]):
            diff = df.iloc[i,0] - df.iloc[0,0]
            if diff.seconds > 600:
                break
            indexCount += 1
        df = df.iloc[:indexCount]
        # df[df.columns[0]] = df[df.columns[0]].dt.strftime("%I:%M:%S %p")
        return df

    elif csvDataSheet.type == "total":
        df = pd.read_csv(csvDataSheet.path, encoding="CP949")
        dropArray = []
        for i in [5,6,7, 12,13,14, 19,20,21, 26,27,28, 33,34,35, 40,41,42, 47,48,49, 54,55,56, 61,62,63, 68,69,70]:
            dropArray.append(df.columns[i])
        df = df.drop(dropArray, axis='columns') # 온도, 습도, 대기압 전체 삭제
        df[df.columns[0]] = pd.to_datetime(df[df.columns[0]], format="%Y-%m-%d-%H-%M-%S") # 측정시간 Datetime으로 변환

        # 시작 시간 추출
        if time.isExist() == False:
            time.set_time("start", df.iloc[0,0])

        indexCount = 0 # 범위에 포함되는 인덱스 잘라내기 위한 변수
        for i in range(df.shape[0]):
            diff = df.iloc[i,0] - df.iloc[0,0]
            if diff.seconds > 600:
                break
            indexCount += 1
        df = df.iloc[:indexCount]
        return df

    elif csvDataSheet.type == "CO2" or csvDataSheet.type == "He":
        # csv 전처리 후 numpy Array로 변환하여 DataFrame 생성
        rows = []
        col = {"exist": False, "array": []}
        with open(csvDataSheet.path, 'r') as f:
            rdr = csv.reader(f)
            for i in rdr:
                if col["exist"] == False:
                    del i[1]
                    col["array"] = i
                    col["array"] = list(filter(len, col["array"]))
                    col["exist"] = True
                    continue
                if len(i) > 6:
                    for _ in range(len(i) - 6): # Column이 6개이므로 그 이상은 버리기
                        i.pop()
                i[2] = f"{i[1]} {i[2]}" # 날짜 시간 합쳐주기
                del i[1]
                rows.append(i)
        ar = np.array(rows)
        df = pd.DataFrame(ar, columns=col["array"])
        df = df.astype({"Temp.":float, "Humidity":float})
        df = df.astype({'CO2':float} if csvDataSheet.type == 'CO2' else {'He':float})
        df.rename(columns={'CO2':'이산화탄소(ppm)'} if csvDataSheet.type == 'CO2' else {'He':'헬륨(%)'}, inplace=True)

        # 시작 시간 추출
        if time.isExist() == False:
            print("실험 시작시간을 입력해주세요 ( 예시[14시 8분] -> 14:8 )")
            _time = input("> ")
            _time = list(map(int, _time.split(':')))
            dateInCell = list(map(int, (df.iloc[0,1].split(" "))[0].split("-")))
            time.set_time("start", datetime(dateInCell[2], dateInCell[0], dateInCell[1], _time[0], _time[1], 0))

        df = df.drop([df.columns[0]], axis='columns') # 인덱스 삭제
        df[df.columns[0]] = pd.to_datetime(df[df.columns[0]], format="%m-%d-%Y %H:%M:%S") # 측정시간 Datetime으로 변환
        
        try:
            startIndexNum = df[df["TIME"] == time.get_time("start")].index[0] # 시작시간 인덱스 가져오기
        except:
            raise TimeInfoNotMatched

        indexCount = 0 # 범위에 포함되는 인덱스 잘라내기 위한 변수
        for i in range(startIndexNum, df.shape[0]):
            diff = df.iloc[i,0] - df.iloc[startIndexNum,0]
            if diff.seconds > 600:
                # print(f"600초를 넘어 정지된 시점 : {df.iloc[i,0]}") # FOR DEBUG
                break
            # print(f"[{startIndexNum + indexCount}] {df.iloc[i,0].strftime('%H:%M:%S')}는 시작시간 {df.iloc[startIndexNum,0].strftime('%H:%M:%S')}과의 차가 {diff}초 입니다.") # FOR DEBUG
            indexCount += 1
        df = df.iloc[startIndexNum:startIndexNum + indexCount]
        return df



def dataframe_to_excel(dataframe:pd.DataFrame):
    try:
        wb = Workbook()
        ws = wb.active
        for i in dataframe_to_rows(dataframe, index=False, header=True):
            ws.append(i)
        return wb

    except:
        print("오류!!! None타입을 받았거나 그 외의 오류 발생")



def formular_process(wb:Workbook):
    try:
        # 열추가, 수식추가
        ws = wb.active
        ws.insert_cols(2, 2)
        ws["C1"] = "경과시간(s)"
        for i in range(2, ws.max_row + 1):
            ws[f"C{i}"] = f"=ROUND((A{i}-$A$2)*24*60*60,0)"
        return wb
    except:
        pass

def firebase_process(wb: Workbook, csvDataSheet: CsvDataSheet, __type: str) -> dict:
    dataDict = {}
    ws = wb.active

    # 경과시간(s)를 int 타입으로 변경
    for i in range(2, ws.max_row + 1):
        diff = ws[f"A{i}"].value - ws["A2"].value
        ws[f"C{i}"] = diff.seconds

    # 모든 날짜를 str 타입으로 변경
    for i in range(2, ws.max_row + 1):
        ws[f"A{i}"] = ws[f"A{i}"].value.strftime('%H:%M:%S')

    if __type == 'total':
        # pandas에서 중복된 column에 붙인 숫자 제거
        alignNum = 4
        for _ in range(10):
            ws[f'{get_column_letter(alignNum)}1'] = 'ID'
            ws[f'{get_column_letter(alignNum+1)}1'] = '산소(%)'
            ws[f'{get_column_letter(alignNum+2)}1'] = '이산화탄소(ppm)'
            ws[f'{get_column_letter(alignNum+3)}1'] = '헬륨(%)'
            alignNum += 4

        for id in range(10):
            all = ws.iter_rows(min_row=0, max_row=ws.max_row, min_col=0, max_col=ws.max_column, values_only=True) # for 문 바깥에 있었는데 enumerate 후 값이 사라져 다시 선언하기로 하였다.
            dataDict[f'DUMMY_ID{id+1}'] = {}
            for i, row in enumerate(all):
                # 0 - 날짜, 2 - 경과시간 / n - ID, n+1 - O2, n+2 - CO2, n+3 - He
                dataDict[f'DUMMY_ID{id+1}'][i] = [row[0],row[2],row[4 + 4*id],row[5 + 4*id],row[6 + 4*id]]
        return dataDict

    elif __type == 'CO2' or __type == 'He':
        all = ws.iter_rows(min_row=0, max_row=ws.max_row, min_col=0, max_col=ws.max_column, values_only=True)
        __si = csvDataSheet.name.rfind('_') + 1
        __ei = csvDataSheet.name.find('.csv')
        dataDict[f'{__type}_{csvDataSheet.name[__si:__ei]}'] = {}
        for i, row in enumerate(all):
            # 0 - 날짜, 2 - 경과시간, 5 - 타겟가스
            dataDict[f'{__type}_{csvDataSheet.name[__si:__ei]}'][i] = [row[0],row[2],row[5]]
        return dataDict



def expression_process(wb:Workbook, csvDataSheet:CsvDataSheet):
    ws = wb.active

    if csvDataSheet.type == "dummy":
        dateStyle = NamedStyle(name="datetime", number_format="[$-x-systime]h:mm:ss AM/PM")
        for i in range(2, ws.max_row + 1):
            ws[f"A{i}"].style = dateStyle
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["C"].width = 10
        ws.column_dimensions["D"].width = 4
        ws.column_dimensions["L"].width = 5
        ws.column_dimensions["W"].width = 1
        return wb
    
    elif csvDataSheet.type == "total":
        dateStyle = NamedStyle(name="datetime", number_format="[$-x-systime]h:mm:ss AM/PM")
        for i in range(2, ws.max_row + 1):
            ws[f"A{i}"].style = dateStyle
        ws.column_dimensions["A"].width = 20

        # pandas에서 중복된 column에 붙인 숫자 제거
        alignNum = 4
        for i in range(1, 11):
            ws[f'{get_column_letter(alignNum)}1'] = 'ID'
            ws[f'{get_column_letter(alignNum+1)}1'] = f'ID{i} 산소(%)'
            ws[f'{get_column_letter(alignNum+2)}1'] = f'ID{i} 이산화탄소(ppm)'
            ws[f'{get_column_letter(alignNum+3)}1'] = f'ID{i} 헬륨(%)'
            alignNum += 4
        return wb
    
    elif csvDataSheet.type == "CO2" or csvDataSheet.type == "He":
        dateStyle = NamedStyle(name="datetime", number_format="[$-x-systime]h:mm:ss AM/PM")
        for i in range(2, ws.max_row + 1):
            ws[f"A{i}"].style = dateStyle
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["C"].width = 10
        return wb
        

def chart_process(wb:Workbook, csvDataSheet:CsvDataSheet):
    if csvDataSheet.type == "dummy":
        try:
            # 차트추가
            ws = wb.active
            valueOxygen = Reference(ws, min_row=1, max_row=ws.max_row, min_col=5, max_col=5)
            valueCarbonDioxide = Reference(ws, min_row=1, max_row=ws.max_row, min_col=6, max_col=6)
            valueHelium = Reference(ws, min_row=1, max_row=ws.max_row, min_col=7, max_col=7)
            valueTime = Reference(ws, min_row=2, max_row=ws.max_row, min_col=3, max_col=3)

            oxygenChart = LineChart()
            oxygenChart.add_data(valueOxygen, titles_from_data=True)
            oxygenChart.set_categories(valueTime)
            # stylesheet
            oxygenChart.x_axis.title = "Time(s)"
            oxygenChart.y_axis.title = "O2(%)"
            oxygenChart.height = 13
            oxygenChart.width = 19
            
            ws.add_chart(oxygenChart, "B6")

            carbonDioxideChart = LineChart()
            carbonDioxideChart.add_data(valueCarbonDioxide, titles_from_data=True)
            carbonDioxideChart.set_categories(valueTime)
            # stylesheet
            carbonDioxideChart.x_axis.title = "Time(s)"
            carbonDioxideChart.y_axis.title = "CO2(ppm)"
            carbonDioxideChart.height = 13
            carbonDioxideChart.width = 19

            ws.add_chart(carbonDioxideChart, "M6")

            heliumChart = LineChart()
            heliumChart.add_data(valueHelium, titles_from_data=True)
            heliumChart.set_categories(valueTime)
            # stylesheet
            heliumChart.x_axis.title = "Time(s)"
            heliumChart.y_axis.title = "He(%)"
            heliumChart.height = 13
            heliumChart.width = 19

            ws.add_chart(heliumChart, "X6")
        except:
            pass

    elif csvDataSheet.type == "total":
        ws = wb.active
        alignNum = 4 # ID칸 시작이 D열이다 (4씩 추가된다)

        valueTime = Reference(ws, min_row=2, max_row=ws.max_row, min_col=3, max_col=3)

        row = 2
        col = 4
        for _ in range(10):
            valueOxygen = Reference(ws, min_row=1, max_row=ws.max_row, min_col=col + 1, max_col=col + 1)
            valueCarbonDioxide = Reference(ws, min_row=1, max_row=ws.max_row, min_col=col + 2, max_col=col + 2)
            valueHelium = Reference(ws, min_row=1, max_row=ws.max_row, min_col=alignNum + 3, max_col=alignNum + 3)

            ws.column_dimensions[get_column_letter(alignNum)].width = 65

            oxygenChart = LineChart()
            oxygenChart.add_data(valueOxygen, titles_from_data=True)
            oxygenChart.set_categories(valueTime)
            # stylesheet
            oxygenChart.x_axis.title = "Time(s)"
            oxygenChart.y_axis.title = "O2(%)"
            oxygenChart.height = 13
            oxygenChart.width = 19
            
            ws.add_chart(oxygenChart, f"{get_column_letter(alignNum)}{row}")

            carbonDioxideChart = LineChart()
            carbonDioxideChart.add_data(valueCarbonDioxide, titles_from_data=True)
            carbonDioxideChart.set_categories(valueTime)
            # stylesheet
            carbonDioxideChart.x_axis.title = "Time(s)"
            carbonDioxideChart.y_axis.title = "CO2(ppm)"
            carbonDioxideChart.height = 13
            carbonDioxideChart.width = 19

            ws.add_chart(carbonDioxideChart, f"{get_column_letter(alignNum)}{row + 22}")

            heliumChart = LineChart()
            heliumChart.add_data(valueHelium, titles_from_data=True)
            heliumChart.set_categories(valueTime)
            # stylesheet
            heliumChart.x_axis.title = "Time(s)"
            heliumChart.y_axis.title = "He(%)"
            heliumChart.height = 13
            heliumChart.width = 19

            ws.add_chart(heliumChart, f"{get_column_letter(alignNum)}{row + 44}")
            
            col += 4
            # 5*6 배열 사용시 주석해제
            if alignNum >= 20:
                alignNum = 4
                row += 68
                continue
            alignNum += 4
            # 5*6 여기까지
            

    elif csvDataSheet.type == "CO2" or csvDataSheet.type == "He":
        # 차트추가
        # try:
            ws = wb.active
            valueCarbonDioxide = Reference(ws, min_row=1, max_row=ws.max_row, min_col=6, max_col=6)
            valueTime = Reference(ws, min_row=2, max_row=ws.max_row, min_col=3, max_col=3)

            carbonDioxideChart = LineChart()
            carbonDioxideChart.add_data(valueCarbonDioxide, titles_from_data=True)
            carbonDioxideChart.set_categories(valueTime)
            # stylesheet
            carbonDioxideChart.x_axis.title = "Time(s)"
            carbonDioxideChart.y_axis.title = '이산화탄소(ppm)' if csvDataSheet.type == 'CO2' else '헬륨(%)'
            carbonDioxideChart.height = 13
            carbonDioxideChart.width = 19

            ws.add_chart(carbonDioxideChart, "B6")






if __name__ == "__main__":
    pass