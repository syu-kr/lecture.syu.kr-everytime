#
#     ____                  __________
#    / __ \_   _____  _____/ __/ / __ \_      __
#   / / / / | / / _ \/ ___/ /_/ / / / / | /| / /
#  / /_/ /| |/ /  __/ /  / __/ / /_/ /| |/ |/ /
#  \____/ |___/\___/_/  /_/ /_/\____/ |__/|__/
# 
#  The copyright indication and this authorization indication shall be
#  recorded in all copies or in important parts of the Software.
# 
#  @author 0verfl0w767
#  @link https://github.com/0verfl0w767
#  @license MIT LICENSE
#
import os
import json
import openpyxl
from openpyxl.styles import Color, PatternFill

YEAR = "2023"
SEMESTER = "2학기 정규"

DATA_FOLDER_NAME = "data/"

if not os.path.exists(os.path.join(os.path.dirname(__file__), DATA_FOLDER_NAME)):
  os.makedirs(os.path.join(os.path.dirname(__file__), DATA_FOLDER_NAME))

API_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), DATA_FOLDER_NAME + "allAPI.json"))

data_path = ""

with open("config.json", "r", encoding = "utf-8") as f:
  JSON_DATA = json.load(f)
  data_path = JSON_DATA["path"]

REAL_PATH = data_path + "\\" + YEAR + "\\" + SEMESTER
ABS_PATH_1 = os.path.abspath(REAL_PATH)

SAVE_NAME = YEAR + "년 " + SEMESTER + " 시간표"
SAVE_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), DATA_FOLDER_NAME + SAVE_NAME + ".xlsx"))

allAPI = []

for college in os.listdir(ABS_PATH_1):
  ABS_PATH_2 = os.path.abspath(ABS_PATH_1 + "\\" + college)
  
  if college == "전체대학":
    continue
  
  if college == "학부(과).json":
    continue
  
  for undergraduate in os.listdir(ABS_PATH_2):
    ABS_PATH_3 = os.path.abspath(ABS_PATH_2 + "\\" + undergraduate)
    
    with open(ABS_PATH_3, "r", encoding = "utf-8") as f:
      DATA = json.load(f)
      
      for realData in DATA["api"]:
        realData["단과대학"] = college
        allAPI.append(realData)

apiJson = {}
apiJson["year"] = YEAR
apiJson["semester"] = SEMESTER
apiJson["api"] = allAPI

with open(API_PATH, "w", encoding = "utf-8") as f:
  json.dump(apiJson, f, ensure_ascii = False, indent = 2)

excelWB = openpyxl.Workbook()
sheet = excelWB.active

sheet.column_dimensions["C"].width = 35 # 과목명
sheet.column_dimensions["D"].width = 18 # 학부(과)
sheet.column_dimensions["G"].width = 15 # 영역구분
sheet.column_dimensions["I"].width = 15 # 교수명
sheet.column_dimensions["J"].width = 15 # 수업시간
sheet.column_dimensions["K"].width = 30 # 장소

sheet["A1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["B1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["C1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["D1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["E1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["F1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["G1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["H1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["I1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["J1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["K1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))
sheet["L1"].fill = PatternFill(fill_type="solid", fgColor=Color("29CDFF"))

sheet.freeze_panes = "A2"

sheet["A1"] = "강좌번호"
sheet["B1"] = "과목코드"
sheet["C1"] = "과목명"
sheet["D1"] = "학부(과)"
sheet["E1"] = "학년"
sheet["F1"] = "이수구분"
sheet["G1"] = "영역구분"
sheet["H1"] = "학점"
sheet["I1"] = "교수명"
sheet["J1"] = "수업시간"
sheet["K1"] = "장소"
sheet["L1"] = "단과대학"

with open(API_PATH, "r", encoding = "utf-8") as f:
  DATA = json.load(f)
  rowCount = 1
  
  for realData in DATA["api"]:
    rowCount += 1
    
    sheet["A" + str(rowCount)] = realData["강좌번호"]
    sheet["B" + str(rowCount)] = realData["과목코드"]
    sheet["C" + str(rowCount)] = realData["과목명"]
    sheet["D" + str(rowCount)] = realData["학부(과)"]
    sheet["E" + str(rowCount)] = realData["학년"]
    sheet["F" + str(rowCount)] = realData["이수구분"]
    sheet["G" + str(rowCount)] = realData["영역구분"]
    sheet["H" + str(rowCount)] = realData["학점"]
    sheet["I" + str(rowCount)] = realData["교수명"]
    sheet["J" + str(rowCount)] = realData["수업시간"]
    sheet["K" + str(rowCount)] = realData["장소"]
    sheet["L" + str(rowCount)] = realData["단과대학"]

excelWB.save(SAVE_PATH)
excelWB.close()