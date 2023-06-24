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

from utils.Logger import Logger

LOGGER = Logger()

LOGGER.logo()

RED_TEXT = "\033[1;91m"
GREEN_TEXT = "\033[1;92m"
YELLOW_TEXT = "\033[1;93m"
RESET_TEXT = "\033[0m"

year = ""
semester = ""
folder_path = ""
data_path = ""

with open("config.json", "r", encoding = "utf-8") as f:
  JSON_DATA = json.load(f)
  year = JSON_DATA["year"]
  semester = JSON_DATA["semester"]
  folder_path = JSON_DATA["folder-path"]
  data_path = JSON_DATA["data-path"]

if not os.path.exists(os.path.join(os.path.dirname(__file__), folder_path)):
  os.makedirs(os.path.join(os.path.dirname(__file__), folder_path))

REAL_PATH = data_path + "\\" + year + "\\" + semester
ABS_PATH_1 = os.path.abspath(REAL_PATH)

FILE_NAME = year + "년 " + semester + " 시간표"
XLSX_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), folder_path + FILE_NAME + ".xlsx"))
JSON_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), folder_path + FILE_NAME + ".json"))

LOGGER.info(FILE_NAME)

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
      PREV_YEAR = int(ABS_PATH_3.split("\\")[7]) - 1
      PREV_PATH = ABS_PATH_3.replace(year, str(PREV_YEAR))
      
      with open(PREV_PATH, "r", encoding = "utf-8") as f:
        PREV_DATA = json.load(f)
        
        if not DATA["api"]:
          if not PREV_DATA["api"]:
            LOGGER.info(os.path.splitext(undergraduate)[0] + " 정보가 " + RED_TEXT + "확인되지 않음. (이전 연도 정보가 확인되지 않음.)" + RESET_TEXT)
          else:
            LOGGER.info(os.path.splitext(undergraduate)[0] + " 정보가 " + RED_TEXT + "확인되지 않음. " + GREEN_TEXT + "(이전 연도 " + str(len(PREV_DATA["api"])) + "개 정보가 확인됨.)" + RESET_TEXT)
          continue
        
        LOGGER.info(os.path.splitext(undergraduate)[0] + " 정보가 " + str(len(DATA["api"])) + "개 " + GREEN_TEXT + "확인 됨. " + YELLOW_TEXT + "(이전 연도 " + str(len(PREV_DATA["api"])) + "개 정보가 확인됨.)" + RESET_TEXT)
      
      for realData in DATA["api"]:
        realData["단과대학"] = college
        allAPI.append(realData)

apiJson = {}
apiJson["year"] = year
apiJson["semester"] = semester
apiJson["api"] = allAPI

with open(JSON_PATH, "w", encoding = "utf-8") as f:
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

with open(JSON_PATH, "r", encoding = "utf-8") as f:
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

excelWB.save(XLSX_PATH)
excelWB.close()

LOGGER.info(XLSX_PATH)
LOGGER.info(JSON_PATH)