import openpyxl
import requests
import xml.etree.ElementTree as ET

EXCEL_FILE_NAME = "366.xlsx"
EXCEL_SHEET_NAME = "Master"
YAHOO_API = "https://jlp.yahooapis.jp/FuriganaService/V1/furigana"
YAHOO_CLIENT_ID_FILE = "yahoo_client_id.txt"
YAHOO_SCHEMALOCATION = "{urn:yahoo:jp:jlp:FuriganaService}"
YAHOO_XML_TAG_WORD = YAHOO_SCHEMALOCATION+"Word"
YAHOO_XML_TAG_SURFACE = YAHOO_SCHEMALOCATION+"Surface"
YAHOO_XML_TAG_FURIGANA = YAHOO_SCHEMALOCATION+"Furigana"
GRADE = "1"

def ruby_func(separator):
    def inner(honbun, furigana):
        return separator[0] + honbun + separator[1] + furigana + separator[2]
    return inner
SEPARATOR = ('[', '/', ']')
#SEPARATOR = ("<ruby>", "<rt>", "</rt></ruby>")
#SEPARATOR = ("<ruby>", "<rp>(</rp><rt>", "</rt><rp>)</rp></ruby>")
ruby = ruby_func(SEPARATOR)
    
#Excelファイルを開いてシートをコピー
wb = openpyxl.load_workbook(EXCEL_FILE_NAME)
ws = wb.copy_worksheet(wb[EXCEL_SHEET_NAME])

#全てのセルに対してYahooのルビ振りAPIを使ってルビを振る
fin = open(YAHOO_CLIENT_ID_FILE, 'rt')
yahoo_client_id = fin.readline()
fin.close()

for row in ws.rows:
    for column in row:
        sentence = column.value
        if sentence:

            # Yahoo API の呼び出し
            params = {"appid": yahoo_client_id, "grade": GRADE, "sentence": sentence}
            response = requests.get(YAHOO_API, params)
            xml_data = response.text

            # APIレスポンスのXMLを処理
            xml_root = ET.fromstring(xml_data)
            rubied_sentence = ""
            for word in xml_root.iter(YAHOO_XML_TAG_WORD):
                surface = word.find(YAHOO_XML_TAG_SURFACE)
                furigana = word.find(YAHOO_XML_TAG_FURIGANA)
                if surface is not None:
                    if furigana is not None:
                        #ルビ振りのタグを追加
                        rubied_sentence += ruby(surface.text, furigana.text)
                    else:
                        rubied_sentence += (surface.text)
            
            #ルビを振ったテキストを書き戻し
            print(rubied_sentence)
            column.value = rubied_sentence

#Excelの変更を保存
wb.save(EXCEL_FILE_NAME)