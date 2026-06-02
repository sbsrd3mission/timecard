import re
import base64
import openpyxl
import io
import sys

def verify():
    sys.stdout.reconfigure(encoding='utf-8')
    content = open('timecard.html', 'r', encoding='utf-8').read()
    m = re.search(r"const EXCEL_TEMPLATE_B64 = '(.*?)';", content, re.DOTALL)
    b = base64.b64decode(m.group(1))
    wb = openpyxl.load_workbook(io.BytesIO(b), data_only=False)
    print("Formula in J21:", wb['マスタ']['J21'].value)

if __name__ == '__main__':
    verify()
