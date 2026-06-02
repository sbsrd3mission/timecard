import base64
import re

def update_b64():
    # Read the updated excel file
    with open('タイムカード入力空バージョン.xlsx', 'rb') as f:
        b64 = base64.b64encode(f.read()).decode('utf-8')
    
    # Read timecard.html
    with open('timecard.html', 'r', encoding='utf-8') as f:
        content = f.read()
        
    # Replace the EXCEL_TEMPLATE_B64
    # The string looks like: const EXCEL_TEMPLATE_B64 = `...`;
    content = re.sub(r'const EXCEL_TEMPLATE_B64 = `.*?`;', f'const EXCEL_TEMPLATE_B64 = `{b64}`;', content, flags=re.DOTALL)
    
    # Write back
    with open('timecard.html', 'w', encoding='utf-8') as f:
        f.write(content)
        
    print("Successfully updated EXCEL_TEMPLATE_B64 in timecard.html")

if __name__ == '__main__':
    update_b64()
