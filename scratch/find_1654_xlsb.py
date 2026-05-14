from pyxlsb import open_workbook
try:
    with open_workbook('PLANILHA BLOCOS.xlsb') as wb:
        with wb.get_sheet(1) as sheet: # Usually the first sheet
            headers = None
            for row in sheet.rows():
                if not headers:
                    headers = [str(c.v).strip().upper() for c in row]
                    print("Headers:", headers)
                    continue
                
                row_dict = dict(zip(headers, [c.v for c in row]))
                # Look for 1654
                bloco_val = str(row_dict.get('BLOCO', '')).split('.')[0]
                if bloco_val == '1654':
                    print("Found block 1654 in PLANILHA BLOCOS:")
                    print(row_dict)
                    break
            else:
                print("Block 1654 NOT FOUND in PLANILHA BLOCOS.")
except Exception as e:
    print("Error:", e)
