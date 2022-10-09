from openpyxl import load_workbook
import ipaddress

xlsFile = "C:/00_STUDY/WORK/01_Projects/IMS PP/cm_sz_nolinks.xlsx"
wb = load_workbook(xlsFile)

def is_ip(address):
    try:
        ipaddress.ip_address(address)
        return True
    except ValueError:
        return False

ip_list=[]

for row in wb['AAA']['G']:
    if is_ip(row.value):
        ip_list.append(row.value)

for row in wb['AAA']['O']:
    if is_ip(row.value):
        ip_list.append(row.value)

ip_list=list(set(ip_list))

print(ip_list)