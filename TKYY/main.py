from FunctionTools import *
import datetime
import os
from concurrent.futures import ThreadPoolExecutor



if __name__ == '__main__':

    date = datetime.date.today().isoformat()
    Path = ('output/'+date)
    out_path = 'output/zip'
    zip_name = f'TKYY_Automatic_Backup_{date}.zip'
    zip_names = out_path +'/'+ zip_name

    if not os.path.exists(Path):
        os.makedirs(Path)

    file_name = 'output/' + 'switch_information_' + date + '.xlsx'
    sheet_name = 'Sheet1'

    create_excel(sheet_name, file_name)

    IF = getip()
    ip     = IF[0]
    type   = IF[1]
    user   = IF[2]
    paswd  = IF[3]

    with ThreadPoolExecutor(max_workers=50) as p:
        p.map(Check,ip,type,user,paswd)

    code = verify_IP()

    if code == 1:
        fillter_data()
        change_xl_style(sheet_name,file_name)
        zip_file_path(Path, out_path, zip_name)
        email(date, zip_names)

    else:
        print(f'Some one device has an outage{date}')



