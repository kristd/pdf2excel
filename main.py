# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

#import pandas
import pdfplumber
import shutil
import re
import os
import glob
import datetime


def get_purchase_order_first_page_text(file_name):
    pdf_object = pdfplumber.open(file_name)
    first_page = pdf_object.pages[1]
    return first_page.extract_text_simple()


def get_file_validation(file_path,file_format):
    return


def split_lines(text):
    return text.splitlines()


def get_brand(text):
    return text


def get_order_type(text):
    if text.find('Online') == -1:
        return 'Store'
    else:
        return 'Online'


def get_season(text):
    return 'S' + re.findall(r'\d+-\d+', text)[0]


def get_order_no(text):
    return re.findall(r'\d+', text)[0]


def get_department_no(text):
    return re.findall(r'\d+', text)[1]


def get_product_no(text):
    return re.findall(r'\d+', text)[2]


def get_date_of_order(text):
    return re.findall(r'\d+', text)[2]


# with pdfplumber.open('/Users/kristd/Downloads/2PDF/667098_PurchaseOrder_20230627_144321.pdf') as pdf:
#     for page in pdf.pages:
#         print(page.extract_text_simple())
#         # 每页打印一分页分隔
#         print('---------- 分页分隔 ----------')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #get PurchaseOrder files list
    base_path = '/Users/Kristd/TJ/Nutstore Files/Project/python/pdf2excel/In/'
    formatted_today = datetime.date.today().strftime('%Y%m%d')
    file_path= '/Users/Kristd/TJ/Nutstore Files/Project/python/pdf2excel/In/' + formatted_today +'/'
    print(file_path)
    f = glob.glob(file_path+'*PurchaseOrder*')
    for i in f:  #PurchaseOrder loop, the outer loop
        file_name = os.path.basename(i)
        orderNum = file_name.split('_',1)[0]
        print(orderNum)


    #logic to move the folder which is older than 1 year to Archive folder
    archive_path = '/Users/Kristd/TJ/Nutstore Files/Project/python/pdf2excel/Archive'
    dirs = next(os.walk(base_path))[1]
    last_year_date = (datetime.date.today()+datetime.timedelta(days=-365)).strftime('%Y%m%d')
    for i_dirs in dirs:
        if i_dirs < last_year_date:
            shutil.move(base_path+i_dirs,archive_path)
            print('have move ' + i_dirs + ' to Archive folder')

    '''
    order_text = get_purchase_order_first_page_text('/Users/kristd/Downloads/2PDF/667098_PurchaseOrder_20230627_144321.pdf')
    order_info_array = split_lines(order_text)

    # --
    print(order_info_array)

    # --
    brand = get_brand(order_info_array[1])
    order_type = get_order_type(order_info_array[0])
    season = get_season(order_info_array[6])
    order_no = get_order_no(order_info_array[3])
    department_no = get_department_no(order_info_array[3])
    product_no = get_product_no(order_info_array[3])

    print(product_no)
    
    '''

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
