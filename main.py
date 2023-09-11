# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


import pdfplumber
import pandas
import shutil
import re
import os
import glob
import datetime
import re


def get_purchase_order_first_page_text(orderNum):
    pdf_object = pdfplumber.open(file_name)
    first_page = pdf_object.pages[1]
    return first_page.extract_text_simple()


def get_file_validation(file_path,orderNum,formatted_today,flag):
    if flag == 1:
        f = glob.glob(file_path + str(orderNum) +'_SizePerColourBreakdown*')
        if len(f) > 0:
            return True
        else:
            return False
    else:
        f = glob.glob(file_path + 'Updated_' + str(orderNum) +'_SizePerColourBreakdown*')
        if len(f) > 0:
            return True
        else:
            return False

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
    file_path= '/Users/Kristd/TJ/Nutstore Files/Project/python/pdf2excel/In/'
    #print(file_path)
    f = glob.glob(file_path+'*PurchaseOrder*')
    pattern = 'UPDATE_*'
    for i in f:  #PurchaseOrder loop, the outer loop
        file_name = os.path.basename(i)
        if re.search(pattern,file_name.upper()) is None:
            orderNum = file_name.split('_',1)[0]
            print(orderNum)
            try:
                #print(get_file_validation(file_path, orderNum, formatted_today,1))
                if get_file_validation(file_path, orderNum, formatted_today,1):
                    # get the PurchaseOrder info and go down to the detail loop!!!!!!!!!!
                    ####New Order Scenario
                    print('NEW Order Scenario')
                    order_text = get_purchase_order_first_page_text(i)
                    order_info_array = split_lines(order_text)

                    brand = get_brand(order_info_array[1])
                    order_type = get_order_type(order_info_array[0])
                    season = get_season(order_info_array[6])
                    order_no = get_order_no(order_info_array[3])
                    department_no = get_department_no(order_info_array[3])
                    product_no = get_product_no(order_info_array[3])

                    print(brand)

                else:
                    # throw error as the child file is not exists.
                    f = open(file_path + str(orderNum) + '_SizePerColourBreakdown*')
            except FileNotFoundError:
                print('File is not found!')
        else:
            orderNum = file_name.split('_', 2)[1]
            print(orderNum)
            try:
                #print(get_file_validation(file_path,orderNum,formatted_today,2))
                if get_file_validation(file_path,orderNum,formatted_today,2):
                    ##get the PurchaseOrder info and go down to the detail loop!!!!!!!!!!
                    ####Update Order Scenarios
                    print('UPDATED Order Senario')
                else:
                    #throw error as the child file is not exists.
                    f = open(file_path + str(orderNum) +'_SizePerColourBreakdown*')
            except FileNotFoundError:
                print('File is not found!')

    #logic to remove files in Archive folder older than 1 year
    archive_path = '/Users/Kristd/TJ/Nutstore Files/Project/python/pdf2excel/Archive/'
    dirs = next(os.walk(base_path))[1]
    last_year_date = (datetime.date.today()+datetime.timedelta(days=-365)).strftime('%Y%m%d')
    f2 = glob.glob(archive_path + '*PurchaseOrder*')
    for i_d in f2:
        file_name = os.path.basename(i_d)
        if re.search(pattern, file_name.upper()) is None:
            l_date = file_name.split('_',4)[2]
            #print(l_date)
        else:
            l_date = file_name.split('_', 5)[3]
            #print(l_date)

        if l_date < last_year_date:
            os.remove(i_d)
            i_sd = i_d.replace('PurchaseOrder','SizePerColourBreakdown')
            os.remove(i_sd)
            print('Files have been removed')

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
