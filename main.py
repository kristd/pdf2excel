# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


import pdfplumber
import pandas
import os
import glob
import datetime
import re
import calendar


def get_purchase_order_first_page_text(i):
    pdf_object = pdfplumber.open(i)
    first_page = pdf_object.pages[0]
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
    dateOfOrder =  re.findall(r'\d+ [a-zA-Z]+, \d{4}',text)[0]
    l_day = re.findall(r'\d{2} ',dateOfOrder)[0].strip()
    l_mon = list(calendar.month_abbr).index(re.findall(r'[a-zA-Z]{3},',dateOfOrder)[0].replace(',',''))
    l_year = re.findall(r'\d{4}',dateOfOrder)[0]
    return str(l_year)+'-' + str(l_mon)+'-' +str(l_day)

def get_product_name(text):
    return re.findall(r'Product Name: .*',text)[0].split(':')[1].strip()

def get_development_no(text):
    return re.findall(r'Development .*',text)[0].split('  ')[1]

def get_no_of_pieces(text):
    return re.findall(r'\d+',text)[0]

# with pdfplumber.open('/Users/kristd/Downloads/2PDF/667098_PurchaseOrder_20230627_144321.pdf') as pdf:
#     for page in pdf.pages:
#         print(page.extract_text_simple())
#         # 每页打印一分页分隔
#         print('---------- 分页分隔 ----------')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #get PurchaseOrder files list
    formatted_today = datetime.date.today().strftime('%Y%m%d')
    file_path= '/Users/Kristd/Documents/github/pdf2excel/In/'
    f = glob.glob(file_path+'*PurchaseOrder*')
    pattern = 'UPDATE_*'
    for i in f:  #PurchaseOrder loop, the outer loop
        file_name = os.path.basename(i)
        if re.search(pattern,file_name.upper()) is None:
            orderNum = file_name.split('_',1)[0]
            ##print(orderNum)
            try:
                #print(get_file_validation(file_path, orderNum, formatted_today,1))
                if get_file_validation(file_path, orderNum, formatted_today,1):
                    # get the PurchaseOrder info and go down to the detail loop!!!!!!!!!!
                    ####New Order Scenario
                    print('NEW Order Scenario')
                    order_text = get_purchase_order_first_page_text(i)
                    order_info_array = split_lines(order_text)
                    brand = get_brand(order_info_array[1]) #Column A
                    order_type = get_order_type(order_info_array[0])  #Column B
                    season = get_season(order_info_array[6]) #Column C
                    order_no = get_order_no(order_info_array[3]) #Column D
                    department_no = get_department_no(order_info_array[3])   #Column E
                    product_no = get_product_no(order_info_array[3])  #Column F
                    date_of_order = get_date_of_order(order_info_array[5]) #Column G
                    product_name = get_product_name(order_info_array[4]) #Column H
                    development_no = get_development_no(order_info_array[9])  #Column I
                    no_of_pieces  = get_no_of_pieces(order_info_array[12])#Column J


                    ### create the price/country mapping information
                    terms_1st_position = 0
                    terms_last_position = 0
                    time_delivery_1st_position = 0
                    time_delivery_last_position = 0
                    price_1st_position = 0
                    price_last_position = 0
                    for p in range(0,len(order_info_array)):
                        if re.search(r'Terms of Delivery',order_info_array[p]) is not None:
                            terms_1st_position = p
                        if re.search(r'Time of Delivery Planning Markets',order_info_array[p]) is not None:
                            terms_last_position = p-1
                            time_delivery_1st_position = p
                        if re.search(r'Total: \d+',order_info_array[p]) is not None:
                            time_delivery_last_position = p-1
                        if re.search(r'Invoice Average Price',order_info_array[p]) is not None:
                            price_1st_position = p
                        if re.search(r'By accepting',order_info_array[p]) is not None:
                            price_last_position = p-1

                    print('')
                    print('-----------')

                else:
                    # throw error as the child file is not exists.
                    f = open(file_path + str(orderNum) + '_SizePerColourBreakdown*')
            except FileNotFoundError:
                print('OrderNum: ' + orderNum)
                print('')
                print('File is not found!')
                print('-----------')
        else:
            orderNum = file_name.split('_', 2)[1]
            ##print(orderNum)
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


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
