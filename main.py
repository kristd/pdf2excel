# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


import pdfplumber
import pandas as pd
import os
import glob
import datetime
import re
import calendar


def get_purchase_order_first_page_text(i):
    pdf_object = pdfplumber.open(i)
    first_page = pdf_object.pages[0]
    return first_page.extract_text_simple()

def get_sizecolourbreakdown_pages_text(i):
    pdf_object = pdfplumber.open(i)
    return pdf_object.pages


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

def get_detail_country_code(text):
    country_code = re.findall(r' [a-zA-Z]{2} \(',text)[0].replace(')','').replace('(','').strip()
    return country_code
def get_delivery_dates_dicts(text,p1, p2):
    time_delivery_dict = {}
    time_delivery_dict.clear()
    for p in range(p1,  p2):
        if re.search(r'\d+', text[p]) is not None:
            if re.search(r'\d+', text[p + 1]) is not None:
                l_day = text[p][0:2]
                l_mon = list(calendar.month_abbr).index(text[p][3:6])
                l_year = text[p][8:12]
                time_delivery_dict[text[p].replace('\xa0', '')[13:].replace(
                    re.findall(r'\d+ .*%', text[p].replace('\xa0', '')[13:])[0], '').strip()] = str(
                    l_year) + '-' + str(l_mon) + '-' + str(l_day)
        else:
            l_day = text[p - 1][0:2]
            l_mon = list(calendar.month_abbr).index(text[p - 1][3:6])
            l_year = text[p - 1][8:12]
            time_delivery_dict[text[p - 1].replace('\xa0', '')[13:].replace(
                re.findall(r'\d+ .*%', text[p - 1].replace('\xa0', '')[13:])[0], '').strip() +
                               text[p].replace('\xa0', '').strip()] = str(l_year) + '-' + str(
                l_mon) + '-' + str(l_day)
    return time_delivery_dict

def get_price_dicts(text,p1,p2):
    price_dict = {}
    price_dict.clear()
    for p in range( p1, p2):
        l_price = text[p].split(' ')[0]
        l_currency = text[p].split(' ')[1]
        l_countries = text[p].replace(text[p].split(' ')[0], '').replace(
            text[p].split(' ')[1], '').strip()
        price_dict[l_countries] = str(l_price) + '-' + str(l_currency)
    return price_dict

def get_term_dicts(text,p1,p2):
    term_dicts={}
    term_dicts.clear()
    df = pd.read_excel("/Users/Kristd/Documents/github/pdf2excel/country_code.xlsx")
    df_li = df.values.tolist()
    country_codes=[]
    for s_li in df_li:
        country_codes.append(s_li[0] + '-' + s_li[1])

    for p in range(p1,p2):
        country_list = text[p].split(',')
        for cl in country_list:
            if re.search(r'/', cl.strip()) is not None:
                for c in country_codes:
                    if  re.search(cl.strip(),c) is not None:
                        term_dicts[c[3:]] = text[p + 1].split(' ')[2]
                        break
            else:
                for c in country_codes:
                    if re.search(cl.strip(),c) is not None:
                        term_dicts[c] = text[p + 1].split(' ')[2]
                        break
    return term_dicts

def get_colourname_dicts(text, p1, p2):
    colourname_dicts = {}
    colourname_dicts.clear()
    for p in range(p1,p2):
        colourname_dicts[text[p].split(' ')[0]] = re.findall(r'[a-zA-Z]+.*[a-zA-Z]+',text[p].replace('Solid','').replace('USD','').replace('All over pattern',''))[0]
    return colourname_dicts

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
                    print(orderNum)
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
                    download_date = datetime.date.today().strftime('%Y-%m-%d') #Column V


                    ### create the price/country mapping information
                    terms_1st_position = 0
                    terms_last_position = 0
                    time_delivery_1st_position = 0
                    time_delivery_last_position = 0
                    price_1st_position = 0
                    price_last_position = 0
                    colourname_1st_position = 0
                    colourname_last_position = 0
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
                        if (re.search(r'By accepting',order_info_array[p]) is not None) or (re.search(r'License Order - Pool Party',order_info_array[p]) is not None):
                            price_last_position = p-1
                        if (re.search(r'Article No H&M Colour Code',order_info_array[p])):
                            colourname_1st_position = p
                        if (re.search(r'Total Quantity:',order_info_array[p])):
                            colourname_last_position = p-1

                    ### create the country/term mapping
                    term_dict = {}
                    term_dict.clear()
                    term_dict = get_term_dicts(order_info_array, terms_1st_position+1,terms_last_position+1)



                    ##create the time delivery mapping
                    time_delivery_dict = {}
                    time_delivery_dict.clear()
                    time_delivery_dict = get_delivery_dates_dicts(order_info_array,time_delivery_1st_position+1,time_delivery_last_position+1)

                    ##create price mapping
                    price_dict = {}
                    price_dict.clear()
                    price_dict = get_price_dicts(order_info_array,price_1st_position+1,price_last_position+1)

                    ##create colourname mapping
                    colourname_dict = {}
                    colourname_dict.clear()
                    colourname_dict = get_colourname_dicts(order_info_array, colourname_1st_position+1,colourname_last_position+1)

                    ###detail loop start here
                    detail_pages = []
                    detail_pages = get_sizecolourbreakdown_pages_text(
                        i.replace('PurchaseOrder', 'SizePerColourBreakdown'))
                    for dp in range(0,len(detail_pages)):
                        country_name = ''  # Column K
                        fright_term = ''  # Column L
                        cost = ''  #Column Q
                        currency = '' # Column R
                        page_text = detail_pages[dp].extract_text_simple()
                        order_detail_array = split_lines(page_text)

                        assortment_1st_position = 0
                        assortment_last_position = 0
                        solid_1st_position = 0
                        solid_last_position = 0

                        for p in range(0,len(order_detail_array)):
                            if (re.search(r'Assortment', order_detail_array[p])):
                                assortment_1st_position = p
                            if (re.search(r'Solid', order_detail_array[p])):
                                solid_1st_position = p
                                assortment_last_position = p - 1
                            if (re.search(r'Total', order_detail_array[p])):
                                solid_last_position = p - 1

                        country_code = get_detail_country_code(order_detail_array[11])
                        ## get term

                        for k in term_dict.keys():
                            if re.search(country_code,str(k)) is not None:
                                country_name = k
                                fright_term = term_dict[k]
                                break
                        ## get cost and currency
                        for k in price_dict.keys():
                            if re.search(country_code,str(k)) is not None:
                                cost = price_dict[k].split('-')[0]
                                currency = price_dict[k].split('-')[1]

                        ##get artical NO & colourcode+colourname
                        artical_list = re.findall(r'\d+.*\d+',order_detail_array[12])[0].split(' ')
                        colourcode_list = re.findall(r'\d+-.*',order_detail_array[13])[0].split(' ')
                        print(country_name)
                        for a in artical_list:
                            artical_no = ''  # Column U
                            colourcode_colourname = ''# Column M
                            size = ''#Column N
                            packing_type = ''# Column O
                            qty_artical = '' # Column P
                            artical_no = a
                            for k in colourname_dict.keys():
                                if a == k:
                                    colourcode_colourname =  colourcode_list[artical_list.index(a)]+' ' + str(colourname_dict[a])
                                    break
                            ## assorment/solid qty information
                            if assortment_1st_position > 0:
                                packing_type = 'Assortment'
                                print(packing_type)
                                no_of_asst_list = []
                                for ap in range(assortment_1st_position+1,assortment_last_position+1):
                                    if re.search(r'No of Asst:',order_detail_array[ap]) is not None:
                                        no_of_asst_list = order_detail_array[ap].replace('No of Asst: ', '').split(' ')
                                        break
                                for ap in range(assortment_1st_position+1,assortment_last_position+1):
                                    if re.search(r'\(.*\)\*', order_detail_array[ap]) is not None:
                                        size = re.findall(r'\(.*\)\*', order_detail_array[ap])[0].replace(')*', '').replace('(','')
                                        if re.findall(r'\* \d+.*', order_detail_array[ap]) == []:
                                            next
                                        else:
                                            if re.findall(r'\* \d+.*', order_detail_array[ap]) is not []:
                                                qty_artical = int(
                                                    re.findall(r'\* \d+.*', order_detail_array[ap])[0].replace('* ',
                                                                                                               '').split(
                                                        ' ')[artical_list.index(a)])
                                                print(size + ': ' + str(qty_artical))
                                                #######write data into excel sheet
                            if solid_1st_position > 0:
                                packing_type = 'Solid'
                                print(packing_type)
                                for sp in range(solid_1st_position+1,solid_last_position+1):
                                    if re.search(r'\(.*\)\*', order_detail_array[sp]) is not None:
                                        size = re.findall(r'\(.*\)\*', order_detail_array[sp])[0].replace(')*', '').replace('(','')
                                        #print(re.findall(r'\* \d+.*', order_detail_array[sp])[0].replace('* ',''))
                                        if re.findall(r'\* \d+.*',order_detail_array[sp]) == []:
                                            next
                                        else:
                                            if re.findall(r'\* \d+.*',order_detail_array[sp]) is not []:
                                                qty_artical = int(re.findall(r'\* \d+.*',order_detail_array[sp])[0].replace('* ','').split(' ')[artical_list.index(a)])
                                                print(size + ': ' + str(qty_artical))
                                                #######write data into excel sheet

                    ######end of detail loop
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

    ##end of outer for loop

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
