import boto3
import json
import ast
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
from botocore.exceptions import ClientError
import re

# config = Config(
#     region_name = 'us-gov-west-1',
#     signature_version = 'v4',
#     retries = {
#         'max_attempts': 10,
#         'mode': 'standard'
#     }
# )
def find_whole_word(word, string):
    return re.search(r'\b' + re.escape(word) + r'\b', string, re.IGNORECASE)

def current_price(ServiceCode,usagecode,regionCode,usagetype, byol=False):

    region_name='us-east-1'
    instance_type='r5a.xlarge'
    preinstalled_software='NA'
    os='Linux'
    tenancy="Shared"
    byol=False
    #usagetype='UGW1-CPUCredits:t3a'
    #regionCode='us-gov-east-1'
    filters1= [{}]

    if usagecode in ['UGW1-BoxUsage']:
            filters1 = [
                {'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
                {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagetype},
                #{'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': region_name},
                #{'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': instance_type},
                {'Type': 'TERM_MATCH', 'Field': 'tenancy', 'Value': tenancy},
                {'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': os},
                {'Type': 'TERM_MATCH', 'Field': 'preInstalledSw','Value': preinstalled_software},
                {'Type': 'TERM_MATCH', 'Field': 'regionCode','Value': regionCode},
                {'Type': 'TERM_MATCH', 'Field': 'licenseModel','Value': 'Bring your own license' if byol else 'No License required'},
                ]
    elif usagecode in ['UGW1-EBS']:
            filters1 = [
                #{'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
                {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagetype},
                #{'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': region_name},
                #{'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': instance_type},
                #{'Type': 'TERM_MATCH', 'Field': 'tenancy', 'Value': tenancy},
                #{'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': os},
                #{'Type': 'TERM_MATCH', 'Field': 'preInstalledSw','Value': preinstalled_software},
                {'Type': 'TERM_MATCH', 'Field': 'regionCode','Value': regionCode},
                #{'Type': 'TERM_MATCH', 'Field': 'licenseModel','Value': 'Bring your own license' if byol else 'No License required'},
                ]
    elif usagecode in ['UGW1-RDS']:
        filters1 = [
                #{'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
                {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagetype},
                #{'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': region_name},
                #{'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': instance_type},
                #{'Type': 'TERM_MATCH', 'Field': 'tenancy', 'Value': tenancy},
                #{'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': os},
                #{'Type': 'TERM_MATCH', 'Field': 'preInstalledSw','Value': preinstalled_software},
                {'Type': 'TERM_MATCH', 'Field': 'regionCode','Value': regionCode},
                #{'Type': 'TERM_MATCH', 'Field': 'licenseModel','Value': 'Bring your own license' if byol else 'No License required'},
                ]
    elif find_whole_word("Fargate", usagecode):
        filters1 = [
                #{'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
                {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagetype},
                #{'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': region_name},
                #{'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': instance_type},
                #{'Type': 'TERM_MATCH', 'Field': 'tenancy', 'Value': tenancy},
                #{'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': os},
                #{'Type': 'TERM_MATCH', 'Field': 'preInstalledSw','Value': preinstalled_software},
                {'Type': 'TERM_MATCH', 'Field': 'regionCode','Value': regionCode},
                #{'Type': 'TERM_MATCH', 'Field': 'licenseModel','Value': 'Bring your own license' if byol else 'No License required'},
                ]
    else:
        #print('Filter not matching')
        pass
    
    try:
        pricing_client = boto3.client('pricing', region_name=region_name) #config=config)
        response = pricing_client.get_products(ServiceCode=ServiceCode, Filters=filters1)
        #with open("data.json", "w") as file:
        #    json.dump(response, file ,indent=4)
        #print(json.dumps(response))
        for price in response['PriceList']:
            price = json.loads(price)

            for on_demand in price['terms']['OnDemand'].values():
                for price_dimensions in on_demand['priceDimensions'].values():
                    return(price_dimensions['pricePerUnit']['USD'])
                #print ({
                #        'usagetype' : usagetype,
                        #'region': region_name,
                        #'os': os,
                        #'preinstalled_software': preinstalled_software,
                        #'tenancy': tenancy,
                        #'byol': byol,
                #        'price': price_dimensions['pricePerUnit']['USD']}
                        #'effective': on_demand['effectiveDate'],
                        #'description': price_dimensions['description']}
                #    )
    except Exception as e:
        #print(e)
        pass


if __name__ == "__main__":
    #filename = 'MonthlyUsage.xlsx'
    filename='MonthlyUsageReport-Multipleaccounts-scrubbed.xlsx'
    wb = load_workbook(filename)
    ws = wb.active

    max_row = ws.max_row
    max_col = ws.max_column
    ws.insert_cols(max_col+1)
    ws.cell(row=1,column=max_col+1).value = "UnitPrice"
    #UnitPrice_col=max_col+1
    #print(UnitPrice_col)
     # print(max_row,max_col)
    wb.save(filename)

     #lifecycle_config={}
    for row_cells in ws.iter_rows(min_row=2, max_row=max_row):        
        counter = 0
        #counter1 = 2
        #print(counter,row_cells.row)
        for i, cell in enumerate(row_cells):
            #print(i,cell.value)
            counter += 1
            #print(get_column_letter(counter) + str(1))
            colname = ws[get_column_letter(counter) + str(1)]
            #print(colname.value,cell.value)
            if colname.value == 'Line Item Product Code':
              ServiceCode=cell.value
              #print(f'ServiceCode is {ServiceCode}')
            if colname.value == 'Line Item Usage Type':
              t1=cell.value
              usagetype= t1
              usagecode=t1.split(':')[0]
              #print(f'usagetype is {usagetype}')
            if colname.value == 'Product Region':
              regionCode=cell.value
              #print(f'regionCode is {regionCode}')
        unitprice=current_price(ServiceCode,usagecode,regionCode,usagetype)
        #print(unitprice,counter)
        for cell in row_cells:
            #print(cell.row)
            ws.cell(row=cell.row,column=9).value=unitprice
        
        if unitprice:
            print(f'Unit Price for {usagetype} is {unitprice}')
            #print(counter)
            #print(ws.cell(row=counter,column=9).value)
            #ws.cell(row=counter+1,column=9).value=unitprice
                
    wb.save(filename)
