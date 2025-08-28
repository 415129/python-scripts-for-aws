# UGW1-IATimedStorage-ET-SmallFiles --> changed to EUC2-IATimedStorage-ET-SmallFiles

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

dedicatedfactor = {
    "UGW1-DedicatedUsage:r6i.large": "64",
    "UGW1-DedicatedUsage:m6i.2xlarge": "16",
    "UGW1-DedicatedUsage:r6i.8xlarge": "4",
    "UGW1-DedicatedUsage:m6i.large": "64",
    "UGW1-DedicatedUsage:m5.large": "24",
    "UGW1-DedicatedUsage:r5.xlarge": "24",
    "UGW1-DedicatedUsage:m5.4xlarge": "3",
    "UGW1-DedicatedUsage:m5.2xlarge": "6",
    "UGW1-DedicatedUsage:m5.xlarge": "12",
    "UGW1-DedicatedUsage:r5.2xlarge": "12",
    "UGW1-DedicatedUsage:r5.large": "48",
    "UGW1-DedicatedUsage:m5a.4xlarge": "3"
}


def rdsheavyusuage(ServiceCode, filters1, region_name):
    pricing_client = boto3.client('pricing', region_name=region_name)
    response = pricing_client.get_products(
        ServiceCode=ServiceCode, Filters=filters1)
    for price in response['PriceList']:
        price = json.loads(price)

        for ReservedOne in price['terms']['Reserved'].values():
            for price_dimensions in ReservedOne['priceDimensions'].values():
                if price_dimensions["description"] == "Upfront Fee":
                    price = (int(float(price_dimensions['pricePerUnit']['USD']))/365/24)
                    return (price)
                # return(price_dimensions['pricePerUnit']['USD'])

def ec2dedicatedusuage(ServiceCode, filters1, region_name,usagevalue):
    pricing_client = boto3.client('pricing', region_name=region_name)
    response = pricing_client.get_products(ServiceCode=ServiceCode, Filters=filters1)
    for price in response['PriceList']:
        price = json.loads(price)

        for ReservedOne in price['terms']['Reserved'].values():
            if ReservedOne["termAttributes"]["LeaseContractLength"] == "1yr":
                for price_dimensions in ReservedOne['priceDimensions'].values():
                    if price_dimensions["description"] == "Upfront Fee" and price_dimensions["unit"] == "Quantity":
                        #print(price_dimensions['pricePerUnit']['USD'] +'/' + dedicatedfactor[usagevalue])
                        price = (int(price_dimensions['pricePerUnit']['USD'])/int(dedicatedfactor[usagevalue]))
                        return (price)
                    # return(price_dimensions['pricePerUnit']['USD'])

def find_whole_word(word, string):
    for w in word:
        # print(w)
        match = re.search(r'\b' + re.escape(w) + r'\b', string, re.IGNORECASE)
        if match is not None:
            return True


def current_price(ServiceCode, usagecode, regionCode, usagevalue, byol='Bring your own license'):

    region_name = 'us-east-1'
    #instance_type = 'r5a.xlarge'
    preinstalled_software = 'NA'
    os = 'Linux'
    tenancy = "Shared"
    byol = False
    # usagetype='UGW1-CPUCredits:t3a'
    # regionCode='us-gov-east-1'
    filters1 = [{}]

    if find_whole_word(["BoxUsage"], usagecode):
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
            # {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': region_name},
            # {'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': instance_type},
            {'Type': 'TERM_MATCH', 'Field': 'tenancy', 'Value': tenancy},
            {'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': os},
            {'Type': 'TERM_MATCH', 'Field': 'preInstalledSw','Value': preinstalled_software},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
            {'Type': 'TERM_MATCH', 'Field': 'licenseModel',
                'Value': 'Bring your own license' if byol else 'No License required'},
        ]
    elif find_whole_word(["DedicatedUsage"], usagecode):
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
            {'Type': 'TERM_MATCH', 'Field': 'tenancy', 'Value': 'Dedicated'},
            {'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': os},
            {"Type": "TERM_MATCH", "Field": "preInstalledSw","Value": "NA"},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
    elif find_whole_word(["SnapshotUsage", "CPUCredits", "VolumeP-IOPS"], usagevalue):
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
    elif find_whole_word(["EBS"], usagecode):
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'productFamily', 'Value': 'Storage'},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
        if find_whole_word(["VolumeIOUsage", "VolumeP-Throughput"], usagevalue):
            filters1.pop(0)
    elif find_whole_word(["RDS", "Multi-AZUsage", "InstanceUsage", "HeavyUsage","Storage"], usagecode) and ServiceCode == 'AmazonRDS':
        filters1 = [
            # {'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagevalue.replace("HeavyUsage", "InstanceUsage")},
            # {"Type": "TERM_MATCH", "Field": "databaseEdition","Value": "Enterprise"},
            {"Type": "TERM_MATCH", "Field": "databaseEngine", "Value": "MySQL"},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
            # {'Type': 'TERM_MATCH', 'Field': 'licenseModel','Value': 'Bring your own license'},
        ]
        if find_whole_word(["PIOPS-Storage"], usagevalue):
            filters1.pop()
            #filters1.append({"Type": "TERM_MATCH", "Field": "volumeName", "Value": "io1"})
        elif usagevalue in ['UGW1-RDS:PIOPS', 'UGW1-RDS:ChargedBackupUsage', 'UGW1-RDS:Multi-AZ-PIOPS']:
            filters1.clear()
            filters1 = [{"Type": "TERM_MATCH", 'Field': "databaseEngine", "Value": "MySQL"},
                        {'Type': 'TERM_MATCH', 'Field': 'regionCode',
                            'Value': regionCode},
                        {'Type': 'TERM_MATCH', 'Field': 'usagetype',
                            'Value': usagevalue}
                        ]
        elif usagevalue in ['UGW1-InstanceUsage:db.r6i.2xl']:
            filters1.clear()
            filters1 = [{"Type": "TERM_MATCH", 'Field': "databaseEngine", "Value": "Aurora MySQL"},
                        {'Type': 'TERM_MATCH', 'Field': 'regionCode','Value': regionCode},
                        {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagevalue}
                        ]
    elif find_whole_word(['Aurora'], usagevalue):
        filters1 = [
            # {'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
            {"Type": "TERM_MATCH", "Field": "databaseEngine", "Value": "Aurora MySQL" if find_whole_word(
                ['Aurora:BackupUsage', 'ServerlessV2Usage'], usagevalue) else "Any"},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
            # {'Type': 'TERM_MATCH', 'Field': 'licenseModel','Value': 'Bring your own license' if byol else 'No License required'},
        ]
    elif find_whole_word(["Fargate"], usagecode) and ServiceCode == 'AmazonECS':
        filters1 = [
            # {'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagevalue.replace("SpotUsage-", "")},
            # {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': region_name},
            # {'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': instance_type},
            # {'Type': 'TERM_MATCH', 'Field': 'tenancy', 'Value': tenancy},
            # {'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': os},
            # {'Type': 'TERM_MATCH', 'Field': 'preInstalledSw','Value': preinstalled_software},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
            # {'Type': 'TERM_MATCH', 'Field': 'licenseModel','Value': 'Bring your own license' if byol else 'No License required'},
        ]
    elif ServiceCode == 'AmazonS3' and not find_whole_word(["DataTransfer","In-Bytes","Out-Bytes","Requests","TimedStorage-ByteHrs"], usagevalue):
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode.replace("global", "us-east-1")},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue.replace("GDA-ByteHrs", "GDA-Staging")},
        ]
        if find_whole_word(["Global-Bucket-Hrs-FreeTier"], usagevalue):
            filters1.pop(0)
        elif find_whole_word(["Tier4"], usagevalue):
            # filters1.pop(0)
            filters1.append({'Type': 'TERM_MATCH', 'Field': 'groupDescription',
                            'Value': "Lifecycle Transition Requests into Intelligent-Tiering"},)
    elif ServiceCode == 'AmazonEFS':
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': re.sub(r'ET-SmallFiles|SmallFiles', 'ByteHrs', usagevalue).replace('UGW1-ArchiveEarlyDelete-ByteHrs', 'UGW1-ArchiveTimedStorage-ByteHrs')},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
    elif find_whole_word(["DataTransfer","In-Bytes","Out-Bytes"], usagevalue):
        ServiceCode="AWSDataTransfer"
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue.replace("AZ","xAZ")}
        ]
    elif find_whole_word(["Resource-Operation-Count"], usagevalue) and ServiceCode == 'AWSCloudFormation':
        ServiceCode="AWSCloudFormation"
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue.replace("Resource-Operation-Count", "Resource-Invocation-Count")}
        ]
    else:
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
        ]

    #print(ServiceCode, usagecode, regionCode, usagevalue)
    #print(filters1)
    try:
        pricing_client = boto3.client('pricing', region_name=region_name)
        response = pricing_client.get_products(ServiceCode=ServiceCode, Filters=filters1)
        #print(json.dumps(response))
        for price in response['PriceList']:
            price = json.loads(price)
            
            for on_demand in price['terms']['OnDemand'].values():
                for price_dimensions in on_demand['priceDimensions'].values():
                    if find_whole_word(["SpotUsage"], usagecode) and ServiceCode == 'AmazonECS':
                        print(f'Found SpotUsage for {ServiceCode} {usagecode}')
                        return ((float(price_dimensions['pricePerUnit']['USD']) / 100) * 30)     # same for ec2
                    elif find_whole_word(["HeavyUsage"], usagecode) and ServiceCode == 'AmazonRDS':
                        return(rdsheavyusuage(ServiceCode, filters1, region_name))
                    elif ServiceCode in ["ComputeSavingsPlans","EC2InstanceSavingsPlans","MachineLearningSavingsPlans"]:
                        return('NA')
                    elif ServiceCode =="AmazonPinpoint":
                        return(0.03)
                    elif not ServiceCode.startswith("Amazon"):
                        return('Marketplace')
                    #elif find_whole_word(["DedicatedUsage"], usagecode) and ServiceCode == 'AmazonEC2':
                    #    return(ec2dedicatedusuage(ServiceCode, filters1, region_name,usagevalue))
                    else:
                        #print(price_dimensions)
                        return (price_dimensions['pricePerUnit']['USD'])
    except Exception as e:
        print(e)
        pass


if __name__ == "__main__":
    # filename = 'MonthlyUsage.xlsx'
    filename = 'MonthlyUsageReport-Multipleaccounts-scrubbed2.xlsx'
    wb = load_workbook(filename)
    ws = wb.active

    max_row = ws.max_row
    max_col = ws.max_column
    ws.insert_cols(max_col+1)
    ws.cell(row=1, column=max_col+1).value = "UnitPrice"
    # UnitPrice_col=max_col+1
    # print(UnitPrice_col)
    # print(max_row,max_col)
    wb.save(filename)

    # lifecycle_config={}
    for row_cells in ws.iter_rows(min_row=2, max_row=max_row):
        counter = 0
        # counter1 = 2
        # print(counter,row_cells.row)
        for i, cell in enumerate(row_cells):
            # print(i,cell.value)
            counter += 1
            # print(get_column_letter(counter) + str(1))
            colname = ws[get_column_letter(counter) + str(1)]
            # print(colname.value,cell.value)
            if colname.value == 'Line Item Product Code':
                ServiceCode = cell.value
                # print(f'ServiceCode is {ServiceCode}')
            if colname.value == 'Line Item Usage Type':
                t1 = cell.value
                usagevalue = t1
                usagecode = t1.split(':')[0]
                # print(f'usagetype is {usagetype}')
            if colname.value == 'Product Region':
                regionCode = cell.value
                # print(f'regionCode is {regionCode}')
        # print(ServiceCode,usagecode,regionCode,usagevalue)
        unitprice = current_price(ServiceCode, usagecode, regionCode, usagevalue)
        # print(unitprice,counter)
        for cell in row_cells:
            # print(cell.row)
            ws.cell(row=cell.row, column=9).value = unitprice

        if unitprice:
            print(f'Unit Price for Service {ServiceCode} {usagevalue} is {unitprice}')
    wb.save(filename)
