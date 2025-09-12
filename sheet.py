###############################################################################
#  AWS Monthly Usage Pricing Script
#  Author: [Your Name/Team]
#  Description: Extracts AWS service unit prices from monthly usage Excel report,
#               fetches current prices via AWS Pricing API, and writes results
#               back to the Excel file for analysis.
#  Last Modified: [Date]
###############################################################################

import boto3
import json
import sys
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
    """
    Fetches RDS heavy usage pricing from AWS Pricing API.
    Returns price per hour for reserved RDS instances.
    """
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

def ec2dedicatedusuage(ServiceCode, filters1, region_name, usagevalue):
    """
    Fetches EC2 dedicated usage pricing from AWS Pricing API.
    Returns price per unit for dedicated EC2 instances.
    """
    pricing_client = boto3.client('pricing', region_name=region_name)
    response = pricing_client.get_products(ServiceCode=ServiceCode, Filters=filters1)
    for price in response['PriceList']:
        price = json.loads(price)
        for ReservedOne in price['terms']['Reserved'].values():
            if ReservedOne["termAttributes"]["LeaseContractLength"] == "1yr":
                for price_dimensions in ReservedOne['priceDimensions'].values():
                    if price_dimensions["description"] == "Upfront Fee" and price_dimensions["unit"] == "Quantity":
                        price = (int(price_dimensions['pricePerUnit']['USD'])/int(dedicatedfactor[usagevalue]))
                        return (price)

def find_whole_word(word, string):
    """
    Checks if any word in the list exists as a whole word in the string.
    Returns True if found, else None.
    """
    for w in word:
        match = re.search(r'\b' + re.escape(w) + r'\b', string, re.IGNORECASE)
        if match is not None:
            return True

def current_price(ServiceCode, usagecode, regionCode, usagevalue, byol='Bring your own license'):
    """
    Main function to determine the current unit price for a given AWS usage line.
    Builds filters, queries AWS Pricing API, and handles special cases.
    Returns the unit price or a marker string.
    """
    region_name = 'us-east-1'
    preinstalled_software = 'NA'
    os = 'Linux'
    tenancy = "Shared"
    byol = False
    filters1 = [{}]

    if find_whole_word(["BoxUsage"], usagecode):
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
            {'Type': 'TERM_MATCH', 'Field': 'tenancy', 'Value': tenancy},
            {'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': os},
            {'Type': 'TERM_MATCH', 'Field': 'preInstalledSw','Value': preinstalled_software},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
            {'Type': 'TERM_MATCH', 'Field': 'licenseModel',
                'Value': 'Bring your own license' if byol else 'No License required'},
        ]
    elif ServiceCode == 'AmazonEC2' and find_whole_word(["SpotUsage"], usagecode):
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'Spot'},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue.replace("-SpotUsage", "")},
            {'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': os},
            {'Type': 'TERM_MATCH', 'Field': 'preInstalledSw','Value': preinstalled_software},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode}
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
    elif find_whole_word(["RDS", "Multi-AZUsage", "InstanceUsage", "HeavyUsage","Storage","Mirror"], usagecode) and ServiceCode == 'AmazonRDS':
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagevalue.replace("HeavyUsage", "InstanceUsage")},
            {"Type": "TERM_MATCH", "Field": "databaseEngine", "Value": "MySQL"},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
        if find_whole_word(["PIOPS-Storage"], usagevalue):
            filters1.pop()
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
        elif find_whole_word(["Mirror"], usagevalue):
            filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagevalue},
            {"Type": "TERM_MATCH", "Field": "databaseEngine", "Value": "Any"},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode}
        ]
    elif find_whole_word(['Aurora'], usagevalue):
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
            {"Type": "TERM_MATCH", "Field": "databaseEngine", "Value": "Aurora MySQL" if find_whole_word(
                ['Aurora:BackupUsage', 'ServerlessV2Usage'], usagevalue) else "Any"},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
    elif find_whole_word(["Fargate"], usagecode) and ServiceCode == 'AmazonECS':
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagevalue.replace("SpotUsage-", "")},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
    elif ServiceCode == 'AmazonS3' and not find_whole_word(["DataTransfer","In-Bytes","Out-Bytes","Requests","TimedStorage-ByteHrs"], usagevalue):
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode.replace("global", "us-east-1")},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue.replace("GDA-ByteHrs", "GDA-Staging")},
        ]
        if find_whole_word(["Global-Bucket-Hrs-FreeTier"], usagevalue):
            filters1.pop(0)
        elif find_whole_word(["Tier4"], usagevalue):
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
    elif ServiceCode == 'AmazonSageMaker':
        if usagevalue.split(':')[1].startswith("ml"):
            new_regionCode = regionCode.replace("gov-", "")
            filters1 = [
                {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': 'us-east-1'},
                {'Type': 'TERM_MATCH', 'Field': 'component', 'Value': 'Notebook'},
                {'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': usagecode + '-Notebook'}
            ]
        else:
            filters1 = [
                {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
                {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue}
            ]
        # If no price found
    else:
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
        ]

    try:
        if ServiceCode == "AmazonPinpoint":
            return(0.03)
        #elif not ServiceCode.startswith("Amazon"):
        #    return('Marketplace')
        elif ServiceCode in ["ComputeSavingsPlans","EC2InstanceSavingsPlans","MachineLearningSavingsPlans"]:
            return('SavingsPlans')
        pricing_client = boto3.client('pricing', region_name=region_name)
        response = pricing_client.get_products(ServiceCode=ServiceCode, Filters=filters1)
        found_price = None
        for price in response['PriceList']:
            price = json.loads(price)
            for on_demand in price['terms']['OnDemand'].values():
                for price_dimensions in on_demand['priceDimensions'].values():
                    if find_whole_word(["SpotUsage"], usagecode) and ServiceCode == 'AmazonECS':
                        print(f'Found SpotUsage for {ServiceCode} {usagecode}')
                        found_price = ((float(price_dimensions['pricePerUnit']['USD']) / 100) * 30)
                    elif find_whole_word(["HeavyUsage"], usagecode) and ServiceCode == 'AmazonRDS':
                        found_price = rdsheavyusuage(ServiceCode, filters1, region_name)
                    else:
                        found_price = price_dimensions['pricePerUnit']['USD']
                    if found_price is not None:
                        #print(f'Found price for {ServiceCode} {usagevalue} in {regionCode} is {found_price}')
                        return found_price
        # If no price found and ServiceCode is AmazonSageMaker or AmazonEC2, try removing 'gov-' from regionCode and retry
        if ServiceCode in ["AmazonSageMaker", "AmazonEC2"] and not found_price and regionCode.startswith("us-gov"):
            new_regionCode = regionCode.replace("gov-", "")
            print(f'Trying with modified regionCode: {new_regionCode} for ServiceCode: {ServiceCode}')
            if ServiceCode == "AmazonSageMaker":                
                filters1 = [
                    {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': 'ca-central-1'},
                    {'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': usagevalue}
                ]
            elif ServiceCode == "AmazonEC2":
                filters1 = [
                    {'Type': 'TERM_MATCH', 'Field': 'termType', 'Value': 'OnDemand'},
                    {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
                    {'Type': 'TERM_MATCH', 'Field': 'tenancy', 'Value': tenancy},
                    {'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': os},
                    {'Type': 'TERM_MATCH', 'Field': 'preInstalledSw','Value': preinstalled_software},
                    {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': new_regionCode},
                    {'Type': 'TERM_MATCH', 'Field': 'licenseModel','Value': 'Bring your own license' if byol else 'No License required'},
                ]
            response = pricing_client.get_products(ServiceCode=ServiceCode, Filters=filters1)
            for price in response['PriceList']:
                price = json.loads(price)
                for on_demand in price['terms']['OnDemand'].values():
                    for price_dimensions in on_demand['priceDimensions'].values():
                        return price_dimensions['pricePerUnit']['USD']
    except Exception as e:
        print(e)
        pass

if __name__ == "__main__":
    default_filename = 'MonthlyUsageReport-Multipleaccounts-scrubbed.xlsx'
    filename = input(f"Please enter full filename path (press Enter to use default: {default_filename}): ").strip()
    if not filename:
        filename = default_filename

    wb = load_workbook(filename)
    ws = wb.active

    max_row = ws.max_row
    max_col = ws.max_column
    ws.insert_cols(max_col+1)
    ws.cell(row=1, column=max_col+1).value = "UnitPrice"
    wb.save(filename)

    for row_cells in ws.iter_rows(min_row=2, max_row=max_row):
        counter = 0
        for i, cell in enumerate(row_cells):
            counter += 1
            colname = ws[get_column_letter(counter) + str(1)]
            if colname.value == 'Line Item Product Code':
                ServiceCode = cell.value
            if colname.value == 'Line Item Usage Type':
                t1 = cell.value
                usagevalue = t1
                usagecode = t1.split(':')[0]
            if colname.value == 'Product Region':
                regionCode = cell.value
        unitprice = current_price(ServiceCode, usagecode, regionCode, usagevalue)
        for cell in row_cells:
            ws.cell(row=cell.row, column=9).value = unitprice
        if unitprice:
            print(f'Unit Price for Service {ServiceCode} {usagevalue} is {unitprice}')
    wb.save(filename)
