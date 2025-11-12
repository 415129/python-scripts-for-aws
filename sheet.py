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
                        {'Type': 'TERM_MATCH', 'Field': 'regionCode','Value': regionCode},
                        {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagevalue}
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
            {"Type": "TERM_MATCH", "Field": "databaseEngine", "Value": "Aurora MySQL" if find_whole_word(['Aurora:BackupUsage', 'ServerlessV2Usage'], usagevalue) else "Any"},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
    elif find_whole_word(["Fargate"], usagecode) and ServiceCode == 'AmazonECS':
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype','Value': usagevalue.replace("SpotUsage-", "")},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
    elif ServiceCode == 'AmazonS3' and not find_whole_word(["DataTransfer","In-Bytes","Out-Bytes","Requests","TimedStorage-ByteHrs","Monitoring-Automation"], usagevalue):
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode.replace("global", "us-east-1")},
            #{'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue.replace("GDA-ByteHrs", "GDA-Staging")},
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue}
        ]
        if find_whole_word(["Global-Bucket-Hrs-FreeTier"], usagevalue):
            filters1.pop(0)
        elif find_whole_word(["Tier4"], usagevalue):
            filters1.append({'Type': 'TERM_MATCH', 'Field': 'groupDescription','Value': "Lifecycle Transition Requests into Intelligent-Tiering"},)
    elif ServiceCode == 'AmazonEFS':
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': re.sub(r'ET-SmallFiles|SmallFiles', 'ByteHrs', usagevalue).replace('UGW1-ArchiveEarlyDelete-ByteHrs', 'UGW1-ArchiveTimedStorage-ByteHrs')},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
    elif find_whole_word(["DataTransfer","In-Bytes","Out-Bytes"], usagevalue):
        ServiceCode="AWSDataTransfer"
        # replace AZ with xAZ but do NOT change existing xAZ occurrences
        safe_usagevalue = re.sub(r'(?<!x)AZ', 'xAZ', usagevalue)
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': safe_usagevalue}
        ]
    elif find_whole_word(["Resource-Operation-Count"], usagevalue) and ServiceCode == 'AWSCloudFormation':
        ServiceCode="AWSCloudFormation"
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue.replace("Resource-Operation-Count", "Resource-Invocation-Count")}
        ]
    elif ServiceCode == 'AmazonSageMaker':
        instance_type_part1 = usagevalue.split(':')[0].split('-')[-1]
        instance_type_part2 = usagevalue.split(':')[-1]
        # Determine component from usagecode (e.g., Notebook, Host, Training) UGW1-Studio: KernelGateway-ml.g4dn.xlarge
        if instance_type_part1 == 'Notebk':            
            component = "Notebook" # Default
        elif instance_type_part1 == 'Host' or instance_type_part1 == 'Hst':
            component = "Hosting"
        elif instance_type_part1 == 'Train':
            component = "Training"
        elif instance_type_part1 == 'Studio':
            component = "Studio-Notebook"
        else:
            component = "Notebook" # Fallback
        
        if instance_type_part2.startswith("VolumeUsage"):
            filters1 = [
                {'Type': 'TERM_MATCH', 'Field': 'productFamily', 'Value': "Storage"},
                {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue}
            ]
        elif instance_type_part2.startswith("Studio") or instance_type_part2.startswith("KernelGateway"):
            filters1 = [
                {'Type': 'TERM_MATCH', 'Field': 'productFamily', 'Value': "ML Instance"},
                {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue}
            ]
        elif instance_type_part1.startswith("Host"): #UGW1-Hosting:ml.g5.24xlarge
            filters1 = [
                {'Type': 'TERM_MATCH', 'Field': 'productFamily', 'Value': "ML Instance"},
                {'Type': 'TERM_MATCH', 'Field': 'component', 'Value': component},
                {'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': instance_type_part2 + '-Hosting'},
                {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode}
            ]
        else:
            filters1 = [
                    {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
                    {'Type': 'TERM_MATCH', 'Field': 'component', 'Value': component},
                    {'Type': 'TERM_MATCH', 'Field': 'instanceName', 'Value': instance_type_part2}
                ]
    elif ServiceCode == "CodeBuild":
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
        ]
    elif find_whole_word(["AmazonElastiCache",'AmazonElasticCache'], ServiceCode):
        ServiceCode='AmazonElastiCache'
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'productFamily', 'Value': "Cache Instance"},
            {'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': usagevalue.split('-')[-1].split(':')[-1]},
            {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': regionCode},
            {"Type": "TERM_MATCH", "Field": "cacheEngine","Value": "Redis"}
        ]
    else:
        filters1 = [
            {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': usagevalue},
        ]

    try:
        if ServiceCode == "AmazonPinpoint":
            return(0.03)
        # Check for numeric ServiceCodes, which usually indicate a Marketplace product.
        elif ServiceCode.startswith(tuple(str(i) for i in range(10))):
            print(f'Skipping non-Amazon service: {ServiceCode}')
            return('Marketplace')
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
        if not found_price and ServiceCode == 'AmazonEC2':
            print(f'Price not found for {ServiceCode} {usagevalue}. Retrying with generic BoxUsage.')
            try:
                instance_type = usagevalue.split(':')[-1]
                new_usage_type = f'BoxUsage:{instance_type}'
                new_filters = [
                    {'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': new_usage_type},
                    {'Type': 'TERM_MATCH', 'Field': 'operatingSystem', 'Value': 'Linux'},
                    {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': 'us-east-1'}
                ]
                response = pricing_client.get_products(ServiceCode=ServiceCode, Filters=new_filters)
                for price in response['PriceList']:
                    price = json.loads(price)
                    for on_demand in price['terms']['OnDemand'].values():
                        for price_dimensions in on_demand['priceDimensions'].values():
                            return price_dimensions['pricePerUnit']['USD']
            except Exception as e:
                print(f"Error during EC2 BoxUsage retry: {e}")
        # If no price is found, and it's a gov region, try searching in other US regions.
        if not found_price and regionCode.startswith("us-gov"):
            us_regions_to_check = ['us-east-1', 'us-east-2', 'us-west-1', 'us-west-2']
            # Create a new list of filters, excluding the original regionCode filter.
            new_filters = [f for f in filters1 if f.get('Field') != 'regionCode']
            new_filters = [f for f in filters1 if f.get('Field') not in ['regionCode', 'usagetype']]

            # Create a generic usage type by removing the regional prefix (e.g., UGE1-)
            generic_usagevalue_parts = usagevalue.split('-', 1)
            generic_usagevalue = usagevalue
            if len(generic_usagevalue_parts) > 1:
                generic_usagevalue = generic_usagevalue_parts[1]
            generic_usagevalue = generic_usagevalue_parts[1] if len(generic_usagevalue_parts) > 1 else usagevalue
            
            # The pricing API doesn't support wildcards in usagetype filters.
            # We will have to rely on the other filters being specific enough.
            new_filters.append({'Type': 'TERM_MATCH', 'Field': 'usagetype', 'Value': generic_usagevalue})
            
            for new_region in us_regions_to_check:
                #print(f'Price not found in {regionCode}. Trying with region: {new_region} for ServiceCode: {ServiceCode} and usagevalue: {usagevalue}')
                print(f'Price not found in {regionCode}. Trying with region: {new_region} for ServiceCode: {ServiceCode} and generic usagevalue: {generic_usagevalue}')
                region_filter = {'Type': 'TERM_MATCH', 'Field': 'regionCode', 'Value': new_region}
                filters_with_new_region = new_filters + [region_filter]
                filters_with_new_region = new_filters + [region_filter] # This creates a copy
                #print(filters_with_new_region)
                response = pricing_client.get_products(ServiceCode=ServiceCode, Filters=filters_with_new_region)
                for price in response['PriceList']:
                    price = json.loads(price)
                    for on_demand in price['terms']['OnDemand'].values():
                        for price_dimensions in on_demand['priceDimensions'].values():
                            return price_dimensions['pricePerUnit']['USD']
    except Exception as e:
        print(e)
        pass

if __name__ == "__main__":
    default_filename = 'sheet32.xlsx'
    filename = input(f"Please enter full filename path (press Enter to use default: {default_filename}): ").strip()
    if not filename:
        filename = default_filename
    #filename = 'MonthlyUsageReport-Multipleaccounts-scrubbed.xlsx'
    wb = load_workbook(filename)
    ws = wb.active

    max_row = ws.max_row
    max_col = ws.max_column
    unit_price_col_idx = max_col + 1
    ws.insert_cols(unit_price_col_idx)
    ws.cell(row=1, column=unit_price_col_idx).value = "UnitPrice"
    wb.save(filename)

    for row_cells in ws.iter_rows(min_row=2, max_row=max_row):
        counter = 0
        for i, cell in enumerate(row_cells):
            counter += 1 # This will now be off by one for columns after the insertion, but it's only used for reading headers.
            colname = ws[get_column_letter(counter) + str(1)]
            if colname.value == 'Line Item Product Code':
                ServiceCode = cell.value
            if colname.value == 'Line Item Usage Type':
                t1 = cell.value
                usagevalue = t1
                usagecode = t1.split(':')[0]
            if colname.value == 'Product Region':
                regionCode = cell.value
        print(f'Fetching price for Service {ServiceCode} with Usage Type {usagevalue} in Region {regionCode}')
        unitprice = current_price(ServiceCode, usagecode, regionCode, usagevalue)
        ws.cell(row=row_cells[0].row, column=unit_price_col_idx).value = unitprice
        if unitprice:
            print(f'Unit Price for Service {ServiceCode} {usagevalue} is {unitprice}')
    wb.save(filename)
