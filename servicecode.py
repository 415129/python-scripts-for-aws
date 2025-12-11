import boto3

# The Pricing API only has two global endpoints: 'us-east-1' and 'ap-south-1'.
# It's highly recommended to use 'us-east-1' for the client region.
REGION_NAME = 'us-east-1'

# Initialize the AWS Pricing client
client = boto3.client('pricing', region_name=REGION_NAME)

def get_all_service_codes():
    print("Fetching all available ServiceCodes from AWS Pricing API...")
    
    # Use a paginator to handle the potentially long list of services
    paginator = client.get_paginator('describe_services')
    
    # Create the page iterator
    pages = paginator.paginate(
        MaxResults=100  # Default is 100
    )
    
    service_codes = []
    
    # Iterate through all pages and extract the ServiceCode from each service object
    for page in pages:
        for service in page.get('Services', []):
            service_codes.append(service.get('ServiceCode'))
            
    return service_codes

# --- Execution ---
try:
    available_codes = get_all_service_codes()
    
    print("-" * 50)
    print(f"Total Service Codes Found: {len(available_codes)}")
    print("-" * 50)
    
    # Display the first 10 service codes as an example
    print("First 10 Service Codes:")
    for code in available_codes[:1000]:
        print(f"* {code}")

    # You can now use any of these codes in your client.list_price_lists() call
    
except Exception as e:
    print(f"An error occurred while calling DescribeServices: {e}")
    

#print (get_all_service_codes (AWSKeyManagementService))