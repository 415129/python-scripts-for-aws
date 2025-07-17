import boto3

# Create a pricing client
pricing_client = boto3.client('pricing', region_name='us-east-1')

# Get all services
response = pricing_client.describe_services()

# Print the ServiceCode for each service
for service in response['Services']:
    print(service['ServiceCode'])