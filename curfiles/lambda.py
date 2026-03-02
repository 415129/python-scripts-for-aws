import boto3
import logging
import os
import gzip
import uuid

# It's a good practice to set up logging for Lambda functions
logger = logging.getLogger()
logger.setLevel(logging.INFO)


def decompress_and_upload_gz(s3_client, source_bucket, source_prefix, source_key, dest_bucket, dest_prefix):
    """
    Downloads a .gz file from S3 to a temporary local file,
    decompresses it, and uploads the result to another S3 location.
    The destination key is derived from the source key by removing the '.gz' extension.
    """
    tmp_download_path = f"/tmp/{uuid.uuid4()}"

    try:
        # Determine destination key
        relative_key = source_key[len(source_prefix):]
        dest_key_base = relative_key
        if relative_key.endswith('.gz'):
            dest_key_base = relative_key[:-len('.gz')]
        dest_key = dest_prefix + dest_key_base
        
        logger.info(f"Decompressing 's3://{source_bucket}/{source_key}' to 's3://{dest_bucket}/{dest_key}'")

        # 1. Download from S3 to /tmp
        s3_client.download_file(source_bucket, source_key, tmp_download_path)

        # 2. Decompress from /tmp and upload stream to S3
        with gzip.open(tmp_download_path, 'rb') as f_gz:
            s3_client.upload_fileobj(f_gz, dest_bucket, dest_key)

        logger.info(f"Successfully uploaded decompressed file to 's3://{dest_bucket}/{dest_key}'")

    except Exception as e:
        logger.error(f"Error processing gzipped file {source_key}: {e}")
        raise
    finally:
        # 4. Clean up the /tmp file
        if os.path.exists(tmp_download_path):
            os.remove(tmp_download_path)


def lambda_handler(event, context):
    """
    Copies all objects recursively from a source S3 prefix to a destination S3 prefix.
    If an object is a '.csv.gz' file, it is decompressed before being copied.
    """
    s3 = boto3.client('s3')

    # Correctly define bucket names and prefixes
    source_bucket = "maximussandboxorgcur"
    source_prefix = "/cur/cur/" 
    dest_bucket = "amazon-sagemaker-423014875898-us-east-2-cxd5ximkhsltpl"
    dest_prefix = "curfiles/"

    try:
        paginator = s3.get_paginator('list_objects_v2')
        
        # Paginate through all objects under the source_prefix.
        # The list_objects_v2 operation is recursive by default when a Prefix is used,
        # so it will find all objects in all sub-folders.
        for page in paginator.paginate(Bucket=source_bucket, Prefix=source_prefix):
            for obj in page.get('Contents', []):
                source_key = obj['Key']

                # S3 can sometimes include the folder itself as an object ending in '/', skip it
                if source_key.endswith('/'):
                    continue

                if source_key.endswith('.csv.gz'):
                    # Decompress and upload gzipped CSV files
                    decompress_and_upload_gz(s3, source_bucket, source_prefix, source_key, dest_bucket, dest_prefix)
                else:
                    # For all other files, just copy them directly
                    relative_key = source_key[len(source_prefix):]
                    dest_key = dest_prefix + relative_key

                    copy_source = {'Bucket': source_bucket, 'Key': source_key}
                    
                    logger.info(f"Copying 's3://{source_bucket}/{source_key}' to 's3://{dest_bucket}/{dest_key}'")

                    s3.copy_object(
                        CopySource=copy_source, 
                        Bucket=dest_bucket, 
                        Key=dest_key
                    )

        logger.info("All objects processed successfully.")
        return {
            'statusCode': 200,
            'body': 'Object processing completed successfully.'
        }

    except Exception as e:
        logger.error(f"An error occurred: {str(e)}")
        # Return an error response to the caller
        return {
            'statusCode': 500,
            'body': f"An error occurred: {str(e)}"
        }