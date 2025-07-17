import boto3  # for AWS to interact with AWS services
from botocore.config import Config
import logging
import pandas as pd
import time
from datetime import datetime
from botocore.exceptions import ClientError
from time import strftime
from tqdm import tqdm  # displays progress bar for loops


# initializing the s3 client (interact w/ s3 buckets)
s3 = boto3.client('s3')

# buckets to ignore (not set up yet)
ignorelist = []


# policy names expect each bucket to have (check)
stdpname = ['MMSDeletionStandardPolicy', 'MMSStdDelMarkerPolicy',
            'AbortIncompleteMultipartUploadsRule']


# get Account ID, to only get the bucket owned by expected account
# STS AWS Security Token Service
def getAccountID():
    sts = boto3.client("sts")
    return sts.get_caller_identity()["Account"]  # gets current account's ID


# gets lifecycle configuration policy for each bucket
def getLCP(Name):
    ownerAccountId = getAccountID()  # get current AWS account ID
    policy = {}  # dict stores lifecycle policies
    try:
        # call s3, get lifecycle config.
        result = s3.get_bucket_lifecycle_configuration(
            Bucket=Name, ExpectedBucketOwner=ownerAccountId)
        Rules = result['Rules']  # list of lifecycle rules

        # print(Rules)
        n = 1
        for r1 in Rules:  # loop over each rule
            vname = Name + '_Rule_' + str(n)  # key names

            # policy[vname] = r1
            # n+= 1
            # print(policy)
            # print(r1['ID'])
            if r1['ID'] not in stdpname:  # add to dictionary if not std policy
                policy[vname] = r1

                # print(policy)
            n += 1

        # print(Rules.type())
        # policy[Name] = Rules
        # print(policy)
    except ClientError as err:
        # print(err.response['Error']['Code'])
        if err.response['Error']['Code'] == 'NoSuchLifecycleConfiguration':
            # if no lifecycle config. found, return empty dict
            policy[Name] = {}

            # return the lifecycle policy information
            return policy

# list buckets and check for missing lifecycle policies


def missinglcpBuckets():
    BucketName = s3.list_buckets()  # gets all buckets in account
    missing_policies = []

    # print(ignorelist)
    # for bucket in BucketName['Buckets']:
    # progress bar shown while iterating buckets
    for bucket in tqdm(BucketName['Buckets']):
        # and bucket['Name'] == 'oraclebkupbucket':
        if bucket['Name'] not in ignorelist:
            Name = bucket['Name']
            lcplist = getLCP(Name)  # lifecycle policy details

            # need to get all of the rule IDs from the lifecycle
            # create empty list
            existing_ids = []

            # looping over
            if lcplist:
                for lcp in lcplist:
                    rule = lcplist[lcp]
                    if isinstance(rule, dict) and 'ID' in rule:
                        # add the rule's ID to the list
                        existing_ids.append(rule['ID'])
            else:
                print(f'[INFO] Bucket {Name} has no lifecycle policies.')
            # compare policy name to ones that we have found in the bucket
            for std_policy in stdpname:
                if std_policy not in existing_ids:
                    print(
                        f"[WARNING] Bucket {'Name'} is missing a lifecycle policy: {std_policy}")
                    # add missing to list
                    missing_policies.append((Name, std_policy))
    return missing_policies


if __name__ == '__main__':
    results = missinglcpBuckets()
    print("\n Summary: Missing Policies")
    for bucket, policy in results:
        print(f"Bucket Name: {bucket}, Missing Policy: {policy}")
