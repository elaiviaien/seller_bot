
import boto3
from botocore.exceptions import NoCredentialsError
import requests
import mimetypes

ACCESS_KEY_ID = 'AKIATGTXOYXK6F5HYMGZ'
SECRET_ACCESS_KEY = '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'


def upload_file(local_file, bucket, s3_file):
    s3 = boto3.client('s3', aws_access_key_id=ACCESS_KEY_ID,
                      aws_secret_access_key=SECRET_ACCESS_KEY)

    try:
        s3.upload_file(local_file, bucket, s3_file)
        print("Upload Successful")
        return True
    except FileNotFoundError:
        print("The file was not found")
        return False
    except NoCredentialsError:
        print("Credentials not available")
        return False

