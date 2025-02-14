import oci
import os
import argparse

# Function to upload files recursively
def upload_to_oci(bucket_name, namespace, folder_path, config):
    # Initialize Object Storage Client
    object_storage_client = oci.object_storage.ObjectStorageClient(config)

    # Ensure the bucket exists
    try:
        object_storage_client.get_bucket(namespace, bucket_name)
    except oci.exceptions.ServiceError as e:
        print(f"Bucket '{bucket_name}' does not exist or is not accessible: {e}")
        return

    # Walk through the directory and upload files
    for root, _, files in os.walk(folder_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)
            object_name = os.path.relpath(file_path, folder_path)  # Preserve folder structure

            print(f"Uploading: {file_path} → {bucket_name}/{object_name}")

            with open(file_path, "rb") as file:
                object_storage_client.put_object(
                    namespace_name=namespace,
                    bucket_name=bucket_name,
                    object_name=object_name,
                    put_object_body=file
                )

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Upload files from a folder to OCI Object Storage")
    parser.add_argument("--bucket-name", required=True, help="Name of the OCI bucket")
    parser.add_argument("--folder-path", required=True, help="Path to the folder to upload")
    
    # Authentication parameters
    parser.add_argument("--user", required=True, help="OCI User OCID")
    parser.add_argument("--tenancy", required=True, help="OCI Tenancy OCID")
    parser.add_argument("--region", required=True, help="OCI Region (e.g., ap-melbourne-1)")
    parser.add_argument("--fingerprint", required=True, help="Fingerprint of the API Key - setup in OCI profile")
    parser.add_argument("--private-key", required=True, help="Path to the private key file (.pem) - download from OCI API key")

    args = parser.parse_args()

    # Load authentication details
    config = {
        "user": args.user,
        "tenancy": args.tenancy,
        "region": args.region,
        "fingerprint": args.fingerprint,
        "key_file": args.private_key
    }

    # Fetch namespace dynamically
    object_storage_client = oci.object_storage.ObjectStorageClient(config)
    namespace = object_storage_client.get_namespace().data  

    upload_to_oci(
        args.bucket_name,
        namespace,
        args.folder_path,
        config
    )
