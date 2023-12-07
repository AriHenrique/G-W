import boto3
from datetime import datetime
from botocore.exceptions import NoCredentialsError


def upload_to_s3(local_file, bucket_name, s3_file):
    s3 = boto3.client('s3', aws_access_key_id='id', aws_secret_access_key='key')

    try:
        # Envia o arquivo para o S3
        s3.upload_file(local_file, bucket_name, s3_file)
        print("Upload bem-sucedido")
        return True
    except FileNotFoundError:
        print("Arquivo não encontrado")
        return False
    except NoCredentialsError:
        print("Credenciais da AWS não disponíveis")
        return False


local_file_path = "C:\SHARMAQ\SHOficina\dados.mdb"
bucket_name = "exemplo-aula-si"
s3_file_path = f"oficina/ano={datetime.now().year}/mes={datetime.now().month}/dia={datetime.now().day}/dados.mdb"

upload_to_s3(local_file_path, bucket_name, s3_file_path)
