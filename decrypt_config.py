from cryptography.fernet import Fernet
import json

with open('key.key', 'rb') as file:
    key = file.read()

with open('config.encrypted', 'rb') as file:
    encrypted = file.read()
fernet = Fernet(key)
decrypted = fernet.decrypt(encrypted.decode('utf-8'))

config = json.loads(decrypted)

user = config['user']
password = config['password']
host = config['host']
port = config['port']
service_name = config['service_name']