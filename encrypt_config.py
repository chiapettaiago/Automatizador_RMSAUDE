from cryptography.fernet import Fernet
import json

key = Fernet.generate_key()

with open('config.json', 'r') as file:
    config = file.read()

fernet = Fernet(key)
encrypted = fernet.encrypt(config.encode('utf-8'))
with open('config.encrypted', 'wb') as file:
    file.write(encrypted)

with open('key.key', 'wb') as file:
    file.write(key)

