from keygen import SECRET_KEY
from cryptography.fernet import Fernet

fernet = Fernet(SECRET_KEY)

key = input("Paste license key: ")

try:
    decrypted = fernet.decrypt(key.encode()).decode()
    print("Decrypted data:", decrypted)
except Exception as e:
    print("Invalid key!", e)
