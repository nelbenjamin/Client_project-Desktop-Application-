from cryptography.fernet import Fernet
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox

# Same key as in your app.py
SECRET_KEY = b"a_WeqpQk65a-PGnLFodDaeL7PnbRUDKv0VXwRr-sKKI="

def generate_license_key(months=12):
    exp_date = (datetime.now() + timedelta(days=30*months)).strftime("%Y-%m-%d")
    fernet = Fernet(SECRET_KEY)
    return fernet.encrypt(exp_date.encode()).decode()

if __name__ == "__main__":
    key = generate_license_key()

    # GUI popup
    root = tk.Tk()
    root.withdraw()  # Hide root window
    root.clipboard_clear()
    root.clipboard_append(key)
    root.update()
    messagebox.showinfo("License Key Generated",
                        f"License key:\n\n{key}\n\n(It has also been copied to clipboard.)")
    root.destroy()
