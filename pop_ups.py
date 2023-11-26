import tkinter as tk
from tkinter import simpledialog

class CustomDialog(simpledialog.Dialog):
    def body(self, master):
        tk.Label(master, text="Store URL:").grid(row=0)
        tk.Label(master, text="Consumer Key:").grid(row=1)
        tk.Label(master, text="Consumer Secret:").grid(row=2)

        self.store_url_entry = tk.Entry(master)
        self.consumer_key_entry = tk.Entry(master)
        self.consumer_secret_entry = tk.Entry(master)

        self.store_url_entry.grid(row=0, column=1)
        self.consumer_key_entry.grid(row=1, column=1)
        self.consumer_secret_entry.grid(row=2, column=1)

        return self.store_url_entry  # initial focus

    def apply(self):
        store_url = self.store_url_entry.get()
        consumer_key = self.consumer_key_entry.get()
        consumer_secret = self.consumer_secret_entry.get()
        self.result = store_url, consumer_key, consumer_secret
