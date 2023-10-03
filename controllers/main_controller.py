from threading import Thread
from queue import Queue
from tkinter import filedialog, END, messagebox

class MainController:
    def __init__(self):
        self.task_queue = Queue()
        
    def start_process(self):
         messagebox.showwarning("Info", "Processo Iniciado")