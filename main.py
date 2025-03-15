# main.py
import tkinter as tk
from view import SistemaView
from controller import SistemaController
from model import SistemaModel

if __name__ == "__main__":
    root = tk.Tk()
    model = SistemaModel()
    controller = SistemaController(model, None)
    view = SistemaView(root, controller)
    controller.configurar_view(view)
    root.mainloop()