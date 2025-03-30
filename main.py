import tkinter as tk
from model import SistemaModel
from view import SistemaView
from controller import SistemaController

def main():
    # Inicializar o modelo
    model = SistemaModel("controle_alunos.db")

    # Inicializar a janela principal
    root = tk.Tk()

    # Inicializar o controller
    controller = SistemaController(model, None)

    # Inicializar a view com o controller
    view = SistemaView(root, controller)

    # Configurar a view no controller
    controller.configurar_view(view)

    # Exibir a tela inicial
    view.tela_inicial()

    # Configurar o evento de fechamento da janela
    def on_closing():
        controller.fechar_conexao()  # Garante que a conexão com o banco seja fechada e o backup seja criado
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    # Iniciar o loop principal
    root.mainloop()

if __name__ == "__main__":
    main()