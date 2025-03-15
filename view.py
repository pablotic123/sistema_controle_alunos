# view.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from PIL import Image, ImageTk

class SistemaView:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.FONT = ("Arial", 12)
        self.FONT_TITULO = ("Arial", 16, "bold")
        self.BG_COLOR = "#f0f0f0"
        self.FG_COLOR = "#333333"
        self.BTN_COLOR = "#4CAF50"
        self.HIGHLIGHT_COLOR = "#0078D4"
        self.SAVE_COLOR = "#28A745"
        self.CANCEL_COLOR = "#6C757D"
        self.DELETE_COLOR = "#DC3545"
        self.current_frame = None
        self.turma_combo = None
        self.scrollable_frame = None
        self.entradas = {}
        self.root.title("Sistema de Controle de Alunos")
        self.root.geometry("1080x768")
        try:
            self.root.iconbitmap(os.path.join(os.path.dirname(__file__), "imagens", "icones", "icon.ico"))
        except tk.TclError:
            pass
        self.style = ttk.Style()
        self.configurar_estilo()
        self.current_frame = None
        self.criar_menu()

    def configurar_estilo(self):
        self.style.theme_use("clam")
        self.style.configure("TLabel", background=self.BG_COLOR, foreground=self.FG_COLOR, font=self.FONT)
        self.style.configure("TButton", font=self.FONT)
        self.style.map("TButton", background=[("active", "#005A9E")], foreground=[("active", "white")])
        self.style.configure("Custom.Treeview", background=self.BG_COLOR, foreground=self.FG_COLOR, fieldbackground=self.BG_COLOR)
        self.style.configure("Custom.Treeview.Heading", font=("Arial", 10, "bold"), background=self.HIGHLIGHT_COLOR, foreground="white")
        self.style.map("Custom.Treeview.Heading", background=[("active", "#005A9E")])
        self.style.configure("oddrow", background="#F0F0F0")
        self.style.configure("evenrow", background="#FFFFFF")
        self.style.configure("Save.TButton", background=self.SAVE_COLOR)
        self.style.map("Save.TButton", background=[("active", "#218838")])
        self.style.configure("Cancel.TButton", background=self.CANCEL_COLOR)
        self.style.map("Cancel.TButton", background=[("active", "#5A6268")])
        self.style.configure("Delete.TButton", background=self.DELETE_COLOR)
        self.style.map("Delete.TButton", background=[("active", "#C82333")])

    def novo_frame(self):
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = tk.Frame(self.root, bg=self.BG_COLOR)
        self.current_frame.pack(fill=tk.BOTH, expand=True)
        return self.current_frame

    def criar_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        cadastro_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Cadastros", menu=cadastro_menu)
        cadastro_menu.add_command(label="Instituição", command=lambda: None)        
        cadastro_menu.add_command(label="Curso", command=lambda: None)
        cadastro_menu.add_command(label="Turma", command=lambda: None)
        cadastro_menu.add_command(label="Professor", command=lambda: None)
        cadastro_menu.add_command(label="Aluno", command=lambda: None)

        consulta_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Consultas", menu=consulta_menu)
        consulta_menu.add_command(label="Instituições", command=lambda: None)        
        consulta_menu.add_command(label="Cursos", command=lambda: None)
        consulta_menu.add_command(label="Turmas", command=lambda: None)
        consulta_menu.add_command(label="Professores", command=lambda: None)
        consulta_menu.add_command(label="Alunos", command=lambda: None)

        carometro_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Carômetro", menu=carometro_menu)
        ##carometro_menu.add_command(label="Visualizar Carômetro", command=lambda: None)
        carometro_menu.add_command(label="Exportar PDF", command=lambda: None)
        carometro_menu.add_command(label="Exportar Excel", command=lambda: None)
        carometro_menu.add_command(label="Exportar Word", command=lambda: None)
        menubar.add_command(label="Sair", command=self.root.quit)

    def consulta_generica(self, titulo, colunas, atualizar_callback, duplo_clique_callback):
        frame = self.novo_frame()
        tk.Label(frame, text=titulo, font=("Arial", 16, "bold"), bg=self.BG_COLOR, fg=self.FG_COLOR).pack(pady=10, fill=tk.X)
        form_frame = tk.Frame(frame, bg=self.BG_COLOR, padx=20, pady=10)
        form_frame.pack(fill=tk.BOTH, expand=True)

        filtro_frame = tk.Frame(form_frame, bg=self.BG_COLOR)
        filtro_frame.pack(pady=5, fill=tk.X)
        filtros = {}
        for col in colunas[1:]:  # Ignora "id"
            tk.Label(filtro_frame, text=f"{col.capitalize()}:", font=self.FONT, bg=self.BG_COLOR, fg=self.FG_COLOR).pack(side=tk.LEFT, padx=(0 if col == colunas[1] else 10, 5))
            entry = tk.Entry(filtro_frame, width=30 if col == "nome" else 20, font=self.FONT)
            entry.pack(side=tk.LEFT, padx=5)
            filtros[col] = entry
            entry.bind("<KeyRelease>", lambda e: atualizar_callback())

        ttk.Button(filtro_frame, text="Aplicar", command=atualizar_callback).pack(side=tk.LEFT, padx=5)

        tree = ttk.Treeview(form_frame, columns=colunas, show="headings", style="Custom.Treeview")
        for col in colunas:
            tree.heading(col, text=col.capitalize())
            tree.column(col, width=0 if col == "id" else (200 if col == "nome" else 150), stretch=tk.NO if col == "id" else tk.YES)
        tree.pack(fill=tk.BOTH, expand=True)
        tree.bind("<Double-1>", duplo_clique_callback)

        return tree, filtros

    def cadastro_generico(self, titulo, campos, salvar_callback, excluir_callback=None, cancelar_callback=None):
        frame = self.novo_frame()
        tk.Label(frame, text=titulo, font=("Arial", 16, "bold"), bg=self.BG_COLOR, fg=self.FG_COLOR).pack(pady=10, fill=tk.X)
        form_frame = tk.Frame(frame, bg=self.BG_COLOR, padx=20, pady=10)
        form_frame.pack(fill=tk.BOTH, expand=True)

        entradas = {}
        foto_label = None
        for i, (label, tipo, opcoes) in enumerate(campos):
            tk.Label(form_frame, text=f"{label}:", font=self.FONT, bg=self.BG_COLOR, fg=self.FG_COLOR).grid(row=i, column=0, pady=5, sticky="e")
            if tipo == "entry":
                entry = tk.Entry(form_frame, width=50, font=self.FONT)
                entry.grid(row=i, column=1, pady=5, sticky="w")
                entradas[label.lower()] = entry
            elif tipo == "combo":
                var = tk.StringVar()
                combo = ttk.Combobox(form_frame, textvariable=var, values=opcoes, font=self.FONT, width=47)
                combo.grid(row=i, column=1, pady=5, sticky="w")
                entradas[label.lower()] = combo
            elif tipo == "foto":
                entry = tk.Entry(form_frame, width=40, font=self.FONT)
                entry.grid(row=i, column=1, pady=5, sticky="w")
                ttk.Button(form_frame, text="Selecionar", command=lambda e=entry: self.selecionar_foto(e, foto_label)).grid(row=i, column=2, padx=5)
                entradas[label.lower()] = entry
                foto_label = tk.Label(form_frame, bg=self.BG_COLOR)
                foto_label.grid(row=i+1, column=1, columnspan=2, pady=5)
                entradas["foto_label"] = foto_label

        btn_frame = tk.Frame(form_frame, bg=self.BG_COLOR)
        btn_frame.grid(row=len(campos)+1 if foto_label else len(campos), column=0, columnspan=3, pady=10)
        ttk.Button(btn_frame, text="Salvar", command=salvar_callback, style="Save.TButton").pack(side=tk.LEFT, padx=5)
        if excluir_callback:
            ttk.Button(btn_frame, text="Excluir", command=excluir_callback, style="Delete.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=cancelar_callback or (lambda: self.tela_inicial()), style="Cancel.TButton").pack(side=tk.LEFT, padx=5)

        return entradas

    def selecionar_foto(self, entry, foto_label):
        arquivo = filedialog.askopenfilename(filetypes=[("Imagens", "*.jpg;*.jpeg;*.png")])
        if arquivo and os.path.exists(arquivo):
            entry.delete(0, tk.END)
            entry.insert(0, arquivo)
            try:
                img = Image.open(arquivo)
                img = img.resize((150, 150), Image.Resampling.LANCZOS)  # Tamanho maior como na versão anterior
                foto = ImageTk.PhotoImage(img)
                foto_label.config(image=foto)
                foto_label.image = foto  # Manter referência
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível carregar a imagem: {e}")

    def exportar_carometro(self, tipo, exportar_callback):
        frame = self.novo_frame()
        tk.Label(frame, text=f"Exportar Carômetro ({tipo})", font=("Arial", 16, "bold"), bg=self.BG_COLOR, fg=self.FG_COLOR).pack(pady=10, fill=tk.X)
        tk.Label(frame, text="Selecione a Turma:", font=self.FONT, bg=self.BG_COLOR, fg=self.FG_COLOR).pack(pady=5)
        turma_var = tk.StringVar()
        combo = ttk.Combobox(frame, textvariable=turma_var, font=self.FONT, width=50)
        combo.pack(pady=5)
        ttk.Button(frame, text=f"Exportar {tipo}", command=exportar_callback, style="Save.TButton").pack(pady=10)
        return combo

    def atualizar_tabela(self, tree, dados):
        for item in tree.get_children():
            tree.delete(item)
        for i, row in enumerate(dados):
            tree.insert("", "end", values=row, tags=("evenrow" if i % 2 == 0 else "oddrow"))

    def tela_inicial(self):
        frame = self.novo_frame()
        tk.Label(frame, text="Bem-vindo ao Sistema de Controle de Alunos", font=("Arial", 20, "bold"), bg=self.BG_COLOR, fg=self.FG_COLOR).pack(pady=20)
        tk.Label(frame, text="Este sistema permite gerenciar instituições, professores, cursos, turmas e alunos de forma eficiente.\nUtilize os menus acima para cadastrar, consultar e exportar dados.", font=("Arial", 10), bg=self.BG_COLOR, fg=self.FG_COLOR, justify="center").pack(pady=10)

    def visualizar_carometro(self, callback_atualizar):
        self.novo_frame()
        self.current_frame.config(bg=self.BG_COLOR)

        tk.Label(self.current_frame, text="Carômetro", font=self.FONT_TITULO, bg=self.BG_COLOR, fg=self.FG_COLOR).pack(pady=10)
        
        tk.Label(self.current_frame, text="Selecione a turma:", font=self.FONT, bg=self.BG_COLOR, fg=self.FG_COLOR).pack()
        turmas = [f"{t[0]} - {t[1]}" for t in self.controller.model.carregar_turmas()]  # Ajustado de self.model para self.controller.model
        turmas.insert(0, "Todas")
        self.turma_combo = ttk.Combobox(self.current_frame, values=turmas, state="readonly", font=self.FONT)
        self.turma_combo.set("Todas")
        self.turma_combo.pack(pady=5)
        
        tk.Button(self.current_frame, text="Visualizar Carômetro", command=callback_atualizar, font=self.FONT, bg=self.BTN_COLOR, fg=self.FG_COLOR).pack(pady=5)
        
        canvas = tk.Canvas(self.current_frame, bg=self.BG_COLOR)
        scrollbar = ttk.Scrollbar(self.current_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg=self.BG_COLOR)
        
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")
        
        return self.turma_combo, self.scrollable_frame 

    def mostrar_carregando(self):
        # Cria um modal de carregamento centralizado
        self.carregando_modal = tk.Toplevel(self.root)
        self.carregando_modal.title("Carregando")
        self.carregando_modal.geometry("200x100")
        self.carregando_modal.transient(self.root)  # Mantém o modal acima da janela principal
        self.carregando_modal.grab_set()  # Bloqueia interação com a janela principal
        self.carregando_modal.resizable(False, False)
        
        # Centraliza o modal
        self.carregando_modal.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (200 // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (100 // 2)
        self.carregando_modal.geometry(f"+{x}+{y}")
        
        tk.Label(self.carregando_modal, text="Carregando...", font=self.FONT, bg=self.BG_COLOR, fg=self.FG_COLOR).pack(pady=20)
        self.carregando_modal.protocol("WM_DELETE_WINDOW", lambda: None)  # Impede fechamento manual
        
        self.root.update()  # Garante que o modal apareça imediatamente
        return self.carregando_modal

    def fechar_carregando(self):
        if hasattr(self, "carregando_modal") and self.carregando_modal.winfo_exists():
            self.carregando_modal.grab_release()
            self.carregando_modal.destroy()    