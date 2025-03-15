# controller.py
import threading
import os
import shutil
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Image as RLImage, Paragraph, Table, TableStyle, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from tkinter import messagebox
from PIL import Image, ImageTk
import tkinter as tk
import re
from datetime import datetime

def validar_nome(nome, max_length=100):
    if not nome or nome.strip() == "":
        return "O campo Nome é obrigatório!"
    if len(nome) > max_length:
        return f"O Nome deve ter no máximo {max_length} caracteres!"
    if not re.match(r"^[A-Za-z0-9À-ÿ\s]+$", nome):
        return "O Nome deve conter apenas letras, números e espaços!"
    return None

def validar_ano(ano):
    try:
        ano_int = int(ano)
        ano_atual = datetime.now().year
        if ano_int < 2000 or ano_int > ano_atual + 1:
            return f"O Ano deve estar entre 2000 e {ano_atual + 1}!"
        return None
    except ValueError:
        return "O Ano deve ser um número inteiro!"
    
def remover_foto(foto_path):
    try:
        if foto_path and os.path.exists(foto_path):
            os.remove(foto_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao remover foto: {e}")

class SistemaController:
    def __init__(self, model, view):
        self.model = model
        self.view = view  # Pode ser None inicialmente, mas não usaremos diretamente aqui
        self.FONT = None  # Inicializa como None, será configurado depois
        self.BG_COLOR = None
        self.FG_COLOR = None
        # Não chame setup_menu aqui ainda

    def configurar_view(self, view):
        """Configura os atributos que dependem da view após ela ser criada."""
        self.view = view
        self.FONT = view.FONT
        self.BG_COLOR = view.BG_COLOR
        self.FG_COLOR = view.FG_COLOR
        self.setup_menu()  # Configura o menu após a view estar disponível

    def setup_menu(self):
        menubar = self.view.root.nametowidget(self.view.root.cget("menu"))
        cadastro_menu = menubar.children["!menu"]
        cadastro_menu.entryconfig("Instituição", command=self.cadastro_instituicao)        
        cadastro_menu.entryconfig("Curso", command=self.cadastro_curso)
        cadastro_menu.entryconfig("Turma", command=self.cadastro_turma)
        cadastro_menu.entryconfig("Professor", command=self.cadastro_professor)
        cadastro_menu.entryconfig("Aluno", command=self.cadastro_aluno)

        consulta_menu = menubar.children["!menu2"]
        consulta_menu.entryconfig("Instituições", command=self.consulta_instituicoes)        
        consulta_menu.entryconfig("Cursos", command=self.consulta_cursos)
        consulta_menu.entryconfig("Turmas", command=self.consulta_turmas)
        consulta_menu.entryconfig("Professores", command=self.consulta_professores)
        consulta_menu.entryconfig("Alunos", command=self.consulta_alunos)

        carometro_menu = menubar.children["!menu3"]
        carometro_menu.add_command(label="Visualizar Carômetro", command=self.visualizar_carometro)
        carometro_menu.entryconfig("Exportar PDF", command=self.exportar_carometro_pdf)
        carometro_menu.entryconfig("Exportar Excel", command=self.exportar_carometro_excel)
        carometro_menu.entryconfig("Exportar Word", command=self.exportar_carometro_word)        

    # Consultas
    def consulta_instituicoes(self):
        def atualizar(): self.atualizar_tabela("instituicoes", ["id", "nome"])
        def duplo_clique(event): self.on_double_click(event, self.cadastro_instituicao)
        self.tree, self.filtros = self.view.consulta_generica("Consulta de Instituições", ["id", "nome"], atualizar, duplo_clique)
        atualizar()

    def consulta_professores(self):
        def atualizar(): self.atualizar_tabela("professores", ["id", "nome", "instituicao"])
        def duplo_clique(event): self.on_double_click(event, self.cadastro_professor)
        self.tree, self.filtros = self.view.consulta_generica("Consulta de Professores", ["id", "nome", "instituicao"], atualizar, duplo_clique)
        atualizar()

    def consulta_cursos(self):
        def atualizar(): self.atualizar_tabela("cursos", ["id", "nome", "instituicao"])
        def duplo_clique(event): self.on_double_click(event, self.cadastro_curso)
        self.tree, self.filtros = self.view.consulta_generica("Consulta de Cursos", ["id", "nome", "instituicao"], atualizar, duplo_clique)
        atualizar()

    def consulta_turmas(self):
        def atualizar(): self.atualizar_tabela("turmas", ["id", "nome", "ano", "curso"])
        def duplo_clique(event): self.on_double_click(event, self.cadastro_turma)
        self.tree, self.filtros = self.view.consulta_generica("Consulta de Turmas", ["id", "nome", "ano", "curso"], atualizar, duplo_clique)
        atualizar()

    def consulta_alunos(self):
        def atualizar(): self.atualizar_tabela("alunos", ["id", "nome", "turma", "curso", "instituicao"])
        def duplo_clique(event): self.on_double_click(event, self.cadastro_aluno)
        self.tree, self.filtros = self.view.consulta_generica("Consulta de Alunos", ["id", "nome", "turma", "curso", "instituicao"], atualizar, duplo_clique)
        atualizar()

    def atualizar_tabela(self, tipo, colunas):
        try:
            filtros = {col: self.filtros[col].get() for col in colunas[1:]}
            metodo = getattr(self.model, f"consulta_{tipo}")
            dados = metodo(filtros)
            self.view.atualizar_tabela(self.tree, dados)
        except AttributeError as e:
            messagebox.showerror("Erro", f"Erro ao consultar {tipo}: {e}")

    def on_double_click(self, event, callback):
        item = self.tree.selection()
        if item:
            id_value = self.tree.item(item[0], "values")[0]
            callback(int(id_value))

    # Cadastros
    def cadastro_instituicao(self, id=None):
        campos = [("Nome", "entry", None)]
        def salvar(): self.salvar_instituicao(id)
        def excluir(): self.excluir_instituicao(id)
        self.entradas = self.view.cadastro_generico(f"Cadastro de Instituição{' - Editar' if id else ''}", campos, salvar, excluir if id else None)
        if id:
            dados = self.model.executar_query("SELECT nome FROM instituicao WHERE id = ?", (id,), fetch=True)
            if dados:
                self.entradas["nome"].insert(0, dados[0][0])

    def cadastro_professor(self, id=None):
        instituicoes = [f"{i[0]} - {i[1]}" for i in self.model.carregar_instituicoes()]
        campos = [("Nome", "entry", None), ("Instituição", "combo", instituicoes), ("Foto", "foto", None)]
        def salvar(): self.salvar_professor(id)
        def excluir(): self.excluir_professor(id)
        self.entradas = self.view.cadastro_generico(f"Cadastro de Professor{' - Editar' if id else ''}", campos, salvar, excluir if id else None)
        if id:
            dados = self.model.executar_query("SELECT nome, instituicao_id, foto FROM professor WHERE id = ?", (id,), fetch=True)
            if dados:
                self.entradas["nome"].insert(0, dados[0][0])
                self.entradas["instituição"].set(f"{dados[0][1]} - {self.model.executar_query('SELECT nome FROM instituicao WHERE id = ?', (dados[0][1],), fetch=True)[0][0]}")
                if dados[0][2]:
                    self.entradas["foto"].insert(0, dados[0][2])
                    try:
                        img = Image.open(dados[0][2])
                        img = img.resize((150, 150), Image.Resampling.LANCZOS)
                        foto = ImageTk.PhotoImage(img)
                        self.entradas["foto_label"].config(image=foto)
                        self.entradas["foto_label"].image = foto
                    except Exception as e:
                        messagebox.showerror("Erro", f"Não foi possível carregar a imagem: {e}")

    def cadastro_curso(self, id=None, tree=None):
        instituicoes = [f"{i[0]} - {i[1]}" for i in self.model.carregar_instituicoes()]
        campos = [("Nome", "entry", None), ("Instituição", "combo", instituicoes)]
        def salvar(): self.salvar_curso(id)
        def excluir(): self.excluir_curso(id)
        self.entradas = self.view.cadastro_generico(f"Cadastro de Curso{' - Editar' if id else ''}", campos, salvar, excluir if id else None)
        self.tree = tree  # Armazena a referência ao tree
        if id:
            dados = self.model.executar_query("SELECT nome, instituicao_id FROM curso WHERE id = ?", (id,), fetch=True)
            if dados:
                self.entradas["nome"].insert(0, dados[0][0])
                self.entradas["instituição"].set(f"{dados[0][1]} - {self.model.executar_query('SELECT nome FROM instituicao WHERE id = ?', (dados[0][1],), fetch=True)[0][0]}")

    def cadastro_turma(self, id=None, tree=None):
        cursos = [f"{c[0]} - {c[1]}" for c in self.model.carregar_cursos()]
        campos = [
            ("Nome", "entry", None),
            ("Ano", "entry", None),
            ("Curso", "combo", cursos)
        ]
        def salvar(): self.salvar_turma(id)
        def excluir(): self.excluir_turma(id)
        self.entradas = self.view.cadastro_generico(f"Cadastro de Turma{' - Editar' if id else ''}", campos, salvar, excluir if id else None)
        self.tree = tree  # Armazena a referência ao tree
        if id:
            dados = self.model.executar_query("SELECT nome, ano, curso_id FROM turma WHERE id = ?", (id,), fetch=True)
            if dados:
                self.entradas["nome"].insert(0, dados[0][0])
                self.entradas["ano"].insert(0, dados[0][1])
                self.entradas["curso"].set(f"{dados[0][2]} - {self.model.executar_query('SELECT nome FROM curso WHERE id = ?', (dados[0][2],), fetch=True)[0][0]}")
        else:
            self.entradas["ano"].insert(0, str(datetime.now().year))

    def cadastro_aluno(self, id=None):
        turmas = [f"{t[0]} - {t[1]}" for t in self.model.carregar_turmas()]
        campos = [("Nome", "entry", None), ("Turma", "combo", turmas), ("Foto", "foto", None)]
        def salvar(): self.salvar_aluno(id)
        def excluir(): self.excluir_aluno(id)
        self.entradas = self.view.cadastro_generico(
            f"Cadastro de Aluno{' - Editar' if id else ''}", 
            campos, 
            salvar, 
            excluir if id else None
        )
        if id:
            dados = self.model.executar_query("SELECT nome, turma_id, foto FROM aluno WHERE id = ?", (id,), fetch=True)
            if dados:
                self.entradas["nome"].insert(0, dados[0][0])
                self.entradas["turma"].set(f"{dados[0][1]} - {self.model.executar_query('SELECT nome FROM turma WHERE id = ?', (dados[0][1],), fetch=True)[0][0]}")
                if dados[0][2]:
                    self.entradas["foto"].insert(0, dados[0][2])
                    try:
                        img = Image.open(dados[0][2])
                        img = img.resize((150, 150), Image.Resampling.LANCZOS)
                        foto = ImageTk.PhotoImage(img)
                        self.entradas["foto_label"].config(image=foto)
                        self.entradas["foto_label"].image = foto
                    except Exception as e:
                        messagebox.showerror("Erro", f"Não foi possível carregar a imagem: {e}")

    def salvar_instituicao(self, id):
        try:
            nome = self.entradas["nome"].get()
            erro = validar_nome(nome)
            if erro:
                messagebox.showwarning("Aviso", erro)
                return
            self.model.salvar_instituicao(id, nome)
            self.view.tela_inicial()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar instituição: {e}")

    def salvar_professor(self, id):
        try:
            nome = self.entradas["nome"].get()
            instituicao = self.entradas["instituição"].get().split(" - ")[0] if self.entradas["instituição"].get() else None
            foto_path = self.entradas["foto"].get()
            erro = validar_nome(nome)
            if erro:
                messagebox.showwarning("Aviso", erro)
                return
            if not instituicao:
                messagebox.showwarning("Aviso", "O campo Instituição é obrigatório!")
                return
            # Move ou renomeia a foto com base no ID e nome
            foto = self.mover_foto(foto_path, "professores", id, nome) if foto_path or id else None
            # Salva o registro (inserção ou atualização)
            id = self.model.salvar_professor(id, nome, int(instituicao), foto)
            self.view.tela_inicial()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar professor: {e}")

    def salvar_curso(self, id):
        try:
            nome = self.entradas["nome"].get()
            instituicao = self.entradas["instituição"].get().split(" - ")[0] if self.entradas["instituição"].get() else None
            erro = validar_nome(nome)
            if erro:
                messagebox.showwarning("Aviso", erro)
                return
            if not instituicao:
                messagebox.showwarning("Aviso", "O campo Instituição é obrigatório!")
                return
            self.model.salvar_curso(id, nome, int(instituicao))
            self.view.tela_inicial()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar curso: {e}")

    def salvar_turma(self, id):
        try:
            nome = self.entradas["nome"].get()
            ano = self.entradas["ano"].get()
            curso = self.entradas["curso"].get().split(" - ")[0] if self.entradas["curso"].get() else None
            erro_nome = validar_nome(nome, max_length=50)
            if erro_nome:
                messagebox.showwarning("Aviso", erro_nome)
                return
            erro_ano = validar_ano(ano)
            if erro_ano:
                messagebox.showwarning("Aviso", erro_ano)
                return
            if not curso:
                messagebox.showwarning("Aviso", "O campo Curso é obrigatório!")
                return
            self.model.salvar_turma(id, nome, int(ano), int(curso))
            self.view.tela_inicial()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar turma: {e}")

    def salvar_aluno(self, id):
        try:
            nome = self.entradas["nome"].get()
            turma = self.entradas["turma"].get().split(" - ")[0] if self.entradas["turma"].get() else None
            foto_path = self.entradas["foto"].get()
            erro = validar_nome(nome)
            if erro:
                messagebox.showwarning("Aviso", erro)
                return
            if not turma:
                messagebox.showwarning("Aviso", "O campo Turma é obrigatório!")
                return
            # Move ou renomeia a foto com base no ID e nome
            foto = self.mover_foto(foto_path, "alunos", id, nome) if foto_path or id else None
            # Salva o registro (inserção ou atualização)
            id = self.model.salvar_aluno(id, nome, int(turma), foto)
            self.view.tela_inicial()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar aluno: {e}")

    def excluir_instituicao(self, id):
        try:
            if not id:
                messagebox.showwarning("Aviso", "Nenhuma instituição selecionada para excluir!")
                return
            if not messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir esta instituição?"):
                return
            self.model.excluir_registro("instituicao", id)
            self.consulta_instituicoes()  # Recarrega a consulta após exclusão
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir instituição: {e}")

    def excluir_professor(self, id):
        try:
            if not id:
                messagebox.showwarning("Aviso", "Nenhum professor selecionado para excluir!")
                return
            if not messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir este professor?"):
                return
            # Remove a foto associada antes de excluir o registro
            dados = self.model.executar_query("SELECT foto FROM professor WHERE id = ?", (id,), fetch=True)
            if dados and dados[0][0]:
                remover_foto(dados[0][0])
            self.model.excluir_registro("professor", id)
            self.consulta_professores()  # Recarrega a consulta após exclusão
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir professor: {e}")

    def excluir_curso(self, id):
        try:
            if not id:
                messagebox.showwarning("Aviso", "Nenhum curso selecionado para excluir!")
                return
            if not messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir este curso?"):
                return
            self.model.excluir_registro("curso", id)
            self.consulta_cursos()  # Recarrega a consulta após exclusão
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir curso: {e}")

    def excluir_turma(self, id):
        try:
            if not id:
                messagebox.showwarning("Aviso", "Nenhuma turma selecionada para excluir!")
                return
            if not messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir esta turma?"):
                return
            self.model.excluir_registro("turma", id)
            self.consulta_turmas()  # Recarrega a consulta após exclusão
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir turma: {e}")

    def excluir_aluno(self, id):
        try:
            if not id:
                messagebox.showwarning("Aviso", "Nenhum aluno selecionado para excluir!")
                return
            if not messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir este aluno?"):
                return
            # Remove a foto associada antes de excluir o registro
            dados = self.model.executar_query("SELECT foto FROM aluno WHERE id = ?", (id,), fetch=True)
            if dados and dados[0][0]:
                remover_foto(dados[0][0])
            self.model.excluir_registro("aluno", id)
            self.consulta_alunos()  # Recarrega a consulta após exclusão
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir aluno: {e}")

    def mover_foto(self, foto_path, tipo, id, nome):
        imagens_dir = os.path.join(os.path.dirname(__file__), "imagens", tipo)
        os.makedirs(imagens_dir, exist_ok=True)
        nome_arquivo = f"{str(id).zfill(5)}-{nome.replace(' ', '_')}.jpg"
        novo_caminho = os.path.join(imagens_dir, nome_arquivo)

        # Define o nome correto da tabela com base no tipo
        tabela = "professor" if tipo == "professores" else "aluno"

        # Verifica se já existe uma foto associada ao registro
        if id:
            dados = self.model.executar_query(f"SELECT foto FROM {tabela} WHERE id = ?", (id,), fetch=True)
            foto_antiga = dados[0][0] if dados and dados[0][0] else None
            if foto_antiga and os.path.exists(foto_antiga) and (not foto_path or foto_path == foto_antiga):
                # Se o nome mudou e não há nova foto, renomeia a foto existente
                if foto_antiga != novo_caminho:
                    try:
                        os.rename(foto_antiga, novo_caminho)
                        return novo_caminho
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao renomear foto: {e}")
                        return foto_antiga
                return foto_antiga

        # Se há uma nova foto ou não havia foto antes, move/copia a nova foto
        if foto_path and os.path.exists(foto_path):
            try:
                if os.path.abspath(foto_path) != os.path.abspath(novo_caminho):
                    shutil.copy(foto_path, novo_caminho)
                return novo_caminho
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao mover foto: {e}")
                return None
        return None

    # Exportações
    def exportar_carometro_pdf(self):
        turmas = ["Todas"] + [f"{t[0]} - {t[1]}" for t in self.model.carregar_turmas()]
        combo = self.view.exportar_carometro("PDF", self.exportar_pdf)
        combo["values"] = turmas
        combo.set("Todas")

    def exportar_carometro_excel(self):
        turmas = ["Todas"] + [f"{t[0]} - {t[1]}" for t in self.model.carregar_turmas()]
        combo = self.view.exportar_carometro("Excel", self.exportar_excel)
        combo["values"] = turmas
        combo.set("Todas")

    def exportar_carometro_word(self):
        turmas = ["Todas"] + [f"{t[0]} - {t[1]}" for t in self.model.carregar_turmas()]
        combo = self.view.exportar_carometro("Word", self.exportar_word)
        combo["values"] = turmas
        combo.set("Todas")

    def visualizar_carometro(self):
        # Chama o método visualizar_carometro no view com todos os callbacks necessários
        turma_combo, scrollable_frame = self.view.visualizar_carometro(self.atualizar_carometro)
        self.turma_combo = turma_combo
        self.scrollable_frame = scrollable_frame

    def atualizar_carometro(self):
        def tarefa():
            try:
                print("Debug: Iniciando tarefa de atualização do carômetro")
                turma_str = self.turma_combo.get()
                print(f"Debug: turma_str = {turma_str}")
                if not turma_str:
                    print("Debug: Nenhuma turma selecionada")
                    self.view.fechar_carregando()
                    return
                # Verifica se é "Todas" antes de tentar converter
                turma_id = None if turma_str == "Todas" else int(turma_str.split(" - ")[0])
                print(f"Debug: turma_id = {turma_id}")
                alunos = self.model.carregar_alunos_por_turma(turma_id)
                print(f"Debug: Alunos retornados = {alunos}")
                
                # Preparar os dados na thread secundária
                dados = []
                for aluno in alunos:
                    print(f"Debug: Processando aluno = {aluno}")
                    foto_path = aluno[5] if aluno[5] and os.path.exists(aluno[5]) else os.path.join(os.path.dirname(__file__), "imagens", "00000-sem imagem.jpg")
                    if os.path.exists(foto_path):
                        img = Image.open(foto_path)
                        img = img.resize((100, 100), Image.Resampling.LANCZOS)
                        photo = ImageTk.PhotoImage(img)
                    else:
                        photo = None
                    dados.append((aluno[1], photo, foto_path))  # aluno[1] é o nome
                    
                print(f"Debug: Dados preparados = {len(dados)} itens")
                # Agendar a atualização da UI na thread principal
                self.view.root.after(0, lambda: atualizar_interface(dados))
            except Exception as e:
                erro = e
                print(f"Debug: Erro capturado em tarefa: {erro}")
                self.view.root.after(0, lambda: [self.view.fechar_carregando(), messagebox.showerror("Erro", f"Erro ao atualizar carômetro: {erro}")])

        def atualizar_interface(dados):
            try:
                print("Debug: Iniciando atualização da interface")
                for widget in self.scrollable_frame.winfo_children():
                    widget.destroy()
                
                for i, (nome, photo, foto_path) in enumerate(dados):
                    aluno_frame = tk.Frame(self.scrollable_frame, bg=self.BG_COLOR, bd=2, relief="groove")
                    aluno_frame.grid(row=i//6, column=i%6, padx=5, pady=5, sticky="nsew")
                    
                    if photo:
                        tk.Label(aluno_frame, image=photo, bg=self.BG_COLOR).pack(pady=5)
                        aluno_frame.image = photo  # Manter referência
                    else:
                        tk.Label(aluno_frame, text="Sem Foto", font=self.FONT, bg=self.BG_COLOR, fg=self.FG_COLOR).pack(pady=5)
                    tk.Label(aluno_frame, text=nome, font=self.FONT, bg=self.BG_COLOR, fg=self.FG_COLOR).pack(pady=5)
                
                print("Debug: Interface atualizada com sucesso")
                self.view.fechar_carregando()
            except Exception as e:
                print(f"Debug: Erro ao atualizar interface: {e}")
                self.view.fechar_carregando()
                messagebox.showerror("Erro", f"Erro ao atualizar interface do carômetro: {e}")

        self.view.mostrar_carregando()
        threading.Thread(target=tarefa, daemon=True).start()

    def exportar_pdf(self):
        def tarefa():
            try:
                print("Debug: Iniciando exportar_pdf")
                turma_str = self.view.current_frame.winfo_children()[2].get()
                print(f"Debug: turma_str = {turma_str}")
                if not turma_str:
                    self.view.fechar_carregando()
                    messagebox.showwarning("Aviso", "Selecione uma turma!")
                    return
                turma_id = None if turma_str == "Todas" else int(turma_str.split(" - ")[0])
                print(f"Debug: turma_id = {turma_id}")
                alunos = self.model.carregar_alunos_por_turma(turma_id)
                print(f"Debug: Alunos retornados = {alunos}")
                if not alunos:
                    self.view.fechar_carregando()
                    messagebox.showinfo("Info", "Nenhum aluno encontrado para a turma selecionada.")
                    return

                agora = datetime.now()
                filtro = "todas" if turma_str == "Todas" else turma_str.split(" - ")[1].replace(" ", "_").lower()
                nome_arquivo = f"carometro-{filtro}-{agora.strftime('%Y-%m-%d_%H-%M-%S')}.pdf"
                documentos_dir = os.path.join(os.path.dirname(__file__), "documentos")
                os.makedirs(documentos_dir, exist_ok=True)
                pdf_path = os.path.join(documentos_dir, nome_arquivo)

                pdf = SimpleDocTemplate(pdf_path, pagesize=A4, leftMargin=9*mm, rightMargin=9*mm, topMargin=20*mm, bottomMargin=30*mm)
                elements = []
                styles = getSampleStyleSheet()
                style = styles["Normal"]
                style.fontSize = 8
                style.alignment = 1

                def build_pdf_footer(canvas, doc):
                    canvas.saveState()
                    page_number = canvas.getPageNumber()
                    data_hora = agora.strftime("%d/%m/%Y às %H:%M")
                    footer_text = f"Página {page_number}   |   Documento gerado em {data_hora}"
                    canvas.setFont("Helvetica", 8)
                    canvas.drawCentredString(A4[0] / 2, 15, footer_text)
                    canvas.restoreState()

                elements.append(Paragraph(f"Carômetro - {turma_str}", styles["Heading1"]))
                elements.append(Spacer(1, 12))

                elementos_linha = []
                for i, aluno in enumerate(alunos):
                    print(f"Debug: Processando aluno = {aluno}")
                    foto_path = aluno[5] if aluno[5] and os.path.exists(aluno[5]) else os.path.join(os.path.dirname(__file__), "imagens", "00000-sem imagem.jpg")
                    print(f"Debug: foto_path = {foto_path}")
                    primeiro_nome = aluno[1].split(" ")[0]
                    if os.path.exists(foto_path):
                        try:
                            img = RLImage(foto_path, width=28*mm, height=28*mm)
                        except Exception as e:
                            print(f"Debug: Erro ao carregar foto {foto_path}: {e}")
                            img = Paragraph("Sem Foto", style)
                    else:
                        print(f"Debug: Foto não encontrada: {foto_path}")
                        img = Paragraph("Sem Foto", style)
                    nome = Paragraph(primeiro_nome, style)
                    elementos_linha.append([img, nome])

                    if (i + 1) % 6 == 0 or i == len(alunos) - 1:
                        num_elementos = len(elementos_linha)
                        num_linhas = (num_elementos + 5) // 6
                        if num_elementos > 0:
                            table_data = []
                            for r in range(num_linhas):
                                inicio = r * 6
                                fim = min(inicio + 6, num_elementos)
                                linha = elementos_linha[inicio:fim]
                                while len(linha) < 6:
                                    linha.append([Paragraph("", style), Paragraph("", style)])
                                table_data.append(linha)

                            table = Table(table_data, colWidths=[32*mm]*6, rowHeights=[32*mm]*num_linhas)
                            table.setStyle(TableStyle([
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                                ('FONTSIZE', (0, 0), (-1, -1), 8),
                            ]))
                            elements.append(table)
                            elementos_linha = []
                            if (i + 1) % 42 == 0 and i != len(alunos) - 1:
                                elements.append(PageBreak())

                pdf.build(elements, onFirstPage=build_pdf_footer, onLaterPages=build_pdf_footer)
                self.view.fechar_carregando()
                messagebox.showinfo("Sucesso", f"Carômetro exportado como {nome_arquivo} na pasta 'documentos'!")
                os.startfile(pdf_path)
                self.view.tela_inicial()
            except Exception as e:
                self.view.fechar_carregando()
                messagebox.showerror("Erro", f"Erro ao exportar PDF: {e}")

        self.view.mostrar_carregando()
        threading.Thread(target=tarefa, daemon=True).start()

    def exportar_excel(self):
        def tarefa():
            try:
                turma_str = self.view.current_frame.winfo_children()[2].get()
                if not turma_str:
                    messagebox.showwarning("Aviso", "Selecione uma turma!")
                    return
                turma_id = None if turma_str == "Todas" else int(turma_str.split(" - ")[0])
                alunos = self.model.carregar_alunos_por_turma(turma_id)
                if not alunos:
                    messagebox.showinfo("Info", "Nenhum aluno encontrado para a turma selecionada.")
                    return

                agora = datetime.now()
                filtro = "todas" if turma_str == "Todas" else turma_str.split(" - ")[1].replace(" ", "_").lower()
                nome_arquivo = f"carometro-{filtro}-{agora.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
                documentos_dir = os.path.join(os.path.dirname(__file__), "documentos")
                os.makedirs(documentos_dir, exist_ok=True)
                excel_path = os.path.join(documentos_dir, nome_arquivo)

                wb = Workbook()
                ws = wb.active
                ws.title = "Carômetro"
                ws.append(["ID", "Nome", "Turma", "Curso", "Instituição", "Foto"])

                # Ajustar tamanho das colunas
                ws.column_dimensions['A'].width = 10  # ID
                ws.column_dimensions['B'].width = 20  # Nome
                ws.column_dimensions['C'].width = 15  # Turma
                ws.column_dimensions['D'].width = 15  # Curso
                ws.column_dimensions['E'].width = 20  # Instituição
                ws.column_dimensions['F'].width = 15  # Foto

                temp_dir = os.path.join(os.path.dirname(__file__), "temp")
                os.makedirs(temp_dir, exist_ok=True)

                for i, aluno in enumerate(alunos, start=2):  # Começa na linha 2 por causa do cabeçalho
                    ws.cell(row=i, column=1, value=aluno[0])  # ID
                    ws.cell(row=i, column=2, value=aluno[1])  # Nome
                    ws.cell(row=i, column=3, value=aluno[2])  # Turma
                    ws.cell(row=i, column=4, value=aluno[3])  # Curso
                    ws.cell(row=i, column=5, value=aluno[4])  # Instituição
                    
                    foto_path = aluno[5] if aluno[5] and os.path.exists(aluno[5]) else os.path.join(os.path.dirname(__file__), "imagens", "00000-sem imagem.jpg")
                    if os.path.exists(foto_path):
                        img = Image.open(foto_path)
                        img = img.resize((100, 100), Image.Resampling.LANCZOS)
                        temp_path = os.path.join(temp_dir, f"temp_image_{i}.jpg")
                        img.save(temp_path)  # Salva em um caminho temporário único
                        excel_img = XLImage(temp_path)
                        excel_img.anchor = f"F{i}"
                        ws.add_image(excel_img)
                        os.remove(temp_path)  # Remove o temporário após adicionar

                    ws.row_dimensions[i].height = 80  # Ajusta a altura da linha para a imagem

                wb.save(excel_path)

                self.view.fechar_carregando()
                messagebox.showinfo("Sucesso", f"Carômetro exportado como {nome_arquivo} na pasta 'documentos'!")
                os.startfile(excel_path)
                self.view.tela_inicial()
            except Exception as e:
                self.view.fechar_carregando()
                messagebox.showerror("Erro", f"Erro ao exportar Excel: {e}")

        self.view.mostrar_carregando()
        threading.Thread(target=tarefa, daemon=True).start()
        
    def exportar_word(self):
        def tarefa():
            try:
                print("Debug: Iniciando exportar_word")
                turma_str = self.view.current_frame.winfo_children()[2].get()
                print(f"Debug: turma_str = {turma_str}")
                if not turma_str:
                    self.view.fechar_carregando()
                    messagebox.showwarning("Aviso", "Selecione uma turma!")
                    return
                turma_id = None if turma_str == "Todas" else int(turma_str.split(" - ")[0])
                print(f"Debug: turma_id = {turma_id}")
                alunos = self.model.carregar_alunos_por_turma(turma_id)
                print(f"Debug: Alunos retornados = {alunos}")
                if not alunos:
                    self.view.fechar_carregando()
                    messagebox.showinfo("Info", "Nenhum aluno encontrado para a turma selecionada.")
                    return

                agora = datetime.now()
                filtro = "todas" if turma_str == "Todas" else turma_str.split(" - ")[1].replace(" ", "_").lower()
                nome_arquivo = f"carometro-{filtro}-{agora.strftime('%Y-%m-%d_%H-%M-%S')}.docx"
                documentos_dir = os.path.join(os.path.dirname(__file__), "documentos")
                os.makedirs(documentos_dir, exist_ok=True)
                word_path = os.path.join(documentos_dir, nome_arquivo)

                doc = Document()
                # Ajustar margens para corresponder ao PDF
                section = doc.sections[0]
                section.left_margin = Inches(0.35)  # 9mm
                section.right_margin = Inches(0.35)  # 9mm
                section.top_margin = Inches(0.79)   # 20mm
                section.bottom_margin = Inches(1.18)  # 30mm

                doc.add_heading(f"Carômetro - {turma_str}", 0)
                doc.add_paragraph()  # Spacer

                table = None
                row = None
                for i, aluno in enumerate(alunos):
                    print(f"Debug: Processando aluno = {aluno}")
                    if i % 6 == 0:
                        if table:
                            doc.add_paragraph()  # Spacer após a tabela
                        table = doc.add_table(rows=2, cols=6)
                        table.style = 'Table Grid'
                        for cell in table.rows[0].cells:
                            cell.width = Inches(1.26)  # Aproximadamente 32mm
                        for cell in table.rows[1].cells:
                            cell.width = Inches(1.26)
                        row = 0

                    foto_path = aluno[5] if aluno[5] and os.path.exists(aluno[5]) else os.path.join(os.path.dirname(__file__), "imagens", "00000-sem imagem.jpg")
                    print(f"Debug: foto_path = {foto_path}")
                    cell_img = table.rows[0].cells[i % 6]
                    cell_text = table.rows[1].cells[i % 6]
                    if os.path.exists(foto_path):
                        try:
                            cell_img.paragraphs[0].add_run().add_picture(foto_path, width=Inches(1.1), height=Inches(1.1))
                            cell_img.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        except Exception as e:
                            print(f"Debug: Erro ao carregar foto no Word {foto_path}: {e}")
                            cell_img.paragraphs[0].add_run("Sem Foto")
                            cell_img.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        print(f"Debug: Foto não encontrada: {foto_path}")
                        cell_img.paragraphs[0].add_run("Sem Foto")
                        cell_img.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    primeiro_nome = aluno[1].split(" ")[0]
                    cell_text.paragraphs[0].add_run(primeiro_nome)
                    cell_text.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cell_text.paragraphs[0].runs[0].font.size = Pt(8)

                    if (i + 1) % 42 == 0 and i != len(alunos) - 1:
                        doc.add_page_break()

                # Adicionar rodapé
                for section in doc.sections:
                    footer = section.footer
                    footer_paragraph = footer.paragraphs[0]
                    footer_paragraph.text = f"Página {footer_paragraph.add_run().add_break()} " \
                                            f"Documento gerado em {agora.strftime('%d/%m/%Y às %H:%M')}"
                    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    footer_paragraph.runs[0].font.size = Pt(8)

                doc.save(word_path)

                self.view.fechar_carregando()
                messagebox.showinfo("Sucesso", f"Carômetro exportado como {nome_arquivo} na pasta 'documentos'!")
                os.startfile(word_path)
                self.view.tela_inicial()
            except Exception as e:
                self.view.fechar_carregando()
                messagebox.showerror("Erro", f"Erro ao exportar Word: {e}")

        self.view.mostrar_carregando()
        threading.Thread(target=tarefa, daemon=True).start()

    def tela_inicial(self):
        frame = self.view.novo_frame()
        tk.Label(frame, text="Bem-vindo ao Sistema de Controle de Alunos", font=("Arial", 20, "bold"), bg="#FFFFFF", fg="#1A1A1A").pack(pady=20)
        tk.Label(frame, text="Este sistema permite gerenciar instituições, professores, cursos, turmas e alunos de forma eficiente.\nUtilize os menus acima para cadastrar, consultar e exportar dados.", font=("Arial", 10), bg="#FFFFFF", fg="#1A1A1A", justify="center").pack(pady=10)
        
        inst_count = len(self.model.carregar_instituicoes())
        prof_count = len(self.model.carregar_professores())
        curso_count = len(self.model.carregar_cursos())
        turma_count = len(self.model.carregar_turmas())
        aluno_count = len(self.model.carregar_alunos())
        
        stats_frame = tk.Frame(frame, bg="#F8F9FA", bd=2, relief="groove")
        stats_frame.pack(pady=10, padx=10, fill=tk.X)
        tk.Label(stats_frame, text=f"Instituições: {inst_count}", font=("Arial", 10), bg="#F8F9FA", fg="#1A1A1A", padx=10, pady=5).pack(side=tk.LEFT)
        tk.Label(stats_frame, text=f"Professores: {prof_count}", font=("Arial", 10), bg="#F8F9FA", fg="#1A1A1A", padx=10, pady=5).pack(side=tk.LEFT)
        tk.Label(stats_frame, text=f"Cursos: {curso_count}", font=("Arial", 10), bg="#F8F9FA", fg="#1A1A1A", padx=10, pady=5).pack(side=tk.LEFT)
        tk.Label(stats_frame, text=f"Turmas: {turma_count}", font=("Arial", 10), bg="#F8F9FA", fg="#1A1A1A", padx=10, pady=5).pack(side=tk.LEFT)
        tk.Label(stats_frame, text=f"Alunos: {aluno_count}", font=("Arial", 10), bg="#F8F9FA", fg="#1A1A1A", padx=10, pady=5).pack(side=tk.LEFT)