# model.py
import sqlite3
import os

class SistemaModel:
    def __init__(self):
        self.db_path = os.path.join(os.path.dirname(__file__), "dados", "controle_alunos.db")
        self.inicializar_banco()

    def inicializar_banco(self):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            cursor.execute('''CREATE TABLE IF NOT EXISTS instituicao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL)''')
            cursor.execute('''CREATE TABLE IF NOT EXISTS professor (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                instituicao_id INTEGER NOT NULL,
                foto TEXT,
                FOREIGN KEY (instituicao_id) REFERENCES instituicao(id) ON DELETE RESTRICT)''')
            cursor.execute('''CREATE TABLE IF NOT EXISTS curso (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                instituicao_id INTEGER NOT NULL,
                FOREIGN KEY (instituicao_id) REFERENCES instituicao(id) ON DELETE RESTRICT)''')
            cursor.execute('''CREATE TABLE IF NOT EXISTS turma (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                ano INTEGER NOT NULL,
                curso_id INTEGER NOT NULL,
                FOREIGN KEY (curso_id) REFERENCES curso(id) ON DELETE RESTRICT)''')
            cursor.execute('''CREATE TABLE IF NOT EXISTS aluno (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                turma_id INTEGER NOT NULL,
                foto TEXT,
                FOREIGN KEY (turma_id) REFERENCES turma(id) ON DELETE RESTRICT)''')

            conn.commit()
        except sqlite3.Error as e:
            print(f"Erro ao inicializar o banco: {e}")
        finally:
            conn.close()

    def executar_query(self, query, params=(), fetch=False):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute(query, params)
            if fetch:
                result = cursor.fetchall()
            else:
                conn.commit()
                result = cursor.lastrowid if "INSERT" in query.upper() else None
            return result
        except sqlite3.Error as e:
            print(f"Erro ao executar query: {e}")
            return None
        finally:
            conn.close()

    def salvar_instituicao(self, id, nome):
        if id:
            self.executar_query("UPDATE instituicao SET nome = ? WHERE id = ?", (nome, id))
        else:
            self.executar_query("INSERT INTO instituicao (nome) VALUES (?)", (nome,))

    def salvar_professor(self, id, nome, instituicao_id, foto):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            if id:
                cursor.execute("UPDATE professor SET nome = ?, instituicao_id = ?, foto = ? WHERE id = ?", (nome, instituicao_id, foto, id))
            else:
                cursor.execute("INSERT INTO professor (nome, instituicao_id, foto) VALUES (?, ?, ?)", (nome, instituicao_id, foto))
                id = cursor.lastrowid  # Captura o ID gerado
            conn.commit()
            return id
        
    def salvar_curso(self, id, nome, instituicao_id):
        if id:
            self.executar_query("UPDATE curso SET nome = ?, instituicao_id = ? WHERE id = ?", (nome, instituicao_id, id))
        else:
            self.executar_query("INSERT INTO curso (nome, instituicao_id) VALUES (?, ?)", (nome, instituicao_id))

    def salvar_turma(self, id, nome, ano, curso_id):
        if id:
            self.executar_query("UPDATE turma SET nome = ?, ano = ?, curso_id = ? WHERE id = ?", (nome, ano, curso_id, id))
        else:
            self.executar_query("INSERT INTO turma (nome, ano, curso_id) VALUES (?, ?, ?)", (nome, ano, curso_id))

    def salvar_aluno(self, id, nome, turma_id, foto):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            if id:
                cursor.execute("UPDATE aluno SET nome = ?, turma_id = ?, foto = ? WHERE id = ?", (nome, turma_id, foto, id))
            else:
                cursor.execute("INSERT INTO aluno (nome, turma_id, foto) VALUES (?, ?, ?)", (nome, turma_id, foto))
                id = cursor.lastrowid  # Captura o ID gerado
            conn.commit()
            return id
        
    def excluir_registro(self, tabela, id):
        self.executar_query(f"DELETE FROM {tabela} WHERE id = ?", (id,))

    def consulta_instituicoes(self, filtros):
        query = "SELECT id, nome FROM instituicao WHERE 1=1"
        params = []
        if filtros.get("nome"):
            query += " AND nome LIKE ?"
            params.append(f"%{filtros['nome']}%")
        return self.executar_query(query, params, fetch=True) or []

    def consulta_professores(self, filtros):
        query = "SELECT p.id, p.nome, i.nome FROM professor p JOIN instituicao i ON p.instituicao_id = i.id WHERE 1=1"
        params = []
        if filtros.get("nome"):
            query += " AND p.nome LIKE ?"
            params.append(f"%{filtros['nome']}%")
        if filtros.get("instituicao"):
            query += " AND i.nome LIKE ?"
            params.append(f"%{filtros['instituicao']}%")
        return self.executar_query(query, params, fetch=True) or []

    def consulta_cursos(self, filtros):
        query = "SELECT c.id, c.nome, i.nome FROM curso c JOIN instituicao i ON c.instituicao_id = i.id WHERE 1=1"
        params = []
        if filtros.get("nome"):
            query += " AND c.nome LIKE ?"
            params.append(f"%{filtros['nome']}%")
        if filtros.get("instituicao"):
            query += " AND i.nome LIKE ?"
            params.append(f"%{filtros['instituicao']}%")
        return self.executar_query(query, params, fetch=True) or []

    def consulta_turmas(self, filtros):
        query = "SELECT t.id, t.nome, t.ano, c.nome FROM turma t JOIN curso c ON t.curso_id = c.id WHERE 1=1"
        params = []
        if filtros.get("nome"):
            query += " AND t.nome LIKE ?"
            params.append(f"%{filtros['nome']}%")
        if filtros.get("ano"):
            query += " AND t.ano = ?"
            params.append(filtros['ano'])
        if filtros.get("curso"):
            query += " AND c.nome LIKE ?"
            params.append(f"%{filtros['curso']}%")
        return self.executar_query(query, params, fetch=True) or []

    def consulta_alunos(self, filtros):
        query = "SELECT a.id, a.nome, t.nome, c.nome, i.nome FROM aluno a JOIN turma t ON a.turma_id = t.id JOIN curso c ON t.curso_id = c.id JOIN instituicao i ON c.instituicao_id = i.id WHERE 1=1"
        params = []
        if filtros.get("nome"):
            query += " AND a.nome LIKE ?"
            params.append(f"%{filtros['nome']}%")
        if filtros.get("turma"):
            query += " AND t.nome LIKE ?"
            params.append(f"%{filtros['turma']}%")
        if filtros.get("curso"):
            query += " AND c.nome LIKE ?"
            params.append(f"%{filtros['curso']}%")
        if filtros.get("instituicao"):
            query += " AND i.nome LIKE ?"
            params.append(f"%{filtros['instituicao']}%")
        return self.executar_query(query, params, fetch=True) or []

    def carregar_instituicoes(self):
        return self.executar_query("SELECT id, nome FROM instituicao", fetch=True) or []

    def carregar_professores(self):
        return self.executar_query("SELECT id, nome, instituicao_id, foto FROM professor", fetch=True) or []

    def carregar_cursos(self):
        return self.executar_query("SELECT id, nome FROM curso", fetch=True) or []

    def carregar_turmas(self):
        return self.executar_query("SELECT id, nome FROM turma", fetch=True) or []

    def carregar_alunos(self):
        return self.executar_query("SELECT id, nome, turma_id, foto FROM aluno", fetch=True) or []

    def carregar_alunos_por_turma(self, turma_id=None):
        if turma_id is not None:
            return self.executar_query("""
                SELECT a.id, a.nome, t.nome as turma, c.nome as curso, i.nome as instituicao, a.foto
                FROM aluno a
                JOIN turma t ON a.turma_id = t.id
                JOIN curso c ON t.curso_id = c.id
                JOIN instituicao i ON c.instituicao_id = i.id
                WHERE a.turma_id = ?
                ORDER BY a.nome
            """, (turma_id,), fetch=True) or []
        else: 
            return self.executar_query("""
                SELECT a.id, a.nome, t.nome as turma, c.nome as curso, i.nome as instituicao, a.foto
                FROM aluno a
                JOIN turma t ON a.turma_id = t.id
                JOIN curso c ON t.curso_id = c.id
                JOIN instituicao i ON c.instituicao_id = i.id
                ORDER BY a.nome
            """, fetch=True) or []