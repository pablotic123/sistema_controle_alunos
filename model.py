import sqlite3
import os
import shutil
import datetime

class SistemaModel:
    _instance = None
    _conn = None

    def __new__(cls, db_path):
        if cls._instance is None:
            cls._instance = super(SistemaModel, cls).__new__(cls)
            cls._instance.db_path = os.path.join(os.path.dirname(__file__), "dados", "controle_alunos.db")
            cls._instance._conn = sqlite3.connect(cls._instance.db_path)
            cls._instance.cache = {
                "instituicoes": None,
                "professores": None,
                "cursos": None,
                "turmas": None,
                "alunos": None
            }
            cls._instance.inicializar_banco()
        return cls._instance

    def inicializar_banco(self):
        try:
            cursor = self._conn.cursor()

            # Tabelas existentes
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

            # Tabela de log para auditoria
            cursor.execute('''CREATE TABLE IF NOT EXISTS log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tabela TEXT NOT NULL,
                acao TEXT NOT NULL,
                id_registro INTEGER NOT NULL,
                data_hora DATETIME DEFAULT CURRENT_TIMESTAMP)''')

            self._conn.commit()
        except sqlite3.Error as e:
            print(f"Erro ao inicializar o banco: {e}")
            raise

    def executar_query(self, query, params=(), fetch=False):
        try:
            cursor = self._conn.cursor()
            cursor.execute(query, params)
            if fetch:
                result = cursor.fetchall()
                return result if result else None
            else:
                self._conn.commit()
                return cursor.lastrowid if "INSERT" in query.upper() else None
        except sqlite3.Error as e:
            print(f"Erro ao executar query: {e}")
            raise
        finally:
            cursor.close()

    def commit(self):
        try:
            self._conn.commit()
        except sqlite3.Error as e:
            print(f"Erro ao forçar commit: {e}")
            raise

    def close(self):
        if self._conn:
            self._conn.commit()
            self._conn.close()
            self._conn = None
            self._instance = None
            self.cache = None

    def backup_db(self):
        try:
            backup_dir = os.path.join(os.path.dirname(__file__), "backups")
            os.makedirs(backup_dir, exist_ok=True)
            backup_path = os.path.join(backup_dir, f"controle_alunos_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
            self._conn.commit()  # Garantir que todas as alterações estejam salvas
            shutil.copy2(self.db_path, backup_path)
            print(f"Debug: Backup criado em {backup_path}")
            return backup_path
        except Exception as e:
            print(f"Erro ao criar backup: {e}")
            raise

    def log_operacao(self, tabela, acao, id_registro):
        try:
            self.executar_query("INSERT INTO log (tabela, acao, id_registro) VALUES (?, ?, ?)", (tabela, acao, id_registro))
            self.commit()
        except Exception as e:
            print(f"Erro ao registrar log: {e}")
            raise

    def salvar_instituicao(self, id, nome):
        try:
            if id:
                self.executar_query("UPDATE instituicao SET nome = ? WHERE id = ?", (nome, id))
                self.log_operacao("instituicao", "atualizado", id)
            else:
                id = self.executar_query("INSERT INTO instituicao (nome) VALUES (?)", (nome,))
                self.log_operacao("instituicao", "criado", id)
            self.cache["instituicoes"] = None  # Invalida o cache
            return id
        except Exception as e:
            print(f"Erro ao salvar instituição: {e}")
            raise

    def salvar_professor(self, id, nome, instituicao_id, foto):
        try:
            if id:
                self.executar_query("UPDATE professor SET nome = ?, instituicao_id = ?, foto = ? WHERE id = ?", (nome, instituicao_id, foto, id))
                self.log_operacao("professor", "atualizado", id)
            else:
                id = self.executar_query("INSERT INTO professor (nome, instituicao_id, foto) VALUES (?, ?, ?)", (nome, instituicao_id, foto))
                self.log_operacao("professor", "criado", id)
            self.cache["professores"] = None  # Invalida o cache
            return id
        except Exception as e:
            print(f"Erro ao salvar professor: {e}")
            raise

    def salvar_curso(self, id, nome, instituicao_id):
        try:
            if id:
                self.executar_query("UPDATE curso SET nome = ?, instituicao_id = ? WHERE id = ?", (nome, instituicao_id, id))
                self.log_operacao("curso", "atualizado", id)
            else:
                id = self.executar_query("INSERT INTO curso (nome, instituicao_id) VALUES (?, ?)", (nome, instituicao_id))
                self.log_operacao("curso", "criado", id)
            self.cache["cursos"] = None  # Invalida o cache
            return id
        except Exception as e:
            print(f"Erro ao salvar curso: {e}")
            raise

    def salvar_turma(self, id, nome, ano, curso_id):
        try:
            if id:
                self.executar_query("UPDATE turma SET nome = ?, ano = ?, curso_id = ? WHERE id = ?", (nome, ano, curso_id, id))
                self.log_operacao("turma", "atualizado", id)
            else:
                id = self.executar_query("INSERT INTO turma (nome, ano, curso_id) VALUES (?, ?, ?)", (nome, ano, curso_id))
                self.log_operacao("turma", "criado", id)
            self.cache["turmas"] = None  # Invalida o cache
            return id
        except Exception as e:
            print(f"Erro ao salvar turma: {e}")
            raise

    def salvar_aluno(self, id, nome, turma_id, foto):
        try:
            if id:
                self.executar_query("UPDATE aluno SET nome = ?, turma_id = ?, foto = ? WHERE id = ?", (nome, turma_id, foto, id))
                self.log_operacao("aluno", "atualizado", id)
            else:
                id = self.executar_query("INSERT INTO aluno (nome, turma_id, foto) VALUES (?, ?, ?)", (nome, turma_id, foto))
                self.log_operacao("aluno", "criado", id)
            self.cache["alunos"] = None  # Invalida o cache
            return id
        except Exception as e:
            print(f"Erro ao salvar aluno: {e}")
            raise

    def excluir_registro(self, tabela, id):
        try:
            self.executar_query(f"DELETE FROM {tabela} WHERE id = ?", (id,))
            self.log_operacao(tabela, "excluido", id)
            self.cache[tabela + "s"] = None  # Invalida o cache
        except Exception as e:
            print(f"Erro ao excluir registro de {tabela}: {e}")
            raise

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
        if self.cache["instituicoes"] is None:
            self.cache["instituicoes"] = self.executar_query("SELECT id, nome FROM instituicao", fetch=True) or []
        return self.cache["instituicoes"]

    def carregar_professores(self):
        if self.cache["professores"] is None:
            self.cache["professores"] = self.executar_query("SELECT id, nome, instituicao_id, foto FROM professor", fetch=True) or []
        return self.cache["professores"]

    def carregar_cursos(self):
        if self.cache["cursos"] is None:
            self.cache["cursos"] = self.executar_query("SELECT id, nome FROM curso", fetch=True) or []
        return self.cache["cursos"]

    def carregar_turmas(self):
        if self.cache["turmas"] is None:
            self.cache["turmas"] = self.executar_query("SELECT id, nome FROM turma", fetch=True) or []
        return self.cache["turmas"]

    def carregar_alunos(self):
        if self.cache["alunos"] is None:
            self.cache["alunos"] = self.executar_query("SELECT id, nome, turma_id, foto FROM aluno", fetch=True) or []
        return self.cache["alunos"]

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