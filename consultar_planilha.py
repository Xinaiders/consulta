# -*- coding: utf-8 -*-
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO ---
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
NOME_ARQUIVO_CREDENCIAS = 'credenciais.json'
NOME_DA_PLANILHA = 'Importrange-matriz'
NOME_DA_ABA = 'import_ran_Matriz'
NOME_ABA_SOLICITACOES = 'import_ran_Solicitacoes'
NOME_ABA_COMPRAS = 'import_ran_Pedido_Compra'
NOME_ABA_USUARIOS = 'Usuarios' # Nome da aba de utilizadores


class Planilha:
    def __init__(self):
        self.client = self._conectar()

    def _conectar(self):
        """Conecta-se à API do Google Sheets."""
        try:
            creds = Credentials.from_service_account_file(
                NOME_ARQUIVO_CREDENCIAS, scopes=SCOPES
            )
            return gspread.authorize(creds)
        except Exception as e:
            print(f"Erro ao conectar com a API: {e}")
            return None

    def carregar_usuarios(self):
        """Busca os dados dos utilizadores da aba 'Usuarios' e normaliza os cabeçalhos."""
        if not self.client:
            return {}
        try:
            spreadsheet = self.client.open(NOME_DA_PLANILHA)
            worksheet = spreadsheet.worksheet(NOME_ABA_USUARIOS)
            user_list = worksheet.get_all_records()
            
            usuarios_formatados = {}
            for user_row in user_list:
                normalized_user = {str(k).lower().strip(): v for k, v in user_row.items()}
                
                if 'username' in normalized_user:
                    username = str(normalized_user['username'])
                    usuarios_formatados[username] = normalized_user
            
            return usuarios_formatados
            
        except gspread.exceptions.WorksheetNotFound:
            print(f"ERRO: A aba de utilizadores '{NOME_ABA_USUARIOS}' não foi encontrada na planilha.")
            return {"admin": {"username": "admin", "password": "senha123", "name": "Admin Padrão"}}
        except Exception as e:
            print(f"Erro ao carregar utilizadores da planilha: {e}")
            return {}

    # --- FUNÇÃO DE OBTER SUGESTÕES ATUALIZADA E MAIS ROBUSTA ---
    def obter_sugestoes(self):
        """Busca a lista de códigos e descrições para o autocompletar de forma robusta."""
        if not self.client:
            return []
        try:
            spreadsheet = self.client.open(NOME_DA_PLANILHA)
            worksheet = spreadsheet.worksheet(NOME_DA_ABA)
            all_values = worksheet.get_all_values()
            if not all_values:
                return []

            header = [h.upper().strip() for h in all_values[0]]
            
            cod_index = -1
            desc_index = -1
            
            # Lógica de busca de colunas mais robusta
            if 'COD' in header:
                cod_index = header.index('COD')
            
            # Tenta várias combinações para a coluna de descrição
            possible_desc_names = ['DESCRIÇÃO COMPLETA', 'DESCRICAO COMPLETA', 'DESCRIÇÃO', 'DESCRICAO']
            for name in possible_desc_names:
                if name in header:
                    desc_index = header.index(name)
                    break # Para assim que encontrar a primeira correspondência
            
            if cod_index == -1 or desc_index == -1:
                print(f"ERRO: Não foi possível encontrar as colunas 'COD' e/ou 'DESCRIÇÃO' na aba principal.")
                print(f"Cabeçalhos encontrados na planilha: {header}") # Ajuda a depurar
                return []

            sugestoes_lista = []
            for row in all_values[1:]:
                if len(row) > cod_index and len(row) > desc_index:
                    codigo = row[cod_index].strip()
                    descricao = row[desc_index].strip()
                    if codigo and descricao:
                        sugestoes_lista.append(f"{codigo} - {descricao}")
            
            return sugestoes_lista
        except Exception as e:
            print(f"Erro ao obter sugestões: {e}")
            return []

    def _coerce_number(self, valor):
        """Converte strings no formato BR/EN para número sem inflar valores."""
        if isinstance(valor, (int, float)): return valor
        if valor is None: return 0
        if not isinstance(valor, str): return valor
        s = valor.strip()
        if s.isdigit():
            try: return int(s)
            except Exception: return s
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
            try: return float(s)
            except Exception: return valor
        if "." in s:
            partes = s.split(".")
            if len(partes[-1]) <= 2:
                try: return float(s)
                except Exception: pass
            s = s.replace(".", "")
            try: return float(s)
            except Exception: return valor
        return valor

    def _get_rows(self, worksheet):
        """Lê a aba e monta uma lista de dicionários (header + linhas)."""
        try:
            values = worksheet.get_all_values()
        except Exception as e:
            print(f"Erro ao ler valores: {e}")
            return []

        if not values: return []
        header = [h.strip() for h in values[0]]
        rows = []
        for raw in values[1:]:
            if len(raw) < len(header):
                raw = raw + [""] * (len(header) - len(raw))
            elif len(raw) > len(header):
                raw = raw[:len(header)]
            rows.append(dict(zip(header, raw)))
        return rows

    def buscar_item(self, termo_buscado):
        """Busca um item na planilha por código."""
        if not self.client: return None
        try:
            spreadsheet = self.client.open(NOME_DA_PLANILHA)
            ws = spreadsheet.worksheet(NOME_DA_ABA)
            dados = self._get_rows(ws)
            termo = str(termo_buscado).strip()

            for item in dados:
                cod_item = next((v for k, v in item.items() if k.upper().strip() == 'COD'), None)
                if str(cod_item).strip() == termo:
                    return item
            return None
        except Exception as e:
            print(f"Erro durante a busca do item: {e}")
            return None

    def buscar_suprimentos(self, codigo_buscado):
        """Busca informações nas abas de Solicitações e Compras."""
        solicitacoes, compras = [], []
        if not self.client: return {'solicitacoes': [], 'compras': []}
        try:
            spreadsheet = self.client.open(NOME_DA_PLANILHA)
            # Solicitações
            ws_sol = spreadsheet.worksheet(NOME_ABA_SOLICITACOES)
            dados_sol = self._get_rows(ws_sol)
            for linha in dados_sol:
                if str(linha.get('COD', '')).strip() == str(codigo_buscado) and \
                   str(linha.get('STATUS SOLICITACAO', '')).strip().upper() == "PENDENTE":
                    solicitacoes.append({
                        'Data da Solicitacao': linha.get('DATA SOLICITACAO', 'N/D'),
                        'Status de Pendencia': linha.get('STATUS SOLICITACAO', 'N/D'),
                        'Nivel de Prioridade': linha.get('NIVEL PRIORIDADE', 'N/D'),
                        'Quantidade': self._coerce_number(linha.get('SOLICITACAO', '0'))
                    })
            # Compras
            ws_comp = spreadsheet.worksheet(NOME_ABA_COMPRAS)
            dados_comp = self._get_rows(ws_comp)
            for linha in dados_comp:
                if str(linha.get('COD', '')).strip() == str(codigo_buscado) and \
                   str(linha.get('STATUS PEDIDO', '')).strip().upper() == "EM ABERTO":
                    compras.append({
                        'Quantidade': self._coerce_number(linha.get('QUANTIDADE', '0')),
                        'Fornecedor': linha.get('FORNECEDOR', 'N/D'),
                        'Data do Pedido': linha.get('DATA DO PEDIDO', 'N/D'),
                        'Previsao de Chegada': linha.get('PREVISAO DE ENTREGA', 'N/D')
                    })
        except Exception as e:
            print(f"Erro ao buscar suprimentos: {e}")
        return {'solicitacoes': solicitacoes, 'compras': compras}

