import os
import json
import datetime
import tkinter as tk
import shutil
from tkinter import ttk, messagebox, simpledialog, filedialog, Canvas, PhotoImage, Label
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageTk


# Arquivos de armazenamento
ITENS_FILE = "itens.json"
FORNECEDORES_FILE = "fornecedores.json"
LOGS_FILE = "logs.json"
BACKUP_FOLDER = "backups"

# Cria pasta de backup se não existir
if not os.path.exists(BACKUP_FOLDER):
    os.makedirs(BACKUP_FOLDER)

# Função para formatar valor em formato brasileiro (ex.: R$1.234,56)
def formatar_valor(valor_float):
    valor_formatado = f"{valor_float:,.2f}"  # ex.: 1,234.56
    valor_formatado = valor_formatado.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R${valor_formatado}"

# Função para gerar o próximo ID sequencial (formato 001, 002, etc.)
def gerar_proximo_id():
    if data_manager.itens:
        ids = [int(item["id"]) for item in data_manager.itens]
        novo_id = max(ids) + 1
    else:
        novo_id = 1
    return f"{novo_id:03d}"

# --- DataManager: Gerencia a leitura e escrita dos dados em JSON ---
class DataManager:
    def __init__(self):
        self.itens = self.load_data(ITENS_FILE)
        self.fornecedores = self.load_data(FORNECEDORES_FILE)
        self.logs = self.load_data(LOGS_FILE)
        

    def load_data(self, filename):
        if os.path.exists(filename):
            with open(filename, "r", encoding="utf-8") as f:
                try:
                    return json.load(f)
                except json.JSONDecodeError:
                    return []
        else:
            return []

    def save_data(self, filename, data):
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def salvar_tudo(self):
        self.save_data(ITENS_FILE, self.itens)
        self.save_data(FORNECEDORES_FILE, self.fornecedores)
        self.save_data(LOGS_FILE, self.logs)

    def adicionar_item(self, novo_item):
        """
        Se já existir um item com o mesmo nome (ignorando maiúsculas/minúsculas), atualiza:
          - Verifica se o novo serial (dentro de novo_item["serials"][0]) já está cadastrado.
          - Se não estiver, acrescenta o serial (com seu status) e soma a quantidade.
        Senão, adiciona um novo item com um novo ID.
        """
        for item in self.itens:
            if item["nome"].strip().lower() == novo_item["nome"].strip().lower():
                # Se o item não possuir a chave "serials", converte o antigo campo
                if "serials" not in item:
                    if "serial" in item:
                        item["serials"] = [{"serial": item.pop("serial"), "status": item.get("status", "Em Uso")}]
                    else:
                        item["serials"] = []
                # Verifica se o serial já foi cadastrado
                novo_serial = novo_item["serials"][0]["serial"]
                for s in item["serials"]:
                    if s["serial"] == novo_serial:
                        messagebox.showerror("Erro", f"O serial '{novo_serial}' já foi cadastrado para o item '{item['nome']}'!")
                        return False
                # Acrescenta o novo serial (com seu status)
                item["serials"].append(novo_item["serials"][0])
                # Soma as quantidades
                item["quantidade"] += novo_item["quantidade"]
                self.salvar_tudo()
                return True
        # Novo item
        self.itens.append(novo_item)
        self.salvar_tudo()
        return True

    def adicionar_fornecedor(self, fornecedor):
        self.fornecedores.append(fornecedor)
        self.salvar_tudo()

    def adicionar_log(self, log):
        self.logs.append(log)
        self.salvar_tudo()

# Instância global do DataManager
data_manager = DataManager()

# --- Função para Backup Manual e Automático ---
def realizar_backup():
    # Criar estrutura de pastas: Dia e Horário
    agora = datetime.datetime.now()
    data_backup = agora.strftime("%d-%m-%Y")
    hora_backup = agora.strftime("%H-%M")
    
    backup_pasta = os.path.join(BACKUP_FOLDER, data_backup, hora_backup)

    if not os.path.exists(backup_pasta):
        os.makedirs(backup_pasta)

    # Lista de arquivos JSON a serem salvos separadamente
    arquivos_json = {
        "itens": data_manager.itens,
        "fornecedores": data_manager.fornecedores,
        "logs": data_manager.logs
    }

    # Salvar cada JSON separadamente
    for nome_arquivo, dados in arquivos_json.items():
        caminho_arquivo = os.path.join(backup_pasta, f"{nome_arquivo}.json")
        try:
            with open(caminho_arquivo, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Erro ao salvar {nome_arquivo}.json: {e}")

    # Backup da pasta 'relatorios'
    relatorios_pasta = "relatorios"
    destino_relatorios = os.path.join(backup_pasta, "relatorios")

    if os.path.exists(relatorios_pasta):
        if not os.path.exists(destino_relatorios):
            os.makedirs(destino_relatorios)

        # Copia apenas arquivos, evitando diretórios protegidos
        for arquivo in os.listdir(relatorios_pasta):
            origem = os.path.join(relatorios_pasta, arquivo)
            destino = os.path.join(destino_relatorios, arquivo)

            if os.path.isfile(origem):  # Garante que é um arquivo
                try:
                    shutil.copy2(origem, destino)
                except PermissionError:
                    print(f"⚠ Permissão negada ao copiar: {origem}")

    # Mensagem de sucesso no console
    print(f"✔ Backup realizado em: {backup_pasta}")

def agendar_backup_automatico(root):
    realizar_backup()
    root.after(1800000, lambda: agendar_backup_automatico(root)) 
    print("⏳ Próximo backup em 30 minutos...")
    

def reduzir_opacidade(imagem_path, alpha):
        """Reduz a opacidade da imagem (0.0 a 1.0)."""
        imagem = Image.open(imagem_path).convert("RGBA")
        nova_imagem = Image.new("RGBA", imagem.size, (255, 255, 255, 0))

        for x in range(imagem.width):
            for y in range(imagem.height):
                r, g, b, a = imagem.getpixel((x, y))
                nova_imagem.putpixel((x, y), (r, g, b, int(a * alpha)))
        
        return nova_imagem

# --- Interface Gráfica com Tkinter ---
class MainMenu(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MS Connect - Controle de Estoque")
        self.geometry("1920x760")
        self.configure(bg="#f0f0f0")
        
        self.criar_fundo()
        self.create_widgets()
        
    
    def criar_fundo(self):
        """Carrega e exibe a imagem de fundo com opacidade."""
        try:
            # Carregar e redimensionar a imagem para caber na tela
            imagem = Image.open("bg.png").convert("RGBA")
            imagem = imagem.resize((2000, 960), Image.Resampling.LANCZOS)  # Ajusta ao tamanho da tela
            
            # Reduzir opacidade da imagem (0.3 = 30% visibilidade)
            imagem_translucida = reduzir_opacidade("bg.png", 0.8)
            
            self.fundo_img = ImageTk.PhotoImage(imagem_translucida)

            # Criar Label para exibir a imagem
            self.fundo_label = Label(self, image=self.fundo_img)
            self.fundo_label.place(x=-30, y=100)  # Cobrir a tela inteira
            self.fundo_label.lower()  # Garante que fique atrás de tudo

        except Exception as e:
            print(f"Erro ao carregar a imagem de fundo: {e}")
        
        
        
    def create_widgets(self):
        # Criando a barra superior ondulada
        self.create_top_bar()

        # Área dos botões centrais
        content_frame = tk.Frame(self, bg="#f0f0f0")
        content_frame.pack(expand=True)

        btn_cadastro_item = tk.Button(content_frame, text="Cadastro de Item", width=30, bg="#7617bb", fg="white", command=self.abrir_cadastro_item)
        btn_cadastro_item.pack(pady=5)

        btn_cadastro_fornecedor = tk.Button(content_frame, text="Cadastro de Fornecedor", width=30, bg="#7617bb", fg="white", command=self.abrir_cadastro_fornecedor)
        btn_cadastro_fornecedor.pack(pady=5)

        btn_visualizar_estoque = tk.Button(content_frame, text="Visualizar Estoque", width=30, bg="#7617bb", fg="white", command=self.abrir_visualizar_estoque)
        btn_visualizar_estoque.pack(pady=5)

        btn_backup = tk.Button(content_frame, text="Backup", width=30, bg="#7617bb", fg="white", command=self.realizar_backup)
        btn_backup.pack(pady=5)
        
        btn_log = tk.Button(content_frame, text="Log de Movimentação", width=30, bg="#7617bb", fg="white", command=self.abrir_log_movimentacao)
        btn_log.pack(pady=5)

        # Criando o rodapé
        self.create_footer()
    
    def create_top_bar(self):
        """Cria uma barra superior ondulada moderna"""
        canvas = Canvas(self, height=30, bg="#7617bb", highlightthickness=0)
        canvas.pack(fill="x")

        # Criando a onda superior
        canvas.create_line(0, 80, 150, 60, 300, 100, 450, 60, 600, 80, smooth=True, fill="#7617bb", width=30)

        # Adicionando o título na barra
        label_title = tk.Label(self, text="Controle de Estoque", font=("Helvetica", 12, "bold"), fg="white", bg="#7617bb")
        label_title.place(x=100, y=3)
    
    
        
        try:
            # Reduzir opacidade da imagem (0.3 = 30% de visibilidade)
            imagem_translucida = reduzir_opacidade("img.png", 1)
            self.img = ImageTk.PhotoImage(imagem_translucida)

            # Criar Label para exibir a imagem
            self.img_label = Label(self, image=self.img)
            self.img_label.place(x=1136, y=31)  # Posiciona no canto superior direito
            
            

        except Exception as e:
            print(f"Erro ao carregar a imagem: {e}")
            
    

    def create_footer(self):
        """Cria um rodapé fino com o nome da empresa"""
        footer = tk.Frame(self, height=30, bg="#7617bb")
        footer.pack(fill="x", side="bottom")

        label_footer = tk.Label(footer, text="MS Connect ®️", font=("Helvetica", 10, "bold"), fg="white", bg="#7617bb")
        label_footer.pack(pady=5)

    def abrir_cadastro_item(self):
        CadastroItem(self)
        print("Abrir Cadastro de Item")

    def abrir_cadastro_fornecedor(self):
        CadastroFornecedor(self)
        print("Abrir Cadastro de Fornecedor")

    def abrir_visualizar_estoque(self):
        VisualizarEstoque(self)
        print("Abrir Visualizar Estoque")

    def realizar_backup(self):
        realizar_backup()
        print("Realizar Backup")
    
    def abrir_log_movimentacao(self):
        LogMovimentacao(self)
        print("Abrir Janela de Log")


class LogMovimentacao(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Log de Movimentação")
        self.geometry("900x500")
        self.create_widgets()
        self.carregar_logs()

    def create_widgets(self):
        # Frame para as barras de busca
        search_frame = tk.Frame(self)
        search_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(search_frame, text="Buscar ID:").pack(side="left", padx=(0,5))
        self.search_id = tk.Entry(search_frame)
        self.search_id.pack(side="left", padx=(0,15), fill="x", expand=True)
        self.search_id.bind("<KeyRelease>", self.filtrar_logs)
        
        tk.Label(search_frame, text="Buscar Serial:").pack(side="left", padx=(0,5))
        self.search_serial = tk.Entry(search_frame)
        self.search_serial.pack(side="left", padx=(0,15), fill="x", expand=True)
        self.search_serial.bind("<KeyRelease>", self.filtrar_logs)
        
        # Treeview para exibir os logs
        colunas = ("item_id", "nome", "serial", "status_antigo", "novo_status", "motivo", "responsavel", "data_alteracao")
        self.tree = ttk.Treeview(self, columns=colunas, show="headings")
        for col in colunas:
            self.tree.heading(col, text=col.replace("_", " ").title())
            self.tree.column(col, width=100)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configurar as tags de cores (tons claros)
        self.tree.tag_configure("incremento", background="#ccffcc")   # Verde claro
        self.tree.tag_configure("decremento", background="#ffcccc")   # Vermelho claro
        self.tree.tag_configure("novo", background="#cce5ff")         # Azul claro
        self.tree.tag_configure("manutencao", background="#ffffcc")   # Amarelo claro
        self.tree.tag_configure("descartado", background="#e6ccb3")    # Marrom claro

    def carregar_logs(self):
        """Carrega os logs a partir do data_manager e os armazena em self.logs."""
        self.logs = data_manager.logs  # Supondo que data_manager.logs seja uma lista de dicionários
        self.atualizar_treeview()

    def atualizar_treeview(self):
        """Limpa e atualiza a Treeview com os logs filtrados e aplica as cores."""
        # Limpa a Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Recupera os valores digitados nos campos de busca
        id_busca = self.search_id.get().strip().lower()
        serial_busca = self.search_serial.get().strip().lower()
        
        # Insere os logs que atendem aos filtros
        for log in self.logs:
            log_id = str(log.get("item_id", "")).lower()
            log_serial = str(log.get("serial", "")).lower()
            
            if id_busca and id_busca not in log_id:
                continue
            if serial_busca and serial_busca not in log_serial:
                continue

            valores = (
                log.get("item_id", ""),
                log.get("nome", ""),
                log.get("serial", ""),
                log.get("status_antigo", ""),
                log.get("novo_status", ""),
                log.get("motivo", ""),
                log.get("responsavel", ""),
                log.get("data_alteracao", "")
            )
            
            tag = self.determinar_tag(log)
            self.tree.insert("", "end", values=valores, tags=(tag,))

    def determinar_tag(self, log):
        novo_status = str(log.get("novo_status", "")).lower()
        motivo = str(log.get("motivo", "")).lower()
        
        if novo_status.isdigit() or (novo_status.startswith("-") and novo_status[1:].isdigit()):
            # Se for um número (incremento ou decremento)
            if float(novo_status) > 0:
                return "incremento"   # Incremento: verde claro
            else:
                return "decremento"   # Decremento: vermelho claro
        elif "novo" in novo_status or "cadastrado" in motivo:
            return "novo"         # Item novo: azul claro
        elif "manutencao" in novo_status or "manutenção" in novo_status or "manutencao" in motivo or "manutenção" in motivo:
            return "manutencao"   # Manutenção: amarelo claro
        elif "descartado" in novo_status or "descartado" in motivo:
            return "descartado"   # Descartado: marrom claro
        else:
            return ""


    def filtrar_logs(self, event=None):
        """Atualiza a visualização dos logs conforme o que foi digitado."""
        self.atualizar_treeview()
        
        
# --- Tela de Cadastro de Item ---
class CadastroItem(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Cadastro de Item")
        self.geometry("400x600")
        self.create_widgets()

    def create_widgets(self):
        # Linha para selecionar um item já cadastrado (opcional)
        tk.Label(self, text="Item já cadastrado (opcional):").pack(pady=2)
        self.cmb_item_existente = ttk.Combobox(self, state="readonly")
        self.atualizar_lista_itens()
        self.cmb_item_existente.current(0)
        self.cmb_item_existente.pack(pady=2, fill="x", padx=10)
        self.cmb_item_existente.bind("<<ComboboxSelected>>", self.item_selecionado)

        # Campo Nome
        tk.Label(self, text="Nome:").pack(pady=2)
        self.ent_nome = tk.Entry(self)
        self.ent_nome.pack(pady=2, fill="x", padx=10)

        # Serial
        tk.Label(self, text="Serial:").pack(pady=2)
        self.ent_serial = tk.Entry(self)
        self.ent_serial.pack(pady=2, fill="x", padx=10)

        # Status (para o serial)
        tk.Label(self, text="Status:").pack(pady=2)
        self.status_var = tk.StringVar(value="Em Uso")
        opcoes_status = ["Em Uso", "Manutenção", "Descartado", "Inativo"]
        self.cmb_status = ttk.Combobox(self, textvariable=self.status_var, values=opcoes_status, state="readonly")
        self.cmb_status.pack(pady=2, fill="x", padx=10)
        
        # Departamento
        tk.Label(self, text="Departamento:").pack(pady=2)
        self.departamento_var = tk.StringVar()
        opcoes_departamento = ["B2C", "B2B", "RH"]
        self.cmb_departamento = ttk.Combobox(self, textvariable=self.departamento_var, values=opcoes_departamento, state="readonly")
        self.cmb_departamento.pack(pady=2, fill="x", padx=10)


        # Fornecedor
        tk.Label(self, text="Fornecedor:").pack(pady=2)
        self.fornecedor_var = tk.StringVar()
        fornecedores_lista = [f["nome"] for f in data_manager.fornecedores]
        self.cmb_fornecedor = ttk.Combobox(self, textvariable=self.fornecedor_var, values=fornecedores_lista, state="readonly")
        self.cmb_fornecedor.pack(pady=2, fill="x", padx=10)

        # Valor Unitário
        tk.Label(self, text="Valor Unitário:").pack(pady=2)
        self.ent_valor = tk.Entry(self)
        self.ent_valor.pack(pady=2, fill="x", padx=10)
        self.ent_valor.bind("<FocusOut>", self.formatar_valor_evento)

        # Quantidade
        tk.Label(self, text="Quantidade:").pack(pady=2)
        self.ent_quantidade = tk.Entry(self)
        self.ent_quantidade.pack(pady=2, fill="x", padx=10)

        # Data de Cadastro
        tk.Label(self, text="Data:").pack(pady=2)
        self.ent_data = tk.Entry(self)
        data_atual = datetime.date.today().strftime("%d/%m/%Y")  # Formato de data com barras
        self.ent_data.insert(0, data_atual)  # Preenche com a data atual
        self.ent_data.pack(pady=2, fill="x", padx=10)
        self.ent_data.bind("<FocusOut>", self.formatar_data_evento)
        
        tk.Button(self, text="Salvar Item", bg="#800080", fg="white", command=self.salvar_item).pack(pady=15)
        
    def formatar_data_evento(self, event):
        data_texto = self.ent_data.get().strip()
        if data_texto:
            # Remove as barras antes de formatar
            data_limpa = data_texto.replace("/", "")
            if len(data_limpa) == 8:  # Se a data tem 8 dígitos (DDMMYYYY)
                # Formata a data para o formato DD/MM/YYYY
                data_formatada = f"{data_limpa[:2]}/{data_limpa[2:4]}/{data_limpa[4:]}"
                self.ent_data.delete(0, tk.END)
                self.ent_data.insert(0, data_formatada)
            else:
                messagebox.showerror("Erro", "Data deve estar no formato DD/MM/YYYY!")
                self.ent_data.delete(0, tk.END)

        

    def atualizar_lista_itens(self):
        lista = ["Novo item"] + [item["nome"] for item in data_manager.itens]
        self.cmb_item_existente['values'] = lista

    def item_selecionado(self, event):
        
        selecionado = self.cmb_item_existente.get()
        if selecionado != "Novo item":
            # Preenche o nome e o valor unitário do item selecionado
            item_selecionado = next((item for item in data_manager.itens if item["nome"] == selecionado), None)
            if item_selecionado:
                self.ent_nome.delete(0, tk.END)
                self.ent_nome.insert(0, item_selecionado["nome"])
                self.ent_nome.config(state="disabled")  # Nome não editável
                self.ent_valor.delete(0, tk.END)
                self.ent_valor.insert(0, formatar_valor(item_selecionado["valor_unitario"]))  # Preenche o valor unitário
                self.ent_valor.config(state="normal")  # Permite editar o valor
        else:
            self.ent_nome.config(state="normal")  # Permite editar o nome
            self.ent_nome.delete(0, tk.END)
            self.ent_valor.config(state="normal")  # Permite editar o valor
            self.ent_valor.delete(0, tk.END)

    def formatar_valor_evento(self, event):
        valor_texto = self.ent_valor.get().strip()
        if valor_texto:
            try:
                valor_limpo = valor_texto.replace("R$", "").replace(".", "").replace(",", ".").strip()
                valor_float = float(valor_limpo)
                self.ent_valor.delete(0, tk.END)
                self.ent_valor.insert(0, formatar_valor(valor_float))
            except ValueError:
                pass

    def salvar_item(self):
        if self.cmb_item_existente.get() != "Novo item":
            nome = self.cmb_item_existente.get().strip()
        else:
            nome = self.ent_nome.get().strip()
        serial = self.ent_serial.get().strip()
        status = self.status_var.get()
        fornecedor = self.fornecedor_var.get()
        valor_texto = self.ent_valor.get().strip()
        quantidade_texto = self.ent_quantidade.get().strip()
        data_cadastro = self.ent_data.get().strip()
        departamento = self.departamento_var.get().strip() 

        if not nome or not serial or not valor_texto or not data_cadastro or not fornecedor or not quantidade_texto:
            messagebox.showerror("Erro", "Todos os campos são obrigatórios!")
            return

        try:
            valor_limpo = valor_texto.replace("R$", "").replace(".", "").replace(",", ".").strip()
            valor_float = float(valor_limpo)
            if valor_float < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Valor Unitário deve ser um número positivo!")
            return

        try:
            quantidade = int(quantidade_texto)
            if quantidade <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Quantidade deve ser um número inteiro positivo!")
            return

        data_cadastro = self.ent_data.get().strip()
        try:
            # Tenta parsear a data no formato DD/MM/YYYY
            data_validada = datetime.datetime.strptime(data_cadastro, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Erro", "Data deve estar no formato DD/MM/YYYY!")
            return

        # Cria o novo item; note que o serial é armazenado como dicionário (com seu status)
        novo_item = {
            "id": gerar_proximo_id(),
            "nome": nome,
            "serials": [{"serial": serial, "status": status}],
            "departamento": departamento,
            "fornecedor": fornecedor,
            "valor_unitario": valor_float,
            "quantidade": quantidade,
            "data_cadastro": data_cadastro
        }

        if data_manager.adicionar_item(novo_item):
            # Cria um log para item novo
            novo_log = {
                "item_id": novo_item["id"],
                "nome": novo_item["nome"],
                "serial": novo_item["serials"][0]["serial"] if novo_item.get("serials") else "",
                "departamento": novo_item.get("departamento", ""),
                "status_antigo": "",
                "novo_status": "Novo",  # Indica que é um item novo
                "motivo": "Item cadastrado",
                "responsavel": "Usuário",  # ou capturar essa informação
                "data_alteracao": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            data_manager.adicionar_log(novo_log)
            data_manager.salvar_tudo()
            messagebox.showinfo("Sucesso", "Item cadastrado/atualizado com sucesso!")
            self.destroy()

# --- Tela de Cadastro de Fornecedor ---
import tkinter as tk
from tkinter import messagebox
import re

class CadastroFornecedor(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Cadastro de Fornecedor")
        self.geometry("400x500")
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="Nome:").pack(pady=2)
        self.ent_nome = tk.Entry(self)
        self.ent_nome.pack(pady=2, fill="x", padx=10)

        tk.Label(self, text="CNPJ:").pack(pady=2)
        self.ent_cnpj = tk.Entry(self)
        self.ent_cnpj.pack(pady=2, fill="x", padx=10)
        self.ent_cnpj.bind("<KeyRelease>", self.formatar_cnpj)

        tk.Label(self, text="Telefone:").pack(pady=2)
        self.ent_telefone = tk.Entry(self)
        self.ent_telefone.pack(pady=2, fill="x", padx=10)
        self.ent_telefone.bind("<KeyRelease>", self.formatar_telefone)

        tk.Label(self, text="E-mail:").pack(pady=2)
        self.ent_email = tk.Entry(self)
        self.ent_email.pack(pady=2, fill="x", padx=10)

        tk.Label(self, text="Endereço:").pack(pady=2)
        self.ent_endereco = tk.Entry(self)
        self.ent_endereco.pack(pady=2, fill="x", padx=10)

        tk.Button(self, text="Salvar Fornecedor", bg="#800080", fg="white", command=self.salvar_fornecedor).pack(pady=15)

    def formatar_cnpj(self, event):
        cnpj = self.ent_cnpj.get()
        cnpj = re.sub(r'\D', '', cnpj)
        if len(cnpj) > 14:
            cnpj = cnpj[:14]
        if len(cnpj) > 12:
            cnpj = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
        elif len(cnpj) > 8:
            cnpj = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:]}"
        elif len(cnpj) > 5:
            cnpj = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:]}"
        elif len(cnpj) > 2:
            cnpj = f"{cnpj[:2]}.{cnpj[2:]}"
        self.ent_cnpj.delete(0, tk.END)
        self.ent_cnpj.insert(0, cnpj)

    def formatar_telefone(self, event):
        
        telefone = self.ent_telefone.get()
        telefone = ''.join(filter(str.isdigit, telefone))
        if len(telefone) > 11:
            telefone = telefone[:11]
        if len(telefone) > 10:
            telefone = f"({telefone[:2]}) {telefone[2:7]}-{telefone[7:]}"
        elif len(telefone) > 6:
            telefone = f"({telefone[:2]}) {telefone[2:6]}-{telefone[6:]}"
        elif len(telefone) > 2:
            telefone = f"({telefone[:2]}) {telefone[2:]}"
        self.ent_telefone.delete(0, tk.END)
        self.ent_telefone.insert(0, telefone)


    def validar_email(self, email):
        """Valida se o e-mail inserido é válido"""
        padrao = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
        return re.match(padrao, email)

    def salvar_fornecedor(self):
        nome = self.ent_nome.get().strip()
        cnpj = self.ent_cnpj.get().strip()
        telefone = self.ent_telefone.get().strip()
        email = self.ent_email.get().strip()
        endereco = self.ent_endereco.get().strip()

        if not nome or not cnpj or not telefone or not email or not endereco:
            messagebox.showerror("Erro", "Todos os campos são obrigatórios!")
            return

        if len(re.sub(r"\D", "", cnpj)) != 14:
            messagebox.showerror("Erro", "CNPJ inválido!")
            return

        if len(re.sub(r"\D", "", telefone)) not in [10, 11]:
            messagebox.showerror("Erro", "Telefone inválido!")
            return

        if not self.validar_email(email):
            messagebox.showerror("Erro", "E-mail inválido!")
            return

        novo_fornecedor = {
            "nome": nome,
            "cnpj": cnpj,
            "telefone": telefone,
            "email": email,
            "endereco": endereco
        }
        data_manager.adicionar_fornecedor(novo_fornecedor)
        messagebox.showinfo("Sucesso", "Fornecedor cadastrado com sucesso!")
        self.destroy()


# --- Tela de Visualização do Estoque ---
class VisualizarEstoque(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Visualizar Estoque")
        self.geometry("800x500")
        self.create_widgets()
        self.carregar_itens()

    def create_widgets(self):
        filtro_frame = tk.Frame(self)
        filtro_frame.pack(pady=5, fill="x")

        tk.Label(filtro_frame, text="Fornecedor:").pack(side="left", padx=5)
        self.filtro_fornecedor = tk.Entry(filtro_frame)
        self.filtro_fornecedor.pack(side="left", padx=5)

        tk.Label(filtro_frame, text="Nome do Item:").pack(side="left", padx=5)
        self.filtro_item = tk.Entry(filtro_frame)
        self.filtro_item.pack(side="left", padx=5)

        tk.Button(filtro_frame, text="Filtrar", bg="#800080", fg="white", command=self.filtrar).pack(side="left", padx=5)
        tk.Button(filtro_frame, text="Exportar para Excel", bg="#800080", fg="white", command=self.exportar_excel).pack(side="left", padx=5)
        
          


        # Colunas: removemos a coluna "Status"
        colunas = ("id", "Nome", "Quantidade", "Fornecedor", "Departamento", "Data")
        self.tree = ttk.Treeview(self, columns=colunas, show="headings")
        for col in colunas:
            self.tree.heading(col, text=col.capitalize())
            if col == "nome":
                self.tree.column(col, width=150)
            else:
                self.tree.column(col, width=100)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        # Vincula o clique com o botão direito para exibir o menu contextual
        self.tree.bind("<Button-3>", self.mostrar_menu_contextual)

    def carregar_itens(self, itens=None):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if itens is None:
            itens = data_manager.itens
        for item in itens:
            self.tree.insert("", "end", values=(
                item.get("id"),
                item.get("nome"),
                item.get("quantidade"),
                item.get("fornecedor"),
                item.get("departamento", ""),
                item.get("data_cadastro")
            ))

    def filtrar(self):
        fornecedor = self.filtro_fornecedor.get().strip().lower()
        nome_item = self.filtro_item.get().strip().lower()
        itens_filtrados = []
        for item in data_manager.itens:
            if fornecedor and fornecedor not in item.get("fornecedor", "").lower():
                continue
            if nome_item and nome_item not in item.get("nome", "").lower():
                continue
            itens_filtrados.append(item)
        self.carregar_itens(itens_filtrados)



    def exportar_excel(self, event=None):
        # Criar diretório 'relatorios' se não existir
        pasta_relatorios = "relatorios"
        if not os.path.exists(pasta_relatorios):
            os.makedirs(pasta_relatorios)

        # Criar subpasta com a data atual
        data_atual = datetime.datetime.now().strftime("%Y-%m-%d")
        pasta_data = os.path.join(pasta_relatorios, data_atual)
        if not os.path.exists(pasta_data):
            os.makedirs(pasta_data)

        # Nome do arquivo Excel
        file_path = os.path.join(pasta_data, f"relatorio_estoque_{data_atual}.xlsx")

        # Criando planilha
        wb = Workbook()
        ws = wb.active
        ws.title = "Estoque"

        # Definir cabeçalhos para a aba de estoque
        # Nova ordem de colunas: Departamento é inserido antes de "Data"
        colunas = ["ID", "Nome", "Serial", "Status", "Valor Unitário", "Departamento", "Data", "Quantidade", "Fornecedor", "Valor Total"]
        ws.append(colunas)

        # Estilo dos cabeçalhos (fonte branca sobre fundo roxo)
        for col in range(1, len(colunas) + 1):
            ws.cell(row=1, column=col).font = Font(bold=True, color="FFFFFF")
            ws.cell(row=1, column=col).fill = PatternFill(start_color="800080", end_color="800080", fill_type="solid")

        try:
            # Iterar sobre todos os itens do banco de dados
            for item in data_manager.itens:
                id_ = item.get("id", "")
                nome = item.get("nome", "")
                # Se a chave "fornecedores" existir, use-a; caso contrário, use "fornecedor"
                fornecedor = item.get("fornecedores", "") if item.get("fornecedores") else item.get("fornecedor", "")
                data_cadastro = item.get("data_cadastro", "")
                valor_unitario = float(item.get("valor_unitario", 0))
                quantidade_total = int(item.get("quantidade", 0))
                serials = item.get("serials", [])
                departamento = item.get("departamento", "")  # Novo campo

                serial_linhas = []
                # Caso especial: se o único serial for "0"
                if len(serials) == 1 and serials[0].get("serial") == "0":
                    total_quantidade = quantidade_total
                    total_valor = valor_unitario * total_quantidade
                    # Cria uma única linha com o serial "0"
                    serial_linhas.append([
                        id_,
                        nome,
                        "0",
                        "",
                        valor_unitario,
                        departamento,
                        data_cadastro,
                        total_quantidade,
                        fornecedor,
                        total_valor
                    ])
                else:
                    # Caso normal: cada serial representa um item único
                    total_quantidade = 0
                    total_valor = 0
                    for i, serial in enumerate(serials):
                        serial_val = serial.get("serial", "")
                        status = serial.get("status", "")
                        total_quantidade += 1  # Contabilizando um item por serial
                        total_valor += valor_unitario  # Somando o valor unitário de cada serial
                        # Na primeira linha, preencher os totais; nas demais, deixar em branco
                        if i == 0:
                            serial_linhas.append([
                                id_,
                                nome,
                                serial_val,
                                status,
                                valor_unitario,
                                departamento,
                                data_cadastro,
                                total_quantidade,   # Quantidade total exibida apenas na primeira linha
                                fornecedor,
                                total_valor         # Valor total exibido apenas na primeira linha
                            ])
                        else:
                            serial_linhas.append([
                                id_,
                                nome,
                                serial_val,
                                status,
                                valor_unitario,
                                departamento,
                                data_cadastro,
                                "",   # Deixa em branco para as linhas subsequentes do mesmo item
                                fornecedor,
                                ""
                            ])
                # Adiciona as linhas referentes a esse item
                for linha in serial_linhas:
                    ws.append(linha)
                # Linha em branco para separar os itens (opcional)
                ws.append(["-" * 10] * len(colunas))

            # --- Criar nova aba para os logs ---
            ws_logs = wb.create_sheet(title="Logs")
            logs_colunas = ["Item ID", "Nome", "Serial", "Departamento", "Status Antigo", "Novo Status", "Motivo", "Responsável", "Data Alteração"]
            ws_logs.append(logs_colunas)
            for col in range(1, len(logs_colunas) + 1):
                ws_logs.cell(row=1, column=col).font = Font(bold=True, color="FFFFFF")
                ws_logs.cell(row=1, column=col).fill = PatternFill(start_color="800080", end_color="800080", fill_type="solid")
            
            # Itera sobre os logs e adiciona-os na aba de Logs
            for log in data_manager.logs:
                ws_logs.append([
                    log.get("item_id", ""),
                    log.get("nome", ""),
                    log.get("serial", ""),
                    log.get("departamento", ""),
                    log.get("status_antigo", ""),
                    log.get("novo_status", ""),
                    log.get("motivo", ""),
                    log.get("responsavel", ""),
                    log.get("data_alteracao", "")
                ])

            # Salva o arquivo XLSX
            wb.save(file_path)
            messagebox.showinfo("Exportação", f"Relatório salvo em: {file_path}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar: {e}")


    def mostrar_menu_contextual(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label="Detalhes", command=lambda: self.abrir_detalhe_serial_status(row_id))
            menu.add_command(label="+ Acrescentar Estoque", command=lambda: self.alterar_estoque(row_id, increment=True))
            menu.add_command(label="- Decrementar Estoque", command=lambda: self.alterar_estoque(row_id, increment=False))
            menu.tk_popup(event.x_root, event.y_root)
    
    def alterar_estoque(self, row_id, increment=True):
        """Exemplo de função para alterar o estoque e registrar log."""
        valores = self.tree.item(row_id, "values")
        item_id = valores[0]
        item = next((i for i in data_manager.itens if i["id"] == item_id), None)
        if not item:
            messagebox.showerror("Erro", "Item não encontrado.")
            return

        # Verifica se o item possui seriais; permite alteração em massa somente se o único serial for "0"
        serials = item.get("serials", [])
        if len(serials) > 0:
            if not (len(serials) == 1 and serials[0].get("serial") == "0"):
                messagebox.showinfo("Atenção", "Este item possui seriais cadastrados.\nPara alterar o estoque, cadastre cada unidade de forma unitária com serial.")
                return

        prompt_text = "Informe a quantidade a acrescentar:" if increment else "Informe a quantidade a decrementar:"
        resposta = simpledialog.askstring("Alterar Estoque", prompt_text)
        if resposta is None or resposta.strip() == "":
            return
        try:
            quantidade_mod = int(resposta)
        except ValueError:
            messagebox.showerror("Erro", "Quantidade inválida!")
            return
        if quantidade_mod < 0:
            messagebox.showerror("Erro", "A quantidade deve ser um valor positivo!")
            return
        if not increment and quantidade_mod > item.get("quantidade", 0):
            messagebox.showerror("Erro", "Quantidade a decrementar excede o estoque atual!")
            return

        estoque_anterior = item.get("quantidade", 0)
        if increment:
            item["quantidade"] = estoque_anterior + quantidade_mod
        else:
            item["quantidade"] = estoque_anterior - quantidade_mod

        # Criação do log com os dados necessários, incluindo departamento
        novo_log = {
            "item_id": item.get("id"),
            "nome": item.get("nome"),
            "serial": "N/A",  # Operação em massa não envolve um serial específico
            "departamento": item.get("departamento", ""),
            # Aqui, em vez de informar o estoque anterior, mostramos o valor modificado:
            "status_antigo": "",  
            "novo_status": quantidade_mod if increment else -quantidade_mod,
            "motivo": "Incremento de estoque" if increment else "Decremento de estoque",
            "responsavel": "Usuário",  # Você pode solicitar essa informação se desejar
            "data_alteracao": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        data_manager.logs.append(novo_log)
        data_manager.salvar_tudo()
        self.carregar_itens()  # Atualiza a visualização

    def abrir_detalhe_serial_status(self, row_id):
        valores = self.tree.item(row_id, "values")
        item_id = valores[0]
        item = next((i for i in data_manager.itens if i["id"] == item_id), None)
        if item:
            DetalharItemStatus(self, item)
        else:
            messagebox.showerror("Erro", "Item não encontrado.")

# --- Janela para Detalhar Seriais e Status de um Item ---
import tkinter as tk
from tkinter import ttk, messagebox

class DetalharItemStatus(tk.Toplevel):
    def __init__(self, master, item):
        super().__init__(master)
        self.title(f"Serial e Status - {item['nome']} (ID: {item['id']})")
        self.item = item
        self.serials_originais = item.get("serials", [])  # Lista original de seriais
        self.serials_filtrados = self.serials_originais.copy()  # Cópia para filtragem
        self.create_widgets()

    def create_widgets(self):
        # Frame para a barra de busca
        search_frame = tk.Frame(self)
        search_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(search_frame, text="Buscar Serial:").pack(side="left", padx=(0, 5))

        self.search_entry = tk.Entry(search_frame)
        self.search_entry.pack(side="left", fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", self.filtrar_serials)  # Filtra automaticamente

        # Tabela de seriais e status
        colunas = ("serial", "status")
        self.tree = ttk.Treeview(self, columns=colunas, show="headings")
        self.tree.heading("serial", text="Serial")
        self.tree.heading("status", text="Status")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        self.carregar_serials()

        # Tabela para contagem de status
        self.status_table = ttk.Treeview(self, columns=("Status", "Quantidade"), show="headings")
        self.status_table.heading("Status", text="Status")
        self.status_table.heading("Quantidade", text="Quantidade")
        self.status_table.pack(fill="both", expand=True, padx=10, pady=10)

        # Label para o valor total
        self.valor_total_label = tk.Label(self, text="Valor Total: R$ 0,00")
        self.valor_total_label.pack(pady=10)

        self.carregar_contagem_status()

        # Botão para alterar status
        tk.Button(self, text="Alterar Status", bg="#800080", fg="white", command=self.abrir_alterar_status).pack(pady=5)

    def carregar_serials(self):
        """Carrega os seriais na tabela, exibindo apenas os filtrados."""
        self.tree.delete(*self.tree.get_children())  # Limpa a tabela
        for s in self.serials_filtrados:
            self.tree.insert("", "end", values=(s["serial"], s["status"]))

    def filtrar_serials(self, event=None):
        """Filtra os seriais conforme o texto digitado."""
        busca = self.search_entry.get().strip().lower()  # Remove espaços extras e deixa minúsculo

        if not busca:
            self.serials_filtrados = self.serials_originais.copy()  # Se a busca estiver vazia, mostra tudo
        else:
            self.serials_filtrados = [s for s in self.serials_originais if busca in s["serial"].lower()]

        self.carregar_serials()  # Atualiza a tabela com os resultados filtrados

    def carregar_contagem_status(self):
        """Conta quantos itens estão em cada status e calcula o valor total."""
        self.status_table.delete(*self.status_table.get_children())  # Limpa a tabela

        status_count = {"Em Uso": 0, "Manutenção": 0, "Descartado": 0, "Inativo": 0}
        total_quantidade = self.item.get("quantidade", 0)
        valor_unitario_item = self.item.get("valor_unitario", 0)

        for s in self.serials_originais:
            status_count[s["status"]] += 1

        valor_total = total_quantidade * valor_unitario_item

        for status, count in status_count.items():
            self.status_table.insert("", "end", values=(status, count))

        self.status_table.insert("", "end", values=("Total", total_quantidade))

        self.valor_total_label.config(text=f"Valor Total: R$ {valor_total:,.2f}")

    def abrir_alterar_status(self):
        """Abre a tela para alterar o status do serial selecionado."""
        selecionado = self.tree.focus()
        if not selecionado:
            messagebox.showerror("Erro", "Selecione um serial para alterar o status.")
            return
        valores = self.tree.item(selecionado, "values")
        serial_atual = valores[0]
        AlterarSerialStatus(self, self.item, serial_atual, self.atualizar_status)

    def atualizar_status(self):
        """Atualiza os dados após alteração de status."""
        self.serials_originais = self.item.get("serials", [])  # Atualiza os seriais originais
        self.filtrar_serials()  # Aplica a filtragem novamente
        self.carregar_contagem_status()



# --- Janela para Alterar o Status de um Serial ---
class AlterarSerialStatus(tk.Toplevel):
    def __init__(self, master, item, serial_atual, callback):
        super().__init__(master)
        self.title(f"Alterar Status - Serial: {serial_atual}")
        self.item = item
        self.serial_atual = serial_atual
        self.callback = callback
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text=f"Serial: {self.serial_atual}").pack(pady=5)
        tk.Label(self, text="Novo Status:").pack(pady=5)
        self.status_var = tk.StringVar(value="Em Uso")
        opcoes_status = ["Em Uso", "Manutenção", "Descartado", "Inativo"]
        self.cmb_status = ttk.Combobox(self, textvariable=self.status_var, values=opcoes_status, state="readonly")
        self.cmb_status.pack(pady=5)

        tk.Label(self, text="Motivo:").pack(pady=5)
        self.ent_motivo = tk.Entry(self)
        self.ent_motivo.pack(pady=5, fill="x", padx=10)

        tk.Label(self, text="Responsável:").pack(pady=5)
        self.ent_responsavel = tk.Entry(self)
        self.ent_responsavel.pack(pady=5, fill="x", padx=10)

        tk.Button(self, text="Salvar", bg="#800080", fg="white", command=self.salvar).pack(pady=10)

    def salvar(self):
        novo_status = self.status_var.get()
        motivo = self.ent_motivo.get().strip()
        responsavel = self.ent_responsavel.get().strip()
        if not motivo or not responsavel:
            messagebox.showerror("Erro", "Motivo e Responsável são obrigatórios!")
            return

        # Procura o serial no item
        for s in self.item.get("serials", []):
            if s.get("serial") == self.serial_atual:
                status_antigo = s.get("status")
                s["status"] = novo_status

                # Cria o log para histórico
                log = {
                    "item_id": self.item.get("id"),
                    "nome": self.item.get("nome"),
                    "serial": self.serial_atual,
                    "departamento": self.item.get("departamento", ""),
                    "status_antigo": status_antigo,
                    "novo_status": novo_status,
                    "motivo": motivo,
                    "responsavel": responsavel,
                    "data_alteracao": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }

                # Adiciona o log à lista de logs e salva tudo
                # Se você possui um método 'adicionar_log', você pode usá-lo.
                # Por exemplo:
                # data_manager.adicionar_log(log)
                # Caso contrário, podemos fazer:
                data_manager.logs.append(log)
                data_manager.salvar_tudo()

                messagebox.showinfo("Sucesso", "Status alterado com sucesso!")
                self.callback()  # Atualiza a tela que chamou essa janela
                self.destroy()
                return

        messagebox.showerror("Erro", "Serial não encontrado!")

# --- Execução da Aplicação ---
if __name__ == "__main__":
    app = MainMenu()
    agendar_backup_automatico(app)
    app.mainloop()
    
