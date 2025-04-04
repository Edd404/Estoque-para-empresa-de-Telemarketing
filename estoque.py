import os
import json
import ctypes
import sys
import re
import datetime
import tkinter as tk
import customtkinter as ctk
import shutil
import sqlite3
try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None
from customtkinter import CTkButton, CTkImage
from tkinter import ttk, messagebox, simpledialog, filedialog, Canvas, PhotoImage, Label
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageTk, ImageDraw, ImageSequence

def inicializar_banco(db_path="db/estoque.db"):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS itens (
        id TEXT PRIMARY KEY,
        nome TEXT NOT NULL,
        fornecedor TEXT,
        tamanho TEXT,
        valor_unitario REAL,
        quantidade INTEGER,
        data_cadastro TEXT
    );
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS fornecedores (
        nome TEXT PRIMARY KEY,
        cnpj TEXT,
        telefone TEXT,
        email TEXT,
        endereco TEXT
    );
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS seriais (
        item_id TEXT,
        serial TEXT,
        status TEXT,
        PRIMARY KEY (item_id, serial),
        FOREIGN KEY (item_id) REFERENCES itens(id)
    );
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS logs (
        item_id TEXT,
        nome TEXT,
        serial TEXT,
        departamento TEXT,
        status_antigo TEXT,
        novo_status TEXT,
        motivo TEXT,
        responsavel TEXT,
        data_alteracao TEXT
    );
    """)
    
    # Criar tabela de notas fiscais
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS notas_fiscais (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        responsavel TEXT NOT NULL,
        filial TEXT NOT NULL,
        fornecedor TEXT,
        data_entrada TEXT,
        data_emissao TEXT,
        serie TEXT,
        numero TEXT,
        chave TEXT UNIQUE,
        movimentacao TEXT,
        status TEXT,
        estoque TEXT,
        observacoes TEXT
    )
    """)
    
    

    # Criar tabela de itens da nota
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS itens_nota (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nota_id INTEGER,
        produto TEXT,
        quantidade INTEGER,
        valor_unitario REAL,
        FOREIGN KEY (nota_id) REFERENCES notas_fiscais(id)
    )
    """)
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS solicitacoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data_solicitacao TEXT,
        id_solicitacao TEXT UNIQUE,
        filial TEXT,
        quantidade INTEGER,
        valor_total REAL,
        status TEXT,
        observacoes TEXT
    );

    """)
    
    

    conn.commit()
    conn.close()
    
    

# Arquivos de armazenamento

BACKUP_FOLDER = "backups"

# Cria pasta de backup se não existir
if not os.path.exists(BACKUP_FOLDER):
    os.makedirs(BACKUP_FOLDER)

# Função para formatar valor em formato brasileiro (ex.: R$1.234,56)
def formatar_valor(valor_float):
    valor_formatado = f"{valor_float:,.2f}"  # ex.: 1,234.56
    valor_formatado = valor_formatado.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R${valor_formatado}"


import os
import sys
import ctypes

def load_font(font_path):
    """
    Tenta carregar a fonte a partir de um arquivo TTF para o processo atual.
    Em sistemas Windows, usa AddFontResourceExW.
    """
    if sys.platform.startswith("win"):
        FR_PRIVATE = 0x10  # Adiciona fonte apenas para o processo atual
        abs_path = os.path.abspath(font_path)
        num_fonts = ctypes.windll.gdi32.AddFontResourceExW(abs_path, FR_PRIVATE, 0)
        if num_fonts == 0:
            print(f"Falha ao carregar a fonte: {abs_path}")
        else:
            print(f"Fonte carregada com sucesso: {abs_path}")
    else:
        print(f"Por favor, instale a fonte manualmente: {font_path}")

class SplashScreen(tk.Toplevel):
    def __init__(self, parent, gif_path, duration=3000, delay=30):
        # parent: janela principal que será exibida após a splash
        super().__init__(parent)
        self.parent = parent
        self.overrideredirect(True)  # Remove bordas e barra de título
        self.duration = duration      # Duração da splash em milissegundos
        self.delay = delay            # Tempo de delay entre os frames (em ms)

        # Tamanho fixo da splash: 300x300, centralizada na tela
        splash_width = 400
        splash_height = 400
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (splash_width // 2)
        y = (screen_height // 2) - (splash_height // 2)
        self.geometry(f"{splash_width}x{splash_height}+{x}+{y}")
        self.configure(bg="#E6E6FA")
        
        # Cria um frame principal com borda suave para efeito 3D
        self.main_frame = tk.Frame(self, bg="#E6E6FA", bd=4, relief="raised")
        self.main_frame.pack(expand=True, fill="both")

        # Carrega os frames do GIF usando ImageSequence
        self.frames = []
        try:
            im = Image.open(gif_path)
            for frame in ImageSequence.Iterator(im):
                frame = frame.copy().convert("RGBA")
                # Redimensiona o frame para preencher a área (opcional)
                frame = frame.resize((300, 300), Image.Resampling.LANCZOS)
                self.frames.append(ImageTk.PhotoImage(frame))
        except Exception as e:
            print(f"Erro ao carregar o GIF: {e}")

         # Label para exibir o GIF (dentro do main_frame)
        self.image_label = tk.Label(self.main_frame, bg="#E6E6FA")
        self.image_label.pack(expand=True, fill="both")
        self.idx = 0
        self.animate_gif()
        
        bottom_frame = tk.Frame(self.main_frame, bg="#E6E6FA")
        bottom_frame.pack(side="bottom", pady=5)

        # Label de status (mensagem durante o carregamento)
        self.status_label = tk.Label(bottom_frame, text="Carregando dados...", bg="#E6E6FA", fg="black", font=("Helvetica", 10))
        self.status_label.pack(side="top", pady=2)
        
        # Dentro do __init__ da SplashScreen, antes de criar a progressbar:
        style = ttk.Style(self)
        style.theme_use('clam')  # ou outro tema que permita customização
        style.configure("purple.Horizontal.TProgressbar",
                        troughcolor="#E6E6FA",   # fundo do progress bar
                        bordercolor="#E6E6FA",
                        background="#800080",    # cor do preenchimento (roxo)
                        lightcolor="#800080",
                        darkcolor="#800080")


        # Barra de progresso (menos padding para posicionar mais para cima)
        self.progress = ttk.Progressbar(bottom_frame, mode="indeterminate", length=200, style="purple.Horizontal.TProgressbar")

        self.progress.pack(side="top", pady=2)
        self.progress.start(10)

        # Após o tempo definido, fecha a splash
        self.after(self.duration, self.close_splash)

    def animate_gif(self):
        # Atualiza a imagem exibida para o próximo frame
        if self.frames:
            self.image_label.config(image=self.frames[self.idx])
            self.idx = (self.idx + 1) % len(self.frames)
        self.after(self.delay, self.animate_gif)

    def close_splash(self):
        self.progress.stop()
        self.destroy()
        self.parent.deiconify()  # Exibe a janela principal


# Função para gerar o próximo ID sequencial (formato 001, 002, etc.)
def gerar_proximo_id(self):
    self.cursor.execute("SELECT MAX(CAST(id as integer)) as max_id FROM itens")
    row = self.cursor.fetchone()
    max_id = row["max_id"] if row["max_id"] is not None else 0
    novo_id = int(max_id) + 1
    return f"{novo_id:03d}"


# --- DataManager: Gerencia a leitura e escrita dos dados em JSON ---
class DataManagerSQL:
    def __init__(self, db_path="db/estoque.db"):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row  # Permite acessar as colunas pelo nome
        self.cursor = self.conn.cursor()

    def fechar(self):
        self.conn.close()
    
    def obter_notas_fiscais(self):
        try:
            self.cursor.execute("SELECT * FROM notas_fiscais")  # Ajuste conforme o nome real da sua tabela
            colunas = [desc[0] for desc in self.cursor.description]  # Obtém os nomes das colunas
            return [dict(zip(colunas, row)) for row in self.cursor.fetchall()]  # Retorna os dados como dicionário
        except Exception as e:
            print(f"Erro ao obter notas fiscais: {e}")
            return []
    
    def obter_itens_nota(self, nota_id):
        self.cursor.execute("SELECT * FROM itens_nota WHERE nota_id = ?", (nota_id,))
        return [dict(row) for row in self.cursor.fetchall()]

    # Métodos para carregar os dados (úteis para alimentar a interface)
    def obter_itens(self):
        self.cursor.execute("SELECT * FROM itens")
        rows = self.cursor.fetchall()
        itens = []
        for row in rows:
            item = dict(row)
            self.cursor.execute("SELECT serial, status FROM seriais WHERE item_id = ?", (item["id"],))
            item["serials"] = [dict(s) for s in self.cursor.fetchall()]
            itens.append(item)
        return itens

    def obter_fornecedores(self):
        self.cursor.execute("SELECT * FROM fornecedores")
        return [dict(row) for row in self.cursor.fetchall()]

    def obter_logs(self):
        self.cursor.execute("SELECT * FROM logs")
        return [dict(row) for row in self.cursor.fetchall()]

    def gerar_proximo_id(self):
        self.cursor.execute("SELECT MAX(CAST(id as integer)) as max_id FROM itens")
        row = self.cursor.fetchone()
        max_id = row["max_id"] if row["max_id"] is not None else 0
        novo_id = int(max_id) + 1
        return f"{novo_id:03d}"

    # Método para adicionar (ou atualizar) um item
    def adicionar_item(self, novo_item):
        # Verifica se já existe item com mesmo nome (ignorando caixa)
        self.cursor.execute("SELECT * FROM itens WHERE lower(nome) = lower(?)", (novo_item["nome"],))
        row = self.cursor.fetchone()
        if row:
            messagebox.showerror("Erro", "Item já cadastrado! Utilize a movimentação para alterar a quantidade.")
            return False
        else:
            novo_id = novo_item.get("id") or self.gerar_proximo_id()
            # Força quantidade para 0
            self.cursor.execute("""
                INSERT INTO itens (id, nome, fornecedor, tamanho, valor_unitario, quantidade, data_cadastro)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (novo_id, novo_item["nome"], novo_item.get("fornecedor"), novo_item.get("tamanho"),
                novo_item.get("valor_unitario"), novo_item.get("quantidade"), novo_item.get("data_cadastro")))

        self.conn.commit()
        return True


    # Método para adicionar fornecedor
    def adicionar_fornecedor(self, fornecedor):
        self.cursor.execute("""
            INSERT OR REPLACE INTO fornecedores (nome, cnpj, telefone, email, endereco)
            VALUES (?, ?, ?, ?, ?)
        """, (fornecedor["nome"], fornecedor["cnpj"], fornecedor["telefone"], fornecedor["email"], fornecedor["endereco"]))
        self.conn.commit()

    # Método para adicionar log
    def adicionar_log(self, log):
        self.cursor.execute("""
            INSERT INTO logs (item_id, nome, serial, departamento, status_antigo, novo_status, motivo, responsavel, data_alteracao)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            log.get("item_id"), log.get("nome"), log.get("serial"), log.get("departamento"),
            log.get("status_antigo"), log.get("novo_status"), log.get("motivo"),
            log.get("responsavel"), log.get("data_alteracao")
        ))
        self.conn.commit()
    
    def obter_solicitacoes(self):
        try:
            self.cursor.execute("SELECT * FROM solicitacoes")
            colunas = [desc[0] for desc in self.cursor.description]
            return [dict(zip(colunas, row)) for row in self.cursor.fetchall()]
        except Exception as e:
            print(f"Erro ao obter solicitações: {e}")
            return []
    
    def gerar_proximo_id_solicitacao(self, modo):
        prefix = "ENV" if modo == "Solicitacao" else "DEV"
        self.cursor.execute(
            "SELECT id_solicitacao FROM solicitacoes WHERE id_solicitacao LIKE ? ORDER BY id_solicitacao DESC LIMIT 1",
            (prefix + "%",)
        )
        row = self.cursor.fetchone()
        if row:
            last_id = row["id_solicitacao"]
            number = int(last_id[len(prefix):])
            new_number = number + 1
        else:
            new_number = 1
        return f"{prefix}{new_number:04d}"

    def adicionar_solicitacao(self, solicitacao):
        try:
            self.cursor.execute("""
                INSERT INTO solicitacoes (data_solicitacao, id_solicitacao, filial, quantidade, valor_total, status, observacoes)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                solicitacao["data_solicitacao"],
                solicitacao["id_solicitacao"],
                solicitacao["filial"],
                solicitacao["quantidade"],
                solicitacao["valor_total"],
                solicitacao["status"],
                solicitacao["observacoes"]
            ))
            self.conn.commit()
            return True
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar solicitação: {e}")
            return False
    
    def obter_filiais(self):
        self.cursor.execute("SELECT DISTINCT filial FROM notas_fiscais")
        return [row["filial"] for row in self.cursor.fetchall()]
    
    def obter_itens_por_filial(self, filial):
        self.cursor.execute("""
            SELECT i_n.*, nf.filial 
            FROM itens_nota i_n 
            JOIN notas_fiscais nf ON i_n.nota_id = nf.id 
            WHERE nf.filial = ? AND nf.estoque = 'Sim'
        """, (filial,))
        return [dict(row) for row in self.cursor.fetchall()]



# Instância global do DataManager
data_manager = DataManagerSQL()


# --- Função para Backup Manual e Automático ---
def realizar_backup():
    # Criar estrutura de pastas: Dia e Horário
    agora = datetime.datetime.now()
    data_backup = agora.strftime("%d-%m-%Y")
    hora_backup = agora.strftime("%H-%M")
    
    backup_pasta = os.path.join(BACKUP_FOLDER, data_backup, hora_backup)

    if not os.path.exists(backup_pasta):
        os.makedirs(backup_pasta)

    # Obter os dados do banco SQLite usando os métodos do DataManagerSQL
    arquivos_db = {
        "itens": data_manager.obter_itens(),
        "fornecedores": data_manager.obter_fornecedores(),
        "logs": data_manager.obter_logs()
    }

    # Salvar cada JSON separadamente
    for nome_arquivo, dados in arquivos_db.items():
        caminho_arquivo = os.path.join(backup_pasta, f"{nome_arquivo}.db")
        try:
            with open(caminho_arquivo, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Erro ao salvar {nome_arquivo}.db: {e}")

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

            if os.path.isfile(origem):
                try:
                    shutil.copy2(origem, destino)
                except PermissionError:
                    print(f"⚠ Permissão negada ao copiar: {origem}")

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
        self.geometry("1920x1080")
        self.configure(bg="#f0f0f0")
        
       # self.criar_fundo()
        self.create_widgets()
        
        
    
   # def criar_fundo(self):
   #     """Carrega e exibe a imagem de fundo com opacidade."""
    #    try:
     #       # Carregar e redimensionar a imagem para caber na tela
      #      imagem = Image.open("bg.png").convert("RGBA")
       #     imagem = imagem.resize((2000, 960), Image.Resampling.LANCZOS)  # Ajusta ao tamanho da tela
        #    
         #   # Reduzir opacidade da imagem (0.3 = 30% visibilidade)
          #  imagem_translucida = reduzir_opacidade("bg.png", 0.8)
           # 
            #self.fundo_img = ImageTk.PhotoImage(imagem_translucida)

            # Criar Label para exibir a imagem
           # self.fundo_label = Label(self, image=self.fundo_img)
           # self.fundo_label.place(x=-30, y=100)  # Cobrir a tela inteira
           # self.fundo_label.lower()  # Garante que fique atrás de tudo

       # except Exception as e:
        #    print(f"Erro ao carregar a imagem de fundo: {e}")
        
        
        
    def create_widgets(self):
    # Cria a barra superior (mantido conforme seu código original)
        self.create_top_bar()
        
        # Cria um frame central com fundo roxo para os botões
        content_frame = tk.Frame(self, bg="#f0f0f0")
        content_frame.pack(expand=True)
        
        # Lista de botões com informações: texto, comando e caminho da imagem
        buttons_info = [
            {"text": "Cadastrar Item", "command": self.abrir_cadastro_item, "img": "Imagens/botoes/cadastro_item.png"},
            {"text": "Cadastrar Fornecedor", "command": self.abrir_cadastro_fornecedor, "img": "Imagens/botoes/cadastro_fornecedor.png"},
            {"text": "Solicitações", "command": self.abrir_solicitacoes, "img": "Imagens/botoes/solicitacoes.png"},  # novo botão
            {"text": "Estoque", "command": self.abrir_visualizar_estoque, "img": "Imagens/botoes/visualizar_estoque.png"},
            {"text": "Backup", "command": self.realizar_backup, "img": "Imagens/botoes/backup.png"},
            {"text": "Movimentação", "command": self.abrir_log_movimentacao, "img": "Imagens/botoes/log.png"}
        ]
        
        
        # Tamanho padrão dos botões
        button_width = 150
        button_height = 150
        # Lista para manter referência das imagens (evita que sejam coletadas pelo garbage collector)
        self.button_images = []
        
        # Definimos que serão 3 botões por linha
        columns = 3
        for index, btn in enumerate(buttons_info):
            row = index // columns
            col = index % columns
            
            # Tenta carregar a imagem do botão; se falhar, usa None
            try:
                img = Image.open(btn["img"]).convert("RGBA")
                # Redimensiona a imagem para um tamanho adequado (ex.: 120x120 pixels)
                img = img.resize((120, 120), Image.Resampling.LANCZOS)
                button_img = CTkImage(light_image=img, dark_image=img, size=(90, 90))
            except Exception as e:
                print(f"Erro ao carregar imagem {btn['img']}: {e}")
                button_img = None
            
            self.button_images.append(button_img)
            
            # Cria o botão com a imagem (acima) e o texto (abaixo)
            ctk.set_appearance_mode("light")  # ou "dark" se preferir
            ctk.set_default_color_theme("blue")  # ou crie um tema personalizado

                # Criar o botão com customtkinter
            button = CTkButton(content_frame,
                text=btn["text"],
                image=button_img,
                compound="top",
                command=btn["command"],
                font=("Helvetica", 13, "bold"),
                fg_color="#7617bb",
                text_color="white",
                hover_color="#a084c9",
                corner_radius=12,
                width=button_width,
                height=button_height,
                
                anchor="bottom"  
            )
            button.grid(row=row, column=col, padx=20, pady=20)
            
           
           
        
        # Cria o rodapé (mantido conforme seu código original)
        self.create_footer()

    def create_top_bar(self):
        """Cria uma barra superior ondulada moderna"""
        canvas = Canvas(self, height=30, bg="#7617bb", highlightthickness=0)
        canvas.pack(fill="x")

        # Criando a onda superior
        canvas.create_line(0, 80, 150, 60, 300, 100, 450, 60, 600, 80, smooth=True, fill="#7617bb", width=30)

        # Adicionando o título na barra
        label_title = tk.Label(self, text="Gestão de Uniformes", font=("Helvetica", 12, "bold"), fg="white", bg="#7617bb")
        label_title.place(x=100, y=3)
    
    
        
        try:
            # Reduzir opacidade da imagem (0.3 = 30% de visibilidade)
            imagem_translucida = reduzir_opacidade("imagens/img.png", 1)
            self.img = ImageTk.PhotoImage(imagem_translucida)

            # Criar Label para exibir a imagem
            self.img_label = Label(self, image=self.img)
            self.img_label.place(x=1136, y=31)  # Posiciona no canto superior direito
            
            

        except Exception as e:
            print(f"Erro ao carregar a imagem: {e}")
               
    def create_footer(self):
        """Rodapé com efeito 3D estilo Neumorphism."""
        footer = tk.Frame(self, height=35, bg="#7617bb", relief="ridge", bd=3)
        footer.pack(fill="x", side="bottom", pady=5)

        label_footer = tk.Label(footer, text="MS Connect ®️  |  Versão 1.0  |  © 2025",
                            font=("Helvetica", 10, "bold"), fg="white", bg="#7617bb")
        label_footer.pack(expand=True)

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
    
    def abrir_solicitacoes(self):
        Solicitacoes(self)
        print("Abrir Solicitacoes")

class Solicitacoes(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Solicitações")
        self.state("zoomed")
        self.configure(bg="white")
        self.create_widgets()
        self.carregar_solicitacoes()

    def create_widgets(self):
        # Frame de Filtros
        filtro_frame = tk.Frame(self, bg="white")
        filtro_frame.pack(fill="x", padx=10, pady=10)
        
        tk.Label(filtro_frame, text="Filtrar:", bg="white", font=("Helvetica", 10)).pack(side="left", padx=5)
        filiais = data_manager.obter_filiais()
        if not filiais:
            filiais = ["Filial Padrão"]
        self.filtro_filial = ttk.Combobox(filtro_frame, state="readonly", values=filiais, width=20)
        self.filtro_filial.current(0)
        self.filtro_filial.pack(side="left", padx=5)
        
        # Botão para nova solicitação (usando customtkinter)
        try:
            img = Image.open("Imagens/botoes/nova_solicitacao.png").convert("RGBA")
            img = img.resize((20,20), Image.Resampling.LANCZOS)
            nova_img = CTkImage(light_image=img, dark_image=img, size=(20,20))
        except Exception as e:
            print("Erro ao carregar imagem para nova solicitação:", e)
            nova_img = None
        self.btn_nova_sol = ctk.CTkButton(filtro_frame,
            text="NOVA SOLICITAÇÃO",
            image=nova_img,
            compound="left",
            fg_color="green",
            text_color="white",
            hover_color="#88cc88",
            corner_radius=12,
            command=self.abrir_nova_solicitacao)
        self.btn_nova_sol.pack(side="right", padx=5)
        
        # Treeview para exibir as solicitações
        colunas = ("data_solicitacao", "id_solicitacao", "filial", "quantidade", "valor_total", "status", "opcoes")
        self.tree = ttk.Treeview(self, columns=colunas, show="headings")
        self.tree.heading("data_solicitacao", text="Data Solicitacao")
        self.tree.heading("id_solicitacao", text="ID Solicitacao")
        self.tree.heading("filial", text="Filial")
        self.tree.heading("quantidade", text="Quantidade")
        self.tree.heading("valor_total", text="Valor Total")
        self.tree.heading("status", text="Status Solicitacao")
        self.tree.heading("opcoes", text="Opcoes")
        
        # Ajuste das larguras das colunas
        self.tree.column("data_solicitacao", width=120)
        self.tree.column("id_solicitacao", width=140)
        self.tree.column("filial", width=100)
        self.tree.column("quantidade", width=80)
        self.tree.column("valor_total", width=120)
        self.tree.column("status", width=150)
        self.tree.column("opcoes", width=150)
        
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Vincular evento de clique para o menu de opções, se desejar
        self.tree.bind("<Button-3>", self.mostrar_menu_opcoes)

    def carregar_solicitacoes(self):
        # Limpa a treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Obter solicitações do banco (você deve criar data_manager.obter_solicitacoes())
        solicitacoes = data_manager.obter_solicitacoes()
        for sol in solicitacoes:
            # Monta cada linha. Por exemplo:
            # sol["id_solicitacao"] deve conter ENVxxxx ou DEVxxxx
            # sol["status"] poderá ser "NAO RECEBIDO", "RECEBIDO" ou "Devolucao"
            # Na coluna 'status', você pode definir um botão dinâmico; para simplificação, aqui mostramos apenas o texto
            self.tree.insert("", "end", values=(
                sol.get("data_solicitacao", ""),
                sol.get("id_solicitacao", ""),
                sol.get("filial", ""),
                sol.get("quantidade", ""),
                formatar_valor(sol.get("valor_total", 0)),
                sol.get("status", ""),
                ""  # Coluna de Opcoes – os botões serão inseridos posteriormente
            ))

    def mostrar_menu_opcoes(self, event):
        # Exemplo: Ao clicar com o botão direito, exibe um menu para alterar status
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label="Marcar como RECEBIDO", command=lambda: self.alterar_status(row_id, "RECEBIDO"))
            menu.add_command(label="Marcar como NAO RECEBIDO", command=lambda: self.alterar_status(row_id, "NAO RECEBIDO"))
            menu.tk_popup(event.x_root, event.y_root)

    def alterar_status(self, row_id, novo_status):
        # Recupera dados da linha e atualiza no banco
        valores = self.tree.item(row_id, "values")
        id_solicitacao = valores[1]
        try:
            conn = sqlite3.connect("db/estoque.db")
            cursor = conn.cursor()
            cursor.execute("UPDATE solicitacoes SET status=? WHERE id_solicitacao=?", (novo_status, id_solicitacao))
            conn.commit()
            conn.close()
            self.carregar_solicitacoes()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar status: {e}")

    def abrir_nova_solicitacao(self):
        # Cria a nova janela para cadastro de solicitação/devolução usando um Notebook
        nova_sol = tk.Toplevel(self)
        nova_sol.title("Nova Solicitação / Devolução")
        nova_sol.geometry("400x500")
        
        style = ttk.Style(nova_sol)
        style.theme_use("default")
        style.configure("TNotebook.Tab", background="#E6E6FF", foreground="black", padding=[10,5])
        style.map("TNotebook.Tab", background=[("selected", "#D8BFD8")])

        notebook = ttk.Notebook(nova_sol)
        notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # Aba Solicitação
        frame_solicitacao = tk.Frame(notebook, bg="#e6ffe6", padx=10, pady=10)
        notebook.add(frame_solicitacao, text="Solicitação")

        # Aba Devolução
        frame_devolucao = tk.Frame(notebook, bg="#ffe6e6", padx=10, pady=10)
        notebook.add(frame_devolucao, text="Devolução")
        
        
        # --- Aba Solicitação ---
        tk.Label(frame_solicitacao, text="Filial:", bg="#e6ffe6", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        filiais = data_manager.obter_filiais()
        if not filiais:
            filiais = ["Filial Padrão"]
        filial_cb = ttk.Combobox(frame_solicitacao, state="readonly", values=filiais, width=20)
        filial_cb.current(0)
        filial_cb.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        tk.Label(frame_solicitacao, text="Itens:", bg="#e6ffe6", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        itens_cb = ttk.Combobox(frame_solicitacao, state="readonly", values=[], width=20)
        itens_cb.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        tk.Label(frame_solicitacao, text="Tamanho:", bg="#e6ffe6", font=("Helvetica", 10, "bold")).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        tamanho_cb = ttk.Combobox(frame_solicitacao, state="readonly", values=["P", "M", "G"], width=10)
        tamanho_cb.current(0)
        tamanho_cb.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        tk.Label(frame_solicitacao, text="Quantidade:", bg="#e6ffe6", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", padx=5, pady=5)
        quantidade_entry = tk.Entry(frame_solicitacao, font=("Helvetica", 10, "bold"), width=10)
        quantidade_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        # Atualiza o combobox de Itens conforme a filial selecionada
        def atualizar_itens_solicitacao(event):
            filial_selecionada = filial_cb.get()
            itens_disponiveis = data_manager.obter_itens_por_filial(filial_selecionada)
            nomes = [item["produto"] for item in itens_disponiveis] if itens_disponiveis else []
            if not nomes:
                nomes = ["Nenhum item disponível"]
            itens_cb['values'] = nomes
            itens_cb.current(0)
        filial_cb.bind("<<ComboboxSelected>>", atualizar_itens_solicitacao)
        atualizar_itens_solicitacao(None)
        
        # --- Aba Devolução ---
        tk.Label(frame_devolucao, text="Filial:", bg="#ffe6e6", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        filial_dev = ttk.Combobox(frame_devolucao, state="readonly", values=filiais, width=20)
        filial_dev.current(0)
        filial_dev.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        tk.Label(frame_devolucao, text="Itens:", bg="#ffe6e6", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        itens_dev = ttk.Combobox(frame_devolucao, state="readonly", values=[], width=20)
        itens_dev.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        def atualizar_itens_devolucao(event):
            filial_selecionada = filial_dev.get()
            itens_disponiveis = data_manager.obter_itens_por_filial(filial_selecionada)
            nomes = [item["produto"] for item in itens_disponiveis] if itens_disponiveis else []
            if not nomes:
                nomes = ["Nenhum item disponível"]
            itens_dev['values'] = nomes
            itens_dev.current(0)
        filial_dev.bind("<<ComboboxSelected>>", atualizar_itens_devolucao)
        atualizar_itens_devolucao(None)
        
        tk.Label(frame_devolucao, text="Tamanho:", bg="#ffe6e6", font=("Helvetica", 10, "bold")).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        tamanho_dev = ttk.Combobox(frame_devolucao, state="readonly", values=["P", "M", "G"], width=10)
        tamanho_dev.current(0)
        tamanho_dev.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        tk.Label(frame_devolucao, text="Quantidade:", bg="#ffe6e6", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", padx=5, pady=5)
        quantidade_dev = tk.Entry(frame_devolucao, font=("Helvetica", 10, "bold"), width=10)
        quantidade_dev.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        tk.Label(frame_devolucao, text="Fornecedor:", bg="#ffe6e6", font=("Helvetica", 10, "bold")).grid(row=4, column=0, sticky="w", padx=5, pady=5)
        fornecedores = [f["nome"] for f in data_manager.obter_fornecedores()]
        if not fornecedores:
            fornecedores = ["Nenhum fornecedor cadastrado"]
        fornecedor_dev = ttk.Combobox(frame_devolucao, state="readonly", values=fornecedores, width=20)
        fornecedor_dev.current(0)
        fornecedor_dev.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        
        tk.Label(frame_devolucao, text="Observação:", bg="#ffe6e6", font=("Helvetica", 10, "bold")).grid(row=5, column=0, sticky="w", padx=5, pady=5)
        observacao_text = tk.Text(frame_devolucao, height=3, font=("Helvetica", 10, "bold"))
        observacao_text.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        
        # Botão de adicionar linha para devolução (placeholder, se necessário)
        btn_add_linha_dev = tk.Button(frame_devolucao, text="+", font=("Helvetica", 10, "bold"), width=3)
        btn_add_linha_dev.grid(row=3, column=2, padx=5, pady=5)
        
        # Botão Salvar (fora do Notebook)
        btn_salvar = tk.Button(nova_sol, text="Salvar", font=("Helvetica", 10, "bold"),
                               command=lambda: salvar_solicitacao(notebook.index(notebook.select())))
        btn_salvar.pack(side="bottom", pady=10)
        
        # Função de salvamento que verifica qual aba está ativa (0: Solicitação, 1: Devolução)
        def salvar_solicitacao(indice_aba):
            if indice_aba == 0:
                # Aba Solicitação
                filial = filial_cb.get()
                item_nome = itens_cb.get()
                tamanho = tamanho_cb.get()
                quantidade_text = quantidade_entry.get()
                if not filial or not item_nome or not tamanho or not quantidade_text:
                    messagebox.showerror("Erro", "Preencha todos os campos obrigatórios para Solicitação!")
                    return
                try:
                    quantidade_solicitada = int(quantidade_text)
                except ValueError:
                    messagebox.showerror("Erro", "Quantidade inválida!")
                    return
                # Verifica disponibilidade
                itens_disponiveis = data_manager.obter_itens_por_filial(filial)
                item_encontrado = None
                for item in itens_disponiveis:
                    if item.get("produto") == item_nome and item.get("tamanho") == tamanho:
                        item_encontrado = item
                        break
                if not item_encontrado:
                    messagebox.showerror("Erro", "Item com o tamanho selecionado não está disponível na filial!")
                    return
                if quantidade_solicitada > item_encontrado.get("quantidade", 0):
                    messagebox.showerror("Erro", "Quantidade solicitada excede a disponível!")
                    return

                data_solicitacao = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                id_solicitacao = data_manager.gerar_proximo_id_solicitacao("Solicitacao")
                valor_unitario = item_encontrado.get("valor_unitario", 0.0)
                valor_total = quantidade_solicitada * valor_unitario
                solicitacao = {
                    "data_solicitacao": data_solicitacao,
                    "id_solicitacao": id_solicitacao,
                    "filial": filial,
                    "quantidade": quantidade_solicitada,
                    "valor_total": valor_total,
                    "status": "NAO RECEBIDO",
                    "observacoes": ""
                }
                if data_manager.adicionar_solicitacao(solicitacao):
                    messagebox.showinfo("Sucesso", "Solicitação salva com sucesso!")
                else:
                    messagebox.showerror("Erro", "Erro ao salvar solicitação!")
            else:
                # Aba Devolução
                filial = filial_dev.get()
                item_nome = itens_dev.get()
                tamanho = tamanho_dev.get()
                quantidade_text = quantidade_dev.get()
                fornecedor = fornecedor_dev.get()
                observacao = observacao_text.get("1.0", "end-1c")
                if not filial or not item_nome or not tamanho or not quantidade_text or not fornecedor or not observacao:
                    messagebox.showerror("Erro", "Preencha todos os campos obrigatórios para Devolução!")
                    return
                try:
                    quantidade_solicitada = int(quantidade_text)
                except ValueError:
                    messagebox.showerror("Erro", "Quantidade inválida!")
                    return

                itens_disponiveis = data_manager.obter_itens_por_filial(filial)
                item_encontrado = None
                for item in itens_disponiveis:
                    if item.get("produto") == item_nome and item.get("tamanho") == tamanho:
                        item_encontrado = item
                        break
                if not item_encontrado:
                    messagebox.showerror("Erro", "Item com o tamanho selecionado não está disponível na filial!")
                    return
                if quantidade_solicitada > item_encontrado.get("quantidade", 0):
                    messagebox.showerror("Erro", "Quantidade solicitada excede a disponível!")
                    return

                data_solicitacao = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                id_solicitacao = data_manager.gerar_proximo_id_solicitacao("Devolucao")
                valor_unitario = item_encontrado.get("valor_unitario", 0.0)
                valor_total = quantidade_solicitada * valor_unitario
                solicitacao = {
                    "data_solicitacao": data_solicitacao,
                    "id_solicitacao": id_solicitacao,
                    "filial": filial,
                    "quantidade": quantidade_solicitada,
                    "valor_total": valor_total,
                    "status": "DEVOLUÇÃO PENDENTE",
                    "observacoes": observacao
                }
                if data_manager.adicionar_solicitacao(solicitacao):
                    messagebox.showinfo("Sucesso", "Devolução salva com sucesso!")
                else:
                    messagebox.showerror("Erro", "Erro ao salvar devolução!")
            nova_sol.destroy()
            self.carregar_solicitacoes()
  
             

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
        self.logs = data_manager.obter_logs()
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
        self.geometry("400x400")
        self.create_widgets()

    def create_widgets(self):
        # Container principal para as duas colunas
        container = tk.Frame(self)
        container.pack(padx=10, pady=10, fill="both", expand=True)

        # Frame para a linha de "Item já cadastrado" (span nas duas colunas)
        tk.Label(container, text="Item já cadastrado (opcional):").grid(row=0, column=0, columnspan=2, pady=5, sticky="w")
        self.cmb_item_existente = ttk.Combobox(container, state="readonly")
        self.atualizar_lista_itens()
        self.cmb_item_existente.current(0)
        self.cmb_item_existente.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        self.cmb_item_existente.bind("<<ComboboxSelected>>", self.item_selecionado)

        # Cria dois frames para as duas colunas
        left_frame = tk.Frame(container)
        left_frame.grid(row=2, column=0, sticky="nsew", padx=(0,10))
        right_frame = tk.Frame(container)
        right_frame.grid(row=2, column=1, sticky="nsew", padx=(10,0))
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)

        # Coluna Esquerda
        # Campo Nome
        tk.Label(left_frame, text="Nome:").grid(row=0, column=0, sticky="w", pady=5)
        self.ent_nome = tk.Entry(left_frame)
        self.ent_nome.grid(row=1, column=0, sticky="ew", pady=5)
        left_frame.grid_columnconfigure(0, weight=1)

        # Campo Valor Unitário
        tk.Label(left_frame, text="Valor Unitário:").grid(row=2, column=0, sticky="w", pady=5)
        self.ent_valor = tk.Entry(left_frame)
        self.ent_valor.grid(row=3, column=0, sticky="ew", pady=5)
        self.ent_valor.bind("<FocusOut>", self.formatar_valor_evento)

        # Campo Data
        tk.Label(left_frame, text="Data:").grid(row=4, column=0, sticky="w", pady=5)
        self.ent_data = tk.Entry(left_frame)
        data_atual = datetime.date.today().strftime("%d/%m/%Y")
        self.ent_data.insert(0, data_atual)
        self.ent_data.grid(row=5, column=0, sticky="ew", pady=5)
        self.ent_data.bind("<FocusOut>", self.formatar_data_evento)

        # Coluna Direita
        # Campo Fornecedor
        tk.Label(right_frame, text="Fornecedor:").grid(row=0, column=0, sticky="w", pady=5)
        self.fornecedor_var = tk.StringVar()
        fornecedores_lista = [f["nome"] for f in data_manager.obter_fornecedores()]
        self.cmb_fornecedor = ttk.Combobox(right_frame, textvariable=self.fornecedor_var, values=fornecedores_lista, state="readonly")
        self.cmb_fornecedor.grid(row=1, column=0, sticky="ew", pady=5)
        right_frame.grid_columnconfigure(0, weight=1)

        # Campo Tamanho (novo campo)
        tk.Label(right_frame, text="Tamanho:").grid(row=2, column=0, sticky="w", pady=5)
        self.tamanho_var = tk.StringVar()
        opcoes_tamanho = ["P", "M", "G", "GG"]
        self.cmb_tamanho = ttk.Combobox(right_frame, textvariable=self.tamanho_var, values=opcoes_tamanho, state="readonly")
        self.cmb_tamanho.current(0)
        self.cmb_tamanho.grid(row=3, column=0, sticky="ew", pady=5)

        # Botão Salvar centralizado (linha abaixo das duas colunas)
        btn_frame = tk.Frame(container)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=15)

        # Tenta carregar a imagem para o botão
        try:
            img = Image.open("imagens/botoes/salvar.png").convert("RGBA")
            img = img.resize((20, 20), Image.Resampling.LANCZOS)
            ctk_img = CTkImage(light_image=img, dark_image=img, size=(20, 20))
        except Exception as e:
            print(f"Erro ao carregar imagem: {e}")
            ctk_img = None

        botao_salvar = ctk.CTkButton(btn_frame,
            text="Salvar Item",
            image=ctk_img,
            compound="left",
            command=self.salvar_item,
            font=("Helvetica", 14, "bold"),
            fg_color="#800080",
            text_color="white",
            hover_color="#a084c9",
            corner_radius=12,
            width=80,
            height=40
        )
        botao_salvar.pack()


    def formatar_data_evento(self, event):
        data_texto = self.ent_data.get().strip()
        if data_texto:
            data_limpa = data_texto.replace("/", "")
            if len(data_limpa) == 8:
                data_formatada = f"{data_limpa[:2]}/{data_limpa[2:4]}/{data_limpa[4:]}"
                self.ent_data.delete(0, tk.END)
                self.ent_data.insert(0, data_formatada)
            else:
                messagebox.showerror("Erro", "Data deve estar no formato DD/MM/YYYY!")
                self.ent_data.delete(0, tk.END)

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

    def atualizar_lista_itens(self):
        itens = data_manager.obter_itens()
        lista = ["Novo item"] + [item["nome"] for item in itens]
        self.cmb_item_existente['values'] = lista

    def item_selecionado(self, event):
        selecionado = self.cmb_item_existente.get()
        if selecionado != "Novo item":
            item_selecionado = next((item for item in data_manager.obter_itens() if item["nome"] == selecionado), None)
            if item_selecionado:
                self.ent_nome.delete(0, tk.END)
                self.ent_nome.insert(0, item_selecionado["nome"])
                self.ent_nome.config(state="disabled")
                self.ent_valor.delete(0, tk.END)
                self.ent_valor.insert(0, formatar_valor(item_selecionado["valor_unitario"]))
                self.ent_valor.config(state="normal")
        else:
            self.ent_nome.config(state="normal")
            self.ent_nome.delete(0, tk.END)
            self.ent_valor.config(state="normal")
            self.ent_valor.delete(0, tk.END)

    def salvar_item(self):
        if self.cmb_item_existente.get() != "Novo item":
            nome = self.cmb_item_existente.get().strip()
        else:
            nome = self.ent_nome.get().strip()
        fornecedor = self.fornecedor_var.get()
        valor_texto = self.ent_valor.get().strip()
        data_cadastro = self.ent_data.get().strip()
        tamanho = self.tamanho_var.get()

        if not nome or not valor_texto or not data_cadastro or not fornecedor:
            messagebox.showerror("Erro", "Todos os campos obrigatórios devem ser preenchidos!")
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
            datetime.datetime.strptime(data_cadastro, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Erro", "Data deve estar no formato DD/MM/YYYY!")
            return

        quantidade = 0  # Sempre 0 no cadastro

        novo_item = {
            "id": data_manager.gerar_proximo_id(),
            "nome": nome,
            "fornecedor": fornecedor,
            "valor_unitario": valor_float,
            "quantidade": quantidade,
            "data_cadastro": data_cadastro,
            "tamanho": tamanho
        }

        if data_manager.adicionar_item(novo_item):
            novo_log = {
                "item_id": novo_item["id"],
                "nome": novo_item["nome"],
                "tamanho": tamanho,
                "status_antigo": "",
                "novo_status": "Novo",
                "motivo": "Item cadastrado",
                "responsavel": "Usuário",
                "data_alteracao": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            data_manager.adicionar_log(novo_log)
            messagebox.showinfo("Sucesso", "Item cadastrado com sucesso!")
            self.destroy()



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


def calcular_compra():
        notas = data_manager.obter_notas_fiscais()
        total_itens = 0
        total_valor = 0.0
        for nota in notas:
            itens_nota = data_manager.obter_itens_nota(nota["id"])
            total_itens += sum(item["quantidade"] for item in itens_nota)
            total_valor += sum(item["quantidade"] * item["valor_unitario"] for item in itens_nota)
        return total_itens, total_valor



def calcular_estoque_atual():
    # Filtra apenas as notas onde o campo "estoque" está marcado como "Sim"
    notas_disponiveis = [nota for nota in data_manager.obter_notas_fiscais() if nota.get("estoque", "Sim") == "Sim"]
    total_itens = 0
    total_valor = 0.0
    for nota in notas_disponiveis:
        itens_nota = data_manager.obter_itens_nota(nota["id"])
        total_itens += sum(item["quantidade"] for item in itens_nota)
        total_valor += sum(item["quantidade"] * item["valor_unitario"] for item in itens_nota)
    return total_itens, total_valor





def calcular_solicitacoes():
        # Se você tiver uma função para obter movimentações, utilize-a
        # Caso contrário, o exemplo abaixo retorna 0
        if hasattr(data_manager, 'obter_movimentacoes'):
            movimentacoes = data_manager.obter_movimentacoes()
            total_solicitacoes = len(movimentacoes)
            total_valor = sum(mov["valor"] for mov in movimentacoes)
        else:
            total_solicitacoes, total_valor = 0, 0.0
        return total_solicitacoes, total_valor
    
# --- Tela de Visualização do Estoque ---
class VisualizarEstoque(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Visualizar Estoque")
        self.state("zoomed") 
        self.create_widgets()
        self.carregar_itens()
    
    
    

    def create_widgets(self):
        
        # ------------------------
        # Retângulos de Informações
        # ------------------------
        info_frame = tk.Frame(self)
        info_frame.pack(pady=5, fill="x")

        # Lista de retângulos com título, conteúdo, imagem e, para o último, indicador de clique
        retangulos = [
            {"title": "Entrada de Mercadoria", 
            "content": ["Compra:"], 
            "img": "Imagens/retangulos/entrada.png"},
            {"title": "Saída de Mercadoria", 
            "content": ["Solicitações:"], 
            "img": "Imagens/retangulos/saida.png"},
            {"title": "Estoque atual:", 
            "content": ["Detalhes do estoque"], 
            "img": "Imagens/retangulos/estoque.png"},
            {"title": "Incluir Nota Fiscal", 
            "content": ["Clique para incluir nota fiscal"], 
            "img": "Imagens/retangulos/nova_fiscal.png",
            "clickable": True}
        ]

        self.retangulo_images = []
        
        

        for i, ret in enumerate(retangulos):
            # Se for retângulo clicável, usamos um CTkButton
            if ret.get("clickable"):
                # Carrega a imagem com um tamanho um pouco maior (mais retangular, ex.: 70x70)
                try:
                    img = Image.open(ret["img"]).convert("RGBA")
                    desired_size = (70, 70)
                    img = img.resize(desired_size, Image.Resampling.LANCZOS)
                    ctk_img = CTkImage(light_image=img, dark_image=img, size=desired_size)
                except Exception as e:
                    print(f"Erro ao carregar imagem {ret['img']}: {e}")
                    ctk_img = None

                # Cria o CTkButton com formato mais retangular (mais largo e menor em altura)
                botao_nf = CTkButton(info_frame,
                    text=ret["title"],              # O título do botão
                    image=ctk_img,                  # Imagem carregada
                    compound="left",                # Imagem à esquerda do texto
                    command=self.abrir_incluir_nova_fiscal,
                    font=("Helvetica", 14, "bold"),
                    fg_color="#7617bb",               # Fundo branco (ou ajuste conforme necessário)
                    text_color="white",           # Texto com cor roxa
                    hover_color="#a084c9",          # Efeito hover
                    corner_radius=12,
                    width=200,                      # Largura fixa
                    height=20)                      # Altura fixa (mais retangular)
                botao_nf.grid(row=0, column=i, padx=10, pady=5, sticky="nsew")
            else:
                # Para os retângulos não clicáveis, criaremos um frame customizado.
                # Em vez de usar o LabelFrame padrão, criaremos um frame com header personalizado.
                frame = tk.Frame(info_frame, bg="#f0f0f0", bd=2, relief="groove")
                frame.grid(row=0, column=i, padx=10, pady=5, sticky="nsew")
                info_frame.grid_columnconfigure(i, weight=1)
                
                # Cria o header: um frame contendo a imagem e o título lado a lado
                header = tk.Frame(frame, bg="#f0f0f0")
                header.pack(fill="x", padx=5, pady=(2,0))
                
                # Carrega a imagem com tamanho 25x25
                try:
                    img = Image.open(ret["img"]).convert("RGBA")
                    desired_size = (30, 30)
                    img = img.resize(desired_size, Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                except Exception as e:
                    print(f"Erro ao carregar imagem {ret['img']}: {e}")
                    photo = None
                self.retangulo_images.append(photo)
                
                # Label com a imagem, alinhada à esquerda
                if photo:
                    lbl_img = tk.Label(header, image=photo, bg="#f0f0f0")
                    lbl_img.pack(side="left", anchor="w")
                
                # Label com o título, à direita da imagem
                lbl_title = tk.Label(header, text=ret["title"], font=("Helvetica", 10, "bold"), fg="#800080", bg="#f0f0f0")
                lbl_title.pack(side="left", anchor="w", padx=5)
                
               
                for line in ret["content"]:
                    frame_line = tk.Frame(frame, bg="#f0f0f0")
                    frame_line.pack(fill="x", padx=5, pady=1)
                    lbl_text = tk.Label(frame_line, text=line, anchor="w", bg="#f0f0f0", font=("Helvetica", 9, "bold"))
                    
                    # Define a cor para cada retângulo
                    if ret["title"] == "Entrada de Mercadoria" and line.startswith("Compra:"):
                        lbl_text.configure(fg="green")
                    elif ret["title"] == "Saída de Mercadoria" and line.startswith("Solicitações:"):
                        lbl_text.configure(fg="red")
                    elif ret["title"] == "Estoque atual:" and line.startswith("Detalhes do estoque"):
                        lbl_text.configure(fg="blue")
                    
                    lbl_text.pack(side="left")
                    
                    if ret["title"] == "Entrada de Mercadoria":
                        if line.startswith("Compra:"):
                            self.lbl_compra_valores = tk.Label(frame_line, text="Calculando...", anchor="e", bg="#f0f0f0", font=("Helvetica", 9), fg="green")
                            self.lbl_compra_valores.pack(side="right")
                    
                    elif ret["title"] == "Saída de Mercadoria":
                        if line.startswith("Solicitações:"):
                            self.lbl_solicitacoes = tk.Label(frame_line, text="Calculando...", anchor="e", bg="#f0f0f0", font=("Helvetica", 9), fg="red")
                            self.lbl_solicitacoes.pack(side="right")
                                            
                    elif ret["title"] == "Estoque atual:":
                        if line.startswith("Detalhes do estoque"):
                            self.lbl_estoque_atual = tk.Label(frame_line, text="Calculando...", anchor="e", bg="#f0f0f0", font=("Helvetica", 9), fg="blue")
                            self.lbl_estoque_atual.pack(side="right")

                
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
        colunas = ("Numero NF", "Data", "Filial", "Fornecedor", "Quantidade", "Valor Total")

        self.tree = ttk.Treeview(self, columns=colunas, show="headings")
        for col in colunas:
            self.tree.heading(col, text=col)  # já vem com o nome correto
            # Ajuste as larguras conforme necessário:
            if col == "Numero NF":
                self.tree.column(col, width=100)
            elif col == "Data":
                self.tree.column(col, width=100)
            elif col == "Filial":
                self.tree.column(col, width=100)
            elif col == "Fornecedor":
                self.tree.column(col, width=150)
            elif col == "Quantidade":
                self.tree.column(col, width=100)
            elif col == "Valor Total":
                self.tree.column(col, width=120)

        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        # Vincula o clique com o botão direito para exibir o menu contextual
        self.tree.bind("<Button-3>", self.mostrar_menu_contextual)
        
    def atualizar_labels_info(self):
        # Atualiza "Compra:" (todas as notas, independentemente de estoque)
        total_itens, total_valor = calcular_compra()
        valor_formatado = f'R$ {total_valor:,.2f}'.replace(",", "X").replace(".", ",").replace("X", ".")
        self.lbl_compra_valores.configure(
            text=f"Qtd: {total_itens} | Valor: {valor_formatado}",
            font=("Helvetica", 9, "bold")
    )

        estoque_atual_itens, estoque_atual_valor = calcular_estoque_atual()
        estoque_atual_valor_formatado = f'R$ {estoque_atual_valor:,.2f}'.replace(",", "X").replace(".", ",").replace("X", ".")
        self.lbl_estoque_atual.configure(text=f"Qtd: {estoque_atual_itens} | Valor: {estoque_atual_valor_formatado}",
                                         font=("Helvetica", 9, "bold"))


    def abrir_incluir_nova_fiscal(self, event=None):
        IncluirNovaFiscal(self)

    def carregar_itens(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        notas = data_manager.obter_notas_fiscais()
        
        for nota in notas:
            itens_nota = data_manager.obter_itens_nota(nota["id"])

            # Calcula a quantidade total de itens da nota
            quantidade_total = sum(item["quantidade"] for item in itens_nota)

            # Calcula o valor total corretamente
            valor_total = sum(item["quantidade"] * item["valor_unitario"] for item in itens_nota)

            # Formata o valor total como moeda brasileira
            valor_total_formatado = f'R$ {valor_total:,.2f}'.replace(",", "X").replace(".", ",").replace("X", ".")

            # Insere na Treeview
            self.tree.insert("", "end", values=(
                nota.get("numero", ""),
                nota.get("data_entrada", ""),
                nota.get("filial", ""),
                nota.get("fornecedor", ""),
                quantidade_total,
                valor_total_formatado  # Aqui usamos o valor formatado
            ))
        self.atualizar_labels_info() 


    def filtrar(self):
        fornecedor = self.filtro_fornecedor.get().strip().lower()
        nome_item = self.filtro_item.get().strip().lower()
        itens_filtrados = []
        for item in data_manager.obter_itens():
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
            for item in data_manager.obter_itens():
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
            for log in data_manager.obter_logs():
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
        item = next((i for i in data_manager.obter_itens() if i["id"] == item_id), None)
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

        data_manager.obter_logs().append(novo_log)
       
        self.carregar_itens()  # Atualiza a visualização

    def abrir_detalhe_serial_status(self, row_id):
        valores = self.tree.item(row_id, "values")
        item_id = valores[0]
        item = next((i for i in data_manager.obter_itens() if i["id"] == item_id), None)
        if item:
            DetalharItemStatus(self, item)
        else:
            messagebox.showerror("Erro", "Item não encontrado.")
            
    
class IncluirNovaFiscal(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Incluir Nova Nota Fiscal")
        self.geometry("700x550")
        self.configure(bg="white")
       # self.resizable(False, False)  # Impede redimensionamento ou maximização

        style = ttk.Style()
        style.theme_use("default")
        # Define a cor de fundo do Notebook (a área que contém as abas)
        style.configure("TNotebook", background="#E6E6FF", borderwidth=0)
        # Configura as abas com um fundo roxo clarinho e define o padding
        style.configure("TNotebook.Tab", background="#E6E6FF", foreground="black", padding=[10,5])
        # Define a cor das abas quando selecionadas
        style.map("TNotebook.Tab", background=[("selected", "#D8BFD8")])

        # Captura a data e horário exatos antes de criar os widgets
        agora = datetime.datetime.now()
        self.data_e_hora = agora.strftime("%d/%m/%Y %H:%M")  # Formato dd/mm/aaaa HH:MM

        # Cria o Notebook e as abas
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # Aba Nota Fiscal: contém todos os campos da nota
        self.aba_nota_fiscal = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.aba_nota_fiscal, text="Nota Fiscal")

        # Aba Itens da Nota: contém os campos para inserir itens (Camiseta, Jaqueta, Calça)
        self.aba_itens = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.aba_itens, text="Itens da Nota")

        # Cria os widgets em cada aba
        self.create_widgets_nota_fiscal()
        self.create_widgets_itens_nota()

    def create_widgets_nota_fiscal(self):
        # 1. Usuário Responsável
        tk.Label(self.aba_nota_fiscal, 
             text="Responsável pela Entrada da Nota:", 
             bg="white", font=("Helvetica", 10)
             ).pack(anchor="w", padx=10, pady=(10, 0))
        vcmd_letras = (self.register(self.validar_nome), '%P')
        self.entry_usuario = tk.Entry(self.aba_nota_fiscal, width=35, validate="key", validatecommand=vcmd_letras)
        self.entry_usuario.pack(anchor="w", padx=10, pady=5)

        # 2. Filial de Recebimento (dropdown)
        tk.Label(self.aba_nota_fiscal, 
                text="Filial de Recebimento:", 
                bg="white", font=("Helvetica", 10)
                ).pack(anchor="w", padx=10, pady=(10, 0))
        self.filial_var = tk.StringVar()
        self.combo_filial = ttk.Combobox(self.aba_nota_fiscal, 
                                        textvariable=self.filial_var, 
                                        state="readonly",
                                        values=["Filial 1", "Filial 2", "Filial 3"],
                                        width=20)
        self.combo_filial.current(0)
        self.combo_filial.pack(anchor="w", padx=10, pady=5)
        
          # 2.1 Fornecedor (combobox, puxando os fornecedores do banco de dados)
        tk.Label(self.aba_nota_fiscal, 
                text="Fornecedor:", 
                bg="white", font=("Helvetica", 10)
                ).pack(anchor="w", padx=10, pady=(10, 0))
        self.fornecedor_nota_var = tk.StringVar()
        # Puxa os fornecedores do banco de dados
        fornecedores = [f["nome"] for f in data_manager.obter_fornecedores()]
        self.combo_fornecedor_nota = ttk.Combobox(self.aba_nota_fiscal, 
                                                textvariable=self.fornecedor_nota_var, 
                                                state="readonly",
                                                values=fornecedores,
                                                width=30)
        if fornecedores:
            self.combo_fornecedor_nota.current(0)
        self.combo_fornecedor_nota.pack(anchor="w", padx=10, pady=5)

        # 3. Data da Entrada (texto e data alinhados na mesma linha)
        frame_data = tk.Frame(self.aba_nota_fiscal, bg="white")
        frame_data.pack(anchor="w", padx=10, pady=5)
        tk.Label(frame_data, 
                text="Data da Entrada:", 
                bg="white", font=("Helvetica", 10, "bold")
                ).pack(side="left", padx=(0, 5))
        self.label_data_entrada = tk.Label(frame_data, 
                                        text=self.data_e_hora, 
                                        bg="white", font=("Helvetica", 10), 
                                        width=15, anchor="w")
        self.label_data_entrada.pack(side="left")

        # 4. Data da Emissão da Nota
        tk.Label(self.aba_nota_fiscal, 
                text="Data da Emissão da Nota:", 
                bg="white", font=("Helvetica", 10)
                ).pack(anchor="w", padx=10, pady=(10, 0))
        if DateEntry:
            self.entry_data_emissao = DateEntry(self.aba_nota_fiscal, 
                                                date_pattern="dd/mm/yyyy", 
                                                width=20)
        else:
            self.entry_data_emissao = tk.Entry(self.aba_nota_fiscal, width=20)
            self.entry_data_emissao.insert(0, datetime.date.today().strftime("%d/%m/%Y"))
        self.entry_data_emissao.pack(anchor="w", padx=10, pady=5)

        # 5. Dados da Nota (Série, Número, Chave) organizados em uma mesma linha
        frame_dados = tk.Frame(self.aba_nota_fiscal, bg="white")
        frame_dados.pack(anchor="w", padx=10, pady=(10,5))
        tk.Label(frame_dados, text="Série:", bg="white", font=("Helvetica", 10)).grid(row=0, column=0, sticky="w", padx=5)
        self.entry_serie = tk.Entry(frame_dados, width=8)
        self.entry_serie.grid(row=0, column=1, padx=5, sticky="w")
        tk.Label(frame_dados, text="Número da Nota:", bg="white", font=("Helvetica", 10)).grid(row=0, column=2, sticky="w", padx=5)
        self.entry_numero = tk.Entry(frame_dados, width=12)
        self.entry_numero.grid(row=0, column=3, padx=5, sticky="w")
        tk.Label(frame_dados, text="Chave:", bg="white", font=("Helvetica", 10)).grid(row=0, column=4, sticky="w", padx=5)
        self.entry_chave = tk.Entry(frame_dados, width=30)
        self.entry_chave.grid(row=0, column=5, padx=5, sticky="w")
        
        # 6. Tipo de Movimentação e Status
        frame_mov = tk.Frame(self.aba_nota_fiscal, bg="white")
        frame_mov.pack(anchor="w", padx=10, pady=(10,5))
        tk.Label(frame_mov, text="Tipo de Movimentação:", bg="white", font=("Helvetica", 10)).grid(row=0, column=0, sticky="w", padx=5)
        self.movimentacao_var = tk.StringVar()
        self.combo_mov = ttk.Combobox(frame_mov, textvariable=self.movimentacao_var, state="readonly",
                                    values=["Entrada", "Saída"], width=15)
        self.combo_mov.current(0)
        self.combo_mov.grid(row=0, column=1, padx=5, sticky="w")
        tk.Label(frame_mov, text="Status:", bg="white", font=("Helvetica", 10)).grid(row=0, column=2, sticky="w", padx=5)
        self.status_var = tk.StringVar()
        self.status_var.set("Novo")
        rb_novo = tk.Radiobutton(frame_mov, text="Novo", variable=self.status_var, value="Novo", bg="white", font=("Helvetica", 10))
        rb_novo.grid(row=0, column=3, padx=5, sticky="w")
        rb_usado = tk.Radiobutton(frame_mov, text="Usado", variable=self.status_var, value="Usado", bg="white", font=("Helvetica", 10))
        rb_usado.grid(row=0, column=4, padx=5, sticky="w")
        
        # 7. Disponível em Estoque
        frame_disp = tk.Frame(self.aba_nota_fiscal, bg="white")
        frame_disp.pack(anchor="w", padx=10, pady=(10,5))
        tk.Label(frame_disp, text="Disponível em Estoque?", bg="white", font=("Helvetica", 10)).grid(row=0, column=0, sticky="w", padx=5)
        self.estoque_var = tk.StringVar()
        self.estoque_var.set("Sim")
        rb_sim = tk.Radiobutton(frame_disp, text="Sim", variable=self.estoque_var, value="Sim", bg="white", font=("Helvetica", 10))
        rb_sim.grid(row=0, column=1, padx=5, sticky="w")
        rb_nao = tk.Radiobutton(frame_disp, text="Não", variable=self.estoque_var, value="Não", bg="white", font=("Helvetica", 10))
        rb_nao.grid(row=0, column=2, padx=5, sticky="w")
        
        
        
        # 8. Observações (campo Text com limite de 200 caracteres)
        tk.Label(self.aba_nota_fiscal, 
                text="Observações (até 200 caracteres):", 
                bg="white", font=("Helvetica", 10)
                ).pack(anchor="w", padx=10, pady=(10,0))
        self.entry_obs = tk.Text(self.aba_nota_fiscal, height=3, wrap="word", font=("Helvetica", 10))
        self.entry_obs.pack(anchor="w", padx=10, pady=5)
        self.entry_obs.bind("<KeyRelease>", self.limit_text)

        
        # 9. Botão para Salvar Nota Fiscal (usando CTkButton)
        self.btn_salvar_nf = CTkButton(self.aba_nota_fiscal,
            text="Salvar Nota Fiscal",
            image=None,  # Se houver um PNG, carregue-o via CTkImage
            compound="left",
            command=self.salvar_nf,
            font=("Helvetica", 14, "bold"),
            fg_color="#7617bb",
            text_color="white",
            hover_color="#a084c9",
            corner_radius=12,
            width=200,
            height=40
        )
        self.btn_salvar_nf.pack(pady=10)
    
    def validar_nome(self, P):
        import re
        # Aceita apenas letras (inclusive acentuadas) e espaços
        return bool(re.fullmatch(r"[A-Za-zÀ-ÖØ-öø-ÿ ]*", P))


    def create_widgets_itens_nota(self):
        """Cria os elementos da aba 'Itens da Nota' com campos: Produto, Quantidade e Valor Unitário,
        permitindo adicionar novas linhas ao clicar no botão '+'."""
        tk.Label(self.aba_itens, text="Cadastro de Itens da Nota", bg="white",
                font=("Helvetica", 12, "bold")).pack(anchor="w", pady=10)
        
        # Frame que conterá todas as linhas de itens
        self.items_frame = tk.Frame(self.aba_itens, bg="white")
        self.items_frame.pack(anchor="w", padx=10, pady=5)
        
        # Lista que armazenará as variáveis de cada linha de item
        self.itens_nota_vars = []  # Cada elemento será um dicionário com as variáveis dos campos
        
        # Cria a primeira linha de item
        self.add_item_row()
        
        # Botão circular para adicionar uma nova linha
    # Tente carregar a imagem para o botão, se existir
        try:
            img = Image.open("Imagens/botoes/add_item.png").convert("RGBA")
            # Redimensiona a imagem para um tamanho adequado (por exemplo, 30x30 pixels)
            img = img.resize((20, 20), Image.Resampling.LANCZOS)
            ctk_img = CTkImage(light_image=img, dark_image=img, size=(20, 20))
        except Exception as e:
            print(f"Erro ao carregar imagem do botão: {e}")
            ctk_img = None

        # Crie o botão usando CTkButton com formato circular
        self.btn_add_item = CTkButton(self.aba_itens,
            text="",
            command=self.add_item_row,
            width=30,      # Largura fixa
            height=30,     # Altura fixa
            corner_radius=20,  # Metade da altura para deixar circular
            fg_color="#7617bb",
            text_color="white",
            hover_color="#a084c9",
            font=("Helvetica", 14, "bold"),
            image=ctk_img,          # Usa a imagem se carregada; caso contrário, None
            compound="left"         # Exibe a imagem à esquerda do texto (pode ajustar conforme desejar)
        )
        self.btn_add_item.pack(anchor="w", pady=10)


    def add_item_row(self):
        """Adiciona uma nova linha para inserir um item da nota fiscal."""
        # Cria um frame para a nova linha
        row_frame = tk.Frame(self.items_frame, bg="white")
        row_frame.pack(anchor="w", fill="x", pady=3)
        
        # Campo: Produto (Combobox)
        tk.Label(row_frame, text="Produto:", bg="white", font=("Helvetica", 10)).grid(row=0, column=0, sticky="w", padx=5)
        produto_var = tk.StringVar()
        # Busca os itens cadastrados
        itens = data_manager.obter_itens()
        lista_produtos = [item["nome"] for item in itens]
        # Se não houver nenhum item cadastrado, pode deixar uma opção padrão, por exemplo "Nenhum item cadastrado"
        if not lista_produtos:
            lista_produtos = ["Nenhum item cadastrado"]
        combo_produto = ttk.Combobox(row_frame, textvariable=produto_var, state="readonly",
                                    values=lista_produtos, width=20)
        combo_produto.current(0)
        combo_produto.grid(row=0, column=1, sticky="w", padx=5)
        # Vincula a seleção para atualizar o valor unitário
        combo_produto.bind("<<ComboboxSelected>>", 
                        lambda event, prod_var=produto_var, val_var=None: self.preencher_valor_unitario(prod_var, row_frame))
        
        # Campo: Quantidade (Entry com validação para dígitos)
        tk.Label(row_frame, text="Quantidade:", bg="white", font=("Helvetica", 10)).grid(row=0, column=2, sticky="w", padx=5)
        quantidade_var = tk.StringVar()
        entry_quantidade = tk.Entry(row_frame, textvariable=quantidade_var, font=("Helvetica", 10), width=10)
        entry_quantidade.grid(row=0, column=3, sticky="w", padx=5)
        # Validação para aceitar apenas dígitos
        vcmd = (self.register(self.validate_digits), '%P')
        entry_quantidade.config(validate="key", validatecommand=vcmd)
        
        tk.Label(row_frame, text="Tamanho:", bg="white", font=("Helvetica", 10)).grid(row=0, column=4, sticky="w", padx=5)
        tamanho_var = tk.StringVar()
        # Você pode definir opções fixas ou, se existir no cadastro, buscar dinamicamente.
        opcoes_tamanho = ["P", "M", "G", "GG"]
        combo_tamanho = ttk.Combobox(row_frame, textvariable=tamanho_var, state="readonly",
                                    values=opcoes_tamanho, width=5)
        combo_tamanho.current(0)
        combo_tamanho.grid(row=0, column=5, sticky="w", padx=5)
        
        # Campo: Valor Unitário (Entry com formatação automática)
        tk.Label(row_frame, text="Valor Unitário:", bg="white", font=("Helvetica", 10)).grid(row=0, column=6, sticky="w", padx=5)
        valor_var = tk.StringVar()
        entry_valor = tk.Entry(row_frame, textvariable=valor_var, font=("Helvetica", 10), width=15)
        entry_valor.grid(row=0, column=7, sticky="w", padx=5)
        entry_valor.bind("<FocusOut>", self.formatar_valor_evento)
        
        # NOVO CAMPO: Tamanho (Combobox)
        
        
        # Armazena as variáveis para esta linha na lista
        self.itens_nota_vars.append({
            "produto": produto_var,
            "quantidade": quantidade_var,
            "valor_unitario": valor_var,
            "tamanho": tamanho_var,
            "frame": row_frame  # referência para remoção futura, se necessário
        })

    def preencher_valor_unitario(self, produto_var, row_frame):
        """Busca o valor unitário do item selecionado e preenche o campo correspondente."""
        nome_produto = produto_var.get()
        # Busca os itens cadastrados
        itens = data_manager.obter_itens()
        # Procura o item com o nome selecionado
        item = next((it for it in itens if it["nome"] == nome_produto), None)
        valor = ""
        if item:
            valor = formatar_valor(item.get("valor_unitario", 0))
        # Atualiza o campo de valor unitário na linha correspondente.
        # Assumindo que a variável "valor_unitario" está na mesma linha em self.itens_nota_vars
        for linha in self.itens_nota_vars:
            if linha["produto"] == produto_var:
                linha["valor_unitario"].set(valor)
                break

    # Função para validar que apenas dígitos sejam digitados na quantidade
    def validate_digits(self, P):
        return P.isdigit() or P == ""

    # Função para formatar o valor unitário ao sair do campo
    def formatar_valor_evento(self, event):
        entry = event.widget
        valor_texto = entry.get().strip()
        if valor_texto:
            try:
                # Remove formatação anterior, se houver
                valor_limpo = valor_texto.replace("R$", "").replace(".", "").replace(",", ".")
                valor_float = float(valor_limpo)
                # Formata o valor no padrão brasileiro (supondo que sua função formatar_valor exista)
                entry.delete(0, tk.END)
                entry.insert(0, formatar_valor(valor_float))
            except Exception as e:
                print("Erro ao formatar valor:", e)



    def limit_text(self, event):
        content = self.entry_obs.get("1.0", "end-1c")
        if len(content) > 200:
            self.entry_obs.delete("1.0", "end")
            self.entry_obs.insert("1.0", content[:200])

    def salvar_nf(self):
        # Capturar os valores dos campos preenchidos
        responsavel = self.entry_usuario.get()
        filial = self.filial_var.get()
        fornecedor = self.fornecedor_nota_var.get()
        data_entrada = self.data_e_hora
        data_emissao = self.entry_data_emissao.get()
        serie = self.entry_serie.get()
        numero = self.entry_numero.get()
        chave = self.entry_chave.get()
        movimentacao = self.movimentacao_var.get()
        status = self.status_var.get()
        estoque = self.estoque_var.get()
        observacoes = self.entry_obs.get("1.0", tk.END).strip()

        # Verifica se os campos essenciais estão preenchidos
        if not (responsavel and filial and serie and numero and chave):
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
            return

        try:
            # Conectar ao banco de dados
            conn = sqlite3.connect("db/estoque.db")
            cursor = conn.cursor()

            # Inserir os dados da nota fiscal
            cursor.execute("""
            INSERT INTO notas_fiscais (
                responsavel, filial, fornecedor, data_entrada, data_emissao,
                serie, numero, chave, movimentacao, status, estoque, observacoes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (responsavel,
            filial,
            fornecedor,  
            data_entrada,
            data_emissao,
            serie,
            numero,
            chave,
            movimentacao,
            status,
            estoque,
            observacoes))

            

            # Obter o ID recém-inserido ANTES de fechar a conexão
            nota_id = cursor.lastrowid  

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()

            # Chamar a função para salvar os itens, utilizando o dicionário itens_nota_vars
            if hasattr(self, 'itens_nota_vars') and self.itens_nota_vars:
                self.salvar_itens_nota(nota_id)
            else:
                print("Aviso: Nenhum item foi adicionado à nota fiscal.")

            # Exibir mensagem de sucesso e fechar a janela
            messagebox.showinfo("Sucesso", "Nota Fiscal salva com sucesso!")
            self.destroy()

        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", "Chave da nota já cadastrada!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar nota: {e}")


    def salvar_itens_nota(self, nota_id):
        try:
            conn = sqlite3.connect("db/estoque.db")
            cursor = conn.cursor()

            # Percorre cada dicionário na lista self.itens_nota_vars
            for item_dict in self.itens_nota_vars:
                produto_var = item_dict["produto"]
                quantidade_var = item_dict["quantidade"]
                valor_unitario_var = item_dict["valor_unitario"]
                tamanho_var = item_dict["tamanho"]

                produto = produto_var.get()
                quantidade_text = quantidade_var.get()
                valor_texto = valor_unitario_var.get()
                tamanho = tamanho_var.get()

                # Converte quantidade para inteiro
                try:
                    quantidade = int(quantidade_text)
                except ValueError:
                    quantidade = 0

                # Converte valor unitário (removendo formatação se necessário)
                try:
                    valor_limpo = valor_texto.replace("R$", "").replace(".", "").replace(",", ".")
                    valor_unitario = float(valor_limpo)
                except ValueError:
                    valor_unitario = 0.0

                # Insere no banco
                cursor.execute("""
                    INSERT INTO itens_nota (nota_id, produto, quantidade, valor_unitario, tamanho)
                    VALUES (?, ?, ?, ?, ?)
                """, (nota_id, produto, quantidade, valor_unitario, tamanho))

            conn.commit()
            conn.close()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar itens da nota: {e}")





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
                data_manager.obter_logs().append(log)
                

                messagebox.showinfo("Sucesso", "Status alterado com sucesso!")
                self.callback()  # Atualiza a tela que chamou essa janela
                self.destroy()
                return

        messagebox.showerror("Erro", "Serial não encontrado!")

# --- Execução da Aplicação ---
if __name__ == "__main__":
    
    inicializar_banco()
    data_manager = DataManagerSQL()
    app = MainMenu()
    app.withdraw()
    splash = SplashScreen(app, gif_path="Imagens/splash.gif", duration=3000)
    agendar_backup_automatico(app)
    app.mainloop()
    
