import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk
import threading
import queue
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, StaleElementReferenceException
from chromedriver_py import binary_path
import time
from datetime import datetime
import sys
import unicodedata
import re
import csv
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# ================== CONFIGURAÇÕES DE ESTILO (OVG) ==================
COR_FUNDO = "#F7F7F7"
COR_PRIMARIA = "#7B4FA3"    # Roxo OVG
COR_SECUNDARIA = "#F06292"  # Rosa OVG
COR_TEXTO = "#1F1F1F"
COR_AVISO = "#D32F2F"       # Vermelho escuro para avisos

# ================== VARIÁVEIS GLOBAIS DE CONTROLE ==================
ARQUIVO_TEMP = "temp_dados_extract.csv"
URL_LOGIN = "https://10.237.1.11/bolsa/login/login"
URL_BUSCA = "https://10.237.1.11/bolsa/index/index"

# Fila para comunicação entre a Thread do Robô e a GUI
fila_mensagens = queue.Queue()

# ================== FUNÇÕES AUXILIARES DO ROBÔ ==================

def safe_click(driver, xpath_locator, wait, timeout=3):
    end_time = time.time() + timeout
    last_err = None
    while time.time() < end_time:
        try:
            elem = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_locator)))
            elem.click()
            return True
        except (StaleElementReferenceException, WebDriverException) as e:
            last_err = e
            time.sleep(0.5)
    raise last_err or Exception(f"Não foi possível clicar em {xpath_locator}")

def safe_get_text(driver, xpath_locator, wait=None):
    for _ in range(3):
        try:
            elem = driver.find_element(By.XPATH, xpath_locator)
            return elem.text.strip()
        except (StaleElementReferenceException, WebDriverException):
            time.sleep(0.5)
    return ""

def str_to_float(valor_str):
    if not valor_str: return 0.0
    limpo = str(valor_str).replace('R$', '').strip()
    try:
        if ',' in limpo:
            limpo = limpo.replace('.', '').replace(',', '.')
        elif '.' in limpo:
            if limpo.count('.') > 1 or (len(limpo) - limpo.rfind('.') > 3):
                 limpo = limpo.replace('.', '')
        return float(limpo)
    except:
        return 0.0

def resetar_navegacao(driver, wait, usuario, senha):
    try:
        driver.implicitly_wait(0.5)
        if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0: return False 
        if len(driver.find_elements(By.XPATH, '//*[@id="campo"]')) > 0:
            driver.implicitly_wait(0.5)
            return True
        driver.implicitly_wait(0.5)
        driver.back()
        try:
            if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0: return False
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="campo"]')))
            return True
        except: pass
        driver.back()
        try:
            if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0: return False
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="campo"]')))
            return True
        except: pass
        driver.get(URL_BUSCA)
        if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0: return False
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="campo"]')))
        return True
    except: return False

def garantir_sessao(driver, wait, usuario, senha):
    try:
        if len(driver.find_elements(By.XPATH, '//*[@id="campo"]')) > 0: return True
        print(" > [SISTEMA] Realizando Login...")
        if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) == 0:
            driver.get(URL_LOGIN)
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="usuario"]'))).clear()
        driver.find_element(By.XPATH, '//*[@id="usuario"]').send_keys(usuario)
        driver.find_element(By.XPATH, '//*[@id="senha"]').clear()
        driver.find_element(By.XPATH, '//*[@id="senha"]').send_keys(senha)
        safe_click(driver, '//*[@id="conteudo"]/form/fieldset/p[3]/input', wait)
        
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="campo"]')))
            return True
        except:
            print(" > [ERRO] Falha no login (senha incorreta?)")
            return False
            
    except Exception as e:
        print(f" > [ERRO LOGIN] {e}")
        return False

def salvar_checkpoint(dados):
    existe = os.path.isfile(ARQUIVO_TEMP)
    # COLUNAS REMOVIDAS: VALOR CALCULADO, DIFERENÇA, STATUS
    colunas = [
        "CPF", "INSCRIÇÃO", "NOME", "CURSO", "FACULDADE", "SITUAÇÃO", "PERIODO", 
        "TIPO DE BOLSA", "DATA COLETA", "VALOR S/ DESCONTO", "VALOR C/ DESCONTO",
        "VALOR BENEFÍCIOS", "INFO BENEFÍCIOS", 
        "DATA LANÇAMENTO", "VALOR DA BOLSA"
    ]
    try:
        with open(ARQUIVO_TEMP, mode='a', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=colunas, delimiter=';')
            if not existe: writer.writeheader()
            writer.writerow(dados)
    except Exception as e:
        print(f"Erro checkpoint: {e}")

def gerar_excel_final(nome_arquivo_destino=None):
    if not os.path.isfile(ARQUIVO_TEMP):
        return None
    
    try:
        df_result = pd.read_csv(ARQUIVO_TEMP, sep=';', dtype=str, encoding='utf-8-sig')
        if 'NOME' in df_result.columns: df_result = df_result.dropna(subset=['NOME'])
        
        # Colunas removidas da lista de conversão
        cols_numericas = ["VALOR S/ DESCONTO", "VALOR C/ DESCONTO", "VALOR BENEFÍCIOS", "VALOR DA BOLSA"]
        for col in cols_numericas:
            if col in df_result.columns:
                df_result[col] = df_result[col].fillna('0')
                df_result[col] = pd.to_numeric(df_result[col], errors='coerce').fillna(0.0)

        if not nome_arquivo_destino:
            nome_arquivo_destino = f"Relatorio_Final_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx"

        with pd.ExcelWriter(nome_arquivo_destino, engine='xlsxwriter') as writer:
            df_result.to_excel(writer, sheet_name="Dados", index=False, startrow=1, header=False)
            workbook  = writer.book
            worksheet = writer.sheets["Dados"]
            worksheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
            
            # Formatos
            fmt_header = workbook.add_format({'bold': True, 'fg_color': '#000000', 'font_color': '#FFFFFF'})
            fmt_currency = workbook.add_format({'num_format': '#,##0.00'}) 
            fmt_center = workbook.add_format({'align': 'center'})
            fmt_text = workbook.add_format({'num_format': '@'}) 
            
            for col_num, value in enumerate(df_result.columns.values):
                worksheet.write(0, col_num, value, fmt_header)
            
            for i, col in enumerate(df_result.columns):
                col_upper = str(col).strip().upper()
                largura = len(col_upper) + 4
                if "INSCRI" in col_upper or "CPF" in col_upper: worksheet.set_column(i, i, largura + 5, fmt_text)
                elif col in cols_numericas: worksheet.set_column(i, i, 18, fmt_currency)
                elif "NOME" in col: worksheet.set_column(i, i, 40)
                else: worksheet.set_column(i, i, largura)
            
            # Formatação Condicional e Legenda REMOVIDAS
            
        return nome_arquivo_destino
    except Exception as e:
        print(f"Erro Excel: {e}")
        return None

def extrair_tabelas_financeiras(driver, mapa_financeiro, coletas):
    IDX_LANC_DATA, IDX_LANC_VALOR_BOLSA, IDX_LANC_TIPO_BOLSA = 2, 4, 9
    IDX_COLETA_MENSAL_SEM, IDX_COLETA_MENSAL_COM, IDX_COLETA_VALOR_BENEFICIO, IDX_COLETA_INFO_BENEFICIO, IDX_COLETA_DATA = 4, 5, 6, 7, 10

    try:
        tbl_lanc = driver.find_element(By.XPATH, "//legend[contains(., 'Lançamento(s) de Pagto(s)')]/following-sibling::table")
        rows = tbl_lanc.find_elements(By.XPATH, ".//tbody/tr")
        for idx_r in range(len(rows)):
            try:
                tbl_ref = driver.find_element(By.XPATH, "//legend[contains(., 'Lançamento(s) de Pagto(s)')]/following-sibling::table")
                cols = tbl_ref.find_elements(By.XPATH, ".//tbody/tr")[idx_r].find_elements(By.TAG_NAME, "td")
                if len(cols) <= IDX_LANC_TIPO_BOLSA: continue
                dt_str = cols[IDX_LANC_DATA].text.strip()
                dt_obj = datetime.strptime(dt_str, '%d/%m/%Y')
                val_pago = str_to_float(cols[IDX_LANC_VALOR_BOLSA].text.strip())
                tipo_bolsa_tab = cols[IDX_LANC_TIPO_BOLSA].text.strip()
                chave = f"{dt_obj.year}-{1 if dt_obj.month <= 6 else 2}"
                if chave not in mapa_financeiro or dt_obj > mapa_financeiro[chave]['dt_obj']:
                    mapa_financeiro[chave] = {'dt_obj': dt_obj, 'dt_str': dt_str, 'valor': val_pago, 'tipo_bolsa_real': tipo_bolsa_tab}
            except: continue
    except: pass 

    try:
        tbl_coleta = driver.find_element(By.XPATH, "//legend[contains(., 'Coleta de Dados')]/following-sibling::table")
        rows = tbl_coleta.find_elements(By.XPATH, ".//tbody/tr")
        for idx_r in range(len(rows)):
            try:
                tbl_ref = driver.find_element(By.XPATH, "//legend[contains(., 'Coleta de Dados')]/following-sibling::table")
                cols = tbl_ref.find_elements(By.XPATH, ".//tbody/tr")[idx_r].find_elements(By.TAG_NAME, "td")
                if len(cols) <= IDX_COLETA_DATA: continue
                dt_str = cols[IDX_COLETA_DATA].text.strip()
                dt_obj = datetime.strptime(dt_str, '%d/%m/%Y')
                val_s_float = str_to_float(cols[IDX_COLETA_MENSAL_SEM].text.strip())
                val_c_float = str_to_float(cols[IDX_COLETA_MENSAL_COM].text.strip())
                val_beneficio = str_to_float(cols[IDX_COLETA_VALOR_BENEFICIO].text.strip())
                info_beneficio = cols[IDX_COLETA_INFO_BENEFICIO].text.strip()
                chave = f"{dt_obj.year}-{1 if dt_obj.month <= 6 else 2}"
                if chave not in coletas or dt_obj > coletas[chave]['dt_obj']:
                    coletas[chave] = {
                        'dt_obj': dt_obj, 'dt_str': dt_str, 
                        'val_s_float': val_s_float, 'val_c_float': val_c_float,
                        'val_beneficio': val_beneficio, 'info_beneficio': info_beneficio
                    }
            except: continue
    except: pass

def processar_e_salvar(chave_busca, valor_busca, nome, cpf, curso, ies, situacao, tipo_bolsa_capa, coletas, mapa_financeiro, inscricao_override=None):
    # Lógica simplificada: Apenas extração, sem auditoria.
    
    inscricao_final = ""
    if chave_busca == "INSCRIÇÃO":
        inscricao_final = valor_busca
    elif inscricao_override:
        inscricao_final = inscricao_override

    for sem in sorted(coletas.keys()):
        dados_col = coletas[sem]
        dados_fin = mapa_financeiro.get(sem, {'valor': 0.0, 'dt_str': '', 'tipo_bolsa_real': ''})
        
        # Prioriza o tipo de bolsa do financeiro, senão usa o da capa
        tipo_bolsa_final = dados_fin['tipo_bolsa_real'] if dados_fin['tipo_bolsa_real'] else tipo_bolsa_capa
        
        valor_pago = dados_fin['valor']
        
        linha_excel = {
            "CPF": valor_busca if chave_busca == "CPF" else cpf,
            "INSCRIÇÃO": inscricao_final,
            "NOME": nome, "CURSO": curso, "FACULDADE": ies, "SITUAÇÃO": situacao, "PERIODO": sem, 
            "TIPO DE BOLSA": tipo_bolsa_final, "DATA COLETA": dados_col['dt_str'],
            "VALOR S/ DESCONTO": dados_col['val_s_float'], "VALOR C/ DESCONTO": dados_col['val_c_float'],
            "VALOR BENEFÍCIOS": dados_col['val_beneficio'], "INFO BENEFÍCIOS": dados_col['info_beneficio'], 
            "DATA LANÇAMENTO": dados_fin['dt_str'], "VALOR DA BOLSA": valor_pago
        }
        salvar_checkpoint(linha_excel)

# ================== LÓGICA PRINCIPAL DO ROBÔ ==================

def run_selenium_logic(modo_operacao, lista_dados, usuario_login, senha_login):
    if os.path.exists(ARQUIVO_TEMP): os.remove(ARQUIVO_TEMP)

    options = Options()
    options.add_argument('--ignore-certificate-errors') 
    options.add_experimental_option("detach", True)
    service = Service(binary_path)
    
    driver = None
    try:
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 10) 
        driver.get(URL_LOGIN)
        
        if not garantir_sessao(driver, wait, usuario_login, senha_login):
            messagebox.showerror("Erro de Login", "Não foi possível logar. Verifique Usuário e Senha.")
            driver.quit()
            fila_mensagens.put("FIM_ERRO") 
            return

        consecutive_failures = 0
        total_items = len(lista_dados)

        for i, item in enumerate(lista_dados):
            if consecutive_failures >= 3:
                driver.delete_all_cookies()
                driver.get(URL_LOGIN)
                garantir_sessao(driver, wait, usuario_login, senha_login)
                consecutive_failures = 0 

            if not resetar_navegacao(driver, wait, usuario_login, senha_login):
                garantir_sessao(driver, wait, usuario_login, senha_login)
            
            try:
                # --- BUSCA ---
                sucesso_busca = False
                for tentativa in range(3):
                    try:
                        if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0: raise Exception("Sessão caiu")
                        select_element = wait.until(EC.element_to_be_clickable((By.NAME, 'opcao')))
                        Select(select_element).select_by_value('uni_cpf' if modo_operacao == "CPF" else 'uni_codigo')
                        
                        campo = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="campo"]')))
                        campo.clear()
                        campo.send_keys(item)
                        safe_click(driver, '//*[@id="conteudo"]/div[2]/form/fieldset/input[2]', wait)
                        sucesso_busca = True
                        break
                    except:
                        time.sleep(1)
                        if tentativa == 2: pass
                        continue
                
                if not sucesso_busca: 
                    # Sem STATUS no CSV, apenas logamos no console para depuração
                    print(f"[{i+1}/{total_items}] {item}: Não encontrado ou erro na busca.")
                    continue

                # --- RESULTADOS ---
                try:
                    table_body_xp = '//*[@id="conteudo"]/div[2]/fieldset/table/tbody'
                    wait.until(EC.presence_of_element_located((By.XPATH, table_body_xp)))
                    
                    if modo_operacao == "CPF":
                        all_rows = driver.find_elements(By.XPATH, f'{table_body_xp}/tr')
                        total_registros = len(all_rows) - 1 
                        if total_registros <= 0: raise TimeoutException()

                        for idx_reg in range(total_registros):
                            if idx_reg > 0:
                                driver.back()
                                wait.until(EC.presence_of_element_located((By.XPATH, table_body_xp)))
                            
                            current_row_idx = idx_reg + 2
                            base_xp = f'{table_body_xp}/tr[{current_row_idx}]'
                            
                            inscricao_real = safe_get_text(driver, f'{base_xp}/td[1]')
                            nome = safe_get_text(driver, f'{base_xp}/td[2]')
                            ies = safe_get_text(driver, f'{base_xp}/td[4]')
                            ies = unicodedata.normalize('NFKD', ies).upper().encode('ascii', errors='ignore').decode('utf-8')
                            ies = re.sub(r'[-.,]', ' ', ies).strip()
                            curso = safe_get_text(driver, f'{base_xp}/td[5]')
                            tipo_bolsa_capa = safe_get_text(driver, f'{base_xp}/td[6]')
                            situacao = safe_get_text(driver, f'{base_xp}/td[11]')
                            
                            safe_click(driver, f'{base_xp}/td[14]/a[1]', wait)
                            wait.until(EC.presence_of_element_located((By.XPATH, "//legend[contains(., 'Coleta de Dados')]")))
                            
                            mapa = {}
                            coletas = {}
                            extrair_tabelas_financeiras(driver, mapa, coletas)
                            
                            processar_e_salvar("CPF", item, nome, item, curso, ies, situacao, tipo_bolsa_capa, coletas, mapa, inscricao_override=inscricao_real)

                    else:
                        # Modo Inscrição
                        base_xp = '//*[@id="conteudo"]/div[2]/fieldset/table/tbody/tr[2]'
                        wait.until(EC.presence_of_element_located((By.XPATH, base_xp)))
                        
                        nome = safe_get_text(driver, f'{base_xp}/td[2]')
                        cpf_raw = safe_get_text(driver, f'{base_xp}/td[3]')
                        cpf = ''.join(filter(str.isdigit, cpf_raw))
                        ies = safe_get_text(driver, f'{base_xp}/td[4]')
                        ies = unicodedata.normalize('NFKD', ies).upper().encode('ascii', errors='ignore').decode('utf-8')
                        ies = re.sub(r'[-.,]', ' ', ies).strip()
                        curso = safe_get_text(driver, f'{base_xp}/td[5]')
                        tipo_bolsa_capa = safe_get_text(driver, f'{base_xp}/td[6]')
                        situacao = safe_get_text(driver, f'{base_xp}/td[11]')
                        
                        safe_click(driver, f'{base_xp}/td[14]/a[1]', wait)
                        wait.until(EC.presence_of_element_located((By.XPATH, "//legend[contains(., 'Coleta de Dados')]")))
                        
                        mapa = {}
                        coletas = {}
                        extrair_tabelas_financeiras(driver, mapa, coletas)
                        
                        processar_e_salvar("INSCRIÇÃO", item, nome, cpf, curso, ies, situacao, tipo_bolsa_capa, coletas, mapa, inscricao_override=item)

                    consecutive_failures = 0

                except TimeoutException:
                    print(f"[{i+1}/{total_items}] {item}: Não encontrado (Timeout).")
                    consecutive_failures += 1
                    resetar_navegacao(driver, wait, usuario_login, senha_login)
                    continue
            
            except Exception as e:
                consecutive_failures += 1
                resetar_navegacao(driver, wait, usuario_login, senha_login)

    except Exception as e:
        print(f"Erro Fatal Selenium: {e}")
    finally:
        if driver:
            try: driver.quit()
            except: pass
        fila_mensagens.put("FIM")

# ================== GUI (FRONTEND) ==================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Organiza Lista do Excel")
        self.geometry("720x550") 
        self.configure(bg=COR_FUNDO)
        self.minsize(720, 550)
        
        self.modo_var = tk.StringVar(value="CPF")
        self.criar_widgets()

    def criar_widgets(self):
        self.columnconfigure(0, weight=1)

        # 1. LOGO
        try:
            caminho_logo = os.path.join(os.path.dirname(__file__), "logo_ovg.png")
            if os.path.exists(caminho_logo):
                imagem = Image.open(caminho_logo)
                imagem = imagem.resize((150, 150))
                self.logo_img = ImageTk.PhotoImage(imagem)
                lbl_logo = tk.Label(self, image=self.logo_img, bg=COR_FUNDO)
                lbl_logo.grid(row=0, column=0, pady=5)
            else:
                tk.Label(self, text="LOGO OVG", font=("Arial", 20, "bold"), bg=COR_FUNDO, fg=COR_PRIMARIA).grid(row=0, column=0, pady=10)
        except: pass

        # 2. BOTÃO INFO
        btn_info = tk.Button(self, text="INFO", command=self.mostrar_info, bg=COR_SECUNDARIA, fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=4)
        btn_info.place(relx=0.95, rely=0.18, anchor="ne")

        # 3. CONTROLES DE ENTRADA
        frame_controles = tk.Frame(self, bg=COR_FUNDO)
        frame_controles.grid(row=1, column=0, padx=40, pady=5, sticky="ew")

        tk.Label(frame_controles, text="Cole o texto abaixo:", bg=COR_FUNDO, fg=COR_TEXTO, font=("Segoe UI", 11)).pack(side="left")
        
        frame_radio = tk.Frame(frame_controles, bg=COR_FUNDO)
        frame_radio.pack(side="right")
        tk.Label(frame_radio, text="Selecione o tipo:", bg=COR_FUNDO, font=("Segoe UI", 10, "bold")).pack(side="left", padx=5)
        
        rb_cpf = tk.Radiobutton(frame_radio, text="CPF", variable=self.modo_var, value="CPF", bg=COR_FUNDO, font=("Segoe UI", 10))
        rb_cpf.pack(side="left")
        rb_insc = tk.Radiobutton(frame_radio, text="Inscrição", variable=self.modo_var, value="INSCRIÇÃO", bg=COR_FUNDO, font=("Segoe UI", 10))
        rb_insc.pack(side="left")

        # 4. ÁREA DE TEXTO
        self.entrada_texto = tk.Text(self, height=8, font=("Segoe UI", 10), relief="solid", bd=1)
        self.entrada_texto.grid(row=2, column=0, padx=40, pady=5, sticky="nsew")

        # 5. BOTÕES DE AÇÃO
        frame_acao = tk.Frame(self, bg=COR_FUNDO)
        frame_acao.grid(row=3, column=0, pady=15)

        self.btn_iniciar = tk.Button(frame_acao, text="Iniciar navegação automatica", command=self.iniciar_automacao, bg=COR_PRIMARIA, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=30, pady=8)
        self.btn_iniciar.pack(side="left", padx=10)

        btn_limpar = tk.Button(frame_acao, text="Limpar", command=self.limpar_campos, bg="#BDBDBD", fg="#333333", font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=8)
        btn_limpar.pack(side="left", padx=10)

        # 6. MENSAGEM DE AVISO
        aviso_texto = (
            "⚠️ ATENÇÃO: Minimize o navegador da automação.\n"
            "Você pode continuar usando o computador normalmente.\n"
            "Para cancelar, feche o navegador."
        )
        lbl_aviso = tk.Label(self, text=aviso_texto, bg=COR_FUNDO, fg=COR_AVISO, font=("Segoe UI", 10, "bold"), justify="center")
        lbl_aviso.grid(row=4, column=0, pady=10)

        self.verificar_fila()

    def mostrar_info(self):
        messagebox.showinfo("INFORMAÇÕES", 
            "Automação Gestão Bolsa OVG\n\n"
            "Desenvolvido por: Ihan Messias N. dos Santos\n"
            "Departamento: GGCI\n"
            "Data: 26/01/2026\n\n"
            "Função: Coleta dados financeiros e cadastrais automaticamente."
        )

    def limpar_campos(self):
        self.entrada_texto.delete("1.0", tk.END)

    def solicitar_credenciais(self):
        dialog = tk.Toplevel(self)
        dialog.title("Autenticação OVG")
        dialog.geometry("300x180")
        dialog.resizable(False, False)
        dialog.grab_set()
        
        x = self.winfo_x() + (self.winfo_width() // 2) - 150
        y = self.winfo_y() + (self.winfo_height() // 2) - 90
        dialog.geometry(f"+{x}+{y}")

        credenciais = {"user": None, "pass": None}

        tk.Label(dialog, text="Usuário:", font=("Segoe UI", 10)).pack(pady=(20, 5))
        entry_user = tk.Entry(dialog, font=("Segoe UI", 10))
        entry_user.pack()
        entry_user.focus()

        tk.Label(dialog, text="Senha:", font=("Segoe UI", 10)).pack(pady=(10, 5))
        entry_pass = tk.Entry(dialog, font=("Segoe UI", 10), show="*")
        entry_pass.pack()

        def confirmar():
            u = entry_user.get().strip()
            p = entry_pass.get().strip()
            if u and p:
                credenciais["user"] = u
                credenciais["pass"] = p
                dialog.destroy()
            else:
                messagebox.showwarning("Aviso", "Preencha usuário e senha!", parent=dialog)

        tk.Button(dialog, text="CONFIRMAR", command=confirmar, bg=COR_PRIMARIA, fg="white", font=("Segoe UI", 9, "bold")).pack(pady=20)
        
        self.wait_window(dialog) 
        return credenciais["user"], credenciais["pass"]

    def iniciar_automacao(self):
        texto = self.entrada_texto.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Cole a lista de dados antes de iniciar.")
            return

        user, pwd = self.solicitar_credenciais()
        if not user or not pwd:
            return 

        modo = self.modo_var.get()
        linhas = [l.strip() for l in texto.split("\n") if l.strip()]
        
        lista_final = []
        if modo == "CPF":
            for l in linhas:
                apenas_nums = re.sub(r'\D', '', l)
                if apenas_nums: lista_final.append(apenas_nums.zfill(11))
        else:
            lista_final = linhas

        lista_final = list(dict.fromkeys(lista_final))

        if not lista_final:
            messagebox.showerror("Erro", "Nenhum dado válido encontrado para processar.")
            return

        self.btn_iniciar.config(state="disabled", text="Processando...", bg="#9E9E9E")
        
        t = threading.Thread(target=run_selenium_logic, args=(modo, lista_final, user, pwd))
        t.daemon = True
        t.start()

    def verificar_fila(self):
        try:
            msg = fila_mensagens.get_nowait()
            if msg == "FIM":
                self.finalizar_processo()
            elif msg == "FIM_ERRO":
                self.btn_iniciar.config(state="normal", text="Iniciar navegação automatica", bg=COR_PRIMARIA)
        except queue.Empty: pass
        finally: self.after(1000, self.verificar_fila)

    def finalizar_processo(self):
        self.btn_iniciar.config(state="normal", text="Iniciar navegação automatica", bg=COR_PRIMARIA)
        
        if not os.path.isfile(ARQUIVO_TEMP):
            messagebox.showwarning("Finalizado", "Processo terminou, mas nenhum dado foi coletado.")
            return

        arquivo_destino = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Salvar Relatório Final Como",
            initialfile=f"Relatorio_{self.modo_var.get()}_{datetime.now().strftime('%H-%M')}.xlsx"
        )

        if arquivo_destino:
            gerar_excel_final(arquivo_destino)
            messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{arquivo_destino}")
            try: os.remove(ARQUIVO_TEMP)
            except: pass

if __name__ == "__main__":
    app = App()
    app.mainloop()