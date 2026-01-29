"""
Automação Gestão Bolsa - Versão Final "Blindagem Global (Zero Erros)"
(Correção: Desativa o erro de 'Número como Texto' na planilha inteira de uma só vez)
"""

# ==============================================================================
# 1. IMPORTAÇÕES
# ==============================================================================
import pandas as pd 
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import (
    TimeoutException, 
    WebDriverException, 
    StaleElementReferenceException,
    NoSuchElementException
)
from chromedriver_py import binary_path 
import time
from datetime import datetime
import sys
import xlsxwriter 
from xlsxwriter.utility import xl_col_to_name
import unicodedata
import re
import csv
import os

# ==============================================================================
# CONFIGURAÇÕES GERAIS
# ==============================================================================
ARQUIVO_TEMP = "temp_dados_intuitivos.csv" 
URL_LOGIN = "https://10.237.1.11/bolsa/login/login"
URL_BUSCA = "https://10.237.1.11/bolsa/index/index" 
USUARIO = "ihan.santos"
SENHA = "Mavis08"

# ==============================================================================
# FUNÇÕES AUXILIARES DE SEGURANÇA
# ==============================================================================

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

def resetar_navegacao(driver, wait):
    try:
        driver.implicitly_wait(0.5)
        if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0:
            return False 
        if len(driver.find_elements(By.XPATH, '//*[@id="campo"]')) > 0:
            driver.implicitly_wait(0.5)
            return True
        driver.implicitly_wait(0.5)
        driver.back()
        try:
            if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0: return False
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="campo"]')))
            time.sleep(0.5)
            return True
        except: pass
        driver.back()
        try:
            if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0: return False
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="campo"]')))
            time.sleep(0.5)
            return True
        except: pass
        driver.get(URL_BUSCA)
        if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0:
            return False
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="campo"]')))
        time.sleep(0.5)
        return True
    except Exception:
        return False

def garantir_sessao(driver, wait):
    try:
        if len(driver.find_elements(By.XPATH, '//*[@id="campo"]')) > 0:
            return True
        print(" > [SISTEMA] Sessão perdida ou Tela de Login detectada. Realizando Re-login...")
        if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) == 0:
            driver.get(URL_LOGIN)
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="usuario"]'))).clear()
        driver.find_element(By.XPATH, '//*[@id="usuario"]').send_keys(USUARIO)
        driver.find_element(By.XPATH, '//*[@id="senha"]').clear()
        driver.find_element(By.XPATH, '//*[@id="senha"]').send_keys(SENHA)
        safe_click(driver, '//*[@id="conteudo"]/form/fieldset/p[3]/input', wait)
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="campo"]')))
        print(" > [SISTEMA] Re-login realizado com sucesso!")
        return True
    except Exception as e:
        if "10061" in str(e) or "refused" in str(e).lower() or "closed" in str(e).lower():
            raise KeyboardInterrupt("Conexão perdida (CTRL+C detectado)")
        print(f" > [ERRO CRÍTICO] Falha ao relogar: {e}")
        return False

def salvar_checkpoint(dados):
    existe = os.path.isfile(ARQUIVO_TEMP)
    colunas = [
        "INSCRIÇÃO", "NOME", "CPF", "CURSO", "FACULDADE", "SITUAÇÃO", "PERIODO", 
        "TIPO DE BOLSA", "DATA COLETA", "VALOR S/ DESCONTO", "VALOR C/ DESCONTO",
        "VALOR BENEFÍCIOS", "INFO BENEFÍCIOS", 
        "DATA LANÇAMENTO", "VALOR DA BOLSA", "VALOR CALCULADO", "DIFERENÇA", "STATUS"
    ]
    try:
        with open(ARQUIVO_TEMP, mode='a', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=colunas, delimiter=';')
            if not existe: writer.writeheader()
            writer.writerow(dados)
    except Exception as e:
        print(f"Erro ao salvar backup: {e}")

# ==============================================================================
# 2. CARREGAMENTO INTELIGENTE
# ==============================================================================
print("="*60)
print("INICIANDO AUTOMAÇÃO - AUTO-RECOVERY ATIVADO")
print("="*60)

lista_total = []
lista_processados = set()

try:
    with open('inscricao.txt', 'r', encoding='utf-8') as f:
        lista_total = [
            linha.strip()
            for linha in f
            if linha.strip()
        ]
except Exception as e:
    print(f"ERRO AO LER inscricao.txt: {e}")
    sys.exit()

if os.path.isfile(ARQUIVO_TEMP):
    try:
        df_temp = pd.read_csv(ARQUIVO_TEMP, sep=';', usecols=['INSCRIÇÃO'], encoding='utf-8-sig')
        lista_processados = set(df_temp['INSCRIÇÃO'].astype(str).unique())
        print(f"> Histórico encontrado: {len(lista_processados)} alunos já processados.")
    except: pass

lista_para_fazer = [aluno for aluno in lista_total if aluno not in lista_processados]

print(f"> Total Original: {len(lista_total)}")
print(f"> Restantes na Fila: {len(lista_para_fazer)}")

if len(lista_para_fazer) == 0:
    print("\n[!] TODOS OS ALUNOS JÁ FORAM PROCESSADOS!")
else:
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument('--ignore-certificate-errors') 
    options.add_experimental_option("detach", True)
    service = Service(binary_path)

    driver = None 
    try:
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 10) 
        print("Realizando Login Inicial...")
        driver.get(URL_LOGIN)
        garantir_sessao(driver, wait)

        # ==============================================================================
        # 3. PROCESSAMENTO
        # ==============================================================================
        IDX_COLETA_SITUACAO = 0       
        IDX_COLETA_MENSAL_SEM = 4     
        IDX_COLETA_MENSAL_COM = 5
        IDX_COLETA_VALOR_BENEFICIO = 6 
        IDX_COLETA_INFO_BENEFICIO = 7  
        IDX_COLETA_DATA = 10          
        IDX_LANC_DATA = 2             
        IDX_LANC_VALOR_BOLSA = 4      
        IDX_LANC_TIPO_BOLSA = 9

        print("-" * 60)
        stop_requested = False
        consecutive_failures = 0 

        try: 
            for i, inscricao in enumerate(lista_para_fazer):
                if stop_requested: break 
                if consecutive_failures >= 3:
                    print(f"\n[ALERTA] {consecutive_failures} falhas consecutivas detectadas!")
                    print(" > Sessão provavelmente travada. Forçando Refresh...")
                    try:
                        driver.delete_all_cookies()
                        driver.get(URL_LOGIN)
                        garantir_sessao(driver, wait)
                        consecutive_failures = 0 
                    except Exception as e:
                        print(f"Erro ao tentar recuperar sessão: {e}")

                if not resetar_navegacao(driver, wait):
                    garantir_sessao(driver, wait)
                
                try:
                    # BUSCA 
                    sucesso_busca = False
                    for tentativa in range(3):
                        try:
                            if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0:
                                raise Exception("Sessão caiu")
                            campo = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="campo"]')))
                            campo.clear()
                            campo.send_keys(inscricao)
                            safe_click(driver, '//*[@id="conteudo"]/div[2]/form/fieldset/input[2]', wait)
                            sucesso_busca = True
                            break
                        except KeyboardInterrupt: raise 
                        except Exception as e:
                            time.sleep(1)
                            if "Sessão caiu" in str(e): raise e 
                            if tentativa == 2: raise e
                            continue
                    
                    if not sucesso_busca: raise Exception("Falha na busca")

                    # CAPA DO ALUNO
                    try:
                        base_xp = '//*[@id="conteudo"]/div[2]/fieldset/table/tbody/tr[2]'
                        wait.until(EC.presence_of_element_located((By.XPATH, base_xp)))
                        
                        nome_univ = safe_get_text(driver, f'{base_xp}/td[2]')
                        cpf_raw = safe_get_text(driver, f'{base_xp}/td[3]')
                        cpf_univ = ''.join(filter(str.isdigit, cpf_raw))
                        curso_univ = safe_get_text(driver, f'{base_xp}/td[5]')
                        ies = safe_get_text(driver, f'{base_xp}/td[4]')
                        ies = unicodedata.normalize('NFKD', ies).upper().encode('ascii', errors='ignore').decode('utf-8')
                        ies = re.sub(r'[-.,]', ' ', ies).strip()
                        tipo_bolsa_capa = safe_get_text(driver, f'{base_xp}/td[6]')
                        situacao = safe_get_text(driver, f'{base_xp}/td[11]')
                        
                        safe_click(driver, f'{base_xp}/td[14]/a[1]', wait)
                        wait.until(EC.presence_of_element_located((By.XPATH, "//legend[contains(., 'Coleta de Dados')]")))
                        consecutive_failures = 0 

                    except TimeoutException:
                        print(f"[{i+1}/{len(lista_para_fazer)}] {inscricao}: [ERRO] Aluno não encontrado")
                        salvar_checkpoint({"INSCRIÇÃO": inscricao, "STATUS": "ALUNO NÃO ENCONTRADO"})
                        consecutive_failures += 1
                        resetar_navegacao(driver, wait)
                        continue
                    
                    # EXTRAÇÃO E CÁLCULOS 
                    mapa_financeiro = {} 
                    try:
                        tbl_lanc = driver.find_element(By.XPATH, "//legend[contains(., 'Lançamento(s) de Pagto(s)')]/following-sibling::table")
                        rows = tbl_lanc.find_elements(By.XPATH, ".//tbody/tr")
                        for idx_r in range(len(rows)):
                            try:
                                tbl_ref = driver.find_element(By.XPATH, "//legend[contains(., 'Lançamento(s) de Pagto(s)')]/following-sibling::table")
                                row_ref = tbl_ref.find_elements(By.XPATH, ".//tbody/tr")[idx_r]
                                cols = row_ref.find_elements(By.TAG_NAME, "td")
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

                    coletas = {}
                    try:
                        tbl_coleta = driver.find_element(By.XPATH, "//legend[contains(., 'Coleta de Dados')]/following-sibling::table")
                        rows = tbl_coleta.find_elements(By.XPATH, ".//tbody/tr")
                        for idx_r in range(len(rows)):
                            try:
                                tbl_ref = driver.find_element(By.XPATH, "//legend[contains(., 'Coleta de Dados')]/following-sibling::table")
                                row_ref = tbl_ref.find_elements(By.XPATH, ".//tbody/tr")[idx_r]
                                cols = row_ref.find_elements(By.TAG_NAME, "td")
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

                    dados_encontrados = False
                    log_divergencias = []
                    total_regular = 0
                    
                    for sem in sorted(coletas.keys()):
                        dados_col = coletas[sem]
                        dados_fin = mapa_financeiro.get(sem, {'valor': 0.0, 'dt_str': '', 'tipo_bolsa_real': ''})
                        tipo_bolsa_final = dados_fin['tipo_bolsa_real'] if dados_fin['tipo_bolsa_real'] else tipo_bolsa_capa
                        
                        curso_upper = curso_univ.upper().strip()
                        curso_saude = curso_upper in ["MEDICINA", "ODONTOLOGIA"] 
                        
                        bolsa_calc = 0.0
                        if "PARCIAL" in tipo_bolsa_final.upper():
                            base = dados_col['val_c_float'] * 0.5
                            teto = 2900.00 if curso_saude else 650.00
                            bolsa_calc = min(base, teto)
                        elif "INTEGRAL" in tipo_bolsa_final.upper():
                            base = dados_col['val_c_float']
                            teto = 5800.00 if curso_saude else 1500.00
                            bolsa_calc = min(base, teto)
                        
                        valor_beneficios = dados_col['val_beneficio']
                        mensalidade_com_desc = dados_col['val_c_float']
                        if (bolsa_calc + valor_beneficios) > mensalidade_com_desc:
                                bolsa_calc = mensalidade_com_desc - valor_beneficios

                        eh_desligado = ("Desligado" in situacao) or ("Abandonou" in situacao)
                        valor_pago = dados_fin['valor']
                        diff = abs(valor_pago - bolsa_calc)
                        
                        status = ""
                        if valor_pago == 0.0:
                            if bolsa_calc == 0.0: status = "REGULAR"
                            elif eh_desligado: status = "CANCELADO (DESLIGADO)"
                            else: status = "PENDENTE DE PAGAMENTO"
                        else:
                            if bolsa_calc == 0.0: status = "PAGAMENTO INDEVIDO (SEM COBRANÇA)"
                            elif diff < 1.00: status = "REGULAR"
                            elif valor_pago > bolsa_calc: status = "DIVERGÊNCIA (PAGO > CALC)"
                            else: status = "DIVERGÊNCIA (PAGO < CALC)"

                        saldo = dados_col['val_s_float'] - valor_pago
                        
                        linha_excel = {
                            "INSCRIÇÃO": inscricao, "NOME": nome_univ, "CPF": cpf_univ, "CURSO": curso_univ,
                            "FACULDADE": ies, "SITUAÇÃO": situacao, "PERIODO": sem, "TIPO DE BOLSA": tipo_bolsa_final,
                            "DATA COLETA": dados_col['dt_str'],
                            "VALOR S/ DESCONTO": dados_col['val_s_float'],
                            "VALOR C/ DESCONTO": dados_col['val_c_float'],
                            "VALOR BENEFÍCIOS": dados_col['val_beneficio'], 
                            "INFO BENEFÍCIOS": dados_col['info_beneficio'], 
                            "DATA LANÇAMENTO": dados_fin['dt_str'],
                            "VALOR DA BOLSA": valor_pago,
                            "VALOR CALCULADO": bolsa_calc,
                            "DIFERENÇA": saldo, "STATUS": status
                        }
                        salvar_checkpoint(linha_excel)
                        dados_encontrados = True
                        
                        if status == "REGULAR" or status == "CANCELADO (DESLIGADO)":
                            total_regular += 1
                        else:
                            log_divergencias.append(f"   ! {sem}: {status} (Pago: {valor_pago:.2f} vs Calc: {bolsa_calc:.2f})")
                    
                    header_log = f"[{i+1}/{len(lista_para_fazer)}] {inscricao}:"
                    if not dados_encontrados: 
                        print(f"{header_log} [AVISO] Sem dados cruzados.")
                        salvar_checkpoint({"INSCRIÇÃO": inscricao, "STATUS": "SEM DADOS CRUZADOS"})
                    else:
                        if len(log_divergencias) == 0:
                            print(f"{header_log} [OK] {total_regular} Semestres verificados (Todos REGULARES/CORRETOS)")
                        else:
                            print(f"{header_log} [ATENÇÃO] Encontradas {len(log_divergencias)} divergências:")
                            for log in log_divergencias:
                                print(log)

                except KeyboardInterrupt: raise 
                except Exception as e:
                    erro_str = str(e).lower()
                    if "10061" in erro_str or "refused" in erro_str or "closed" in erro_str:
                        print("\n!!! DETECTADO CORTE DE CONEXÃO (CTRL+C) !!!")
                        stop_requested = True
                        break
                    print(f"[{i+1}/{len(lista_para_fazer)}] {inscricao}: [ERRO DE CODE] {e}")
                    consecutive_failures += 1
                    resetar_navegacao(driver, wait)

        except KeyboardInterrupt: print("\n\n!!! PARADA SOLICITADA PELO USUÁRIO (CTRL+C) !!!")
    except Exception as e: print(f"Erro Fatal: {e}")
    finally:
        if driver:
            print("\nEncerrando Driver...")
            try: driver.quit()
            except: pass

# ==============================================================================
# 5. EXPORTAÇÃO EXCEL (VERSÃO NUCLEAR GLOBAL)
# ==============================================================================
print("\n" + "="*60)

if os.path.isfile(ARQUIVO_TEMP):
    try:
        # Encoding utf-8-sig para limpar caracteres estranhos
        df_result = pd.read_csv(ARQUIVO_TEMP, sep=';', dtype=str, encoding='utf-8-sig')
        
        if 'NOME' in df_result.columns: df_result = df_result.dropna(subset=['NOME'])
        
        cols_numericas = [
            "VALOR S/ DESCONTO", "VALOR C/ DESCONTO", 
            "VALOR BENEFÍCIOS", "VALOR DA BOLSA", 
            "VALOR CALCULADO", "DIFERENÇA"
        ]
        
        # Converte valores financeiros para float
        for col in cols_numericas:
            if col in df_result.columns:
                df_result[col] = df_result[col].fillna('0')
                df_result[col] = pd.to_numeric(df_result[col], errors='coerce').fillna(0.0)

        arquivo_saida = f"Relatorio_Final_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx"
        print(f"Gerando Excel: {arquivo_saida} ...")
        
        with pd.ExcelWriter(arquivo_saida, engine='xlsxwriter') as writer:
            df_result.to_excel(writer, sheet_name="Auditoria", index=False, startrow=1, header=False)
            
            workbook  = writer.book
            worksheet = writer.sheets["Auditoria"]
            
            # --- BLINDAGEM SUPREMA ---
            # Ignora o erro "Número armazenado como texto" em TODA A PLANILHA (A1 até o fim)
            # Isso impede que o triangulo verde apareça em qualquer lugar, independente da coluna.
            worksheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
            # -------------------------
            
            # Formatos
            fmt_header = workbook.add_format({'bold': True, 'fg_color': '#000000', 'font_color': '#FFFFFF'})
            fmt_currency = workbook.add_format({'num_format': '#,##0.00'}) 
            fmt_center = workbook.add_format({'align': 'center'})
            fmt_text = workbook.add_format({'num_format': '@'}) # Força texto
            fmt_vermelho = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) 
            fmt_amarelo = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'}) 
            fmt_laranja = workbook.add_format({'bg_color': '#FFD7B5', 'font_color': '#833C0C'})
            
            # Cabeçalho
            for col_num, value in enumerate(df_result.columns.values):
                worksheet.write(0, col_num, value, fmt_header)
            
            last_row = len(df_result) + 1
            
            # Aplicação de Formatação Coluna por Coluna
            for i, col in enumerate(df_result.columns):
                col_upper = str(col).strip().upper()
                largura = len(col_upper) + 4
                
                # Para ID (CPF/Inscrição), mantemos o formato TEXTO (@) para segurar os Zeros
                if "INSCRI" in col_upper or "CPF" in col_upper:
                    worksheet.set_column(i, i, largura, fmt_text)
                
                elif col in cols_numericas:
                    worksheet.set_column(i, i, 18, fmt_currency)
                elif col == "STATUS": 
                    worksheet.set_column(i, i, 35, fmt_center) 
                elif "NOME" in col: 
                    worksheet.set_column(i, i, 40)
                else: 
                    worksheet.set_column(i, i, largura)
            
            # Formatação Condicional
            rng = f'B2:{xl_col_to_name(len(df_result.columns) - 1)}{last_row}'
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'PAGAMENTO INDEVIDO', 'format': fmt_vermelho})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'PAGO > CALC', 'format': fmt_vermelho})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'PENDENTE', 'format': fmt_amarelo})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'PAGO < CALC', 'format': fmt_laranja})

            # Legenda
            dados_legenda = [
                {"STATUS": "REGULAR", "DESCRIÇÃO": "Tudo certo. Diferença < 1.00"},
                {"STATUS": "PENDENTE DE PAGAMENTO", "DESCRIÇÃO": "Ativo com bolsa calculada, mas sem pagamento."},
                {"STATUS": "CANCELADO (DESLIGADO)", "DESCRIÇÃO": "Desligado e sem pagamento (Correto)."},
                {"STATUS": "PAGAMENTO INDEVIDO (SEM COBRANÇA)", "DESCRIÇÃO": "Pagou, mas cálculo deu R$ 0,00."},
                {"STATUS": "DIVERGÊNCIA (PAGO > CALC)", "DESCRIÇÃO": "Pagou a MAIS que o devido."},
                {"STATUS": "DIVERGÊNCIA (PAGO < CALC)", "DESCRIÇÃO": "Pagou a MENOS que o devido."},
                {"STATUS": "ALUNO NÃO ENCONTRADO", "DESCRIÇÃO": "Inscrição não localizada."},
                {"STATUS": "SEM DADOS CRUZADOS", "DESCRIÇÃO": "Datas não batem."}
            ]
            df_legenda = pd.DataFrame(dados_legenda)
            df_legenda.to_excel(writer, sheet_name="Legenda", index=False, startrow=1, header=False)
            ws_legenda = writer.sheets["Legenda"]
            ws_legenda.set_column(0, 0, 35, fmt_center)
            ws_legenda.set_column(1, 1, 80)
            ws_legenda.write(0, 0, "STATUS", fmt_header)
            ws_legenda.write(0, 1, "DESCRIÇÃO", fmt_header)

        print(f"SUCESSO! Arquivo salvo: {arquivo_saida}")
        
    except KeyboardInterrupt: print("\n!!! Exportação CANCELADA !!! (CSV mantido)")
    except Exception as e: print(f"Erro na geração do Excel: {e}")
else:
    print("Nenhum dado capturado.")