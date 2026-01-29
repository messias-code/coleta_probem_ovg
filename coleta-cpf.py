"""
Automação Gestão Bolsa - Versão "Blindagem Global"
(MODIFICADO: Suporte a Multiplas Capas/Vínculos por CPF)
"""

# ==============================================================================
# 1. IMPORTAÇÕES
# ==============================================================================
import pandas as pd 
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
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
ARQUIVO_TEMP = "temp_dados_audit.csv" 
URL_LOGIN = "https://10.237.1.11/bolsa/login/login"
URL_BUSCA = "https://10.237.1.11/bolsa/index/index" 
USUARIO = "ihan.santos"
SENHA = "Mavis08"

# ==============================================================================
# FUNÇÕES AUXILIARES
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
        print(" > [SISTEMA] Sessão perdida. Realizando Re-login...")
        if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) == 0:
            driver.get(URL_LOGIN)
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="usuario"]'))).clear()
        driver.find_element(By.XPATH, '//*[@id="usuario"]').send_keys(USUARIO)
        driver.find_element(By.XPATH, '//*[@id="senha"]').clear()
        driver.find_element(By.XPATH, '//*[@id="senha"]').send_keys(SENHA)
        safe_click(driver, '//*[@id="conteudo"]/form/fieldset/p[3]/input', wait)
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="campo"]')))
        return True
    except Exception as e:
        if "10061" in str(e) or "refused" in str(e).lower() or "closed" in str(e).lower():
            raise KeyboardInterrupt("Conexão perdida")
        print(f" > [ERRO CRÍTICO] Falha ao relogar: {e}")
        return False

def salvar_checkpoint(dados):
    existe = os.path.isfile(ARQUIVO_TEMP)
    colunas = [
        "CPF", "INSCRIÇÃO", "NOME", "CURSO", "FACULDADE", "SITUAÇÃO", "PERIODO", 
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
# 2. CARREGAMENTO E EXECUÇÃO
# ==============================================================================
print("="*60)
print("INICIANDO AUTOMAÇÃO - MODO MULTI-VÍNCULOS")
print("="*60)

lista_total = []
lista_processados = set()

# --- LER CPF.TXT COM LIMPEZA E FORMATAÇÃO ---
try:
    with open('cpf.txt', 'r', encoding='utf-8') as f:
        for linha in f:
            apenas_numeros = re.sub(r'\D', '', linha)
            if apenas_numeros:
                cpf_formatado = apenas_numeros.zfill(11)
                lista_total.append(cpf_formatado)
    lista_total = list(dict.fromkeys(lista_total))
except FileNotFoundError:
    print("ERRO: 'cpf.txt' não encontrado.")
    sys.exit()

if os.path.isfile(ARQUIVO_TEMP):
    try:
        df_temp = pd.read_csv(ARQUIVO_TEMP, sep=';', usecols=['CPF'], encoding='utf-8-sig')
        lista_processados = set(df_temp['CPF'].astype(str).str.zfill(11).unique())
        print(f"> Histórico: {len(lista_processados)} CPFs já processados.")
    except: pass

lista_para_fazer = [item for item in lista_total if item not in lista_processados]

print(f"> Total CPFs Carregados: {len(lista_total)}")
print(f"> Restantes na Fila: {len(lista_para_fazer)}")

if len(lista_para_fazer) == 0:
    print("\n[!] TODOS OS CPFs JÁ FORAM PROCESSADOS!")
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
        driver.get(URL_LOGIN)
        garantir_sessao(driver, wait)

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

        for i, cpf_input in enumerate(lista_para_fazer):
            if stop_requested: break 
            
            if consecutive_failures >= 3:
                print(" > Resetando Sessão...")
                driver.delete_all_cookies()
                driver.get(URL_LOGIN)
                garantir_sessao(driver, wait)
                consecutive_failures = 0 

            if not resetar_navegacao(driver, wait):
                garantir_sessao(driver, wait)
            
            try:
                # --- BUSCA POR CPF ---
                sucesso_busca = False
                for tentativa in range(3):
                    try:
                        if len(driver.find_elements(By.XPATH, '//*[@id="usuario"]')) > 0:
                            raise Exception("Sessão caiu")
                        
                        select_element = wait.until(EC.element_to_be_clickable((By.NAME, 'opcao')))
                        Select(select_element).select_by_value('uni_cpf')
                        
                        campo = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="campo"]')))
                        campo.clear()
                        campo.send_keys(cpf_input)
                        
                        safe_click(driver, '//*[@id="conteudo"]/div[2]/form/fieldset/input[2]', wait)
                        sucesso_busca = True
                        break
                    except Exception as e:
                        time.sleep(1)
                        if tentativa == 2: raise e
                        continue
                
                if not sucesso_busca: raise Exception("Falha na busca")

                # --- LOOP DE RESULTADOS (CAPAS MÚLTIPLAS) ---
                try:
                    # XPath base da tabela
                    table_body_xp = '//*[@id="conteudo"]/div[2]/fieldset/table/tbody'
                    wait.until(EC.presence_of_element_located((By.XPATH, table_body_xp)))
                    
                    # Conta quantas linhas existem (ignorando cabeçalho se houver)
                    # Assumimos que a primeira TR é cabeçalho, então dados começam no index 2 (XPath)
                    all_rows = driver.find_elements(By.XPATH, f'{table_body_xp}/tr')
                    total_registros = len(all_rows) - 1 # Remove 1 do cabeçalho
                    
                    if total_registros <= 0:
                         # Caso raro onde tabela existe mas sem linhas de dados
                         raise TimeoutException("Sem linhas de dados")

                    # Loop por cada registro encontrado para este CPF
                    for idx_reg in range(total_registros):
                        
                        # Se não for o primeiro registro, precisamos voltar para a lista
                        if idx_reg > 0:
                            driver.back()
                            wait.until(EC.presence_of_element_located((By.XPATH, table_body_xp)))
                        
                        # Calcula o XPath da linha atual (2 é a primeira linha de dados)
                        current_row_idx = idx_reg + 2
                        base_xp = f'{table_body_xp}/tr[{current_row_idx}]'
                        
                        # Extrai dados da Capa (Linha da tabela)
                        # Nota: td[1] contém a inscrição (texto ou h2)
                        inscricao_real = safe_get_text(driver, f'{base_xp}/td[1]')
                        nome_univ = safe_get_text(driver, f'{base_xp}/td[2]')
                        curso_univ = safe_get_text(driver, f'{base_xp}/td[5]')
                        ies = safe_get_text(driver, f'{base_xp}/td[4]')
                        ies = unicodedata.normalize('NFKD', ies).upper().encode('ascii', errors='ignore').decode('utf-8')
                        ies = re.sub(r'[-.,]', ' ', ies).strip()
                        tipo_bolsa_capa = safe_get_text(driver, f'{base_xp}/td[6]')
                        situacao = safe_get_text(driver, f'{base_xp}/td[11]')
                        
                        # Clica no detalhe (Última coluna Ações)
                        safe_click(driver, f'{base_xp}/td[14]/a[1]', wait)
                        wait.until(EC.presence_of_element_located((By.XPATH, "//legend[contains(., 'Coleta de Dados')]")))
                        
                        # --- EXTRAÇÃO DE DETALHES (Mantida igual) ---
                        mapa_financeiro = {} 
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

                        coletas = {}
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

                        # --- CÁLCULO E SALVAMENTO ---
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
                                "CPF": cpf_input, "INSCRIÇÃO": inscricao_real, "NOME": nome_univ,
                                "CURSO": curso_univ, "FACULDADE": ies, "SITUAÇÃO": situacao, "PERIODO": sem, 
                                "TIPO DE BOLSA": tipo_bolsa_final, "DATA COLETA": dados_col['dt_str'],
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
                                log_divergencias.append(f"   ! {sem}: {status}")
                        
                        header_log = f"[{i+1}/{len(lista_para_fazer)}] CPF {cpf_input} (Reg {idx_reg+1}/{total_registros}):"
                        if not dados_encontrados: 
                            print(f"{header_log} [AVISO] Sem dados cruzados.")
                            salvar_checkpoint({"CPF": cpf_input, "INSCRIÇÃO": inscricao_real, "STATUS": "SEM DADOS CRUZADOS"})
                        else:
                            if len(log_divergencias) == 0:
                                print(f"{header_log} [OK] {total_regular} Semestres.")
                            else:
                                print(f"{header_log} [ATENÇÃO] {len(log_divergencias)} divergências.")
                                for log in log_divergencias: print(log)
                    
                    # Fim do loop de registros para este CPF
                    consecutive_failures = 0

                except TimeoutException:
                    print(f"[{i+1}/{len(lista_para_fazer)}] CPF {cpf_input}: [AVISO] Não encontrado")
                    salvar_checkpoint({"CPF": cpf_input, "STATUS": "CPF NÃO ENCONTRADO"})
                    consecutive_failures += 1
                    resetar_navegacao(driver, wait)
                    continue

            except KeyboardInterrupt: raise 
            except Exception as e:
                if "10061" in str(e):
                    print("\n!!! CORTE DE CONEXÃO DETECTADO !!!")
                    stop_requested = True
                    break
                print(f"[{i+1}/{len(lista_para_fazer)}] CPF {cpf_input}: [ERRO] {e}")
                consecutive_failures += 1
                resetar_navegacao(driver, wait)

    except KeyboardInterrupt: print("\n\n!!! PARADA SOLICITADA !!!")
    except Exception as e: print(f"Erro Fatal: {e}")
    finally:
        if driver:
            try: driver.quit()
            except: pass

# ==============================================================================
# 5. EXPORTAÇÃO EXCEL (ATUALIZADA)
# ==============================================================================
print("\n" + "="*60)

if os.path.isfile(ARQUIVO_TEMP):
    try:
        df_result = pd.read_csv(ARQUIVO_TEMP, sep=';', dtype=str, encoding='utf-8-sig')
        
        if 'NOME' in df_result.columns: df_result = df_result.dropna(subset=['NOME'])
        
        cols_numericas = [
            "VALOR S/ DESCONTO", "VALOR C/ DESCONTO", 
            "VALOR BENEFÍCIOS", "VALOR DA BOLSA", 
            "VALOR CALCULADO", "DIFERENÇA"
        ]
        
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
            worksheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
            
            fmt_header = workbook.add_format({'bold': True, 'fg_color': '#000000', 'font_color': '#FFFFFF'})
            fmt_currency = workbook.add_format({'num_format': '#,##0.00'}) 
            fmt_center = workbook.add_format({'align': 'center'})
            fmt_text = workbook.add_format({'num_format': '@'}) 
            fmt_vermelho = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) 
            fmt_amarelo = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'}) 
            fmt_laranja = workbook.add_format({'bg_color': '#FFD7B5', 'font_color': '#833C0C'})
            
            for col_num, value in enumerate(df_result.columns.values):
                worksheet.write(0, col_num, value, fmt_header)
            
            last_row = len(df_result) + 1
            
            for i, col in enumerate(df_result.columns):
                col_upper = str(col).strip().upper()
                largura = len(col_upper) + 4
                
                if "INSCRI" in col_upper or "CPF" in col_upper:
                    worksheet.set_column(i, i, largura + 5, fmt_text)
                elif col in cols_numericas:
                    worksheet.set_column(i, i, 18, fmt_currency)
                elif col == "STATUS": 
                    worksheet.set_column(i, i, 35, fmt_center) 
                elif "NOME" in col: 
                    worksheet.set_column(i, i, 40)
                else: 
                    worksheet.set_column(i, i, largura)
            
            rng = f'B2:{xl_col_to_name(len(df_result.columns) - 1)}{last_row}'
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'PAGAMENTO INDEVIDO', 'format': fmt_vermelho})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'PAGO > CALC', 'format': fmt_vermelho})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'PENDENTE', 'format': fmt_amarelo})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'PAGO < CALC', 'format': fmt_laranja})

            dados_legenda = [
                {"STATUS": "REGULAR", "DESCRIÇÃO": "Tudo certo. Diferença < 1.00"},
                {"STATUS": "PENDENTE DE PAGAMENTO", "DESCRIÇÃO": "Ativo com bolsa calculada, mas sem pagamento."},
                {"STATUS": "CANCELADO (DESLIGADO)", "DESCRIÇÃO": "Desligado e sem pagamento (Correto)."},
                {"STATUS": "PAGAMENTO INDEVIDO (SEM COBRANÇA)", "DESCRIÇÃO": "Pagou, mas cálculo deu R$ 0,00."},
                {"STATUS": "DIVERGÊNCIA (PAGO > CALC)", "DESCRIÇÃO": "Pagou a MAIS que o devido."},
                {"STATUS": "DIVERGÊNCIA (PAGO < CALC)", "DESCRIÇÃO": "Pagou a MENOS que o devido."},
                {"STATUS": "CPF NÃO ENCONTRADO", "DESCRIÇÃO": "CPF não localizado na busca."},
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