import tkinter as tk  # Biblioteca que cria a janela, botões e textos na tela
from tkinter import filedialog, messagebox  # Ferramentas para abrir janelas de "Salvar/Abrir" e avisos pop-up
import pandas as pd  # A biblioteca principal (Pandas). É ela que lê e mexe no Excel.
import os  # Biblioteca que conversa com o Windows (para checar se arquivos existem)

# ==============================================================================
# FUNÇÕES DO PROGRAMA (O CÉREBRO)
# ==============================================================================

def selecionar_arquivo_entrada():
    """
    O que isso faz: Abre aquela janela do Windows para você clicar no arquivo Excel.
    """
    # 1. Abre a janelinha de "Abrir Arquivo"
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel para análise",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    
    # 2. Se você escolheu algo (não cancelou):
    if arquivo:
        # Limpa o campo de texto (caso tenha algo escrito antes)
        entry_arquivo_entrada.delete(0, tk.END)
        # Escreve o caminho do arquivo (ex: C:/MeusDocs/Planilha.xlsx) na caixinha branca
        entry_arquivo_entrada.insert(0, arquivo)

def processar_e_salvar():
    """
    O que isso faz: É o motor principal. Lê, filtra, limpa e cria o Excel novo bonitão.
    """
    
    # --- 1. SEGURANÇA (Verifica se tem arquivo) ---
    caminho_entrada = entry_arquivo_entrada.get() # Pega o texto da caixinha
    
    # Se estiver vazio ou o arquivo não existir no computador:
    if not caminho_entrada or not os.path.exists(caminho_entrada):
        messagebox.showwarning("Atenção", "Por favor, selecione um arquivo de entrada válido primeiro.")
        return # Para tudo e volta

    try:
        # --- 2. LEITURA E FILTROS (Mexe nos dados) ---
        
        # Carrega o Excel para a memória do Python (chamamos isso de DataFrame ou df_)
        df_ = pd.read_excel(caminho_entrada)
        
        # Define qual é a frase exata que queremos manter
        tipo_desejado = "CONTRATO DE PRESTAÇÃO DE SERVIÇOS EDUCACIONAIS OU COMPROVANTE DE MATRÍCULA"
        
        # Manda o Python jogar fora tudo que NÃO for o tipo acima
        df_ = df_[df_['Documento Tipo'] == tipo_desejado]
        
        # Se existir a coluna 'Status Obs', joga ela fora (drop)
        if 'Status Obs' in df_.columns:
            df_ = df_.drop(columns=['Status Obs'])
            
        # Limpeza pesada no nome da Faculdade (Tirar acento, deixar maiúsculo, tirar traço)
        df_['Faculdade'] = (
            df_['Faculdade']
            .str.upper() # Tudo MAIÚSCULO
            .str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8') # Remove acentos (ã -> a)
            .str.replace(r'[-.,]', ' ', regex=True) # Troca traço e ponto por espaço
            .str.replace(r'\s+', ' ', regex=True)   # Remove espaços duplos
            .str.strip() # Remove espaço sobrando no começo e fim
        )

        # Arruma CPF e Inscrição (Coloca os zeros à esquerda de volta)
        # Ex: Se estava 123, vira 00000000123
        df_['CPF'] = df_['CPF'].astype(str).str.zfill(11)
        df_['Inscrição'] = df_['Inscrição'].astype(str).str.zfill(7)
        
        # Organiza a lista alfabeticamente (Faculdade primeiro, depois Nome do aluno)
        df_ = df_.sort_values(by=['Faculdade', 'Bolsista'])
        
        # --- 3. SALVAR (A parte visual do Excel) ---
        
        # Pergunta onde quer salvar o arquivo novo
        caminho_saida = filedialog.asksaveasfilename(
            title="Salvar arquivo processado como...",
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile="analise_contratos_processados.xlsx"
        )

        # Se escolheu onde salvar:
        if caminho_saida:
            # Liga o motor "XlsxWriter" (que nos deixa pintar células e mexer no visual)
            with pd.ExcelWriter(caminho_saida, engine='xlsxwriter') as writer:
                
                # A. Joga os dados na planilha, MAS pula a primeira linha (startrow=1)
                # e NÃO escreve o cabeçalho padrão feio (header=False)
                df_.to_excel(writer, sheet_name="analise_contratos", index=False, startrow=1, header=False)
                
                # Pega as ferramentas de desenho do Excel
                workbook  = writer.book  # O arquivo em si
                worksheet = writer.sheets["analise_contratos"] # A aba da planilha
                
                # B. Cria a "Tinta" para o cabeçalho: Fundo Preto, Letra Branca, Sem Borda
                formato_cabecalho = workbook.add_format({
                    'bold': True,              # Negrito
                    'text_wrap': False,        # Não quebrar linha
                    'valign': 'top',           # Alinhado no topo
                    'fg_color': '#000000',     # FUNDO PRETO
                    'font_color': '#FFFFFF',   # LETRA BRANCA
                    'border': 0                # SEM BORDA
                })
                
                # C. Escreve os nomes das colunas manualmente na linha 0 usando a tinta preta criada acima
                for col_num, value in enumerate(df_.columns.values):
                    worksheet.write(0, col_num, value, formato_cabecalho)
                
                # D. Arruma os erros de "Número como Texto" (as bandeirinhas verdes)
                
                # Cria um formato que diz pro Excel: "Isso é texto puro (@)"
                formato_texto = workbook.add_format({'num_format': '@'}) 
                
                try:
                    # Descobre em qual coluna (número) está o CPF e a Inscrição
                    idx_cpf = df_.columns.get_loc('CPF')
                    idx_insc = df_.columns.get_loc('Inscrição')
                    
                    # Aplica o formato de texto nessas colunas e aumenta a largura delas
                    worksheet.set_column(idx_cpf, idx_cpf, 15, formato_texto)
                    worksheet.set_column(idx_insc, idx_insc, 12, formato_texto)
                    
                    # O TRUQUE MÁGICO: Manda o Excel ignorar o erro "Número armazenado como texto"
                    # Isso remove o triângulo verde chato
                    worksheet.ignore_errors({'number_stored_as_text': 'A:XFD'})
                    
                    # (Extra) Tenta ajustar a largura das outras colunas pra não ficarem espremidas
                    for i, col in enumerate(df_.columns):
                        # Se não for CPF nem Inscrição (que já arrumamos), ajusta o tamanho
                        if i != idx_cpf and i != idx_insc: 
                            worksheet.set_column(i, i, len(col) + 2) # Tamanho do texto + um pouquinho
                            
                except Exception as ex_fmt:
                    # Se der erro só na formatação visual, avisa no console mas não trava o programa
                    print(f"Erro ao formatar colunas (não crítico): {ex_fmt}")

            # Mostra o aviso de SUCESSO na tela
            messagebox.showinfo("Concluído", f"Sucesso!\nArquivo salvo em: {caminho_saida}")

    except Exception as e:
        # Se der erro grave (arquivo corrompido, etc), mostra um X vermelho e o erro
        messagebox.showerror("Erro", f"Ocorreu um erro ao processar o arquivo:\n{str(e)}")

# ==============================================================================
# CONFIGURAÇÃO DA TELA (VISUAL)
# ==============================================================================

# Cria a janela principal
janela = tk.Tk()
janela.title("Processador de Documentos IA") # Título lá no topo
janela.geometry("500x200") # Tamanho: 500 largura x 200 altura
janela.resizable(False, False) # Trava o tamanho (não deixa esticar)

# Cria um painel principal para dar uma margem nas bordas
main_frame = tk.Frame(janela, padx=20, pady=20)
main_frame.pack(fill="both", expand=True)

# Texto: "1. Selecione o arquivo..."
lbl_instrucao = tk.Label(main_frame, text="1. Selecione o arquivo Excel original:", font=("Arial", 10))
lbl_instrucao.pack(anchor="w") # w = West (Oeste/Esquerda)

# Cria uma linha para agrupar a caixinha branca e o botão buscar
frame_busca = tk.Frame(main_frame)
frame_busca.pack(fill="x", pady=5)

# A caixinha branca de texto
entry_arquivo_entrada = tk.Entry(frame_busca)
entry_arquivo_entrada.pack(side="left", fill="x", expand=True, padx=(0, 10))

# O botão cinza "Buscar Arquivo..."
btn_buscar = tk.Button(frame_busca, text="Buscar Arquivo...", command=selecionar_arquivo_entrada)
btn_buscar.pack(side="right")

# Um espaço vazio só para separar
tk.Label(main_frame, text="").pack() 

# O Botão Verde "GERAR E SALVAR"
btn_gerar = tk.Button(
    main_frame, text="GERAR E SALVAR", font=("Arial", 12, "bold"), 
    bg="#4CAF50", # Verde
    fg="white",   # Letra branca
    height=2,
    command=processar_e_salvar # Quando clicar, roda a função principal lá de cima
)
btn_gerar.pack(fill="x", pady=10)

# Mantém a janela aberta esperando você clicar
janela.mainloop()