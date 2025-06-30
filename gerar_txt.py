import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os

class ECFProcessor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Processador ECF")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # Variáveis
        self.arquivo_selecionado = tk.StringVar()
        
        self.setup_interface()
        
    def setup_interface(self):
        # Container principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Título
        titulo = ttk.Label(main_frame, text="Processador ECF", font=('Arial', 16, 'bold'))
        titulo.grid(row=0, column=0, pady=(0, 20))
        
        # Seleção de arquivo
        ttk.Label(main_frame, text="Arquivo:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        arquivo_frame = ttk.Frame(main_frame)
        arquivo_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        arquivo_frame.grid_columnconfigure(0, weight=1)
        
        self.entry_arquivo = ttk.Entry(arquivo_frame, textvariable=self.arquivo_selecionado, width=50)
        self.entry_arquivo.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        btn_selecionar = ttk.Button(arquivo_frame, text="Procurar", command=self.selecionar_arquivo)
        btn_selecionar.grid(row=0, column=1)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Selecione uma planilha ECF")
        self.status_label.grid(row=3, column=0, pady=(0, 10))
        
        # Barra de progresso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # Botões
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=5, column=0, pady=(0, 10))
        
        btn_processar = ttk.Button(btn_frame, text="Processar ECF", command=self.processar_ecf)
        btn_processar.grid(row=0, column=0, padx=(0, 10))
        
        btn_limpar = ttk.Button(btn_frame, text="Limpar", command=self.limpar_dados)
        btn_limpar.grid(row=0, column=1, padx=(0, 10))
        
        btn_sair = ttk.Button(btn_frame, text="Sair", command=self.root.quit)
        btn_sair.grid(row=0, column=2)
        
        # Informações
        info_text = tk.Text(main_frame, height=8, width=70, wrap=tk.WORD)
        info_text.grid(row=6, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        info_scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=info_text.yview)
        info_scrollbar.grid(row=6, column=1, sticky=(tk.N, tk.S), pady=(10, 0))
        info_text.configure(yscrollcommand=info_scrollbar.set)
        
        info_content = """BLOCOS PROCESSADOS:
• M001: Bloco Inicial (colunas A-E) - processado uma vez
• M010: Bloco de Abertura (colunas A-U, coluna P formatada) - processado uma vez
• M030: Indicador de Meses (colunas B-J)
• M300: Bloco Principal (coluna C + filhos M305/M310)
• M305: Filho do M300 (coluna D, valores coluna I) - CORRIGIDO: verifica coluna I
• M310: Filho do M300 (coluna D, valores coluna K)
• M350: Bloco Secundário (colunas C-Q + filhos M355/M360) - CORRIGIDO
• M355: Filho do M350 (coluna D, valores coluna I)
• M360: Filho do M350 (coluna D, valores coluna K)
• M500: Bloco Adicional (colunas E-AA, valores P e X) - por período
• M510: Bloco Adicional (colunas C-AA, valores P) - por período

"""
        
        info_text.insert(tk.END, info_content)
        info_text.configure(state='disabled')
        
        # Configurar redimensionamento
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(6, weight=1)
        
    def limpar_dados(self):
        """Limpa os dados inseridos"""
        self.arquivo_selecionado.set("")
        self.status_label.config(text="Selecione uma planilha ECF")
        
    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione a planilha ECF",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if arquivo:
            self.arquivo_selecionado.set(arquivo)
            self.status_label.config(text=f"Arquivo: {os.path.basename(arquivo)}")
    
    def processar_ecf(self):
        if not self.arquivo_selecionado.get():
            messagebox.showerror("Erro", "Selecione uma planilha ECF primeiro!")
            return
        
        try:
            self.status_label.config(text="Iniciando processamento...")
            self.progress.start()
            self.root.update()
            
            arquivo = self.arquivo_selecionado.get()
            
            try:
                self.status_label.config(text="Carregando abas...")
                self.root.update()
                
                # Ler como string para preservar formato original
                df_m001 = pd.read_excel(arquivo, sheet_name='M001', dtype=str)
                df_m010 = pd.read_excel(arquivo, sheet_name='M010', dtype=str)
                df_m030 = pd.read_excel(arquivo, sheet_name='M030', dtype=str)
                df_m300 = pd.read_excel(arquivo, sheet_name='M300', dtype=str)
                df_m305 = pd.read_excel(arquivo, sheet_name='M305', dtype=str)  
                df_m310 = pd.read_excel(arquivo, sheet_name='M310', dtype=str)
                df_m350 = pd.read_excel(arquivo, sheet_name='M350', dtype=str)
                df_m355 = pd.read_excel(arquivo, sheet_name='M355', dtype=str)
                df_m360 = pd.read_excel(arquivo, sheet_name='M360', dtype=str)
                df_m500 = pd.read_excel(arquivo, sheet_name='M500', dtype=str)
                df_m510 = pd.read_excel(arquivo, sheet_name='M510', dtype=str)
                
                self.status_label.config(text="Processando valores...")
                self.root.update()
                
                # Arredondar valores específicos
                df_m010 = self.arredondar_coluna_valor(df_m010, 'P')  # Coluna P do M010
                df_m300 = self.arredondar_coluna_valor(df_m300, 'N')
                df_m310 = self.arredondar_coluna_valor(df_m310, 'K')  
                df_m305 = self.arredondar_coluna_valor(df_m305, 'I')
                df_m350 = self.arredondar_coluna_valor(df_m350, 'N')
                df_m355 = self.arredondar_coluna_valor(df_m355, 'I')
                df_m360 = self.arredondar_coluna_valor(df_m360, 'K')
                
                # Arredondar valores dos blocos M500 e M510
                df_m500 = self.arredondar_coluna_valor(df_m500, 'P')  # Coluna P do M500
                df_m500 = self.arredondar_coluna_valor(df_m500, 'X')  # Coluna X do M500
                df_m510 = self.arredondar_coluna_valor(df_m510, 'P')  # Coluna P do M510
                
            except Exception as e:
                self.progress.stop()
                self.status_label.config(text="Erro ao carregar planilha!")
                messagebox.showerror("Erro", f"Erro ao carregar as abas: {str(e)}\n\nVerifique se as abas M001, M010, M030, M300, M305, M310, M350, M355, M360, M500 e M510 existem.")
                return
            
            self.status_label.config(text="Processando dados...")
            self.root.update()
            
            # Processar os dados
            dados_resultado = self.processar_dados(df_m001, df_m010, df_m030, df_m300, df_m305, df_m310, df_m350, df_m355, df_m360, df_m500, df_m510)
            
            self.status_label.config(text="Salvando arquivo...")
            self.root.update()
            
            # Salvar em arquivo TXT
            arquivo_txt = self.salvar_txt(arquivo, dados_resultado)
            
            self.progress.stop()
            self.status_label.config(text="Processamento concluído!")
            messagebox.showinfo("Sucesso", f"Dados processados com sucesso!\n\nArquivo salvo: {os.path.basename(arquivo_txt)}")
            
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Erro durante processamento!")
            messagebox.showerror("Erro", f"Erro ao processar arquivo: {str(e)}")
    
    def tem_valor_significativo(self, valor):
        """Verifica se um valor é significativo (não é vazio, zero ou NaN)"""
        if pd.isna(valor) or valor == 'nan' or valor == '':
            return False
        
        try:
            valor_float = float(valor)
            return valor_float != 0.0
        except (ValueError, TypeError):
            valor_str = str(valor).strip()
            if valor_str in ['0', '0.0', '0.00']:
                return False
            return bool(valor_str)
    
    def arredondar_coluna_valor(self, df, letra_coluna):
        """Arredonda valores numéricos em uma coluna específica para 2 casas decimais com vírgula"""
        indice_coluna = ord(letra_coluna.upper()) - ord('A')
        
        if indice_coluna >= len(df.columns):
            return df
        
        nome_coluna = df.columns[indice_coluna]
        
        def arredondar_valor(valor):
            if pd.isna(valor) or valor == 'nan' or valor == '':
                return ''
            
            try:
                valor_float = float(valor)
                if valor_float == 0.0:
                    return ''
                valor_arredondado = round(valor_float, 2)
                # Formatar com vírgula em vez de ponto
                return f"{valor_arredondado:.2f}".replace('.', ',')
            except (ValueError, TypeError):
                if str(valor).strip() in ['0', '0.0', '0.00']:
                    return ''
                return valor
        
        df[nome_coluna] = df[nome_coluna].apply(arredondar_valor)
        return df
    
    def concatenar_linha(self, serie):
        """Concatena os valores de uma série (linha) sem separador"""
        valores = []
        for valor in serie:
            if pd.isna(valor) or valor == 'nan':
                valores.append('')
            else:
                valores.append(str(valor))
        return ''.join(valores)
    
    def obter_periodos_ordenados(self, df_m030):
        """Obtém os períodos únicos ordenados com dezembro duplicado no final"""
        if df_m030.empty or len(df_m030.columns) == 0:
            return []
        
        periodos_unicos = df_m030.iloc[:, 0].dropna().unique().tolist()
        
        # Separar dezembro, data especial e outros períodos
        dezembro = None
        data_especial = None
        outros_periodos = []
        
        for periodo in periodos_unicos:
            periodo_str = str(periodo).strip().upper()
            if periodo_str in ['12', 'DEZ', 'DEZEMBRO', 'DEC']:
                dezembro = periodo
            elif '01/01/2025' in str(periodo):
                data_especial = periodo
            else:
                outros_periodos.append(periodo)
        
        # Ordenar os outros períodos (assumindo que são numéricos ou seguem ordem alfabética)
        try:
            # Tentar ordenar como números
            outros_periodos_num = []
            for p in outros_periodos:
                try:
                    outros_periodos_num.append((int(p), p))
                except:
                    outros_periodos_num.append((float('inf'), p))
            outros_periodos_num.sort()
            outros_periodos = [p[1] for p in outros_periodos_num]
        except:
            # Se falhar, ordenar alfabeticamente
            outros_periodos.sort()
        
        # Construir lista final: outros períodos, dezembro, dezembro duplicado (data especial)
        periodos_ordenados = []
        periodos_ordenados.extend(outros_periodos)  # Outros meses
        
        if dezembro is not None:
            periodos_ordenados.append(dezembro)  # Dezembro normal
        
        if data_especial is not None:
            periodos_ordenados.append(data_especial)  # Dezembro duplicado com data especial
        
        return periodos_ordenados
    
    def obter_valores_dezembro_a12(self, df_m030):
        """Obtém os valores A12 da coluna I do M030 para dezembro"""
        valores_a12 = []
        
        # Buscar por dezembro (pode ser 12, DEZ, etc.)
        for _, linha in df_m030.iterrows():
            periodo = str(linha.iloc[0]).strip().upper()
            if periodo in ['12', 'DEZ', 'DEZEMBRO', 'DEC']:
                # Coluna I é índice 8 (I = 9ª coluna, índice 8)
                if len(linha) > 8:
                    coluna_i = linha.iloc[8]  # Coluna I
                    if not pd.isna(coluna_i) and str(coluna_i).strip():
                        # Verificar se é A12
                        if 'A12' in str(coluna_i):
                            valores_a12.append(str(coluna_i))
        
        return valores_a12
    
    def converter_a12_para_a00(self, valor_a12):
        """Converte valor A12 para A00"""
        if pd.isna(valor_a12) or not str(valor_a12).strip():
            return ''
        
        valor_str = str(valor_a12)
        # Substituir A12 por A00
        valor_a00 = valor_str.replace('A12', 'A00')
        return valor_a00
    
    def processar_dados(self, df_m001, df_m010, df_m030, df_m300, df_m305, df_m310, df_m350, df_m355, df_m360, df_m500, df_m510):
        resultado = []
        
        # Fazer cópia dos DataFrames
        df_m001 = df_m001.copy()
        df_m010 = df_m010.copy()
        df_m030 = df_m030.copy()
        df_m300 = df_m300.copy()
        df_m305 = df_m305.copy()
        df_m310 = df_m310.copy()
        df_m350 = df_m350.copy()
        df_m355 = df_m355.copy()
        df_m360 = df_m360.copy()
        df_m500 = df_m500.copy()
        df_m510 = df_m510.copy()
        
        # Obter valores A12 de dezembro para duplicação
        valores_dezembro_a12 = self.obter_valores_dezembro_a12(df_m030)
        
        # 1. PROCESSAR M001 (BLOCO INICIAL) - PRIMEIRO
        self.status_label.config(text="Processando M001...")
        self.root.update()
        
        if not df_m001.empty:
            for _, linha_m001 in df_m001.iterrows():
                linha_m001_concat = self.concatenar_linha(linha_m001.iloc[0:5])  # A até E
                resultado.append(linha_m001_concat)
        
        # 2. PROCESSAR M010 (ABERTURA DO TEMPLATE) - SEGUNDO
        self.status_label.config(text="Processando M010...")
        self.root.update()
        
        if not df_m010.empty:
            for _, linha_m010 in df_m010.iterrows():
                linha_m010_concat = self.concatenar_linha(linha_m010.iloc[0:21])  # A até U
                resultado.append(linha_m010_concat)
        
        # 3. PROCESSAR DEMAIS BLOCOS POR PERÍODO COM DEZEMBRO DUPLICADO
        periodos_ordenados = self.obter_periodos_ordenados(df_m030)
        
        # Processar cada período na ordem: JAN → FEV → ... → DEZ → DEZ(duplicado com 01/01/2025)
        for i, periodo in enumerate(periodos_ordenados):
            # Verificar se é o dezembro duplicado (com data especial)
            eh_dezembro_duplicado = ('01/01/2025' in str(periodo))
            
            if eh_dezembro_duplicado:
                self.status_label.config(text=f"Processando dezembro duplicado (01/01/2025)...")
            else:
                self.status_label.config(text=f"Processando período {periodo}...")
            self.root.update()
            
            # M030
            df_m030_periodo = df_m030[df_m030.iloc[:, 0] == periodo]
            for _, linha_m030 in df_m030_periodo.iterrows():
                linha_m030_concat = self.concatenar_linha(linha_m030.iloc[1:10])  # B até J
                resultado.append(linha_m030_concat)
            
            # NOVA FUNCIONALIDADE: Duplicar valores A12 como A00 após M030 de 01/01/2025
            if eh_dezembro_duplicado and valores_dezembro_a12:
                self.status_label.config(text="Adicionando valores A12 convertidos para A00...")
                self.root.update()
                
                for valor_a12 in valores_dezembro_a12:
                    valor_a00 = self.converter_a12_para_a00(valor_a12)
                    if valor_a00:
                        resultado.append(valor_a00)
            
            # M300, M305, M310
            df_m300_periodo = df_m300[df_m300.iloc[:, 0] == periodo] if not df_m300.empty and len(df_m300.columns) > 0 else pd.DataFrame()
            df_m305_periodo = df_m305[df_m305.iloc[:, 0] == periodo] if not df_m305.empty and len(df_m305.columns) > 0 else pd.DataFrame()
            df_m310_periodo = df_m310[df_m310.iloc[:, 0] == periodo] if not df_m310.empty and len(df_m310.columns) > 0 else pd.DataFrame()
            
            # Processar M300 com seus filhos
            for _, linha_m300 in df_m300_periodo.iterrows():
                codigo_ecf = linha_m300.iloc[1] if len(linha_m300) > 1 else ''
                
                # M300 pai
                linha_m300_concat = self.concatenar_linha(linha_m300.iloc[2:])  # C em diante
                resultado.append(linha_m300_concat)
                
                # M305 filhos - CORRIGIDO: verificar coluna I (índice 8) em vez de J (índice 9)
                df_m305_codigo = df_m305_periodo[df_m305_periodo.iloc[:, 1] == codigo_ecf] if not df_m305_periodo.empty else pd.DataFrame()
                for _, linha_m305 in df_m305_codigo.iterrows():
                    if len(linha_m305) > 8 and self.tem_valor_significativo(linha_m305.iloc[8]):  # Coluna I (índice 8)
                        linha_m305_concat = self.concatenar_linha(linha_m305.iloc[3:])  # D em diante
                        resultado.append(linha_m305_concat)
                
                # M310 filhos
                df_m310_codigo = df_m310_periodo[df_m310_periodo.iloc[:, 1] == codigo_ecf] if not df_m310_periodo.empty else pd.DataFrame()
                for _, linha_m310 in df_m310_codigo.iterrows():
                    if len(linha_m310) > 10 and self.tem_valor_significativo(linha_m310.iloc[10]):  # Coluna K
                        linha_m310_concat = self.concatenar_linha(linha_m310.iloc[3:])  # D em diante
                        resultado.append(linha_m310_concat)
            
            # M350, M355, M360
            df_m350_periodo = df_m350[df_m350.iloc[:, 0] == periodo] if not df_m350.empty and len(df_m350.columns) > 0 else pd.DataFrame()
            df_m355_periodo = df_m355[df_m355.iloc[:, 0] == periodo] if not df_m355.empty and len(df_m355.columns) > 0 else pd.DataFrame()
            df_m360_periodo = df_m360[df_m360.iloc[:, 0] == periodo] if not df_m360.empty and len(df_m360.columns) > 0 else pd.DataFrame()
            
            # Processar M350 com seus filhos
            for _, linha_m350 in df_m350_periodo.iterrows():
                codigo_ecf = linha_m350.iloc[1] if len(linha_m350) > 1 else ''
                
                # M350 pai - CORRIGIDO: concatenar apenas da coluna C até Q (índices 2 até 16)
                linha_m350_concat = self.concatenar_linha(linha_m350.iloc[2:17])  # C até Q (índice 2 até 16, exclusive 17)
                resultado.append(linha_m350_concat)
                
                # M355 filhos
                df_m355_codigo = df_m355_periodo[df_m355_periodo.iloc[:, 1] == codigo_ecf] if not df_m355_periodo.empty else pd.DataFrame()
                for _, linha_m355 in df_m355_codigo.iterrows():
                    if len(linha_m355) > 8 and self.tem_valor_significativo(linha_m355.iloc[8]):  # Coluna I
                        linha_m355_concat = self.concatenar_linha(linha_m355.iloc[3:])  # D em diante
                        resultado.append(linha_m355_concat)
                
                # M360 filhos
                df_m360_codigo = df_m360_periodo[df_m360_periodo.iloc[:, 1] == codigo_ecf] if not df_m360_periodo.empty else pd.DataFrame()
                for _, linha_m360 in df_m360_codigo.iterrows():
                    if len(linha_m360) > 10 and self.tem_valor_significativo(linha_m360.iloc[10]):  # Coluna K
                        linha_m360_concat = self.concatenar_linha(linha_m360.iloc[3:])  # D em diante
                        resultado.append(linha_m360_concat)
                        
            
            # M500 e M510 - iterados apenas por período (não por código ECF)
            df_m500_periodo = df_m500[df_m500.iloc[:, 0] == periodo] if not df_m500.empty and len(df_m500.columns) > 0 else pd.DataFrame()
            df_m510_periodo = df_m510[df_m510.iloc[:, 0] == periodo] if not df_m510.empty and len(df_m510.columns) > 0 else pd.DataFrame()
            
            # M500 (colunas E até AA - índices 4 até 26)
            for _, linha_m500 in df_m500_periodo.iterrows():
                linha_m500_concat = self.concatenar_linha(linha_m500.iloc[4:27])  # E até AA
                resultado.append(linha_m500_concat)
            
            # M510 (colunas C até AA - índices 2 até 26)
            for _, linha_m510 in df_m510_periodo.iterrows():
                linha_m510_concat = self.concatenar_linha(linha_m510.iloc[2:27])  # C até AA
                resultado.append(linha_m510_concat)
        
        return resultado
    
    def salvar_txt(self, arquivo_original, dados):
        """Salva os dados processados em arquivo TXT"""
        nome_base = os.path.splitext(arquivo_original)[0]
        arquivo_txt = f"{nome_base}_processado.txt"
        
        with open(arquivo_txt, 'w', encoding='utf-8') as f:
            for linha in dados:
                f.write(linha + '\n')
        
        return arquivo_txt
    
    def executar(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ECFProcessor()
    app.executar()