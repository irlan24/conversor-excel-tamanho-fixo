import pandas as pd
from typing import List, Dict
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
import re
import openpyxl

# -------------------------------------------------------------
# FUNÇÕES DE PROCESSAMENTO
# -------------------------------------------------------------

def limpar_nome_coluna(nome: str) -> str:
    """Remove espaços extras, quebras de linha e caracteres invisíveis."""
    if not isinstance(nome, str):
        nome = str(nome)
    # Remove quebras de linha, tabs e espaços extras
    nome = re.sub(r'[\n\r\t]+', ' ', nome)
    # Remove espaços duplicados
    nome = re.sub(r'\s+', ' ', nome)
    # Remove espaços no início e fim (incluindo non-breaking spaces)
    nome = nome.strip().strip('\xa0').strip()
    return nome

def normalizar_colunas_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza os nomes das colunas do DataFrame."""
    df.columns = [limpar_nome_coluna(col) for col in df.columns]
    return df

def formatar_coluna(valor, nome_coluna: str, tamanho: int, zfill_cols: Dict[str, int]) -> str:
    """Formata uma célula para o tamanho fixo especificado."""
    if pd.isna(valor):
        valor_str = ""
    else:
        valor_str = str(valor).strip()
    
    if nome_coluna in zfill_cols:
        comprimento_zfill = zfill_cols[nome_coluna]
        valor_str = valor_str.zfill(comprimento_zfill)
    
    if len(valor_str) > tamanho:
        valor_str = valor_str[:tamanho]
    
    return valor_str.ljust(tamanho, ' ')

def formatar_linha_tamanho_fixo(df: pd.DataFrame, tamanhos: Dict[str, int], zfill_cols: Dict[str, int]) -> List[str]:
    """Processa todas as linhas do DataFrame."""
    linhas_formatadas = []
    nomes_colunas = list(tamanhos.keys())

    for row_tuple in df.itertuples(index=False):
        linha_saida = ""
        for i, nome_coluna in enumerate(nomes_colunas):
            if nome_coluna not in tamanhos:
                continue
            valor_original = row_tuple[i]
            tamanho_desejado = tamanhos[nome_coluna]
            campo_formatado = formatar_coluna(valor_original, nome_coluna, tamanho_desejado, zfill_cols)
            linha_saida += campo_formatado
        linhas_formatadas.append(linha_saida)
    
    return linhas_formatadas

def encontrar_coluna_similar(nome_procurado: str, colunas_df: List[str]) -> str:
    """Tenta encontrar uma coluna similar ignorando case e espaços."""
    nome_limpo = limpar_nome_coluna(nome_procurado).lower()
    
    for col in colunas_df:
        col_limpo = limpar_nome_coluna(col).lower()
        if nome_limpo == col_limpo:
            return col
    return None

# -------------------------------------------------------------
# INTERFACE GRÁFICA
# -------------------------------------------------------------

class ConversorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Conversor Excel - Tamanho Fixo")
        self.geometry("900x790")
        self.minsize(800, 650)
        
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.arquivo_selecionado = None
        self.colunas_excel = []  # Colunas do Excel carregado
        self.colunas_config = []
        self.criar_interface()
    
    def criar_interface(self):
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        titulo = ctk.CTkLabel(
            main_frame, 
            text="Conversor de Excel para Arquivo de Tamanho Fixo",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        titulo.pack(pady=(0, 15))
        
        # Frame de seleção de arquivo
        file_frame = ctk.CTkFrame(main_frame)
        file_frame.pack(fill="x", pady=5)
        
        self.label_arquivo = ctk.CTkLabel(
            file_frame,
            text="Nenhum arquivo selecionado",
            font=ctk.CTkFont(size=12)
        )
        self.label_arquivo.pack(pady=5)
        
        btn_selecionar = ctk.CTkButton(
            file_frame,
            text="Selecionar Arquivo Excel",
            command=self.selecionar_arquivo,
            font=ctk.CTkFont(size=13),
            height=35
        )
        btn_selecionar.pack(pady=5)
        
        # Frame para mostrar colunas do Excel
        self.excel_info_frame = ctk.CTkFrame(file_frame)
        self.excel_info_frame.pack(fill="x", padx=10, pady=5)
        
        self.label_colunas_excel = ctk.CTkLabel(
            self.excel_info_frame,
            text="",
            font=ctk.CTkFont(size=10),
            wraplength=800,
            justify="left"
        )
        self.label_colunas_excel.pack(pady=5)
        
        # Frame para adicionar colunas
        add_frame = ctk.CTkFrame(main_frame)
        add_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(add_frame, text="Configurar Colunas:", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=5)
        
        input_frame = ctk.CTkFrame(add_frame)
        input_frame.pack(fill="x", padx=10, pady=5)
        
        # ComboBox para selecionar coluna do Excel
        ctk.CTkLabel(input_frame, text="Coluna do Excel:", font=ctk.CTkFont(size=11)).grid(row=0, column=0, padx=5, pady=2)
        self.combo_coluna = ctk.CTkComboBox(input_frame, width=250, values=[], state="readonly")
        self.combo_coluna.grid(row=1, column=0, padx=5, pady=2)
        self.combo_coluna.set("Selecione...")
        
        # Ou digitar manualmente
        ctk.CTkLabel(input_frame, text="Ou digite:", font=ctk.CTkFont(size=11)).grid(row=0, column=1, padx=5, pady=2)
        self.entry_nome = ctk.CTkEntry(input_frame, width=200, placeholder_text="Nome manual")
        self.entry_nome.grid(row=1, column=1, padx=5, pady=2)
        
        ctk.CTkLabel(input_frame, text="Tamanho:", font=ctk.CTkFont(size=11)).grid(row=0, column=2, padx=5, pady=2)
        self.entry_tamanho = ctk.CTkEntry(input_frame, width=70, placeholder_text="Ex: 11")
        self.entry_tamanho.grid(row=1, column=2, padx=5, pady=2)
        
        self.var_zfill = ctk.BooleanVar(value=False)
        self.check_zfill = ctk.CTkCheckBox(
            input_frame, 
            text="Zeros à esquerda", 
            variable=self.var_zfill,
            font=ctk.CTkFont(size=11)
        )
        self.check_zfill.grid(row=1, column=3, padx=10, pady=2)
        
        btn_adicionar = ctk.CTkButton(
            input_frame,
            text="+ Adicionar",
            command=self.adicionar_coluna,
            width=100,
            height=30,
            fg_color="green",
            hover_color="darkgreen"
        )
        btn_adicionar.grid(row=1, column=4, padx=10, pady=2)
        
        # Frame para lista de colunas
        lista_frame = ctk.CTkFrame(main_frame)
        lista_frame.pack(fill="both", expand=True, pady=5)
        
        ctk.CTkLabel(
            lista_frame, 
            text="Colunas Configuradas:", 
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(pady=5)
        
        header_frame = ctk.CTkFrame(lista_frame)
        header_frame.pack(fill="x", padx=10)
        
        ctk.CTkLabel(header_frame, text="Nº", width=30, font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=2)
        ctk.CTkLabel(header_frame, text="Nome da Coluna", width=250, font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=2)
        ctk.CTkLabel(header_frame, text="Tamanho", width=70, font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=2)
        ctk.CTkLabel(header_frame, text="Preenchimento", width=120, font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=2)
        ctk.CTkLabel(header_frame, text="Ações", width=100, font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=2)
        
        self.scroll_frame = ctk.CTkScrollableFrame(lista_frame, height=150)
        self.scroll_frame.pack(fill="both", expand=True, padx=10, pady=3)
        
        action_frame = ctk.CTkFrame(main_frame)
        action_frame.pack(fill="x", pady=(1,4))
        
        btn_limpar = ctk.CTkButton(
            action_frame,
            text="Limpar Todas",
            command=self.limpar_colunas,
            width=120,
            height=35,
            fg_color="gray",
            hover_color="darkgray"
        )
        btn_limpar.pack(side="left", padx=10, pady=5)
        
        self.btn_converter = ctk.CTkButton(
            action_frame,
            text="Converter Arquivo",
            command=self.converter_arquivo,
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            width=200,
            state="disabled"
        )
        self.btn_converter.pack(side="right", padx=10, pady=5)
        
        self.label_status = ctk.CTkLabel(
            main_frame,
            text="Selecione um arquivo Excel para começar",
            font=ctk.CTkFont(size=12)
        )
        self.label_status.pack(pady=5)
        
        self.label_total = ctk.CTkLabel(
            main_frame,
            text="Tamanho total da linha: 0 caracteres",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.label_total.pack(pady=5)
    
    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[
                ("Arquivos Excel", "*.xlsx *.xls"),
                ("Todos os arquivos", "*.*")
            ]
        )
        
        if arquivo:
            try:
                # Lê o Excel e normaliza as colunas
                df_temp = pd.read_excel(arquivo)
                df_temp = normalizar_colunas_df(df_temp)
                
                self.colunas_excel = df_temp.columns.tolist()
                self.arquivo_selecionado = arquivo
                
                nome_arquivo = os.path.basename(arquivo)
                self.label_arquivo.configure(text=f"Arquivo: {nome_arquivo}")
                
                # Atualiza ComboBox com colunas do Excel
                self.combo_coluna.configure(values=self.colunas_excel)
                self.combo_coluna.set("Selecione...")
                
                # Mostra colunas encontradas
                colunas_texto = f"Colunas encontradas ({len(self.colunas_excel)}): " + " | ".join(self.colunas_excel)
                self.label_colunas_excel.configure(text=colunas_texto, text_color="lightgreen")
                
                self.label_status.configure(text="Arquivo carregado! Adicione as colunas desejadas.", text_color="green")
                self.verificar_botao_converter()
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel:\n\n{str(e)}")
                self.label_status.configure(text=f"Erro ao ler arquivo", text_color="red")
    
    def adicionar_coluna(self):
        # Prioriza ComboBox, senão usa entrada manual
        nome_combo = self.combo_coluna.get()
        nome_manual = self.entry_nome.get().strip()
        
        if nome_combo and nome_combo != "Selecione...":
            nome = nome_combo
        elif nome_manual:
            nome = limpar_nome_coluna(nome_manual)
            # Tenta encontrar coluna similar no Excel
            if self.colunas_excel:
                similar = encontrar_coluna_similar(nome, self.colunas_excel)
                if similar and similar != nome:
                    if messagebox.askyesno("Coluna Similar", 
                        f"Você digitou: '{nome}'\n\n"
                        f"Foi encontrada uma coluna similar: '{similar}'\n\n"
                        f"Deseja usar '{similar}' ao invés?"
                    ):
                        nome = similar
        else:
            messagebox.showwarning("Atenção", "Selecione ou digite o nome da coluna!")
            return
        
        tamanho_str = self.entry_tamanho.get().strip()
        usar_zfill = self.var_zfill.get()
        
        if not tamanho_str:
            messagebox.showwarning("Atenção", "Informe o tamanho da coluna!")
            return
        
        try:
            tamanho = int(tamanho_str)
            if tamanho <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showwarning("Atenção", "O tamanho deve ser um número inteiro positivo!")
            return
        
        for col in self.colunas_config:
            if col['nome'] == nome:
                messagebox.showwarning("Atenção", f"A coluna '{nome}' já foi adicionada!")
                return
        
        self.colunas_config.append({
            'nome': nome,
            'tamanho': tamanho,
            'zfill': usar_zfill
        })
        
        self.combo_coluna.set("Selecione...")
        self.entry_nome.delete(0, 'end')
        self.entry_tamanho.delete(0, 'end')
        self.var_zfill.set(False)
        
        self.atualizar_lista_colunas()
        self.verificar_botao_converter()
    
    def remover_coluna(self, index):
        if 0 <= index < len(self.colunas_config):
            self.colunas_config.pop(index)
            self.atualizar_lista_colunas()
            self.verificar_botao_converter()
    
    def mover_cima(self, index):
        if index > 0:
            self.colunas_config[index], self.colunas_config[index-1] = \
                self.colunas_config[index-1], self.colunas_config[index]
            self.atualizar_lista_colunas()
    
    def mover_baixo(self, index):
        if index < len(self.colunas_config) - 1:
            self.colunas_config[index], self.colunas_config[index+1] = \
                self.colunas_config[index+1], self.colunas_config[index]
            self.atualizar_lista_colunas()
    
    def atualizar_lista_colunas(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        
        for i, col in enumerate(self.colunas_config):
            row_frame = ctk.CTkFrame(self.scroll_frame)
            row_frame.pack(fill="x", pady=2)
            
            ctk.CTkLabel(row_frame, text=str(i+1), width=30).pack(side="left", padx=2)
            ctk.CTkLabel(row_frame, text=col['nome'], width=250, anchor="w").pack(side="left", padx=2)
            ctk.CTkLabel(row_frame, text=str(col['tamanho']), width=70).pack(side="left", padx=2)
            
            tipo = "Zeros (zfill)" if col['zfill'] else "Espaços (ljust)"
            cor = "orange" if col['zfill'] else "lightblue"
            ctk.CTkLabel(row_frame, text=tipo, width=120, text_color=cor).pack(side="left", padx=2)
            
            btn_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
            btn_frame.pack(side="left", padx=2)
            
            ctk.CTkButton(btn_frame, text="↑", width=25, height=25, command=lambda idx=i: self.mover_cima(idx)).pack(side="left", padx=1)
            ctk.CTkButton(btn_frame, text="↓", width=25, height=25, command=lambda idx=i: self.mover_baixo(idx)).pack(side="left", padx=1)
            ctk.CTkButton(btn_frame, text="X", width=25, height=25, fg_color="red", hover_color="darkred", command=lambda idx=i: self.remover_coluna(idx)).pack(side="left", padx=1)
        
        total = sum(col['tamanho'] for col in self.colunas_config)
        self.label_total.configure(text=f"Tamanho total da linha: {total} caracteres")
    
    def limpar_colunas(self):
        if self.colunas_config:
            if messagebox.askyesno("Confirmar", "Deseja remover todas as colunas configuradas?"):
                self.colunas_config.clear()
                self.atualizar_lista_colunas()
                self.verificar_botao_converter()
    
    def verificar_botao_converter(self):
        if self.arquivo_selecionado and self.colunas_config:
            self.btn_converter.configure(state="normal")
            self.label_status.configure(text="Pronto para converter!", text_color="green")
        else:
            self.btn_converter.configure(state="disabled")
            if not self.colunas_config:
                self.label_status.configure(text="Adicione pelo menos uma coluna", text_color="orange")
            elif not self.arquivo_selecionado:
                self.label_status.configure(text="Selecione um arquivo Excel", text_color="orange")
    
    def converter_arquivo(self):
        if not self.arquivo_selecionado:
            messagebox.showerror("Erro", "Nenhum arquivo selecionado!")
            return
        
        if not self.colunas_config:
            messagebox.showerror("Erro", "Configure pelo menos uma coluna!")
            return
        
        try:
            self.label_status.configure(text="Processando arquivo...", text_color="orange")
            self.update()
            
            tamanhos = {col['nome']: col['tamanho'] for col in self.colunas_config}
            zfill_cols = {col['nome']: col['tamanho'] for col in self.colunas_config if col['zfill']}
            nomes_colunas = [col['nome'] for col in self.colunas_config]
            
            # Lê e normaliza o DataFrame
            df = pd.read_excel(self.arquivo_selecionado)
            df = normalizar_colunas_df(df)
            
            # Verifica colunas faltantes
            colunas_faltantes = []
            for col in nomes_colunas:
                if col not in df.columns:
                    similar = encontrar_coluna_similar(col, df.columns.tolist())
                    if similar:
                        colunas_faltantes.append(f"'{col}' (similar: '{similar}')")
                    else:
                        colunas_faltantes.append(f"'{col}'")
            
            if colunas_faltantes:
                messagebox.showerror(
                    "Erro",
                    f"Colunas não encontradas no arquivo:\n\n" + "\n".join(colunas_faltantes) +
                    f"\n\nColunas disponíveis:\n" + "\n".join(df.columns.tolist())
                )
                self.label_status.configure(text="Erro: Colunas faltantes", text_color="red")
                return
            
            df_ordenado = df[nomes_colunas]
            linhas_de_saida = formatar_linha_tamanho_fixo(df_ordenado, tamanhos, zfill_cols)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_saida = f"arquivo_convertido_{timestamp}.txt"
            
            caminho_saida = os.path.join(os.path.dirname(self.arquivo_selecionado), nome_saida)
            with open(caminho_saida, 'w', encoding='utf-8') as f:
                f.write('\n'.join(linhas_de_saida))
            
            tamanho_linha = sum(col['tamanho'] for col in self.colunas_config)
            messagebox.showinfo(
                "Sucesso!",
                f"Arquivo convertido com sucesso!\n\n"
                f"Linhas processadas: {len(linhas_de_saida)}\n"
                f"Tamanho por linha: {tamanho_linha} caracteres\n"
                f"Arquivo salvo em:\n{caminho_saida}"
            )
            
            self.label_status.configure(
                text=f"Conversão concluída! {len(linhas_de_saida)} linhas processadas",
                text_color="green"
            )
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo:\n\n{str(e)}")
            self.label_status.configure(text=f"Erro: {str(e)}", text_color="red")

# -------------------------------------------------------------
# EXECUÇÃO
# -------------------------------------------------------------

if __name__ == "__main__":
    app = ConversorApp()
    app.mainloop()