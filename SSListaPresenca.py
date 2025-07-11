# @Author: Giancarlo Medeiros de Almeida
# Versão 1.2
# Data: 2025-07-10 22:03
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pypdf import PdfReader, PdfWriter
import pdfplumber
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import black
from io import BytesIO
import re
import os

# --- Variável global para a janela de progresso ---
progress_window = None
progress_label = None

# --- Função para criar a janela de progresso ---
def criar_janela_progresso(titulo):
    global progress_window, progress_label
    progress_window = tk.Toplevel()
    progress_window.title(titulo)
    # Impede a interação com a janela principal enquanto esta estiver aberta
    progress_window.grab_set()
    # Impede o fechamento da janela de progresso pelo botão "X"
    progress_window.protocol("WM_DELETE_WINDOW", lambda: None)
    progress_label = ttk.Label(progress_window, text="")
    progress_label.pack(padx=20, pady=20)
    # Centraliza a janela
    window_width = progress_window.winfo_reqwidth()
    window_height = progress_window.winfo_reqheight()
    position_right = int(progress_window.winfo_screenwidth()/2 - window_width/2)
    position_down = int(progress_window.winfo_screenheight()/2 - window_height/2)
    progress_window.geometry("+{}+{}".format(position_right, position_down))
    progress_window.update()

# --- Função para atualizar a mensagem de progresso ---
def atualizar_progresso(mensagem):
    global progress_label, progress_window
    if progress_label and progress_window:
        progress_label.config(text=mensagem)
        progress_window.update()

# --- Função para fechar a janela de progresso ---
def fechar_janela_progresso():
    global progress_window
    if progress_window:
        progress_window.destroy()
        progress_window = None
        progress_label = None

# --- Funções de UI (modificadas para usar a janela de progresso) ---
def selecionar_arquivo(titulo, tipos_arquivo):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    root.lift()
    caminho = filedialog.askopenfilename(title=titulo, filetypes=tipos_arquivo)
    root.destroy()
    return caminho

def selecionar_pasta_saida(titulo, default_path=""):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    root.lift()
    pasta = filedialog.askdirectory(title=titulo, initialdir=default_path)
    root.destroy()
    return pasta

def ler_csv(caminho_csv):
     atualizar_progresso("Lendo o arquivo CSV...")
     faltosos = []
     if not caminho_csv:
         messagebox.showerror("Erro de Arquivo", "Nenhum arquivo CSV selecionado.")
         return []

     for encoding in ['utf-8', 'latin-1', 'utf-8-sig']:
         try:
             with open(caminho_csv, newline='', encoding=encoding) as csvfile:
                 leitor = csv.DictReader(csvfile, delimiter=';')

                 original_fieldnames = [f.strip().strip('"') for f in leitor.fieldnames]
                 normalized_fieldnames = [f.lower().replace('matrícula', 'matricula').replace('aluno', 'aluno').replace('check', 'check') for f in original_fieldnames]

                 matricula_col_idx = -1
                 aluno_col_idx = -1
                 check_col_idx = -1

                 for idx, col_name in enumerate(normalized_fieldnames):
                     if 'matricula' in col_name:
                         matricula_col_idx = idx
                     if 'aluno' in col_name:
                         aluno_col_idx = idx
                     if 'check' in col_name:
                         check_col_idx = idx

                 if matricula_col_idx == -1 or aluno_col_idx == -1 or check_col_idx == -1:
                      raise ValueError("❌ O arquivo CSV não contém as colunas esperadas: 'Matrícula', 'Aluno' e 'Check'.")

                 for linha in leitor:
                     try:
                         # CORREÇÃO AQUI: Usar [] para acessar elementos de dicionários e listas
                         matricula = linha[original_fieldnames[matricula_col_idx]].strip()
                         aluno = linha[original_fieldnames[aluno_col_idx]].strip()
                         check = linha[original_fieldnames[check_col_idx]].strip().upper()
                     except IndexError:
                         print(f"⚠️ Aviso: Linha malformada no CSV, pulando: {linha}")
                         continue
                     except KeyError as e:
                         print(f"⚠️ Aviso: Coluna não encontrada na linha: {e}. Linha: {linha}")
                         continue

                     if check == 'F':
                         faltosos.append({'matricula': matricula, 'nome': aluno})

                 print(f"✅ {len(faltosos)} alunos com falta carregados do CSV.")
                 return faltosos
         except UnicodeDecodeError:
             continue
         except ValueError as e:
             messagebox.showerror("Erro de Colunas CSV", str(e))
             return []
         except Exception as e:
             messagebox.showerror("Erro na Leitura do CSV", f"Ocorreu um erro inesperado ao ler o CSV: {e}")
             return []

     messagebox.showerror("Erro de Codificação", "❌ Não foi possível ler o arquivo CSV com as codificações conhecidas (UTF-8, Latin-1, UTF-8-SIG).")
     return []

def criar_overlay_buffer(textos_com_posicoes):
    buffer = BytesIO()
    can = canvas.Canvas(buffer, pagesize=letter)
    can.setFont("Helvetica-Bold", 9)
    can.setFillColor(black)

    for texto, x, y in textos_com_posicoes:
        can.drawString(x, y, texto)
    can.save()
    buffer.seek(0)
    return PdfReader(buffer)

def marcar_faltas(caminho_pdf, faltosos, saida_pdf):
     if not caminho_pdf or not faltosos or not saida_pdf:
         messagebox.showinfo("Operação Cancelada", "Caminhos inválidos ou lista de faltosos vazia.")
         return

     leitor = None
     escritor = None
     pdfplumber_doc = None

     alunos_faltosos_nao_encontrados_final = set()

     try:
         leitor = PdfReader(caminho_pdf)
         escritor = PdfWriter()

         pdfplumber_doc = pdfplumber.open(caminho_pdf)
         total_paginas = len(pdfplumber_doc.pages)

         atualizar_progresso(f"Processando 0 de {total_paginas} páginas...")

         alunos_do_csv_para_rastrear = {
             (f['matricula'], f['nome'].lower()) for f in faltosos
         }
         alunos_encontrados_alguma_vez = set()

         for i in range(total_paginas):
             atualizar_progresso(f"Processando página {i + 1} de {total_paginas}...")
             # CORREÇÃO AQUI: Usar [] para acessar elementos de listas
             pagina_plumber = pdfplumber_doc.pages[i]
             pagina_pypdf = leitor.pages[i]

             textos_para_inserir_nesta_pagina = []

             page_height_pypdf = pagina_pypdf.mediabox.height

             for faltoso in faltosos:
                 # CORREÇÃO AQUI: Usar [] para acessar elementos de dicionários
                 matricula_original = faltoso['matricula']
                 nome_original = faltoso['nome']

                 normalized_nome_aluno = re.sub(r'[ÁÀÂÃÄáàâãä]', 'a', nome_original, flags=re.IGNORECASE)
                 normalized_nome_aluno = re.sub(r'[ÉÈÊËéèêë]', 'e', normalized_nome_aluno, flags=re.IGNORECASE)
                 normalized_nome_aluno = re.sub(r'[ÍÌÎÏíìîï]', 'i', normalized_nome_aluno, flags=re.IGNORECASE)
                 normalized_nome_aluno = re.sub(r'[ÓÒÔÕÖóòôõö]', 'o', normalized_nome_aluno, flags=re.IGNORECASE)
                 normalized_nome_aluno = re.sub(r'[ÚÙÛÜúùûü]', 'u', normalized_nome_aluno, flags=re.IGNORECASE)
                 normalized_nome_aluno = re.sub(r'[Çç]', 'c', normalized_nome_aluno, flags=re.IGNORECASE)

                 found_word = None
                 try:
                     words_on_page = pagina_plumber.extract_words()

                     for word_obj in words_on_page:
                         # CORREÇÃO AQUI: Usar [] para acessar elementos de dicionários
                         if matricula_original in word_obj['text'] or normalized_nome_aluno in word_obj['text'].lower():
                             found_word = word_obj
                             break
                 except Exception as e:
                     print(f"⚠️ Aviso: Erro ao extrair palavras da página {i+1} ou buscar por '{nome_original}'/'{matricula_original}': {e}")
                     continue

                 if found_word:
                     # CORREÇÃO AQUI: Usar [] para acessar elementos de dicionários
                     x0_plumber = found_word['x0']
                     bottom_plumber = found_word['bottom']

                     y_base_pypdf_coords = page_height_pypdf - bottom_plumber

                     vertical_offset_para_falta = 3

                     y_final_para_reportlab = y_base_pypdf_coords + vertical_offset_para_falta
                     x_final_para_reportlab = x0_plumber + 540

                     textos_para_inserir_nesta_pagina.append(("FALTA", x_final_para_reportlab, y_final_para_reportlab))
                     print(f"📝 Página {i+1}: '{nome_original}' (Matrícula: {matricula_original}) → Marcado 'FALTA' em x={x_final_para_reportlab:.1f}, y={y_final_para_reportlab:.1f}")

                     alunos_encontrados_alguma_vez.add((matricula_original, nome_original.lower()))

             try:
                 if textos_para_inserir_nesta_pagina:
                     overlay_reader = criar_overlay_buffer(textos_para_inserir_nesta_pagina)
                     if overlay_reader.pages:
                         # CORREÇÃO AQUI: Usar [] para acessar elementos de listas
                         overlay_page = overlay_reader.pages[0]
                         pagina_pypdf.merge_page(overlay_page)
                         print(f"✅ Overlay mesclado na página {i+1}.")
                     else:
                         print(f"⚠️ Overlay vazio na página {i+1}, pulando mesclagem.")
                 escritor.add_page(pagina_pypdf)
             except Exception as e:
                 print(f"❌ Erro crítico ao processar/mesclar página {i+1}: {e}. Adicionando página original.")
                 escritor.add_page(leitor.pages[i])

         atualizar_progresso("Salvando o arquivo PDF...")
         try:
             alunos_faltosos_nao_encontrados_final = alunos_do_csv_para_rastrear - alunos_encontrados_alguma_vez
             if alunos_faltosos_nao_encontrados_final:
                 print(f"⚠️ Aviso Final: {len(alunos_faltosos_nao_encontrados_final)} alunos com falta do CSV não foram encontrados em nenhuma página do PDF: {alunos_faltosos_nao_encontrados_final}")

             with open(saida_pdf, "wb") as f:
                 escritor.write(f)
        
             messagebox.showinfo("Sucesso", f"✅ PDF gerado com faltas em:\n{saida_pdf}")
         except Exception as e:
             messagebox.showerror("Erro ao Salvar o PDF Final", f"Ocorreu um erro ao salvar o PDF: {e}\n"
                                                                "Verifique se o arquivo não está aberto em outro programa e se você tem permissão de escrita na pasta.")

     except Exception as e:
         messagebox.showerror("Erro Inesperado no Processamento do PDF", f"Ocorreu um erro crítico durante o processamento do PDF: {e}\n"
                                                   "Verifique se os arquivos estão corretos e se você tem permissão de escrita.")
     finally:
         if pdfplumber_doc:
             pdfplumber_doc.close()
             
def main():
    root_main = tk.Tk()
    root_main.withdraw()

    messagebox.showinfo("Início", "Selecione o arquivo CSV e o PDF de lista de chamada.")

    caminho_csv = selecionar_arquivo("Selecione o arquivo de faltas (CSV)", [("Arquivos CSV", "*.csv")])
    if not caminho_csv:
        return

    faltosos = ler_csv(caminho_csv)
    if not faltosos:
        return

    caminho_pdf = selecionar_arquivo("Selecione o arquivo da lista de chamada (PDF)", [("Arquivos PDF", "*.pdf")])
    if not caminho_pdf:
        return

    default_output_dir = os.path.dirname(caminho_pdf)
    initial_pdf_name = os.path.splitext(os.path.basename(caminho_pdf))[0]
    output_filename = f"{initial_pdf_name}_com_faltas.pdf"

    saida_pdf = os.path.join(default_output_dir, output_filename)

    criar_janela_progresso("Marcando Faltas...")
    marcar_faltas(caminho_pdf, faltosos, saida_pdf)
    fechar_janela_progresso()

if __name__ == "__main__":
    main()