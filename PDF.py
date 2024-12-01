import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from pdf2image import convert_from_path
from pdf2docx import Converter
import fitz
import PyPDF2

def selecionar_pdfs():
    """Abre uma interface gráfica para o usuário selecionar arquivos PDF."""
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos PDF", 
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    return list(arquivos)

def converter_pdf_para_png(pdf_path, output_folder):
    """Converte um arquivo PDF em imagens PNG, uma para cada página."""
    try:
        doc = fitz.open(pdf_path)
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            image = page.get_pixmap()
            image.save(f"{output_folder}/pagina_{page_num + 1}.png")
        return True
    except Exception as e:
        print(f"Erro ao converter PDF para PNG: {e}")
        return False

def converter_pdf_para_docx(pdf_path, output_path):
    """Converte um arquivo PDF em um arquivo DOCX."""
    try:
        cv = Converter(pdf_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        return True
    except Exception as e:
        print(f"Erro ao converter PDF para DOCX: {e}")
        return False

def merge_pdfs(pdf_paths, output_path):
    """Mescla múltiplos arquivos PDF em um único arquivo PDF."""
    try:
        pdf_writer = PyPDF2.PdfWriter()
        for pdf_path in pdf_paths:
            pdf_reader = PyPDF2.PdfReader(pdf_path)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                pdf_writer.add_page(page)
        with open(output_path, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)
        return True
    except Exception as e:
        print(f"Erro ao mesclar PDFs: {e}")
        return False

def on_dnd_drop(event):
    """Função para manipular arquivos arrastados para a interface."""
    arquivos = event.data.split()
    for arquivo in arquivos:
        print(f"Arquivo Dropped: {arquivo}")
    processar_arquivos(arquivos)

def processar_arquivos(arquivos, tipo_conversao):
    """Processa os arquivos PDF selecionados e realiza as conversões e o merge."""
    for pdf in arquivos:
        output_folder = os.path.dirname(pdf)
        if tipo_conversao == 'png':
            if converter_pdf_para_png(pdf, output_folder):
                messagebox.showinfo("Sucesso", f"Conversão de {pdf} para PNG concluída!")
            else:
                messagebox.showerror("Erro", f"Falha ao converter {pdf} para PNG.")
        elif tipo_conversao == 'docx':
            docx_output = os.path.splitext(pdf)[0] + ".docx"
            if converter_pdf_para_docx(pdf, docx_output):
                messagebox.showinfo("Sucesso", f"Conversão de {pdf} para DOCX concluída!")
            else:
                messagebox.showerror("Erro", f"Falha ao converter {pdf} para DOCX.")
            
def processar_merge():
    """Função para processar o merge de PDFs."""
    arquivos = selecionar_pdfs()
    if arquivos:
        output_pdf = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")],
            title="Salvar PDF mesclado"
        )
        if output_pdf:
            if merge_pdfs(arquivos, output_pdf):
                messagebox.showinfo("Sucesso", f"PDFs mesclados com sucesso em {output_pdf}")
            else:
                messagebox.showerror("Erro", "Falha ao mesclar PDFs.")

def criar_interface():
    """Cria a interface gráfica do aplicativo."""
    root = TkinterDnD.Tk()
    root.title("Conversor de PDF")

    # Definir o tamanho da janela
    root.geometry("500x500")

    # Label informativo
    label = tk.Label(root, text="Selecione os arquivos PDF para conversão ou mesclagem", font=("Helvetica", 14))
    label.grid(row=0, column=0, columnspan=3, pady=20)

    # Botão de conversão para PNG
    btn_conversao_png = tk.Button(root, text="Converter para PNG", command=lambda: processar_arquivos(selecionar_pdfs(), 'png'))
    btn_conversao_png.grid(row=1, column=0, padx=20, pady=10)

    # Botão de conversão para DOCX
    btn_conversao_docx = tk.Button(root, text="Converter para DOCX", command=lambda: processar_arquivos(selecionar_pdfs(), 'docx'))
    btn_conversao_docx.grid(row=2, column=0, padx=20, pady=10)

    # Botão de mesclagem de PDFs
    btn_merge = tk.Button(root, text="Mesclar PDFs", command=processar_merge)
    btn_merge.grid(row=3, column=0, padx=20, pady=10)

    # Área de Drop (Drag-and-Drop)
    drop_area = tk.Label(root, text="Arraste e solte os arquivos PDF aqui", relief="solid", width=40, height=10, bg="lightgrey")
    drop_area.grid(row=4, column=0, columnspan=3, pady=20)
    drop_area.drop_target_register(DND_FILES)
    drop_area.dnd_bind('<<Drop>>', on_dnd_drop)
    
    # Botão para sair
    btn_sair = tk.Button(root, text="Sair", command=root.quit)
    btn_sair.grid(row=5, column=0, columnspan=3, pady=10)

    # Executar a interface
    root.mainloop()

if __name__ == "__main__":
    criar_interface()
