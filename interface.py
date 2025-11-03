import main
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path



def selecionar_lote():
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo em lote",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
    if caminho:
        entry_lote.delete(0, tk.END)
        entry_lote.insert(0, caminho)


def selecionar_modelo():
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo modelo (teste em lote)",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
    if caminho:
        entry_modelo.delete(0, tk.END)
        entry_modelo.insert(0, caminho)


# Janela principal
root = tk.Tk()
root.title("Gerador de Lotes de Integração")
root.geometry("650x420")

# Linha: arquivo em lote
tk.Label(root, text="Arquivo em lote (LOTE 99):").grid(row=0, column=0, sticky="w", padx=10, pady=5)
entry_lote = tk.Entry(root, width=60)
entry_lote.grid(row=0, column=1, padx=10, pady=5)
btn_lote = tk.Button(root, text="Selecionar...", command=selecionar_lote)
btn_lote.grid(row=0, column=2, padx=5, pady=5)

# Linha: arquivo modelo
tk.Label(root, text="Arquivo modelo (teste em lote):").grid(row=1, column=0, sticky="w", padx=10, pady=5)
entry_modelo = tk.Entry(root, width=60)
entry_modelo.grid(row=1, column=1, padx=10, pady=5)
btn_modelo = tk.Button(root, text="Selecionar...", command=selecionar_modelo)
btn_modelo.grid(row=1, column=2, padx=5, pady=5)

# Separador visual
tk.Label(root, text="Nomes das colunas no arquivo em lote (LOTE 99):", font=("Segoe UI", 10, "bold")).grid(
    row=2, column=0, columnspan=3, sticky="w", padx=10, pady=(15, 5)
)

# Linha: coluna PROCESSO
tk.Label(root, text="Coluna do PROCESSO:").grid(row=3, column=0, sticky="w", padx=10, pady=3)
entry_col_processo = tk.Entry(root, width=30)
entry_col_processo.grid(row=3, column=1, sticky="w", padx=10, pady=3)
entry_col_processo.insert(0, "Número do Processo")  # sugestão padrão

# Linha: coluna HC30%
tk.Label(root, text="Coluna para HC30%:").grid(row=4, column=0, sticky="w", padx=10, pady=3)
entry_col_hc30 = tk.Entry(root, width=30)
entry_col_hc30.grid(row=4, column=1, sticky="w", padx=10, pady=3)
entry_col_hc30.insert(0, "Contratual - 30%")

# Linha: coluna HCP
tk.Label(root, text="Coluna para HCP:").grid(row=5, column=0, sticky="w", padx=10, pady=3)
entry_col_hcp = tk.Entry(root, width=30)
entry_col_hcp.grid(row=5, column=1, sticky="w", padx=10, pady=3)
entry_col_hcp.insert(0, "Contratual CHM")

# Linha: coluna CALCS
tk.Label(root, text="Coluna para CALCS:").grid(row=6, column=0, sticky="w", padx=10, pady=3)
entry_col_calcs = tk.Entry(root, width=30)
entry_col_calcs.grid(row=6, column=1, sticky="w", padx=10, pady=3)
entry_col_calcs.insert(0, "Agosto.2025 - SUCUMBENCIA")

# Linha: coluna HSP
tk.Label(root, text="Coluna para HSP:").grid(row=7, column=0, sticky="w", padx=10, pady=3)
entry_col_hsp = tk.Entry(root, width=30)
entry_col_hsp.grid(row=7, column=1, sticky="w", padx=10, pady=3)
entry_col_hsp.insert(0, "Sucumb. Preço")

# Linha: coluna CALCP
tk.Label(root, text="Coluna para CALCP:").grid(row=8, column=0, sticky="w", padx=10, pady=3)
entry_col_calcp = tk.Entry(root, width=30)
entry_col_calcp.grid(row=8, column=1, sticky="w", padx=10, pady=3)
entry_col_calcp.insert(0, "Agosto.2025 - PRINCIPAL")

# Linha: SOLICITADO_POR
tk.Label(root, text="SOLICITADO_POR:").grid(row=9, column=0, sticky="w", padx=10, pady=10)
entry_solicitado_por = tk.Entry(root, width=15)
entry_solicitado_por.grid(row=9, column=1, sticky="w", padx=10, pady=10)
entry_solicitado_por.insert(0, "45270")

# Botão Gerar
btn_gerar = tk.Button(root, text="Gerar Arquivos", command=gerar_arquivos, bg="#4CAF50", fg="white", width=20)
btn_gerar.grid(row=10, column=0, columnspan=3, pady=20)

root.mainloop()