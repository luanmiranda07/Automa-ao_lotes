import pandas as pd
from datetime import date
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import unicodedata, re, difflib
import numpy as np

# ------------------------------
# CONFIGURAÇÃO: ARQUIVO MODELO PADRÃO
# ------------------------------
# Deixe esse arquivo na mesma pasta do script (ou coloque o caminho completo)
ARQUIVO_MODELO = "testesLotes.xlsx"

# ------------------------------
# Funções de leitura e geração
# ------------------------------

def _norm(s: str) -> str:
    """Normaliza string: remove NBSP, trim, colapsa espaços, remove acentos, lower."""
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = s.replace("\u00A0", " ")            # NBSP -> espaço normal
    s = re.sub(r"\s+", " ", s).strip()      # colapsa e tira espaços nas pontas
    s_nfkd = unicodedata.normalize("NFKD", s)
    s_noacc = "".join(c for c in s_nfkd if not unicodedata.combining(c))
    return s_noacc.lower()

def _normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica strip/colapso de espaço em todos os cabeçalhos."""
    df = df.copy()
    new_cols = []
    for c in df.columns:
        cs = str(c).replace("\u00A0", " ")
        cs = re.sub(r"\s+", " ", cs).strip()
        new_cols.append(cs)
    df.columns = new_cols
    return df

def _find_best_column(df: pd.DataFrame, user_text: str | None) -> str | None:
    """Tenta mapear o texto digitado pelo usuário para uma coluna do df."""
    if not user_text:
        return None

    cols = list(df.columns)
    # 1) Igualdade direta
    if user_text in cols:
        return user_text

    # 2) Igualdade case-insensitive + strip
    us = user_text.replace("\u00A0", " ").strip()
    for c in cols:
        if c.replace("\u00A0", " ").strip().lower() == us.lower():
            return c

    # 3) Sem acentos
    col_norm_map = {_norm(c): c for c in cols}
    usn = _norm(user_text)
    if usn in col_norm_map:
        return col_norm_map[usn]

    # 4) Fuzzy match (normalizado)
    match = difflib.get_close_matches(usn, list(col_norm_map.keys()), n=1, cutoff=0.8)
    if match:
        return col_norm_map[match[0]]

    # 5) Substring (normalizado)
    for kn, orig in col_norm_map.items():
        if usn in kn or kn in usn:
            return orig

    return None



def carregar_lote(caminho_lote: str, nome_coluna_processo: str | None = None) -> pd.DataFrame:
    """
    Lê a planilha LOTE e tenta descobrir automaticamente qual linha é o cabeçalho.
    - Primeiro tenta achar a linha onde está a coluna do processo.
    - Se não achar, pega a primeira linha "cheia" de dados como cabeçalho.
    Em seguida, normaliza os cabeçalhos (strip/colapso de espaços).
    """
    raw = pd.read_excel(caminho_lote, header=None)

    alvo = (nome_coluna_processo.strip().lower() if nome_coluna_processo
            else "número do processo")

    header_row_idx = None

    # 1) Pela coluna de processo
    for i, row in raw.iterrows():
        linha_str = row.astype(str).str.strip().str.lower()
        if linha_str.eq(alvo).any():
            header_row_idx = i
            break

    # 2) Fallback: primeira linha "cheia"
    if header_row_idx is None:
        for i, row in raw.iterrows():
            if row.count() >= 3:
                header_row_idx = i
                break

    if header_row_idx is None:
        raise ValueError("Não consegui localizar automaticamente a linha de cabeçalho no arquivo de lote.")

    header = raw.iloc[header_row_idx]
    dados = raw.iloc[header_row_idx + 1:].copy()
    dados.columns = header
    dados = dados.dropna(axis=1, how="all").dropna(how="all")

    # >>> NORMALIZA CABEÇALHOS (remove espaços/acentos problemáticos)
    dados = _normalize_headers(dados)

    return dados




def carregar_modelo(caminho_modelo: str) -> list:
    """Lê o modelo para obter a estrutura de colunas."""
    modelo = pd.read_excel(caminho_modelo)
    return list(modelo.columns)


def montar_saida(dados_lote, colunas_modelo, coluna_processo,
                 evento_integracao_val, evento_map, solicitado_por):
    """
    Cria o DataFrame de saída com base no evento informado,
    usando os nomes de colunas informados pelo usuário, com correspondência robusta.
    """
    saida = pd.DataFrame(columns=colunas_modelo)

    # PROCESSO
    col_proc = _find_best_column(dados_lote, coluna_processo)
    if col_proc:
        saida["PROCESSO"] = dados_lote[col_proc].values
        print(f"PROCESSO <- '{col_proc}'")
    else:
        print(f"⚠️ Coluna de processo '{coluna_processo}' não encontrada no lote.")

    # EVENTO (valor financeiro/coluna específica do lote)
    if evento_integracao_val in evento_map:
        coluna_origem_digitada = evento_map[evento_integracao_val]
        col_evt = _find_best_column(dados_lote, coluna_origem_digitada)
        if col_evt:
            saida["EVENTO"] = dados_lote[col_evt].values
            print(f"EVENTO[{evento_integracao_val}] <- '{col_evt}' (entrada: '{coluna_origem_digitada}')")
        else:
            print(f"⚠️ Coluna '{coluna_origem_digitada}' não encontrada no lote para o evento {evento_integracao_val}.")
    else:
        print(f"⚠️ Evento '{evento_integracao_val}' não está mapeado.")

    # DATA atual
    if "DATA" in saida.columns:
        saida["DATA"] = date.today().strftime("%d/%m/%Y")

    # RESULT sempre "OK"
    if "RESULT" in saida.columns:
        saida["RESULT"] = "OK"

    # SOLICITADO_POR
    if "SOLICITADO_POR" in saida.columns:
        saida["SOLICITADO_POR"] = solicitado_por

    # EVENTO_INTEGRACAO = nome do evento
    if "EVENTO_INTEGRACAO" in saida.columns:
        saida["EVENTO_INTEGRACAO"] = evento_integracao_val

    return saida



def gerar_arquivos():
    """Função acionada pelo botão da interface."""
    caminho_lote = entry_lote.get().strip()

    if not caminho_lote:
        messagebox.showerror("Erro", "Selecione o arquivo em lote.")
        return

    # Verifica se o modelo padrão existe
    caminho_modelo = Path(ARQUIVO_MODELO)
    if not caminho_modelo.is_file():
        messagebox.showerror(
            "Erro",
            f"Arquivo modelo padrão '{ARQUIVO_MODELO}' não foi encontrado.\n"
            f"Deixe-o na mesma pasta do script ou ajuste o caminho em ARQUIVO_MODELO."
        )
        return

    try:
        dados_lote = carregar_lote(caminho_lote)
        colunas_modelo = carregar_modelo(str(caminho_modelo))
    except Exception as e:
        messagebox.showerror("Erro ao ler arquivos", str(e))
        return

    # Pega os nomes das colunas digitados
    coluna_processo = entry_col_processo.get().strip()
    col_hc30 = entry_col_hc30.get().strip()
    col_hcp = entry_col_hcp.get().strip()
    col_calcs = entry_col_calcs.get().strip()
    col_hsp = entry_col_hsp.get().strip()
    col_calcp = entry_col_calcp.get().strip()
    solicitado_por = entry_solicitado_por.get().strip() or "45270"

    # Mapa evento -> coluna correspondente no lote
    evento_map = {
        "HC30%": col_hc30,
        "HCP": col_hcp,
        "CALCS": col_calcs,
        "HSP": col_hsp,
        "CALCP": col_calcp,
    }

    # Pasta base para salvar (mesma pasta do modelo padrão)
    base_path = caminho_modelo
    pasta_saida = base_path.parent
    nome_base = base_path.stem  # sem extensão

    eventos_integracao = ["HC30%", "HCP", "CALCS", "HSP", "CALCP"]

    try:
        for evento in eventos_integracao:
            saida = montar_saida(
                dados_lote,
                colunas_modelo,
                coluna_processo=coluna_processo,
                evento_integracao_val=evento,
                evento_map=evento_map,
                solicitado_por=solicitado_por,
            )

            arquivo_saida = pasta_saida / f"1 Cópia de modelo rb 03 - {evento} preenchido.xlsx"
            saida.to_excel(arquivo_saida, index=False)
            print(f"✅ Arquivo gerado: {arquivo_saida}")

        messagebox.showinfo("Sucesso", "Arquivos gerados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro ao gerar arquivos", str(e))


# ------------------------------
# Interface Tkinter
# ------------------------------

def selecionar_lote():
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo em lote",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
    if caminho:
        entry_lote.delete(0, tk.END)
        entry_lote.insert(0, caminho)


# Janela principal
root = tk.Tk()
root.title("Gerador de Lotes de Integração")
root.geometry("650x380")

# Linha: arquivo em lote
tk.Label(root, text="Arquivo (LOTE):").grid(row=0, column=0, sticky="w", padx=10, pady=5)
entry_lote = tk.Entry(root, width=60)
entry_lote.grid(row=0, column=1, padx=10, pady=5)
btn_lote = tk.Button(root, text="Selecionar...", command=selecionar_lote)
btn_lote.grid(row=0, column=2, padx=5, pady=5)

# (Opcional) aviso do modelo padrão
tk.Label(
    root,
    text=f"Usando modelo padrão: {ARQUIVO_MODELO}",
    fg="gray"
).grid(row=1, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 10))

# Separador visual
tk.Label(root, text="Nomes das colunas no arquivo em lote :", font=("Segoe UI", 10, "bold")).grid(
    row=2, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 5)
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
