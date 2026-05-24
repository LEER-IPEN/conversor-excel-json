import json
import pandas as pd

caminho_planilha = r"planilha/planilha.xlsx"
caminho_json = r"json/resultado.json"

def converter_valores_seguro(valor, e_json=False):
    """Trata valores nulos (NaN/None), remove espaços e converte strings JSON."""
    # Se o valor for nulo, NaN ou uma string escrito 'nan'
    if pd.isna(valor) or valor is None or str(valor).lower().strip() == "nan":
        if e_json:
            return {}  # Retorna dicionário vazio para Majoritarios/Traco
        return ""  # Retorna string vazia para campos de texto comuns

    # Se for texto, limpa espaços invisíveis nas pontas
    if isinstance(valor, str):
        valor = valor.strip()

    # Se a coluna precisar ser um Objeto/JSON (como Majoritarios e Traco)
    if e_json:
        if isinstance(valor, (int, float)):
            return {
                "Valor": valor
            }  # Caso tenha digitado só um número por engano

        if isinstance(valor, str) and valor.startswith("{"):
            try:
                # Corrige aspas inteligentes do Excel (“ ”) para aspas retas (")
                valor_limpo = valor.replace("“", '"').replace("”", '"')
                return json.loads(valor_limpo)
            except json.JSONDecodeError:
                return {}  # Se o texto do JSON estiver inválido, evita travar e retorna vazio
        elif isinstance(valor, str) and not valor:
            return {}

    # Trata formatação de números inteiros para não virem como float (ex: ano 2010 em vez de 2010.0)
    if isinstance(valor, float) and valor.is_integer():
        return int(valor)

    return valor

# 1. LER A PLANILHA
# Usamos o argumento (skiprows) caso sua planilha tenha linhas vazias no topo antes do cabeçalho real.
# Se o seu cabeçalho estiver logo na primeira linha, o Pandas ignora linhas totalmente vazias automaticamente.
df = pd.read_excel(caminho_planilha)

# Limpeza opcional nos nomes das colunas da planilha para garantir que não tenham espaços escondidos
df.columns = df.columns.str.strip()

json_final = []

# 2. MAPEAMENTO MANUAL DOS CAMPOS
for _, linha in df.iterrows():
    # Cria a estrutura exata exigida pelo seu JSON exemplo
    estrutura = {
        "id": str(converter_valores_seguro(linha.get("id"))),
        "filename": str(converter_valores_seguro(linha.get("filename"))),
        "tecnica": {
            "sigla": str(converter_valores_seguro(linha.get("tecnica_sigla"))),
            "nome": str(converter_valores_seguro(linha.get("tecnica_nome"))),
        },
        "nome": str(converter_valores_seguro(linha.get("nome"))),
        "rotulo": str(converter_valores_seguro(linha.get("rotulo"))),
        "decada": str(converter_valores_seguro(linha.get("decada"))),
        "medium": str(converter_valores_seguro(linha.get("medium"))),
        "elementos": {
            "Ag": {
                "percentual_traco": converter_valores_seguro(
                    linha.get("ag_percentual")
                ),
                "Majoritarios": converter_valores_seguro(
                    linha.get("ag_majoritarios"), e_json=True
                ),
                "Traco": converter_valores_seguro(
                    linha.get("ag_traco"), e_json=True
                ),
            },
            "Rh": {
                "percentual_traco": converter_valores_seguro(
                    linha.get("rh_percentual")
                ),
                "Majoritarios": converter_valores_seguro(
                    linha.get("rh_majoritarios"), e_json=True
                ),
                "Traco": converter_valores_seguro(
                    linha.get("rh_traco"), e_json=True
                ),
            },
            "Au": {
                "percentual_traco": converter_valores_seguro(
                    linha.get("au_percentual")
                ),
                "Majoritarios": converter_valores_seguro(
                    linha.get("au_majoritarios"), e_json=True
                ),
                "Traco": converter_valores_seguro(
                    linha.get("au_traco"), e_json=True
                ),
            },
        },
        "imagem": str(converter_valores_seguro(linha.get("imagem"))),
        "espectros": {
            "Ag": str(converter_valores_seguro(linha.get("espectros_ag"))),
            "Rh": str(converter_valores_seguro(linha.get("espectros_rh"))),
            "Au": str(converter_valores_seguro(linha.get("espectros_au"))),
        },
    }
    json_final.append(estrutura)

# 3. SALVAR O ARQUIVO JSON
with open(caminho_json, "w", encoding="utf-8") as f:
    json.dump(json_final, f, indent=4, ensure_ascii=False)

print("JSON gerado com sucesso")
