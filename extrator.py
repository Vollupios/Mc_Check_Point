#!/usr/bin/env python3
"""
Extrator de dados de ponto - Marcos Contabilidade

Aceita dois tipos de entrada:

1. Planilha do Cartão Ponto exportada do sistema (.xls / .xlsx)
   - Um bloco por funcionário (Empresa, Nome, Nº Folha, Departamento,
     Horário de Trabalho SEG..DOM e as linhas diárias Ent./Saí.).
   - Os dados são separados por mês (AAAA-MM).

2. HTML (formato antigo): o próprio ponto.html com JavaScript incorporado
   ou HTML tradicional com tabelas.

Uso:
    python extrator.py ponto_ate_m_s_10.xls
        -> gera pontos_todos_funcionarios.json (dados agrupados por mês)

    python extrator.py ponto_ate_m_s_10.xls --html ponto.html
        -> além do JSON, grava os meses no ponto.html (substitui meses já
           existentes com a mesma chave e mantém os demais)

    python extrator.py ponto.html
        -> extrai os dados já incorporados no HTML (comportamento antigo)
"""

import datetime
import io
import json
import os
import re
import sys

MESES_PT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
            "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

DIAS_SEMANA = {"DOM": 0, "SEG": 1, "TER": 2, "QUA": 3, "QUI": 4, "SEX": 5, "SÁB": 6, "SAB": 6}

# Colunas das linhas diárias: Data | Ent.1 | Saí.1 | Ent.2 | Saí.2 | Ent.3 | Saí.3
COLUNAS_SAIDA = (1, 3, 5)  # posições (dentro das 6 batidas) que são saídas

MARCADOR_INICIO = "// <<DADOS_EXTRAIDOS_INICIO>>"
MARCADOR_FIM = "// <<DADOS_EXTRAIDOS_FIM>>"

RE_DATA = re.compile(r"^(\d{2})/(\d{2})/(\d{2})")
RE_HORA = re.compile(r"^\d{1,2}:\d{2}")


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def is_data_valida(texto):
    """Verifica se o texto começa com uma data no formato DD/MM/YY"""
    return isinstance(texto, str) and RE_DATA.match(texto.strip()) is not None


def normalizar_data(data_str):
    """Converte data de DD/MM/YY (com ou sem sufixo ' - seg') para YYYY-MM-DD"""
    m = RE_DATA.match(str(data_str).strip())
    if not m:
        return data_str
    dia, mes, ano = m.groups()
    # Anos 00-30 são 2000-2030, 31-99 são 1931-1999
    ano_completo = ("20" if int(ano) <= 30 else "19") + ano
    return f"{ano_completo}-{mes}-{dia}"


def formatar_batida(valor):
    """Converte uma célula de batida para texto.

    - datetime.time -> "7:23" (mesmo formato dos dados já existentes)
    - "07:30*" (batida lançada manualmente) é mantida como está
    - códigos (FÉRIAS, Feriado, ATESTAD, ...) são devolvidos como texto
    - vazio -> ""
    """
    if valor is None:
        return ""
    if isinstance(valor, datetime.datetime):
        valor = valor.time()
    if isinstance(valor, datetime.time):
        return f"{valor.hour}:{valor.minute:02d}"
    if isinstance(valor, datetime.timedelta):
        total = int(valor.total_seconds() // 60)
        return f"{total // 60}:{total % 60:02d}"
    return str(valor).strip()


def formatar_horario_previsto(valor):
    """Horário previsto sempre com dois dígitos: "07:30"."""
    texto = formatar_batida(valor)
    if RE_HORA.match(texto):
        h, m = texto.rstrip("*").split(":")[:2]
        return f"{int(h):02d}:{m[:2]}"
    return ""


def is_hora(texto):
    return bool(texto) and RE_HORA.match(texto) is not None


def nome_mes(chave):
    ano, mes = chave.split("-")
    return f"{MESES_PT[int(mes) - 1]} {ano}"


# ---------------------------------------------------------------------------
# Planilha do Cartão Ponto (.xls / .xlsx)
# ---------------------------------------------------------------------------

def ler_linhas_planilha(caminho):
    """Lê todas as linhas da primeira aba da planilha como tuplas de valores."""
    with open(caminho, "rb") as f:
        conteudo = f.read()

    # O sistema exporta com extensão .xls, mas o arquivo costuma ser .xlsx
    # (começa com "PK"). O openpyxl recusa a extensão .xls, por isso o
    # arquivo é aberto a partir da memória.
    if conteudo[:2] == b"PK":
        try:
            import openpyxl
        except ImportError:
            sys.exit("Erro: instale o openpyxl  ->  pip install openpyxl")
        wb = openpyxl.load_workbook(io.BytesIO(conteudo), data_only=True, read_only=True)
        ws = wb.worksheets[0]
        linhas = [tuple(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return linhas

    # .xls antigo (formato binário do Excel 97-2003)
    try:
        import xlrd
    except ImportError:
        sys.exit("Erro: arquivo .xls antigo. Instale o xlrd (pip install xlrd) "
                 "ou salve a planilha como .xlsx.")
    book = xlrd.open_workbook(file_contents=conteudo)
    sh = book.sheet_by_index(0)
    linhas = []
    for i in range(sh.nrows):
        linha = []
        for c in sh.row(i):
            if c.ctype == xlrd.XL_CELL_DATE:
                linha.append(xlrd.xldate.xldate_as_datetime(c.value, book.datemode))
            elif c.ctype == xlrd.XL_CELL_EMPTY:
                linha.append(None)
            else:
                linha.append(c.value)
        linhas.append(tuple(linha))
    return linhas


def celula(linha, i):
    return linha[i] if i < len(linha) else None


def processar_planilha(caminho):
    """Processa a planilha do Cartão Ponto.

    Retorna {"2026-01": {"NOME": {...dados do funcionário...}}, ...}
    """
    linhas = ler_linhas_planilha(caminho)

    funcionarios = []
    atual = None
    coletando = False

    for linha in linhas:
        t0 = celula(linha, 0)
        t0s = str(t0).strip() if t0 is not None else ""

        # Cada bloco de funcionário começa na linha "Empresa"
        if t0s == "Empresa":
            atual = {"nome": None, "departamento": None, "folha": None,
                     "admissao": None, "horarios": {}, "dias": []}
            funcionarios.append(atual)
            coletando = False

        if atual is None:
            continue

        # Grade "Horário de Trabalho" (coluna G = dia; H..M = Ent1..Sai3)
        dia_semana = str(celula(linha, 6) or "").strip().upper()
        if dia_semana in DIAS_SEMANA:
            horarios = [formatar_horario_previsto(celula(linha, c)) for c in range(7, 13)]
            horarios = [h for h in horarios if h]
            if horarios:
                atual["horarios"][DIAS_SEMANA[dia_semana]] = {
                    "entry": horarios[0], "leave": horarios[-1]}

        if t0s == "Nome":
            atual["nome"] = str(celula(linha, 1) or "").strip()
        elif t0s == "Departamento":
            atual["departamento"] = str(celula(linha, 1) or "").strip()
        elif t0s.startswith("Nº Folha"):
            atual["folha"] = str(celula(linha, 1) or "").strip()
        elif t0s == "CTPS":
            adm = celula(linha, 4)
            if isinstance(adm, datetime.datetime):
                atual["admissao"] = adm.date()
        elif t0s == "Data":
            coletando = True
        elif coletando and is_data_valida(t0s):
            batidas = [formatar_batida(celula(linha, c)) for c in range(1, 7)]
            atual["dias"].append((normalizar_data(t0s), batidas))
        elif t0s.startswith("(*)"):
            coletando = False

    por_mes = {}
    for f in funcionarios:
        if not f["nome"]:
            continue

        semana = f["horarios"]
        # Horário padrão = o de segunda (ou o primeiro dia útil cadastrado)
        padrao = semana.get(1) or next((semana[d] for d in sorted(semana) if d not in (0, 6)), None)

        dias_por_mes = {}
        for data, batidas in f["dias"]:
            dias_por_mes.setdefault(data[:7], []).append((data, batidas))

        for mes, dias in dias_por_mes.items():
            # Funcionário entra no mês só se houver algo lançado (batida ou
            # ocorrência) e se já tiver sido admitido até o fim do mês.
            if not any(any(b) for _, b in dias):
                continue
            if f["admissao"] and f["admissao"].strftime("%Y-%m") > mes:
                continue

            registros = []
            for data, batidas in dias:
                ent1 = batidas[0]
                if not is_hora(ent1):
                    continue  # férias, feriado, atestado, folga, externo...
                # Saída = última batida do dia, desde que seja uma coluna de saída
                ultima = max((i for i, b in enumerate(batidas) if is_hora(b)), default=None)
                sai2 = batidas[ultima] if ultima in COLUNAS_SAIDA else ""
                registros.append({"date": data, "ent1": ent1, "sai2": sai2})

            dados = {
                "department": f["departamento"],
                "sheetNumber": f["folha"],
                "scheduledEntryTime": padrao["entry"] if padrao else None,
                "scheduledLeaveTime": padrao["leave"] if padrao else None,
                "records": registros,
            }
            # Horário diferente em algum dia da semana (ex.: quinta e sexta)
            if padrao and any(h != padrao for d, h in semana.items() if d not in (0, 6)):
                dados["weeklySchedule"] = {str(d): h for d, h in sorted(semana.items())}

            por_mes.setdefault(mes, {})[f["nome"]] = dados

    return {mes: dict(sorted(por_mes[mes].items())) for mes in sorted(por_mes)}


# ---------------------------------------------------------------------------
# Atualização do ponto.html
# ---------------------------------------------------------------------------

def _json_mes(funcionarios):
    """Um funcionário por linha, como no restante do arquivo."""
    linhas = [f"                {json.dumps(nome, ensure_ascii=False)}: "
              f"{json.dumps(dados, ensure_ascii=False, separators=(', ', ': '))}"
              for nome, dados in funcionarios.items()]
    return "{\n" + ",\n".join(linhas) + "\n            }"


def _ler_meses_extraidos(html):
    """Lê o bloco de meses já gravado entre os marcadores (se existir)."""
    ini = html.find(MARCADOR_INICIO)
    fim = html.find(MARCADOR_FIM)
    if ini == -1 or fim == -1:
        return None
    trecho = html[ini + len(MARCADOR_INICIO):fim]
    m = re.search(r"const extractedMonthsData\s*=\s*(\{.*\});", trecho, re.S)
    return json.loads(m.group(1)) if m else {}


def atualizar_html(caminho_html, dados_por_mes):
    with open(caminho_html, "r", encoding="utf-8") as f:
        html = f.read()

    existentes = _ler_meses_extraidos(html)
    if existentes is None:
        sys.exit(f"Erro: marcadores {MARCADOR_INICIO} / {MARCADOR_FIM} não "
                 f"encontrados em {caminho_html}.")

    meses = {k: v["data"] for k, v in existentes.items()}
    meses.update(dados_por_mes)

    blocos = [f'            "{mes}": {{ "name": "{nome_mes(mes)}", "data": {_json_mes(meses[mes])} }}'
              for mes in sorted(meses)]
    novo = (MARCADOR_INICIO + "\n"
            "            const extractedMonthsData = {\n" + ",\n".join(blocos) + "\n            };\n"
            "            " + MARCADOR_FIM)

    ini = html.find(MARCADOR_INICIO)
    fim = html.find(MARCADOR_FIM) + len(MARCADOR_FIM)
    html = html[:ini] + novo + html[fim:]

    with open(caminho_html, "w", encoding="utf-8") as f:
        f.write(html)
    return sorted(meses)


# ---------------------------------------------------------------------------
# HTML (formato antigo)
# ---------------------------------------------------------------------------

def extrair_texto(td):
    return td.get_text(strip=True) if td else ""


def processar_html(html_file_path):
    """Processa arquivo HTML e extrai dados de funcionários"""
    try:
        with open(html_file_path, "r", encoding="utf-8") as f:
            html_content = f.read()
    except Exception as e:
        print(f"Erro ao ler arquivo: {e}")
        return {}

    meses = _ler_meses_extraidos(html_content)
    if meses:
        print("Dados extraídos do objeto extractedMonthsData.")
        return {mes: v["data"] for mes, v in meses.items()}

    funcionarios = extrair_dados_javascript(html_content)
    if funcionarios:
        print("Dados extraídos do JavaScript incorporado.")
        return funcionarios

    print("Tentando parser HTML tradicional...")
    return processar_html_tradicional(html_content)


def extrair_dados_javascript(html_content):
    """Extrai dados de funcionários do JavaScript incorporado no HTML"""
    patterns = [
        r'const allEmployeesData\s*=\s*(\{[^;]+\});',
        r'allEmployeesData\s*=\s*(\{[^;]+\});',
        r'employeesData\s*=\s*(\{[^;]+\});'
    ]
    for pattern in patterns:
        match = re.search(pattern, html_content, re.DOTALL)
        if match:
            try:
                funcionarios = json.loads(match.group(1))
                if funcionarios and isinstance(funcionarios, dict):
                    return funcionarios
            except json.JSONDecodeError as e:
                print(f"Erro ao decodificar JSON: {e}")
    return {}


def processar_html_tradicional(html_content):
    """Processa HTML tradicional com tabelas"""
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        sys.exit("Erro: instale o beautifulsoup4  ->  pip install beautifulsoup4")

    soup = BeautifulSoup(html_content, 'html.parser')
    funcionarios = {}
    atual = {}
    registros = []
    coletando = False

    def salvar():
        if atual.get("name"):
            funcionarios[atual["name"]] = {
                "department": atual.get("department"),
                "sheetNumber": atual.get("sheetNumber"),
                "scheduledEntryTime": atual.get("scheduledEntryTime"),
                "scheduledLeaveTime": atual.get("scheduledLeaveTime"),
                "records": registros,
            }

    for row in soup.find_all("tr"):
        cells = row.find_all("td")
        if not cells:
            continue
        t0 = extrair_texto(cells[0])

        if t0 == "Nome":
            salvar()
            atual = {"name": extrair_texto(cells[1])}
            registros = []
            coletando = False
        elif t0 == "Departamento" and atual:
            atual["department"] = extrair_texto(cells[1])
        elif "Nº Folha" in t0 and atual:
            atual["sheetNumber"] = extrair_texto(cells[1])
        elif t0 == "SEG" and atual and not atual.get("scheduledEntryTime"):
            atual["scheduledEntryTime"] = extrair_texto(cells[1])
            atual["scheduledLeaveTime"] = extrair_texto(cells[4]) if len(cells) > 4 else ""
        elif t0 == "Data":
            coletando = True
        elif coletando and is_data_valida(t0):
            ent1 = extrair_texto(cells[1]) if len(cells) > 1 else ""
            sai2 = extrair_texto(cells[4]) if len(cells) > 4 else ""
            if ent1 and sai2 and "Feriado" not in ent1:
                registros.append({"date": normalizar_data(t0), "ent1": ent1, "sai2": sai2})

    salvar()
    return funcionarios


# ---------------------------------------------------------------------------

def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return

    entrada = args[0]
    html_destino = None
    if "--html" in args:
        i = args.index("--html")
        html_destino = args[i + 1] if i + 1 < len(args) else "ponto.html"

    if not os.path.exists(entrada):
        sys.exit(f"Erro: arquivo {entrada} não encontrado.")

    print(f"Processando arquivo: {entrada}")
    saida = "pontos_todos_funcionarios.json"

    if entrada.lower().endswith((".xls", ".xlsx", ".xlsm")):
        dados = processar_planilha(entrada)
        if not dados:
            sys.exit("Nenhum funcionário encontrado na planilha.")
        for mes, funcs in dados.items():
            total = sum(len(f["records"]) for f in funcs.values())
            print(f"  {nome_mes(mes):<16} {len(funcs):>3} funcionários  {total:>5} registros")
        with open(saida, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
        print(f"Arquivo salvo em: {saida}")

        if html_destino:
            meses = atualizar_html(html_destino, dados)
            print(f"{html_destino} atualizado. Meses disponíveis: "
                  + ", ".join(nome_mes(m) for m in meses))
        return

    funcionarios = processar_html(entrada)
    if not funcionarios:
        sys.exit("Nenhum funcionário encontrado ou erro no processamento.")
    print(f"Funcionários encontrados: {len(funcionarios)}")
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(funcionarios, f, ensure_ascii=False, indent=2)
    print(f"Arquivo salvo em: {saida}")


if __name__ == "__main__":
    main()
