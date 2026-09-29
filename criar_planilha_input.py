# -*- coding: utf-8 -*-
"""
Gera a planilha de alimentação das Salas VIP
============================================
Saída: C:\\Users\\<user>\\SalaVIP_Dashboard\\Base_SalaVIP_Input.xlsx

Estrutura (definida com o usuário):
  · Uma aba por sala VIP  -> GRU DOM, GRU INTER, GIG DOM, GIG INTER
    Se surgir uma sala nova, basta criar outra aba com o mesmo layout:
    o leitor descobre sozinho, não precisa mexer em código.
    Colunas: SALA VIP | PARCEIRA | ANO | MÊS | PAX | CUSTO | MOEDA
    TODAS as companhias aparecem em TODAS as salas. As que não operam
    naquela sala vêm com "-" em PAX e CUSTO e ficam ocultas no relatório.
    Para ativar uma companhia, basta trocar o "-" por um número.
  · PAX PAGANTES -> alimentada pelo usuário (total da sala menos GOL)
  · APOIO        -> médias mensais de EUR/USD (saída do fx_collector.py)
  · LEIA-ME      -> instruções

Todo o histórico existente já vem preenchido, para não perder continuidade.
"""
import openpyxl, os, sys, io, json, glob, collections, datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# Source folder for the VIP Lounge base workbook.
# Configure via the VIP_LOUNGE_SOURCE_PATH environment variable; the fallback
# below is a generic placeholder, not a real network path.
PASTA_G = os.environ.get(
    "VIP_LOUNGE_SOURCE_PATH",
    r"G:\Commercial\Sales_Performance\International\VIP_Lounge",
)


def base_mais_recente():
    """
    Sempre a base mais nova da pasta do G:.

    Antes o caminho estava fixo numa base específica, e cada mês novo salvo
    no G: era simplesmente ignorado — foi o que fez julho/2026 não aparecer.
    A ordenação é pelo nome (AAMMDD_...), com a data de modificação como
    desempate, então uma base recém-salva sempre vence.
    """
    cands = glob.glob(os.path.join(PASTA_G, "*BaseSalaVIP - USD.xlsx"))
    cands = [c for c in cands if not os.path.basename(c).startswith("~$")]
    if not cands:
        raise SystemExit("Nenhuma base '*BaseSalaVIP - USD.xlsx' em %s" % PASTA_G)
    return max(cands, key=lambda c: (os.path.basename(c), os.path.getmtime(c)))


BASE = base_mais_recente()

# Output folder for the generated input workbook and support files.
# Configure via the VIP_LOUNGE_OUTPUT_PATH environment variable.
_OUT_ROOT = os.environ.get(
    "VIP_LOUNGE_OUTPUT_PATH",
    r"F:\Commercial\Sales_Performance\Daily_Sales_Tracking\VIP_Lounge",
)
SAIDA = os.path.join(_OUT_ROOT, "3. Apoio")
DEST = os.path.join(_OUT_ROOT, "1. Base", "Base_SalaVIP_Input.xlsx")
FXJSON = os.path.join(SAIDA, "fx_eurusd.json")

MESES = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"]
SALAS = ["GRU DOM", "GRU INTER", "GIG DOM", "GIG INTER"]
PARCEIRAS = ["AVIANCA", "AIR FRANCE", "KLM", "BRITISH AIRWAYS", "ANGOLA AIRLINES",
             "UNITED AIRLINES", "ETHIOPIAN AIRLINES", "AIR EUROPA", "AEROMEXICO"]

# Uma companhia entra como "-" numa sala apenas se NUNCA teve acesso ali.
# Qualquer histórico, por menor que seja, mantém a companhia ativa — assim a
# planilha nunca descarta um dado que a base já tem (ex.: Ethiopian na GRU DOM,
# que aparece em poucos meses mas existe). O usuário pode marcar "-" à mão.
MIN_MESES = 1

LARANJA = "FF7020"; CINZA = "F2F2F2"; AZUL = "1F4E79"
f_hdr = Font(bold=True, color="FFFFFF", size=10)
p_hdr = PatternFill("solid", fgColor=LARANJA)
p_alt = PatternFill("solid", fgColor="FBFBF9")
p_na  = PatternFill("solid", fgColor=CINZA)
p_inp = PatternFill("solid", fgColor="FFF7E6")
thin  = Side(style="thin", color="D9D9D9")
bd    = Border(left=thin, right=thin, top=thin, bottom=thin)


def ler_base():
    wb = openpyxl.load_workbook(BASE, data_only=True, read_only=True)
    linhas = list(wb["Planilha1"].iter_rows(values_only=True)); wb.close()
    hdr = [str(h).strip() for h in linhas[0] if h is not None]
    c = {h: i for i, h in enumerate(hdr)}
    out = []
    for r in linhas[1:]:
        if r[0] is None: continue
        out.append({"sala": str(r[c["SALA VIP"]]).strip(),
                    "parc": str(r[c["PARCEIRA"]]).strip(),
                    "ano": int(r[c["ANO"]]), "mes": str(r[c["MÊS"]]).strip(),
                    "pax": int(r[c["PAX"]] or 0), "fee": float(r[c["CUSTO"]] or 0),
                    "moeda": str(r[c["MOEDA"]] or "USD").strip()})
    return out


def ler_pax_pagantes_do_input():
    """
    Le o que JA ESTA na aba PAX PAGANTES da planilha de alimentacao.

    Este script reescreve a planilha do zero a partir da base do G:. Sem esta
    leitura, tudo que o usuario digitou a mao na aba PAX PAGANTES (que NAO
    existe na base do G:) seria perdido a cada regeracao. O que o usuario
    digitou tem prioridade sobre o historico da Planilha3.
    """
    if not os.path.exists(DEST):
        return {}
    try:
        wb = openpyxl.load_workbook(DEST, data_only=True, read_only=True)
        if "PAX PAGANTES" not in wb.sheetnames:
            wb.close(); return {}
        ws = wb["PAX PAGANTES"]
        cab = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
        salas = [str(c).strip() for c in cab[2:] if c]
        out = {}
        for r in ws.iter_rows(min_row=5, values_only=True):
            if not r[0] or not r[1]:
                continue
            ano, mes = int(r[0]), str(r[1]).strip()
            for i, sala in enumerate(salas):
                v = r[2 + i] if 2 + i < len(r) else None
                if isinstance(v, (int, float)) and v > 0:
                    out[(sala, ano, mes)] = int(v)
        wb.close()
        return out
    except Exception as e:
        print("  AVISO: nao consegui ler os PAX PAGANTES atuais (%s)." % e)
        print("  Preservacao desligada nesta rodada - confira a aba antes de salvar.")
        return {}


def ler_pax_pagantes():
    """Puxa a tabela que ja existe na Planilha3 (r92..r107) para nao perder historico."""
    import zipfile, re
    import xml.etree.ElementTree as ET
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    NSR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    z = zipfile.ZipFile(BASE)
    wbx = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rmap = {r.get("Id"): r.get("Target") for r in rels}
    alvo = None
    for sh in wbx.find(NS + "sheets"):
        if sh.get("name") == "Planilha3":
            t = rmap[sh.get(NSR + "id")]
            alvo = t if t.startswith("xl/") else "xl/" + t.lstrip("/")
    shared = ["".join(t.text or "" for t in si.iter(NS + "t"))
              for si in ET.fromstring(z.read("xl/sharedStrings.xml"))]
    cel = collections.defaultdict(dict)
    for row in ET.fromstring(z.read(alvo)).iter(NS + "row"):
        r = int(row.get("r"))
        if not (90 <= r <= 108): continue
        for c in row:
            col = re.match(r"([A-Z]+)", c.get("r")).group(1)
            t, v = c.get("t"), c.find(NS + "v")
            if v is None: continue
            val = v.text
            if t == "s":
                try: val = shared[int(val)]
                except: pass
            else:
                try: val = float(val)
                except: pass
            cel[r][col] = val
    COLS = ["B","C","D","E","F","G","H","I","J","K","L","M"]
    MAPA = {"GRU INTER": (93, 100), "GRU DOM": (94, 101),
            "GIG DOM": (95, 102), "GIG INTER": (96, 103)}
    out = {}
    for sala, (r25, r26) in MAPA.items():
        for ano, rr in ((2025, r25), (2026, r26)):
            for i, col in enumerate(COLS):
                v = cel.get(rr, {}).get(col)
                if isinstance(v, float) and v > 0:
                    out[(sala, ano, MESES[i])] = int(v)
    return out


def ler_apoio_legado():
    """
    Médias mensais EUR/USD como estão hoje na aba Apoio da BaseSalaVIP.
    São ELAS que reproduzem o PPT já publicado (erros de digitação inclusive),
    então viram a coluna 'MÉDIA (como a planilha antiga)'. A coluna corrigida
    vem do fx_collector. Meses futuros, sem histórico, usam só a corrigida.
    """
    import zipfile, re
    import xml.etree.ElementTree as ET
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    NSR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    z = zipfile.ZipFile(BASE)
    wbx = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rmap = {r.get("Id"): r.get("Target") for r in rels}
    alvo = None
    for sh in wbx.find(NS + "sheets"):
        if sh.get("name") == "Apoio":
            t = rmap[sh.get(NSR + "id")]
            alvo = t if t.startswith("xl/") else "xl/" + t.lstrip("/")
    shared = ["".join(t.text or "" for t in si.iter(NS + "t"))
              for si in ET.fromstring(z.read("xl/sharedStrings.xml"))]
    out = {}
    for row in ET.fromstring(z.read(alvo)).iter(NS + "row"):
        f = g = None
        for c in row:
            col = re.match(r"([A-Z]+)", c.get("r")).group(1)
            if col not in ("F", "G"):
                continue
            t, v = c.get("t"), c.find(NS + "v")
            if v is None:
                continue
            val = v.text
            if t == "s":
                try: val = shared[int(val)]
                except: pass
            else:
                try: val = float(val)
                except: pass
            if col == "F": f = val
            else: g = val
        if isinstance(f, str) and "/" in f and isinstance(g, float):
            out[f] = round(g, 6)
    return out


def main():
    print("Base do G: %s" % os.path.basename(BASE))
    dados = ler_base()
    # historico da Planilha3 + o que o usuario ja digitou (este ultimo vence)
    pagantes = ler_pax_pagantes()
    digitados = ler_pax_pagantes_do_input()
    if digitados:
        novos = {k: v for k, v in digitados.items()
                 if pagantes.get(k) != v}
        pagantes.update(digitados)
        print("PAX PAGANTES preservados da planilha atual: %d células "
              "(%d diferentes do histórico)" % (len(digitados), len(novos)))
    legado = ler_apoio_legado()
    print("Médias EUR/USD herdadas da Apoio original: %d meses" % len(legado))
    print("Base lida: %d linhas | PAX pagantes historicos: %d celulas"
          % (len(dados), len(pagantes)))

    # --- quem e regular em cada sala ---
    meses_ativos = collections.defaultdict(set)
    fee_de = {}
    for d in dados:
        if d["pax"] > 0:
            meses_ativos[(d["sala"], d["parc"])].add((d["ano"], d["mes"]))
        if d["fee"]:
            fee_de[d["parc"]] = (d["fee"], d["moeda"])
    regular = {k: len(v) >= MIN_MESES for k, v in meses_ativos.items()}

    print("\nCompanhias regulares por sala (>= %d meses com acesso):" % MIN_MESES)
    for s in SALAS:
        reg = [p for p in PARCEIRAS if regular.get((s, p))]
        na  = [p for p in PARCEIRAS if not regular.get((s, p))]
        print("  %-10s ativas: %s" % (s, ", ".join(reg)))
        print("  %-10s  ' - ' : %s" % ("", ", ".join(na)))

    existente = {(d["sala"], d["parc"], d["ano"], d["mes"]): d for d in dados}
    periodos = sorted({(d["ano"], MESES.index(d["mes"])) for d in dados})

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- LEIA-ME ----------------
    ws = wb.create_sheet("LEIA-ME")
    txt = [
        ("COMO ALIMENTAR ESTA PLANILHA", 14, True),
        ("", 10, False),
        ("1. ABAS DE SALA VIP  (GRU DOM, GRU INTER, GIG DOM, GIG INTER)", 11, True),
        ("   Preencha somente as colunas PAX e CUSTO. As demais ja vem prontas.", 10, False),
        ("   Cada mes novo: copie o bloco do mes anterior, troque ANO/MES e atualize PAX.", 10, False),
        ("", 10, False),
        ("   CELULAS COM ' - '  =  a companhia nao opera naquela sala.", 10, True),
        ("   Elas NAO aparecem no relatorio. Se a companhia passar a vender,", 10, False),
        ("   basta trocar o ' - ' por um numero: ela passa a ser exibida automaticamente.", 10, False),
        ("", 10, False),
        ("2. SALA NOVA", 11, True),
        ("   Crie uma aba com o nome da sala e o mesmo layout de colunas.", 10, False),
        ("   O relatorio le sozinho, nao precisa alterar nada no codigo.", 10, False),
        ("", 10, False),
        ("3. ABA 'PAX PAGANTES'", 11, True),
        ("   Alimentada por voce. E o total de acessos da sala no mes MENOS os acessos GOL.", 10, False),
        ("   Exemplo: sala com 23.100 acessos no mes e 5.462 da GOL -> pagantes = 17.638.", 10, False),
        ("   Esse numero e o denominador do indicador 'Paying pax %' do relatorio.", 10, False),
        ("   Sem ele preenchido, o indicador daquele mes simplesmente nao e exibido.", 10, False),
        ("", 10, False),
        ("4. ABA 'APOIO'  (cambio EUR/USD)", 11, True),
        ("   Gerada pelo script fx_collector.py - nao preencha a mao.", 10, False),
        ("   Traz duas colunas: a media como a planilha antiga calculava e a media corrigida", 10, False),
        ("   (sem data inexistente, sem dia repetido e com os valores gravados como texto).", 10, False),
        ("   O relatorio mostra as duas versoes lado a lado.", 10, False),
        ("", 10, False),
        ("REGRA DE OURO: nunca apague colunas nem renomeie cabecalhos.", 11, True),
    ]
    for i, (t, sz, b) in enumerate(txt, start=1):
        c = ws.cell(i, 1, t); c.font = Font(size=sz, bold=b,
                                            color=AZUL if b and sz >= 11 else "000000")
    ws.column_dimensions["A"].width = 100

    # ---------------- abas de sala ----------------
    HDR = ["SALA VIP", "PARCEIRA", "ANO", "MÊS", "PAX", "CUSTO", "MOEDA"]
    for sala in SALAS:
        ws = wb.create_sheet(sala)
        for j, h in enumerate(HDR, start=1):
            c = ws.cell(1, j, h); c.font = f_hdr; c.fill = p_hdr
            c.alignment = Alignment(horizontal="center"); c.border = bd
        r = 2
        for (ano, mi) in periodos:
            mes = MESES[mi]
            for parc in PARCEIRAS:
                reg = regular.get((sala, parc), False)
                d = existente.get((sala, parc, ano, mes))
                fee, moeda = fee_de.get(parc, (0, "USD"))
                vals = [sala, parc, ano, mes,
                        (d["pax"] if d else 0) if reg else "-",
                        (d["fee"] if d and d["fee"] else fee) if reg else "-",
                        (d["moeda"] if d else moeda)]
                for j, v in enumerate(vals, start=1):
                    c = ws.cell(r, j, v); c.border = bd
                    if j in (5, 6):
                        c.alignment = Alignment(horizontal="right")
                        c.fill = p_na if not reg else p_inp
                    elif r % 2 == 0:
                        c.fill = p_alt
                r += 1
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:G{r-1}"
        for col, w in zip("ABCDEFG", (13, 22, 8, 8, 10, 10, 9)):
            ws.column_dimensions[col].width = w
        dv = DataValidation(type="list", formula1='"USD,EUR"', allow_blank=False)
        ws.add_data_validation(dv); dv.add(f"G2:G{r-1}")
        ws.conditional_formatting.add(
            f"E2:F{r-1}",
            CellIsRule(operator="equal", formula=['"-"'],
                       fill=PatternFill("solid", fgColor=CINZA),
                       font=Font(color="A6A6A6")))
        print("  aba %-10s -> %d linhas" % (sala, r - 2))

    # ---------------- PAX PAGANTES ----------------
    ws = wb.create_sheet("PAX PAGANTES")
    ws.cell(1, 1, "PAX PAGANTES POR SALA  (total de acessos da sala no mês MENOS os acessos GOL)"
            ).font = Font(bold=True, size=12, color=AZUL)
    ws.cell(2, 1, "Preencha as células laranja. Mês sem número fica sem o indicador de "
                  "paying pax no relatório.").font = Font(size=9, italic=True, color="808080")
    HDR2 = ["ANO", "MÊS"] + SALAS
    for j, h in enumerate(HDR2, start=1):
        c = ws.cell(4, j, h); c.font = f_hdr; c.fill = p_hdr
        c.alignment = Alignment(horizontal="center"); c.border = bd
    r = 5
    for (ano, mi) in periodos:
        mes = MESES[mi]
        ws.cell(r, 1, ano).border = bd
        ws.cell(r, 2, mes).border = bd
        for j, sala in enumerate(SALAS, start=3):
            v = pagantes.get((sala, ano, mes))
            c = ws.cell(r, j, v if v else None)
            c.border = bd; c.fill = p_inp
            c.alignment = Alignment(horizontal="right")
        r += 1
    ws.freeze_panes = "A5"
    for col, w in zip("ABCDEF", (8, 8, 14, 14, 14, 14)):
        ws.column_dimensions[col].width = w
    print("  aba PAX PAGANTES -> %d linhas (%d ja preenchidas)"
          % (r - 5, len(pagantes)))

    # ---------------- APOIO ----------------
    ws = wb.create_sheet("APOIO")
    ws.cell(1, 1, "CÂMBIO EUR/USD - MÉDIA MENSAL").font = Font(bold=True, size=12, color=AZUL)
    ws.cell(2, 1, "Gerado por fx_collector.py. Não editar à mão."
            ).font = Font(size=9, italic=True, color="808080")
    HDR3 = ["MÊS/ANO", "MÉDIA (como a planilha antiga)", "MÉDIA CORRIGIDA",
            "DIAS", "FONTE", "ALERTAS"]
    for j, h in enumerate(HDR3, start=1):
        c = ws.cell(4, j, h); c.font = f_hdr; c.fill = p_hdr
        c.alignment = Alignment(horizontal="center", wrap_text=True); c.border = bd
    r = 5
    if os.path.exists(FXJSON):
        fx = json.load(open(FXJSON, encoding="utf-8"))
        for m in fx["meses"]:
            cor = m["media_corrigida"]
            # a coluna "publicada" é a taxa que a planilha antiga já usava — é ela
            # que amarra o relatório no PPT. Mês novo, sem histórico, herda a corrigida.
            pub = legado.get(m["mes_ano"], cor)
            ws.cell(r, 1, m["mes_ano"]).border = bd
            ws.cell(r, 2, pub).number_format = "0.000000"
            ws.cell(r, 3, cor).number_format = "0.000000"
            ws.cell(r, 4, m["dias"])
            ws.cell(r, 5, m["fonte"])
            ws.cell(r, 6, " | ".join(m["alertas"]) if m["alertas"] else "")
            for j in range(1, 7): ws.cell(r, j).border = bd
            if m["alertas"]:
                ws.cell(r, 6).font = Font(color="C00000", size=9)
            r += 1
        print("  aba APOIO -> %d meses" % (r - 5))
    else:
        ws.cell(5, 1, "Rode fx_collector.py primeiro.").font = Font(color="C00000")
    ws.freeze_panes = "A5"
    for col, w in zip("ABCDEF", (12, 26, 18, 8, 14, 60)):
        ws.column_dimensions[col].width = w

    # backup antes de sobrescrever: este script reescreve a planilha inteira
    if os.path.exists(DEST):
        import shutil
        bkp_dir = os.path.join(os.path.dirname(DEST), "_backup")
        os.makedirs(bkp_dir, exist_ok=True)
        bkp = os.path.join(bkp_dir, "Base_SalaVIP_Input_%s.xlsx"
                           % datetime.datetime.now().strftime("%y%m%d_%H%M%S"))
        shutil.copy2(DEST, bkp)
        print("Backup da planilha anterior: %s" % os.path.basename(bkp))
        # mantem só os 10 backups mais recentes
        antigos = sorted(glob.glob(os.path.join(bkp_dir, "Base_SalaVIP_Input_*.xlsx")))
        for a in antigos[:-10]:
            try: os.remove(a)
            except OSError: pass

    wb.save(DEST)
    print("\nOK -> %s  (%.0f KB)" % (DEST, os.path.getsize(DEST) / 1024))


if __name__ == "__main__":
    main()
