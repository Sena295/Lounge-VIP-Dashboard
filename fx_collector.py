# -*- coding: utf-8 -*-
"""
Coletor de câmbio EUR/USD para a aba Apoio
==========================================
Substitui a captura manual do investing.com.

Ordem de tentativa (definida com o usuário):
  1. investing.com  -> mesma fonte usada hoje; costuma ser barrada por Cloudflare
  2. BCE            -> taxa oficial diária, API pública, sem chave, histórico completo

O script SEMPRE registra qual fonte alimentou cada mês, para o relatório poder
mostrar isso e ninguém confundir a origem do número.

Auditoria embutida (os 3 problemas encontrados na aba Apoio atual):
  - datas inexistentes  (ex.: 31.04.2026, 31.06.2026 - meses de 30 dias)
  - datas duplicadas    (ex.: 06.06.2026 lançado duas vezes)
  - valores como TEXTO  (ex.: "1,1721<TAB>" - o AVERAGE do Excel ignora em silêncio)

Saída: dois valores por mês
  media_publicada -> reproduz o que a planilha atual calcula (bug inclusive)
  media_corrigida -> remove data inválida/duplicada e converte texto para número
"""
import csv, io, json, os, re, sys, datetime, calendar, collections
import urllib.request, urllib.error

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# Output folder for the collected FX rates.
# Configure via the VIP_LOUNGE_OUTPUT_PATH environment variable.
SAIDA = os.path.join(
    os.environ.get(
        "VIP_LOUNGE_OUTPUT_PATH",
        r"F:\Commercial\Sales_Performance\Daily_Sales_Tracking\VIP_Lounge",
    ),
    "3. Apoio",
)
MESES = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN",
         "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")


# --------------------------------------------------------------- fonte 1
def de_investing(ini, fim):
    """Tenta o investing.com. Retorna {date: rate} ou levanta excecao."""
    url = ("https://api.investing.com/api/financialdata/historical/1?"
           "start-date=%s&end-date=%s&time-frame=Daily"
           % (ini.isoformat(), fim.isoformat()))
    req = urllib.request.Request(url, headers={
        "User-Agent": UA, "Accept": "application/json",
        "domain-id": "br", "Referer": "https://br.investing.com/currencies/eur-usd-historical-data"})
    with urllib.request.urlopen(req, timeout=25) as r:
        payload = json.loads(r.read().decode("utf-8", "replace"))
    out = {}
    for linha in payload.get("data", []):
        cru = linha.get("rowDateTimestamp") or linha.get("rowDate")
        d = None
        # o campo ora vem como epoch, ora como ISO ("2026-08-10T00:00:00Z")
        try:
            d = datetime.date.fromtimestamp(int(cru))
        except (TypeError, ValueError):
            m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(cru))
            if m:
                d = datetime.date(*map(int, m.groups()))
        if d is None:
            continue
        bruto = str(linha.get("last_close", "")).replace(".", "").replace(",", ".")
        try:
            out[d] = float(bruto)
        except ValueError:
            continue
    if not out:
        raise RuntimeError("investing.com respondeu sem linhas utilizaveis")
    return out


# --------------------------------------------------------------- fonte 2
def do_bce(ini, fim):
    """Taxa de referencia diaria do Banco Central Europeu (EUR->USD)."""
    url = ("https://data-api.ecb.europa.eu/service/data/EXR/D.USD.EUR.SP00.A"
           "?format=csvdata&startPeriod=%s&endPeriod=%s" % (ini.isoformat(), fim.isoformat()))
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=30) as r:
        texto = r.read().decode("utf-8", "replace")
    out = {}
    for linha in csv.DictReader(io.StringIO(texto)):
        try:
            out[datetime.date.fromisoformat(linha["TIME_PERIOD"])] = float(linha["OBS_VALUE"])
        except (ValueError, KeyError, TypeError):
            continue
    if not out:
        raise RuntimeError("BCE respondeu sem linhas")
    return out


def coletar(ini, fim):
    for nome, fn in (("investing.com", de_investing), ("BCE", do_bce)):
        try:
            dados = fn(ini, fim)
            print("  fonte OK: %s  (%d cotacoes)" % (nome, len(dados)))
            return dados, nome
        except Exception as e:
            print("  fonte %s indisponivel: %s" % (nome, str(e)[:90]))
    raise SystemExit("Nenhuma fonte de cambio respondeu. Preencha a aba Apoio manualmente.")


# --------------------------------------------------------------- auditoria
def auditar(lancamentos, ano, mes_idx):
    """
    lancamentos: lista de (texto_data, valor_bruto) exatamente como estao na planilha.
    Devolve (media_publicada, media_corrigida, lista_de_problemas).
    """
    ndias = calendar.monthrange(ano, mes_idx + 1)[1]
    problemas, vistos = [], set()
    numeros_excel, numeros_corrigidos = [], []

    for texto, bruto in lancamentos:
        # --- o que o Excel enxerga: texto nao entra no AVERAGE ---
        if isinstance(bruto, (int, float)):
            numeros_excel.append(float(bruto))
            valor = float(bruto)
        else:
            problemas.append("valor como TEXTO em %s (%r) - AVERAGE ignora" % (texto, bruto))
            try:
                valor = float(str(bruto).replace(",", ".").strip())
            except ValueError:
                problemas.append("valor ilegivel em %s (%r) - descartado" % (texto, bruto))
                continue

        # --- validacao da data ---
        m = re.match(r"^(\d{1,2})[./](\d{1,2})[./](\d{4})", str(texto).strip())
        if not m:
            problemas.append("data ilegivel: %r" % texto)
            continue
        dia, mm, aa = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if dia > ndias:
            problemas.append("data inexistente: %s (%s/%d tem %d dias)"
                             % (texto, MESES[mes_idx], ano, ndias))
            continue
        chave = (aa, mm, dia)
        if chave in vistos:
            problemas.append("data duplicada: %s" % texto)
            continue
        vistos.add(chave)
        numeros_corrigidos.append(valor)

    med = lambda v: (sum(v) / len(v)) if v else None
    return med(numeros_excel), med(numeros_corrigidos), problemas


# --------------------------------------------------------------- principal
def main(desde="2024-01-01"):
    ini = datetime.date.fromisoformat(desde)
    fim = datetime.date.today()
    print("Coletando EUR/USD de %s a %s" % (ini, fim))
    diarias, fonte = coletar(ini, fim)

    porm = collections.defaultdict(list)
    for d, v in sorted(diarias.items()):
        porm[(d.year, d.month - 1)].append((d.strftime("%d.%m.%Y"), v))

    linhas = []
    print("\n%-10s %8s %14s %14s %s" % ("MES/ANO", "DIAS", "MEDIA", "FONTE", "ALERTAS"))
    print("-" * 78)
    for (ano, mi) in sorted(porm):
        pub, cor, probs = auditar(porm[(ano, mi)], ano, mi)
        rot = "%s/%d" % (MESES[mi], ano)
        linhas.append({"mes_ano": rot, "ano": ano, "mes": MESES[mi],
                       "media_publicada": round(pub, 6) if pub else None,
                       "media_corrigida": round(cor, 6) if cor else None,
                       "dias": len(porm[(ano, mi)]), "fonte": fonte,
                       "alertas": probs})
        print("%-10s %8d %14.6f %14s %s"
              % (rot, len(porm[(ano, mi)]), cor or 0, fonte,
                 ("%d alerta(s)" % len(probs)) if probs else "-"))
        for p in probs:
            print(" " * 12 + "! " + p)

    os.makedirs(SAIDA, exist_ok=True)
    dest = os.path.join(SAIDA, "fx_eurusd.json")
    with open(dest, "w", encoding="utf-8") as f:
        json.dump({"gerado": datetime.datetime.now().isoformat(timespec="seconds"),
                   "fonte": fonte, "meses": linhas}, f, ensure_ascii=False, indent=1)
    print("\nOK -> %s  (%d meses)" % (dest, len(linhas)))
    return linhas


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "2024-01-01")
