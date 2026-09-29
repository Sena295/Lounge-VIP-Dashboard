# -*- coding: utf-8 -*-
r"""
VIP Lounges Performance - gerador do relatório HTML
===================================================
Lê  ..\1. Base\Base_SalaVIP_Input.xlsx
Emite ..\4. Relatorio\VIP_Lounges_Performance_YYMMDD.html

Cadeia de cálculo (idêntica à da planilha original):
    unit = SE(MOEDA="EUR"; taxa_do_mês * CUSTO; CUSTO)
    amount = unit * PAX

Duas versões de câmbio convivem no relatório:
    'published' -> taxa como a planilha antiga calculava (amarra no PPT)
    'corrected' -> taxa auditada pelo fx_collector
"""
import openpyxl, json, os, sys, io, datetime, collections

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENTRADA = os.path.join(RAIZ, "1. Base", "Base_SalaVIP_Input.xlsx")
SAIDA = os.path.join(RAIZ, "4. Relatorio")

MESES = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"]
MES_EN = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
MES_LONGO = ["January","February","March","April","May","June",
             "July","August","September","October","November","December"]
NAO_SALA = {"LEIA-ME", "PAX PAGANTES", "APOIO"}

SALA_EN = {"GRU DOM": "GRU Domestic", "GRU INTER": "GRU International",
           "GIG DOM": "GIG Domestic", "GIG INTER": "GIG International"}
PARC_EN = {"AVIANCA": "Avianca", "AIR FRANCE": "Air France",
           "ANGOLA AIRLINES": "TAAG Angola", "UNITED AIRLINES": "United",
           "BRITISH AIRWAYS": "British Airways", "KLM": "KLM",
           "ETHIOPIAN AIRLINES": "Ethiopian", "AIR EUROPA": "Air Europa",
           "AEROMEXICO": "AeroMexico"}


def ler():
    wb = openpyxl.load_workbook(ENTRADA, data_only=True, read_only=True)

    fx = {}
    for r in wb["APOIO"].iter_rows(min_row=5, values_only=True):
        if r[0] and isinstance(r[1], (int, float)):
            fx[str(r[0]).strip()] = {"pub": float(r[1]),
                                     "cor": float(r[2]) if isinstance(r[2], (int, float)) else float(r[1]),
                                     "fonte": r[4] or "", "alerta": r[5] or ""}

    pagantes = {}
    ws = wb["PAX PAGANTES"]
    cab = [c for c in next(ws.iter_rows(min_row=4, max_row=4, values_only=True)) if c]
    salas_pg = [str(c).strip() for c in cab[2:]]
    for r in ws.iter_rows(min_row=5, values_only=True):
        if not r[0]:
            continue
        ano, mes = int(r[0]), str(r[1]).strip()
        for i, sala in enumerate(salas_pg):
            v = r[2 + i]
            if isinstance(v, (int, float)) and v > 0:
                pagantes["%s|%d|%s" % (sala, ano, mes)] = int(v)

    linhas, ignoradas, salas = [], 0, []
    for nome in wb.sheetnames:
        if nome.strip().upper() in NAO_SALA:
            continue
        salas.append(nome.strip())
        for r in wb[nome].iter_rows(min_row=2, values_only=True):
            if not r[0]:
                continue
            sala, parc, ano, mes, pax, fee, moeda = r[:7]
            # "-" marca companhia que não opera nessa sala: fora do visual
            if isinstance(pax, str) or isinstance(fee, str) or pax is None or fee is None:
                ignoradas += 1
                continue
            mes = str(mes).strip()
            chave = "%s/%d" % (mes, int(ano))
            taxa = fx.get(chave, {})
            reg = {"sala": str(sala).strip(), "parc": str(parc).strip(),
                   "ano": int(ano), "mes": mes, "pax": int(pax or 0),
                   "fee": float(fee or 0), "moeda": str(moeda or "USD").strip()}
            for v, k in (("pub", "pub"), ("cor", "cor")):
                t = taxa.get(v)
                unit = (t * reg["fee"]) if (reg["moeda"] == "EUR" and t) else reg["fee"]
                reg["unit_" + k] = round(unit, 6)
                reg["usd_" + k] = round(unit * reg["pax"], 6)
            linhas.append(reg)
    wb.close()
    return linhas, fx, pagantes, salas, ignoradas


def main():
    linhas, fx, pagantes, salas, ignoradas = ler()
    print("Linhas: %d  (ignoradas por '-': %d)" % (len(linhas), ignoradas))
    print("Salas detectadas: %s" % ", ".join(salas))

    anos = sorted({r["ano"] for r in linhas})
    parcs = sorted({r["parc"] for r in linhas},
                   key=lambda p: -sum(r["usd_pub"] for r in linhas if r["parc"] == p))
    tot = collections.defaultdict(lambda: [0.0, 0])
    for r in linhas:
        tot[r["ano"]][0] += r["usd_pub"]
        tot[r["ano"]][1] += r["pax"]
    for a in sorted(tot):
        print("  %d: USD %13s   PAX %7s" % (a, "{:,.2f}".format(tot[a][0]),
                                            "{:,}".format(tot[a][1])))

    contratos = {}
    for r in sorted(linhas, key=lambda x: (x["ano"], MESES.index(x["mes"]))):
        if r["fee"]:
            contratos[r["parc"]] = {"fee": r["fee"], "moeda": r["moeda"]}
    for p in parcs:
        contratos.setdefault(p, {"fee": 0, "moeda": "USD"})
        contratos[p]["pax_hist"] = sum(r["pax"] for r in linhas if r["parc"] == p)

    alertas_fx = [{"mes": k, "alerta": v["alerta"]} for k, v in fx.items() if v["alerta"]]

    meta = {"gerado": datetime.datetime.now().strftime("%d %b %Y · %H:%M"),
            "anos": anos, "salas": salas, "parcs": parcs,
            "meses": MESES, "mes_en": MES_EN, "mes_longo": MES_LONGO,
            "sala_en": SALA_EN, "parc_en": PARC_EN, "contratos": contratos,
            "pagantes": pagantes, "fx": fx, "alertas_fx": alertas_fx,
            "fonte_fx": (list(fx.values())[0]["fonte"] if fx else "")}

    # Inter + IBM Plex Mono embutidas em base64 (mesmas do Painel Integrado).
    # Assim a tipografia é idêntica mesmo com o Google Fonts bloqueado na rede.
    fontes = os.path.join(RAIZ, "3. Apoio", "fontes_embutidas.css")
    css_fontes = ""
    if os.path.exists(fontes):
        with open(fontes, encoding="utf-8") as f:
            css_fontes = f.read()
        print("Fontes embutidas: %.0f KB" % (os.path.getsize(fontes) / 1024))
    else:
        print("AVISO: fontes_embutidas.css nao encontrado - o relatorio vai "
              "depender do Google Fonts. Rode baixa_fontes.py.")

    html = TEMPLATE.replace("/*__FONTES__*/", css_fontes)
    html = html.replace("/*__DADOS__*/",
                        json.dumps(linhas, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("/*__META__*/", json.dumps(meta, ensure_ascii=False))
    os.makedirs(SAIDA, exist_ok=True)
    dest = os.path.join(SAIDA, "VIP_Lounges_Performance_%s.html"
                        % datetime.datetime.now().strftime("%y%m%d"))
    with open(dest, "w", encoding="utf-8") as f:
        f.write(html)
    print("\nOK -> %s  (%.0f KB)" % (dest, os.path.getsize(dest) / 1024))


TEMPLATE = r"""<!doctype html>
<html lang="pt-BR"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Salas VIP · Performance</title>
<style>
/* Inter + IBM Plex Mono embutidas — as mesmas do Painel Integrado.
   Vão no arquivo para o relatório abrir offline e sempre com a mesma letra. */
/*__FONTES__*/
</style>
<style>
/* tokens herdados do Painel Integrado (Daily Sales) */
:root{
  color-scheme:light;
  /* fundo e neutros vindos do modelo de PPT da GOL (#FCF6F0 / #37322C / #9A9187) */
  --bg:#FCF6F0; --surface:#ffffff; --surface2:#F7F1EA;
  --border:#E6DFD6; --border2:#DAD0C5;
  --primary:#FF7020; --primary-dark:#DB5014; --primary-light:#F36F21;
  --accent:#FFB414; --warn:#e85d5d; --ok:#2d9f5d;
  --blue:#007895; --purple:#8a63d2; --teal:#009A8B;
  --text:#37322C; --text-secondary:#65605B; --text-muted:#9A9187;
  --grid:#EDE5DB; --axis:#9A9187; --band:rgba(55,50,44,.028);
  --sh:0 1px 2px rgba(55,50,44,.04),0 4px 14px rgba(55,50,44,.06);
  --sh2:0 2px 8px rgba(55,50,44,.08),0 14px 34px rgba(55,50,44,.12);
  --up:#2d9f5d; --dn:#e85d5d;
  /* série: matizes oficiais da GOL reposicionados na banda válida e validados */
  --p1:#e25600; --p2:#0099c1; --p3:#30a15f; --p4:#b57d00;
  --p5:#00a290; --p6:#966fdf; --p7:#c96482;
  --q1:#ffe6d1; --q2:#ffc79c; --q3:#ffa563; --q4:#f5822e; --q5:#d15f00; --q6:#9c4300;
}
@media (prefers-color-scheme:dark){:root:where(:not([data-theme="light"])){
  color-scheme:dark;
  --bg:#14120F; --surface:#1F1C18; --surface2:#191612;
  --border:#332E28; --border2:#453F37;
  --primary:#FF7020; --primary-dark:#DB5014; --accent:#FFB414;
  --text:#F5EFE7; --text-secondary:#C9C1B6; --text-muted:#9A9187;
  --grid:#2C271F; --axis:#6E665B; --band:rgba(255,255,255,.03);
  --sh:0 1px 2px rgba(0,0,0,.45),0 4px 14px rgba(0,0,0,.3);
  --sh2:0 2px 8px rgba(0,0,0,.55),0 14px 34px rgba(0,0,0,.45);
  --up:#3ec13e; --dn:#ef6b6b;
  --p1:#d45000; --p2:#0090b8; --p3:#229856; --p4:#aa7500;
  --p5:#009887; --p6:#8d66d5; --p7:#be5b79;
  --q1:#4a2000; --q2:#7a3600; --q3:#a84c00; --q4:#d15f00; --q5:#f5822e; --q6:#ffa563;
}}
:root[data-theme="dark"]{
  color-scheme:dark;
  --bg:#14120F; --surface:#1F1C18; --surface2:#191612;
  --border:#332E28; --border2:#453F37;
  --text:#F5EFE7; --text-secondary:#C9C1B6; --text-muted:#9A9187;
  --grid:#2C271F; --axis:#6E665B; --band:rgba(255,255,255,.03);
  --sh:0 1px 2px rgba(0,0,0,.45),0 4px 14px rgba(0,0,0,.3);
  --sh2:0 2px 8px rgba(0,0,0,.55),0 14px 34px rgba(0,0,0,.45);
  --up:#3ec13e; --dn:#ef6b6b;
  --p1:#d45000; --p2:#0090b8; --p3:#229856; --p4:#aa7500;
  --p5:#009887; --p6:#8d66d5; --p7:#be5b79;
  --q1:#4a2000; --q2:#7a3600; --q3:#a84c00; --q4:#d15f00; --q5:#f5822e; --q6:#ffa563;
}
*{box-sizing:border-box;min-width:0}
/* Declaração idêntica à do Painel Integrado. Sem -webkit-font-smoothing:
   ele deixava o traço mais fino que o do painel, e o painel não usa. */
body{margin:0;background:var(--bg);color:var(--text);
  font-family:'Inter',sans-serif;line-height:1.5}
/* Mono só onde o painel usa: números densos de eixo, tabela e tooltip.
   O valor grande do KPI é Inter 800, como o .kpi-value de lá. */
.mono,.tick,.dlab,.tk2,td.num,th.num,#tip .r{
  font-family:'IBM Plex Mono',monospace;
  font-variant-numeric:tabular-nums lining-nums}
.wrapper{max-width:1240px;margin:0 auto;padding:20px 24px 64px}

@keyframes rise{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
@keyframes fade{from{opacity:0}to{opacity:1}}
@keyframes gY{from{transform:scaleY(0)}to{transform:scaleY(1)}}
@keyframes gX{from{transform:scaleX(0)}to{transform:scaleX(1)}}
@keyframes draw{to{stroke-dashoffset:0}}

.rise{animation:rise .5s cubic-bezier(.2,.9,.3,1) both;animation-delay:calc(var(--i,0)*50ms)}
rect.by{transform-box:fill-box;transform-origin:50% 100%;animation:gY .6s cubic-bezier(.2,.9,.3,1) both}
rect.bx{transform-box:fill-box;transform-origin:0% 50%;animation:gX .6s cubic-bezier(.2,.9,.3,1) both}
rect.bxr{transform-box:fill-box;transform-origin:100% 50%;animation:gX .6s cubic-bezier(.2,.9,.3,1) both}
path.dl{animation:draw .95s cubic-bezier(.4,0,.2,1) both}
circle.pt,text.dlab,g.endlab{animation:fade .4s ease .55s both}
@media (prefers-reduced-motion:reduce){*,*::before,*::after{animation:none!important;transition:none!important}}

/* ---------- HEADER: mesmo padrão do Painel Integrado ---------- */
.header{background:linear-gradient(135deg,var(--primary) 0%,#d45500 100%);
  border-radius:16px;padding:30px 32px;margin-bottom:20px;color:#fff;
  position:relative;overflow:hidden;box-shadow:0 4px 20px rgba(255,105,0,.15)}
.header-top{position:relative;z-index:3;display:flex;align-items:flex-start;
  gap:24px;flex-wrap:wrap}
.header .eyebrow{font-size:10px;font-weight:700;letter-spacing:.2em;
  text-transform:uppercase;color:rgba(255,255,255,.8);margin-bottom:8px}
.header h1{margin:0;font-size:38px;font-weight:800;letter-spacing:-.035em;line-height:1.02}
.header h1 small{display:block;font-size:14.5px;margin-top:6px;font-weight:400;letter-spacing:0;
  color:rgba(255,255,255,.9);margin-top:3px}
.header .period{margin-top:12px;font-size:12.5px;color:rgba(255,255,255,.9);
  display:inline-flex;gap:9px;align-items:center;background:rgba(255,255,255,.16);
  border-radius:99px;padding:5px 14px;backdrop-filter:blur(6px)}
.header .period b{font-weight:700;color:#fff}
.hspace{flex:1;min-width:20px}
.hstat{position:relative;z-index:3;display:flex;gap:30px;flex-wrap:wrap;margin-top:4px}
.hstat .k{font-size:9.5px;letter-spacing:.13em;text-transform:uppercase;
  color:rgba(255,255,255,.72);font-weight:700}
.hstat .v{font-size:23px;font-weight:700;letter-spacing:-.025em;margin-top:3px}
.hbtns{position:relative;z-index:3;display:flex;gap:8px;margin-top:2px}
.hbtns button{background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.32);
  color:#fff;font-size:12px;font-weight:600;border-radius:9px;padding:7px 13px;
  cursor:pointer;backdrop-filter:blur(6px);transition:all .17s}
.hbtns button:hover{background:rgba(255,255,255,.3)}
button,select,input{font:inherit;color:var(--text);background:var(--surface);
  border:1px solid var(--border);border-radius:9px;padding:7px 12px;cursor:pointer;
  transition:all .17s cubic-bezier(.2,.9,.3,1)}
button:hover{border-color:var(--border2);transform:translateY(-1px);box-shadow:var(--sh)}

/* ---------- FILTROS ---------- */
/* A regra da antiga barra horizontal foi removida: o `margin:0 -24px` dela
   sobrevivia à regra nova (que não redefine margin) e empurrava o painel
   24px para fora da tela. */
.fr{display:flex;gap:9px;align-items:center;flex-wrap:wrap}
.dd{position:relative}
.dd>button{display:flex;align-items:center;gap:8px;font-size:12.5px;white-space:nowrap}
.dd>button .lb{color:var(--text-muted);font-size:10px;font-weight:700;
  letter-spacing:.09em;text-transform:uppercase}
.dd>button .vl{font-weight:600;max-width:180px;overflow:hidden;
  text-overflow:ellipsis;white-space:nowrap}
.dd>button .cv{color:var(--text-muted);font-size:9px}
.dd.open>button{border-color:var(--primary);box-shadow:0 0 0 3px rgba(255,112,32,.13)}
.pan{display:none;position:absolute;bottom:0;left:calc(100% + 8px);z-index:90;
  background:var(--surface);border:1px solid var(--border);border-radius:11px;
  box-shadow:var(--sh2);padding:9px;min-width:224px;max-height:330px;overflow:auto}
.dd.open .pan{display:block;animation:rise .18s ease both}
.fbar .pan{max-height:70vh;overflow:auto}
@keyframes slideRight{from{opacity:0;transform:translateX(-8px)}to{opacity:1;transform:none}}
.dd.open .pan{animation-name:slideRight}
@media(max-width:1180px){.pan{bottom:auto;top:calc(100% + 6px);left:0}}
.opt{display:flex;align-items:center;gap:9px;padding:7px 9px;border-radius:7px;
  font-size:12.5px;cursor:pointer;user-select:none;transition:background .12s}
.opt:hover{background:var(--surface2)}
.opt .bx{width:15px;height:15px;border-radius:4px;border:1.5px solid var(--border2);
  flex:none;display:grid;place-items:center;font-size:10px;color:#fff;line-height:1}
.opt.on .bx{background:var(--primary);border-color:var(--primary)}
.opt .sw{width:9px;height:9px;border-radius:2.5px;flex:none}
.panfoot{display:flex;gap:7px;border-top:1px solid var(--grid);margin-top:7px;padding-top:8px}
.panfoot button{flex:1;font-size:11px;padding:5px 8px}
.seg{display:flex;background:var(--surface2);border:1px solid var(--border);
  border-radius:9px;padding:2px;gap:2px}
.seg button{border:none;background:none;font-size:12px;padding:5px 13px;
  border-radius:7px;font-weight:600;color:var(--text-muted)}
.seg button:hover{transform:none;box-shadow:none;color:var(--text)}
.seg button.on{background:var(--primary);color:#fff;box-shadow:0 2px 8px rgba(255,112,32,.3)}
.tgl{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--text-secondary);
  cursor:pointer;user-select:none;padding:7px 12px;border:1px solid var(--border);
  border-radius:9px;background:var(--surface)}
.tgl .sl{width:30px;height:17px;border-radius:99px;background:var(--border2);
  position:relative;transition:background .2s;flex:none}
.tgl .sl::after{content:"";position:absolute;top:2px;left:2px;width:13px;height:13px;
  border-radius:50%;background:#fff;transition:transform .2s}
.tgl.on .sl{background:var(--primary)} .tgl.on .sl::after{transform:translateX(13px)}

/* ---------- SEÇÕES ---------- */
section{margin:34px 0 0}
.sh{display:flex;align-items:center;gap:12px;margin-bottom:16px;flex-wrap:wrap;
  font-size:12px;letter-spacing:.12em;text-transform:uppercase;
  color:var(--text-secondary);font-weight:700}
.sh::after{content:'';flex:1;min-width:24px;height:1px;background:var(--border);order:9}
.sh .bar{width:3px;height:16px;border-radius:2px;background:var(--primary);flex:none}
.sh h2{margin:0;font-size:12px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--text-secondary)}
.sh p{margin:0;color:var(--text-muted);font-size:12px;flex:1;min-width:180px}
.grid{display:grid;gap:14px;align-items:stretch}
.card{display:flex;flex-direction:column}

.g2{grid-template-columns:repeat(2,minmax(0,1fr))}
.ins-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px}
.g4{grid-template-columns:repeat(4,minmax(0,1fr))}
.gA{grid-template-columns:minmax(0,1.9fr) minmax(0,1fr)}
.gB{grid-template-columns:minmax(0,1.05fr) minmax(0,1.15fr) minmax(0,1fr)}
@media(max-width:1100px){.gB{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:900px){.g4{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:980px){.g2,.g4,.gA,.gB{grid-template-columns:minmax(0,1fr)}}

.card{background:var(--surface);border:2px solid var(--border);border-radius:14px;
  padding:22px;box-shadow:none;overflow:hidden;
  transition:box-shadow .22s,transform .22s}
.card:hover{box-shadow:var(--sh2);transform:translateY(-2px)}

.cardhead{display:flex;align-items:center;justify-content:space-between;gap:10px;
  flex-wrap:wrap;margin-bottom:2px}
.cardhead h3{margin:0;font-size:13.5px;font-weight:700;letter-spacing:-.01em}
.card>h3{margin:0 0 3px;font-size:13.5px;font-weight:700;letter-spacing:-.01em}
.cardhead{display:flex;align-items:center;justify-content:space-between;gap:10px;
  flex-wrap:wrap;margin-bottom:8px}
.cardhead h3{margin:0;font-size:13.5px;font-weight:700;letter-spacing:-.01em}

/* estado vazio: some com tudo e explica, em vez de desenhar gráficos zerados */
#vazio{display:none;padding:64px 24px;text-align:center}
/* Sem dados: o vazio estica e o rodape desce para o pe da tela. */
body.sem-dados .wrapper{display:flex;flex-direction:column;
  min-height:calc(100vh - var(--filtro-gap)*2)}
body.sem-dados #vazio{flex:1 1 auto;display:flex;flex-direction:column;
  align-items:center;justify-content:center}
body.sem-dados footer{margin-top:auto}
body.sem-dados #vazio{display:block}
body.sem-dados section{display:none}
#vazio .ic{font-size:34px;line-height:1;opacity:.35;margin-bottom:14px}
#vazio .msg{font-size:15px;font-weight:600;color:var(--text);max-width:460px;
  margin:0 auto 6px}
#vazio .sub{font-size:12px;color:var(--text-muted);max-width:460px;margin:0 auto}

/* Painel de filtros ancorado à esquerda, na parte de baixo da tela.
   A largura da coluna vive numa variável para o conteúdo reservar exatamente
   o mesmo espaço — sem isso os dois saem de sincronia. */
:root{--filtro-w:216px; --filtro-gap:22px}
.fbar{position:fixed;left:var(--filtro-gap);bottom:var(--filtro-gap);top:auto;right:auto;
  z-index:80;width:var(--filtro-w);
  background:var(--surface);border:2px solid var(--border);border-radius:14px;
  padding:12px;box-shadow:0 8px 28px rgba(26,26,26,.14);
  /* nada de overflow aqui: o painel do dropdown abre para FORA da caixa e
     era recortado, o que deixava o filtro inutilizável */
  overflow:visible;
  max-height:calc(100vh - var(--filtro-gap)*2)}
.fbar .fr{flex-direction:column;align-items:stretch;gap:7px;padding:0;max-width:none}
.fbar .fr>*{width:100%}
.fbar .dd>button{width:100%;justify-content:space-between}
.fbar .seg{width:100%} .fbar .seg button{flex:1}
.fbar .tgl{justify-content:space-between}
.ftit{font-size:9px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;
  color:var(--text-muted);margin-bottom:2px}
/* O conteúdo fica SEMPRE no centro exato da tela. Quando a largura não permite
   1240px sem invadir o painel, é a coluna que estreita — não o centro que sai
   do lugar. O painel ocupa até 238px (gap + largura); com 24px de respiro de
   cada lado, a coluna pode ter no máximo 100vw − 524px. */
.wrapper{max-width:min(1240px, calc(100vw - 524px));padding-bottom:40px}
@media(max-width:1180px){.wrapper{max-width:1240px}}
/* sem largura para a coluna: o painel volta a ser uma faixa acima do conteúdo */
@media(max-width:1180px){
  body{padding-left:0}
  .fbar{position:static;width:auto;max-height:none;margin:0 0 12px;
    box-shadow:none;left:auto;bottom:auto}
  .fbar .fr{flex-direction:row;flex-wrap:wrap;align-items:center}
  .fbar .fr>*{width:auto}
  .ftit{display:none}
}
.card>p.cs{margin:0 0 14px;font-size:12px;color:var(--text-muted);line-height:1.5}

/* ---------- KPI ---------- */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px}
@media(max-width:760px){.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
/* espelha .kpi / .kpi-label / .kpi-value / .kpi-sub do Painel Integrado */
.kpi{position:relative;padding:18px;border:2px solid var(--border);border-radius:12px}
.kpi:hover{border-color:var(--primary-light);transform:translateY(-2px);
  box-shadow:0 4px 12px rgba(255,105,0,.08)}
.kpi .lbl{font-size:10.5px;color:var(--text-muted);text-transform:uppercase;
  letter-spacing:.08em;margin-bottom:6px;font-weight:700;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.kpi .val{font-size:26px;font-weight:800;line-height:1.15;
  overflow-wrap:anywhere;word-break:break-word}
.kpi .val .cur{font-size:15px;color:var(--text-muted);font-weight:700;margin-right:2px}
.kpi .dlt{font-size:12px;font-weight:700;min-height:18px;margin-top:4px}
.kpi .dlt em{font-style:normal;font-weight:400;color:var(--text-secondary)}
.kpi .note{font-size:11.5px;color:var(--text-secondary);margin-top:3px;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.up{color:var(--up)} .dn{color:var(--dn)} .flat{color:var(--text-muted)}

/* Nada vazio ocupa espaço. Cobre o que saiu do HTML e o que é esvaziado em
   runtime: nota de KPI sem comparativo, escopo sem filtro, legenda de série
   única, rodapé de gráfico sem texto. */
.sh p:empty,.card>p.cs:empty,.kpi .note:empty,.kpi .dlt:empty,
.rdscope:empty,.note-s:empty,.legend:empty,.tv:empty,.pill-w:empty{display:none}
.sh p{margin:0}
/* cartão sem descrição: o gráfico encosta no título com o respiro certo */
.card>h3+div,.card>h3+.tw{margin-top:10px}

.tw{overflow:auto}
table{border-collapse:collapse;width:100%;font-size:12.5px}
th,td{text-align:left;padding:9px 10px;border-bottom:1px solid var(--grid);white-space:nowrap}
th{font-size:10.5px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
  color:var(--text-muted);position:sticky;top:0;background:var(--surface);z-index:2}
th.num,td.num{text-align:right;font-variant-numeric:tabular-nums}
tbody tr{transition:background .13s} tbody tr:hover{background:var(--surface2)}
.nm{max-width:160px;overflow:hidden;text-overflow:ellipsis}
.sw{width:9px;height:9px;border-radius:2.5px;display:inline-block;margin-right:7px;flex:none}
.legend{display:flex;gap:14px;flex-wrap:wrap;margin-top:14px;font-size:12px;
  color:var(--text-secondary)}
.legend .it{display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;
  transition:opacity .16s}
.legend .it.off{opacity:.3}
#tip{position:fixed;z-index:999;pointer-events:none;background:var(--surface);
  color:var(--text);border:1px solid var(--border);border-radius:10px;padding:10px 13px;
  font-size:11.5px;box-shadow:var(--sh2);opacity:0;transition:opacity .12s;max-width:280px}
#tip .t{font-weight:700;margin-bottom:6px;font-size:12px}
#tip .r{display:flex;justify-content:space-between;gap:18px;
  font-variant-numeric:tabular-nums;line-height:1.6}
#tip .r span:first-child{color:var(--text-muted)}
#tip .r span:last-child{font-weight:600}
svg{display:block;overflow:hidden;max-width:100%}
.gridline{stroke:var(--grid);stroke-width:1;shape-rendering:crispEdges}
.axisline{stroke:var(--axis);stroke-width:1;opacity:.5;shape-rendering:crispEdges}
.tick{fill:var(--text-muted);font-size:11px;font-variant-numeric:tabular-nums}
.tk2{fill:var(--text-secondary);font-size:12px;font-weight:600}
.dlab{font-size:11.5px;font-weight:700}
.hit{fill:transparent;cursor:pointer}
/* mesmo desenho do .ins do Painel Integrado: título + corpo, borda à esquerda */
.ins{background:var(--surface);border:2px solid var(--border);border-left-width:5px;
  border-radius:12px;padding:18px}
.ins.alta{border-left-color:var(--ok)}
.ins.queda{border-left-color:var(--warn)}
.ins.atencao{border-left-color:var(--accent)}
.ins.info{border-left-color:var(--blue)}
.ins-title{font-size:13.5px;font-weight:700;margin-bottom:6px}
.ins-body{font-size:12.5px;color:var(--text-secondary);line-height:1.55}
.ins-body strong{color:var(--text);font-weight:700}
.tv{display:none} .tv.show{display:block;margin-top:14px;max-height:290px;overflow:auto}
.note-s{font-size:11.5px;color:var(--text-muted);margin:12px 0 0;line-height:1.55}
.pill-w{display:inline-flex;align-items:center;gap:6px;font-size:10.5px;font-weight:600;
  color:#a05a00;background:rgba(255,180,20,.16);border-radius:99px;padding:3px 10px;margin-top:10px}
.rdhead{display:flex;align-items:center;gap:10px;margin-bottom:8px;flex-wrap:wrap}
.rdchip{font-size:10.5px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;
  color:var(--primary);background:rgba(255,112,32,.10);border-radius:99px;padding:4px 12px}
.rdscope{font-size:11.5px;color:var(--text-muted)}
/* o botão fica na mesma linha do título, depois da régua (order:9) */
.rdbtn{order:10;flex:none;display:inline-flex;align-items:center;justify-content:center;
  width:26px;height:26px;padding:0;border-radius:8px;font-size:14px;line-height:1;
  color:var(--text-secondary);background:var(--surface);border:2px solid var(--border)}
.rdbtn:hover{border-color:var(--primary);color:var(--primary)}
.rdbtn.spin{transform:rotate(360deg);transition:transform .5s cubic-bezier(.2,.9,.3,1)}
.rdbtn .ico{display:inline-block;transition:transform .5s cubic-bezier(.2,.9,.3,1)}
.rdbtn.spin .ico{transform:rotate(360deg)}
footer{margin-top:28px;padding-top:12px;border-top:1px solid var(--border);
  color:var(--text-muted);font-size:11px;display:flex;gap:20px;flex-wrap:wrap;align-items:center}
.fspace{flex:1}
.copy{margin-top:14px;padding-top:14px;border-top:1px solid var(--border);
  text-align:center;color:var(--text-muted);font-size:11px}

/* ---------- celular (por ultimo, para vencer as regras base) ---------- */
@media(max-width:560px){
  body{padding-left:0}
  .wrapper{padding:12px 12px 20px}
  .header{padding:12px 14px;border-radius:12px}
  .header h1{font-size:17px} .header h1 small{font-size:10.5px}
  .kpis,#c-idx{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:8px}
  .kpi{padding:10px 11px} .kpi .val{font-size:18px} .kpi .lbl{font-size:8.5px}
  .kpi .note,.kpi .dlt{font-size:9.5px}
  .card{padding:12px 13px;border-radius:11px}
  .grid{gap:9px}
  .sh h2{font-size:13px}
  table{font-size:10px} th,td{padding:3px 5px}
  .callout{padding:8px 10px;font-size:11px}
  .fbar{padding:8px 10px;border-radius:12px}
  .fbar .fr{padding:0 10px;gap:6px;flex-wrap:nowrap;overflow-x:auto;
    -webkit-overflow-scrolling:touch;scrollbar-width:none}
  .fbar .fr::-webkit-scrollbar{display:none}
  .dd>button,.tgl,.fr>button{padding:6px 9px;font-size:11px;white-space:nowrap}
  .dd>button .vl{max-width:96px}
  .seg button{padding:4px 9px;font-size:11px}
  .pan{max-height:56vh}
  #vazio{padding:40px 14px} #vazio .msg{font-size:13.5px}
}

/* ---------- impressão / PDF ---------- */
@page{size:A4 landscape;margin:11mm}
@media print{
  :root{--bg:#fff;--surface:#fff;--surface2:#fafaf8;--text:#111;--border:#dcdad4}
  html,body{background:#fff!important}
  *,*::before,*::after{animation:none!important;transition:none!important;
    box-shadow:none!important}
  .fbar,.hbtns,.rdbtn,.legend .it{display:none!important}
  .wrapper{max-width:none;padding:0}
  .header{border-radius:10px;padding:16px 20px;margin-bottom:12px;
    -webkit-print-color-adjust:exact;print-color-adjust:exact}
  .header h1{font-size:21px} .header h1 small{font-size:12px}
  .card{break-inside:avoid;page-break-inside:avoid;border-radius:10px;
    padding:12px 14px;box-shadow:none;border:1px solid #e2e0da}
  .card:hover{transform:none}
  /* a seção pode quebrar entre páginas; o cartão nunca parte no meio */
  section{margin:16px 0 0;break-inside:auto;page-break-inside:auto}
  .sh{break-after:avoid;page-break-after:avoid}
  .sh h2{font-size:14px}
  .callout,.kpi,tr{break-inside:avoid;page-break-inside:avoid}
  thead{display:table-header-group}
  .kpis{grid-template-columns:repeat(5,minmax(0,1fr))!important;gap:8px}
  #c-idx{grid-template-columns:repeat(4,minmax(0,1fr))!important;gap:8px}
  .kpi .val{font-size:19px}
  .g4{grid-template-columns:repeat(4,minmax(0,1fr))!important}
  .gA,.gB,.g2{grid-template-columns:repeat(2,minmax(0,1fr))!important}
  .grid{gap:10px}
  .tv{display:block!important;max-height:none!important;overflow:visible!important}
  .tw{overflow:visible!important;max-height:none!important}
  table{font-size:9.5px} th,td{padding:3px 5px}
  .callout{break-inside:avoid;padding:9px 12px;font-size:11px}
  footer,.copy{break-inside:avoid}
  .copy{font-weight:600}
  svg{max-width:100%!important;height:auto!important}
}
</style></head><body>

<div id="tip"></div>
<div class="fbar" id="fbar"><div class="fr">
  <div class="ftit">Filtros</div>
  <div class="dd" id="dd-per"></div>
  <div class="dd" id="dd-ano"></div>
  <div class="dd" id="dd-parc"></div>
  <div class="dd" id="dd-sala"></div>
  <div class="seg" id="seg-met"></div>
  <div class="tgl" id="tgl-tab"><span class="sl"></span><span>Tabelas</span></div>
  <button id="clearbtn">Limpar</button>
  <button id="themebtn">Dark</button>
</div></div>
<div class="wrapper">

<div class="header">  <div class="header-top">
    <div style="min-width:0">
      <div class="eyebrow">Gerência de Estratégia Comercial · T2M</div>
      <h1>Salas VIP<small>Performance de Acessos às Salas VIP</small></h1>
    </div>
    <div class="hspace"></div>
  </div>
</div>

<div id="vazio">
  <div class="ic">⌀</div>
  <div class="msg">Não há dados disponíveis</div>
  <div class="sub" id="vazio-sub"></div>
</div>

<section>
  <div class="sh"><span class="bar"></span><h2>GERAL</h2></div>
  <div class="kpis" id="kpis"></div>
</section>

<section>
  <div class="sh"><span class="bar"></span><h2>Índices</h2></div>
  <div class="kpis" id="c-idx" ></div>
</section>

<section>
  <div class="sh"><span class="bar"></span><h2>Leitura Executiva</h2><button class="rdbtn" id="rd-next" title="Ver outra leitura">⟳</button></div>
  <div class="ins-grid" id="c-ins"></div>
</section>

<section>
  <div class="sh"><span class="bar"></span><h2>Evolução Mensal</h2></div>
  <div class="grid">
    <div class="card rise" style="--i:0">
      <h3>Evolução Mensal <span id="t-met" style="color:var(--text-muted);font-weight:400"></span></h3>
      <div id="c-line"></div><div class="legend" id="l-line"></div>
      <div class="tv" id="tv-line"></div>
    </div>
  </div>
  <div class="grid gA" style="margin-top:16px">
    <div class="card rise" style="--i:1">
      <h3>Composição mensal por parceira <span id="t-stk" style="color:var(--text-muted);font-weight:400"></span></h3>
      <div id="c-stack"></div><div class="legend" id="l-stack"></div>
    </div>
    <div class="card rise" style="--i:2">
      <h3>Acumulado no ano · YTD</h3>
      
      <div id="c-ytd"></div>
    </div>
  </div>
</section>

<section>
  <div class="sh"><span class="bar"></span><h2>Parceiras</h2></div>
  <div class="grid g2">
    <div class="card rise" style="--i:0">
      <h3>Ranking</h3>
      <div id="c-rank"></div><div class="tv" id="tv-rank"></div>
    </div>
    <div class="card rise" style="--i:1">
      <h3>Evolução ano a ano</h3>
      <div id="c-contrib"></div><div class="tv" id="tv-contrib"></div>
    </div>
  </div>
</section>

<section>
  <div class="sh"><span class="bar"></span><h2>Salas VIP</h2>
    <p>Custo por Sala e Passageiros</p></div>
  <div class="grid">
    <div class="card rise" style="--i:0">
      <h3>Participação por sala</h3>
      <div id="c-split"></div>
    </div>
  </div>
</section>


<section>
  <div class="sh"><span class="bar"></span><h2>L3M</h2></div>
  <div class="grid gA">
    <div class="card rise" style="--i:0">
      <h3 id="t-l3m">L3M por parceira</h3>
      <p class="cs" id="cs-l3m"></p>
      <div id="c-l3m"></div><div class="legend" id="l-l3m"></div><div class="tv" id="tv-l3m"></div>
    </div>
    <div class="card rise" style="--i:1">
      <div class="cardhead"><h3>Movimentos L3M</h3><span id="mov-sala"></span></div>
      <div id="c-mov"></div><div class="tv" id="tv-mov"></div>
    </div>
  </div>
  <div class="grid g4" style="margin-top:16px" id="c-l3msalas"></div>
</section>

<section id="sec-detparc" style="display:none">
  <div class="sh"><span class="bar"></span><h2>Detalhamento · <span id="t-detparc"></span></h2></div>
  <div class="card"><div id="c-detparc"></div></div>
</section>

<footer>
  <span>Fonte: <strong style="color:var(--text-secondary)">Knime Smiles</strong></span>
  <span>Valores em USD</span>
  <span id="ffx"></span>
  <span class="fspace"></span>
  <span id="fgen"></span>
</footer>
<div class="copy">© 2026 GOL Linhas Aéreas — Todos os direitos reservados.<br><span style="opacity:.75">Desenvolvido por Carlos Sena</span></div>
</div>

<script>
const ROWS = /*__DADOS__*/;
const META = /*__META__*/;
const MK=META.meses, MEN=META.mes_en, MLG=META.mes_longo;
const SER=["--p1","--p2","--p3","--p4","--p5","--p6","--p7"];
const SEQ=["--q1","--q2","--q3","--q4","--q5","--q6"];
const cv = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();

const MPT=["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"];
const MPTL=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto",
            "Setembro","Outubro","Novembro","Dezembro"];

/* cor por ENTIDADE, fixada uma vez; sem reciclagem de matiz */
const PC={}; let _s=0;
META.parcs.forEach(p=>{
  const ativa=(META.contratos[p]||{}).pax_hist>0;
  PC[p]=(ativa&&_s<SER.length)?SER[_s++]:'--text-muted';
});
/* Cores de marca das companhias, usadas só na composição mensal.
   Quatro das sete são azuis na vida real (Air France, United, British, KLM) e,
   postas lado a lado, ficam indistinguíveis. Mantive o matiz de cada marca e
   separei por tom dentro de cada família — mais o valor escrito na barra. */
const CIA={
  'AVIANCA':'#E01933',            /* vermelho Avianca      */
  'ANGOLA AIRLINES':'#8F1A2A',    /* vermelho TAAG, tom escuro para separar */
  'AIR FRANCE':'#0B2E6F',         /* azul-marinho Air France */
  'UNITED AIRLINES':'#2F7CCC',    /* azul United, tom médio  */
  'BRITISH AIRWAYS':'#7FB2E5',    /* azul British, tom claro */
  'KLM':'#00A1DE',                /* ciano KLM              */
  'ETHIOPIAN AIRLINES':'#00A651'  /* verde Ethiopian        */
};
const ciaCor=p=>CIA[p]||cv(PC[p]);
const pn=p=>META.parc_en[p]||p, ln=s=>META.sala_en[s]||s;
const mi=m=>MK.indexOf(m);
const mn=m=>MPT[mi(m)]||m, mlg=m=>MPTL[mi(m)]||m;

const n0=new Intl.NumberFormat('pt-BR',{maximumFractionDigits:0});
const n2=new Intl.NumberFormat('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
const fmt=(v,m)=>m==='pax'?n0.format(Math.round(v)):n2.format(v);
const money=v=>'US$ '+n0.format(Math.round(v));
/* Uma casa decimal abaixo de 100 mil: arredondar 55.665 para "56 mil" fazia o
   rótulo do gráfico brigar com o número exato mostrado no KPI. */
const kort=v=>{const a=Math.abs(v);
  return a>=1e6?(v/1e6).toFixed(1).replace('.',',')+' mi'
       :a>=1e5?n0.format(Math.round(v/1000))+' mil'
       :a>=1000?((v%1000===0)?(v/1000)+' mil':(v/1000).toFixed(1).replace('.',',')+' mil')
       :n0.format(Math.round(v));};
const pct=(a,b)=>(b&&isFinite(a/b))?(a-b)/Math.abs(b)*100:null;
const pctS=v=>v===null?'—':(v>0?'+':v<0?'−':'')+Math.abs(v).toFixed(1).replace('.',',')+'%';
const cls=v=>v===null?'flat':v>0?'up':v<0?'dn':'flat';
const cut=(s,n)=>s.length>n?s.slice(0,n-1)+'…':s;

const S={ano:new Set(META.anos), sala:new Set(META.salas), parc:new Set(META.parcs),
         m1:0,m2:11, modo:'unico', met:'usd', fx:'pub',
         hide:new Set(), tables:false, rd:0, rot:true, movSala:'todas'};

/* Último mês com dado — o relatório abre nele, já filtrado.
   A partir daí a pessoa escolhe o que quiser nos filtros. */
const ULT_ANO=Math.max(...ROWS.map(r=>r.ano));
const ULT_MES=Math.max(...ROWS.filter(r=>r.ano===ULT_ANO).map(r=>mi(r.mes)));
S.ano=new Set([ULT_ANO]); S.m1=ULT_MES; S.m2=ULT_MES;
/* meses que existem em cada ano, para o seletor não oferecer mês vazio */
const MESES_DO_ANO={};
ROWS.forEach(r=>{(MESES_DO_ANO[r.ano]=MESES_DO_ANO[r.ano]||new Set()).add(mi(r.mes));});
const mesesValidos=()=>{
  const s=new Set();
  [...S.ano].forEach(a=>(MESES_DO_ANO[a]||new Set()).forEach(i=>s.add(i)));
  return [...s].sort((x,y)=>x-y);
};
const val=(r,m)=>(m||S.met)==='pax'?r.pax:r['usd_'+S.fx];
const metName=()=>S.met==='pax'?'PAX':'Valor (US$)';
const unit=()=>S.met==='pax'?'':'US$ ';

function filtra(o={}){
  const A=o.ano||S.ano,L=o.sala||S.sala,P=o.parc||S.parc;
  return ROWS.filter(r=>A.has(r.ano)&&L.has(r.sala)&&P.has(r.parc)
    &&mi(r.mes)>=S.m1&&mi(r.mes)<=S.m2);
}
const sum=(rs,m)=>rs.reduce((a,r)=>a+val(r,m),0);
const sumP=rs=>rs.reduce((a,r)=>a+r.pax,0);
const sumU=rs=>rs.reduce((a,r)=>a+r['usd_'+S.fx],0);

/* ================= svg ================= */
const NS='http://www.w3.org/2000/svg';
let _cid=0;
function mk(t,a={}){const e=document.createElementNS(NS,t);for(const k in a)e.setAttribute(k,a[k]);return e;}
function canvas(box,h){
  const w=Math.max(240,Math.floor(box.clientWidth));
  const s=mk('svg',{viewBox:`0 0 ${w} ${h}`,width:w,height:h});
  s.style.cssText='width:100%;height:auto';
  return [s,w];
}
function clip(s,x,y,w,h){
  const id='clp'+(++_cid);
  const d=mk('defs'); const cp=mk('clipPath',{id});
  cp.appendChild(mk('rect',{x,y,width:Math.max(0,w),height:Math.max(0,h)}));
  d.appendChild(cp); s.appendChild(d);
  const g=mk('g',{'clip-path':`url(#${id})`}); s.appendChild(g); return g;
}
function T(s,x,y,t,a={}){const e=mk('text',Object.assign({x,y},a));e.textContent=t;s.appendChild(e);return e;}
/* O laço tem de gerar a marca que COBRE o máximo, e não parar na anterior.
   Com max=65.435 e passo 25.000 a versão antiga devolvia [0, 25k, 50k]: o topo
   do eixo ficava ABAIXO da maior barra e o gráfico aparecia cortado. */
function ticks(max,n=4){
  if(!(max>0))return [0];
  const raw=max/n,mag=Math.pow(10,Math.floor(Math.log10(raw)));
  const st=[1,2,2.5,5,10].map(m=>m*mag).find(s=>s>=raw)||10*mag;
  const o=[];
  for(let v=0;;v+=st){
    o.push(v);
    if(v>=max-st*1e-9||o.length>40) break;
  }
  return o;
}
function suave(pts){
  if(pts.length<2) return pts.length?`M${pts[0][0]} ${pts[0][1]}`:'';
  let d=`M${pts[0][0]} ${pts[0][1]}`;
  for(let i=0;i<pts.length-1;i++){
    const p0=pts[i-1]||pts[i],p1=pts[i],p2=pts[i+1],p3=pts[i+2]||p2;
    d+=`C${p1[0]+(p2[0]-p0[0])/6} ${p1[1]+(p2[1]-p0[1])/6},`
     + `${p2[0]-(p3[0]-p1[0])/6} ${p2[1]-(p3[1]-p1[1])/6},${p2[0]} ${p2[1]}`;
  }
  return d;
}
/* anima o traço; se getTotalLength falhar, mostra a linha em vez de escondê-la */
function anima(p,delay){
  try{
    const L=p.getTotalLength();
    if(!L||!isFinite(L)) return;
    p.style.strokeDasharray=L; p.style.strokeDashoffset=L;
    p.classList.add('dl'); p.style.animationDelay=(delay||0)+'ms';
  }catch(e){ p.style.strokeDasharray=''; p.style.strokeDashoffset=''; }
}
const tip=document.getElementById('tip');
function show(ev,t,ls){
  tip.innerHTML='<div class="t">'+t+'</div>'+ls.map(l=>
    '<div class="r"><span>'+l[0]+'</span><span>'+l[1]+'</span></div>').join('');
  tip.style.opacity=1;
  const r=tip.getBoundingClientRect();
  let x=ev.clientX+15,y=ev.clientY+15;
  if(x+r.width>innerWidth-8)x=ev.clientX-r.width-15;
  if(y+r.height>innerHeight-8)y=ev.clientY-r.height-15;
  tip.style.left=Math.max(8,x)+'px';tip.style.top=Math.max(8,y)+'px';
}
const hide=()=>tip.style.opacity=0;
function hov(el,t,ls){
  el.addEventListener('mousemove',e=>show(e,t,ls));
  el.addEventListener('mouseleave',hide);
  el.setAttribute('tabindex','0');
  el.setAttribute('aria-label',t+': '+ls.map(l=>l[0]+' '+l[1]).join(', '));
  el.addEventListener('focus',()=>{const b=el.getBoundingClientRect();
    show({clientX:b.left+b.width/2,clientY:b.top},t,ls);});
  el.addEventListener('blur',hide);
  /* clicar em qualquer dado fixa os valores à vista em todos os gráficos */
  el.addEventListener('click',ev=>{ ev.stopPropagation();
    S.rot=!S.rot;
    render(); });
  el.style.cursor='pointer';
}
/* rótulo de valor sobre a marca, exibido quando os valores estão à vista */
function rotulo(s,x,y,txt,cor,tam){
  if(!S.rot) return;
  const g=mk('g',{class:'vlab'});
  const w=String(txt).length*(tam||9)*0.62+8;
  g.appendChild(mk('rect',{x:x-w/2,y:y-(tam||9)-4,width:w,height:(tam||9)+6,rx:3,
    fill:cv('--surface'),opacity:.86}));
  const t=mk('text',{x,y:y-2,class:'dlab','text-anchor':'middle',
    fill:cor||cv('--text'),style:'font-size:'+(tam||9)+'px'});
  t.textContent=txt; g.appendChild(t); s.appendChild(g);
}

/* Comparativo contra o mesmo período do ano anterior.
   O ano de comparação NÃO depende do filtro de ano: se o usuário deixa só
   2026 marcado (ou só julho), continua havendo contra o que comparar. Antes
   isso devolvia null e a leitura executiva ficava vazia. */
function ytd(){
  const anos=[...S.ano].sort();
  if(!anos.length) return null;
  const cur=anos[anos.length-1];
  const disp=[...new Set(ROWS.map(r=>r.ano))].sort();
  const pre=disp.filter(a=>a<cur).pop();
  if(pre===undefined) return null;
  const mc=new Set(ROWS.filter(r=>r.ano===cur).map(r=>r.mes));
  const ok=r=>mc.has(r.mes)&&mi(r.mes)>=S.m1&&mi(r.mes)<=S.m2;
  const semAno=o=>ROWS.filter(r=>o.has(r.ano)&&S.sala.has(r.sala)&&S.parc.has(r.parc)
    &&mi(r.mes)>=S.m1&&mi(r.mes)<=S.m2);
  return {cur,pre,
    A:semAno(new Set([cur])).filter(ok),
    B:semAno(new Set([pre])).filter(ok),
    meses:MK.filter(m=>mc.has(m)&&mi(m)>=S.m1&&mi(m)<=S.m2)};
}

/* Mês único: linha ligando o mesmo mês ao longo dos anos (jul/24 → jul/25 → …).
   É a leitura de tendência do mês fechado. */
function cMesUnico(box){
  const mes=MK[S.m1];
  const anos=[...new Set(ROWS.filter(r=>r.mes===mes).map(r=>r.ano))].sort();
  const dados=anos.map(a=>[a, sum(ROWS.filter(r=>r.ano===a&&r.mes===mes
    &&S.sala.has(r.sala)&&S.parc.has(r.parc)))]).filter(e=>e[1]>0);
  if(!dados.length){box.innerHTML='<p class="note-s">Sem dados neste mês.</p>';return;}
  const H=290, base=H-48, top=42;
  const [s,W]=canvas(box,H);
  const max=Math.max(...dados.map(e=>e[1]));
  const tk=ticks(max,4), topo=tk[tk.length-1]||1;
  const ML=Math.max(46,Math.max(...tk.map(x=>(unit()+kort(x)).length))*6.4+14);
  tk.forEach(x=>{const y=base-(base-top)*(x/topo);
    s.appendChild(mk('line',{x1:ML,x2:W,y1:y,y2:y,class:'gridline'}));
    T(s,ML-8,y+3.5,unit()+kort(x),{class:'tick','text-anchor':'end'});});
  s.appendChild(mk('line',{x1:ML,x2:W,y1:base,y2:base,class:'axisline'}));
  const gp=(W-ML)/dados.length;
  const X=i=>ML+i*gp+gp/2, Y=v=>base-(v/topo)*(base-top);
  const pts=dados.map((e,i)=>[X(i),Y(e[1])]);
  const G=clip(s,ML-4,top-24,W-ML+8,base-top+30);
  const p=mk('path',{d:suave(pts),fill:'none',stroke:cv('--primary'),'stroke-width':2.8,
    'stroke-linejoin':'round','stroke-linecap':'round'});
  G.appendChild(p); anima(p,0);
  dados.forEach((e,i)=>{
    const [ano,v]=e, cx=X(i), cy=Y(v);
    G.appendChild(mk('circle',{cx,cy,r:5,fill:cv('--primary'),
      stroke:cv('--surface'),'stroke-width':2.5,class:'pt'}));
    T(s,cx,cy-13,unit()+kort(v),{class:'dlab','text-anchor':'middle',fill:cv('--text')});
    T(s,cx,base+17,String(ano),{class:'tick','text-anchor':'middle'});
    const ant=dados[i-1];
    if(ant){const d=pct(v,ant[1]);
      T(s,cx,base+32,pctS(d),{class:'dlab','text-anchor':'middle',
        fill:d>0?cv('--up'):d<0?cv('--dn'):cv('--text-muted')});}
    const hit=mk('circle',{cx,cy,r:16,class:'hit'}); s.appendChild(hit);
    const sel=ROWS.filter(r=>r.ano===ano&&r.mes===mes&&S.sala.has(r.sala)&&S.parc.has(r.parc));
    hov(hit,MPTL[S.m1]+' de '+ano,[['Valor',money(sumU(sel))],['PAX',n0.format(sumP(sel))],
      ['Custo por acesso',sumP(sel)?'US$ '+n2.format(sumU(sel)/sumP(sel)):'—']]);});
  box.appendChild(s);
  let h='<div class="tw"><table><thead><tr><th>Ano</th><th class="num">'+
    metName()+'</th><th class="num">Variação</th></tr></thead><tbody>';
  dados.forEach((e,i)=>{const d=i?pct(e[1],dados[i-1][1]):null;
    h+='<tr><td>'+e[0]+'</td><td class="num">'+fmt(e[1],S.met)+
      '</td><td class="num '+cls(d)+'">'+pctS(d)+'</td></tr>';});
  document.getElementById('tv-line').innerHTML=h+'</tbody></table></div>';
}

/* Intervalo de meses: barras agrupadas, um grupo por mês, uma barra por ano. */
function cLine(){
  const box=document.getElementById('c-line'); box.innerHTML='';
  if(S.modo==='unico'){
    document.getElementById('t-met').textContent='· '+MPTL[S.m1]+' ao longo dos anos';
    document.getElementById('l-line').innerHTML='';
    cMesUnico(box); return;}
  const anos=[...S.ano].sort();
  /* Uma barra por ano é a leitura padrão. Mas quando o filtro isola algumas
     parceiras dentro de um único ano, o que interessa é comparar essas
     parceiras entre si — então cada uma vira sua própria barra. */
  const sub=META.parcs.filter(p=>S.parc.has(p));
  const porParc = anos.length===1 && sub.length>1 && sub.length<META.parcs.length;
  const ser = porParc ? sub : anos;
  const cor = porParc ? (k=>ciaCor(ser[k])) : (k=>cv(SER[k%SER.length]));
  const nome = porParc ? (k=>pn(ser[k])) : (k=>String(ser[k]));
  document.getElementById('t-met').textContent =
    porParc ? '· por parceira · '+anos[0] : '(US$)';
  const meses=[]; for(let i=S.m1;i<=S.m2;i++) meses.push(i);
  /* linhas de uma série num mês, respeitando todos os outros filtros */
  const dados=(k,mIdx)=> porParc
    ? filtra().filter(r=>r.parc===ser[k]&&r.mes===MK[mIdx])
    : filtra({ano:new Set([ser[k]])}).filter(r=>r.mes===MK[mIdx]);
  const d={}; let max=0;
  ser.forEach((sk,k)=>{
    d[sk]=meses.map(i=>{const rs=dados(k,i); return rs.length?sum(rs):null;});
    if(!S.hide.has(sk)) d[sk].forEach(v=>{if(v!==null&&v>max)max=v;});});
  if(!(max>0)){box.innerHTML='<p class="note-s">Sem dados no recorte.</p>';
    document.getElementById('l-line').innerHTML='';
    document.getElementById('tv-line').innerHTML=''; return;}
  const H=300, base=H-46, top=30;
  const [s,W]=canvas(box,H);
  const tk=ticks(max,4), topo=tk[tk.length-1]||1;
  const ML=Math.max(50,Math.max(...tk.map(x=>(unit()+kort(x)).length))*6.4+14);
  tk.forEach(x=>{const y=base-(base-top)*(x/topo);
    s.appendChild(mk('line',{x1:ML,x2:W,y1:y,y2:y,class:'gridline'}));
    T(s,ML-8,y+3.5,unit()+kort(x),{class:'tick','text-anchor':'end'});});
  s.appendChild(mk('line',{x1:ML,x2:W,y1:base,y2:base,class:'axisline'}));
  const vis=ser.map((sk,k)=>k).filter(k=>!S.hide.has(ser[k]));
  const gp=(W-ML)/meses.length;
  const bw=Math.max(4,Math.min(22,(gp-10)/Math.max(1,vis.length)));
  const G=clip(s,ML,top-18,W-ML,base-top+22);
  /* o valor só cabe sobre a barra com grupo largo e poucas séries */
  const rot=gp>=52 && vis.length<=3;
  meses.forEach((mIdx,i)=>{
    const cx=ML+i*gp+gp/2;
    const larg=vis.length*bw+(vis.length-1)*2;
    vis.forEach((k,j)=>{
      const v=d[ser[k]][i]; if(v===null) return;
      const h=(v/topo)*(base-top);
      const x=cx-larg/2+j*(bw+2);
      const rc=mk('rect',{x,y:base-h,width:bw,height:Math.max(1,h),rx:3,
        fill:cor(k),class:'by'});
      rc.style.animationDelay=((i*vis.length+j)*35)+'ms'; G.appendChild(rc);
      /* alterna a altura entre as barras do grupo: com valores parecidos os
         rótulos vizinhos se escreviam por cima um do outro */
      if(rot&&S.rot) T(s,x+bw/2,Math.max(top-4,base-h-4-(j%2?11:0)),kort(v),
        {class:'dlab','text-anchor':'middle',fill:cv('--text'),style:'font-size:8.5px'});
      const rs=dados(k,mIdx);
      hov(rc,nome(k)+' · '+MPTL[mIdx]+(porParc?' '+anos[0]:''),
        [['Valor',money(sumU(rs))],['PAX',n0.format(sumP(rs))],
         ['Custo por acesso',sumP(rs)?'US$ '+n2.format(sumU(rs)/sumP(rs)):'—']]);});
    T(s,cx,base+16,MPT[mIdx],{class:'tick','text-anchor':'middle'});});
  box.appendChild(s);
  const lg=document.getElementById('l-line'); lg.innerHTML='';
  ser.forEach((sk,k)=>{
    const el=document.createElement('div');
    el.className='it'+(S.hide.has(sk)?' off':'');
    el.innerHTML='<span class="sw" style="background:'+cor(k)+'"></span>'+nome(k);
    /* nunca deixar o clique apagar a última série visível */
    el.onclick=()=>{ if(S.hide.has(sk)) S.hide.delete(sk);
      else if(vis.length>1) S.hide.add(sk); cLine();};
    lg.appendChild(el);});
  let h='<div class="tw"><table><thead><tr><th>Mês</th>'+
    ser.map((sk,k)=>'<th class="num">'+nome(k)+'</th>').join('')+'</tr></thead><tbody>';
  meses.forEach((mIdx,i)=>{h+='<tr><td>'+MPTL[mIdx]+'</td>'+ser.map(sk=>
    '<td class="num">'+(d[sk][i]===null?'—':fmt(d[sk][i],S.met))+'</td>').join('')+'</tr>';});
  document.getElementById('tv-line').innerHTML=h+'</tbody></table></div>';
}

/* ================= YTD — colunas EM PÉ ================= */
function cYTD(){
  const box=document.getElementById('c-ytd'); box.innerHTML='';
  const p=ytd();
  if(!p){
    box.innerHTML='<p class="note-s">Selecione ao menos dois anos para comparar.</p>';return;}
  const blocos=[{t:'Valor (US$)',b:sumU(p.B),a:sumU(p.A),m:'usd'},
                {t:'PAX',b:sumP(p.B),a:sumP(p.A),m:'pax'}];
  /* 'top' abaixo do título e teto de 92% da área: o rótulo do topo da barra
     tem faixa própria e nunca sobe até o texto "Valor (US$)". */
  const H=230, base=H-50, top=48, TIT=15;
  const [s,W]=canvas(box,H);
  const gw=W/blocos.length;
  blocos.forEach((bl,i)=>{
    const cx=i*gw+gw/2, mx=Math.max(bl.a,bl.b)||1;
    const bw=Math.min(38,gw/2-30);
    T(s,cx,TIT,bl.t,{class:'tk2','text-anchor':'middle'});
    [[bl.b,p.pre,cv('--text-muted'),.45,-1],[bl.a,p.cur,cv('--primary'),1,1]].forEach((rw,k)=>{
      const [v,ano,c,op,side]=rw;
      const h=(v/mx)*(base-top)*0.92, x=cx+(side<0?-bw-14:14);
      const rc=mk('rect',{x,y:base-h,width:bw,height:Math.max(2,h),rx:5,
        fill:c,opacity:op,class:'by'});
      rc.style.animationDelay=((i*2+k)*80)+'ms'; s.appendChild(rc);
      T(s,x+bw/2,Math.max(top-4,base-h-6),kort(v),
        {class:'dlab','text-anchor':'middle',fill:cv('--text')});
      T(s,x+bw/2,base+15,String(ano),{class:'tick','text-anchor':'middle'});
      hov(rc,bl.t+' · '+ano,[[bl.t,(bl.m==='usd'?'US$ ':'')+fmt(v,bl.m)]]);
    });
    const d=pct(bl.a,bl.b);
    T(s,cx,base+34,pctS(d),{class:'dlab','text-anchor':'middle',
      fill:d>0?cv('--up'):d<0?cv('--dn'):cv('--text-muted'),style:'font-size:13px'});
  });
  s.appendChild(mk('line',{x1:0,x2:W,y1:base,y2:base,class:'axisline'}));
  box.appendChild(s);
  const cA=sumP(p.A)?sumU(p.A)/sumP(p.A):0, cB=sumP(p.B)?sumU(p.B)/sumP(p.B):0;
}

/* ================= ranking ================= */
function cRank(){
  const box=document.getElementById('c-rank'); box.innerHTML='';
  const rs=filtra(), tot=sum(rs)||1, ag={};
  rs.forEach(r=>ag[r.parc]=(ag[r.parc]||0)+val(r));
  const ord=Object.entries(ag).filter(e=>e[1]>0).sort((a,b)=>b[1]-a[1]);
  if(!ord.length){box.innerHTML='<p class="note-s">Sem dados no recorte.</p>';return;}
  const mx=ord[0][1], rh=30;
  const [s,W]=canvas(box,ord.length*rh+6);
  /* Ambas as faixas laterais são medidas pelo texto real que vai ocupá-las:
     o nome mais longo à esquerda e o maior valor à direita. Com largura fixa,
     a barra passava por cima do nome da parceira e do valor. */
  const nomeMax=Math.max(...ord.map(e=>cut(pn(e[0]),13).length));
  const lw=Math.min(120,Math.round(nomeMax*6.3)+10);
  const larg=Math.max(...ord.map(e=>(unit()+kort(e[1])).length));
  const vw=Math.round(larg*6.4)+52;      /* valor + coluna de % + folga */
  const bw=Math.max(40,W-lw-vw);
  ord.forEach((e,i)=>{
    const [p,v]=e,y=i*rh+3;
    T(s,0,y+16,cut(pn(p),13),{class:'tick',style:'font-size:11.5px'});
    const rc=mk('rect',{x:lw,y:y+4,width:Math.max(2,(v/mx)*bw),height:16,rx:4,
      fill:cv(PC[p]),class:'bx'});
    rc.style.animationDelay=(i*50)+'ms'; s.appendChild(rc);
    T(s,W-46,y+16,unit()+kort(v),{class:'tick','text-anchor':'end'});
    T(s,W,y+16,(v/tot*100).toFixed(1).replace('.',',')+'%',
      {class:'dlab','text-anchor':'end',fill:cv('--text-secondary')});
    const sel=rs.filter(r=>r.parc===p);
    hov(rc,pn(p),[['Valor',money(sumU(sel))],['PAX',n0.format(sumP(sel))],
      ['Custo por acesso',sumP(sel)?'US$ '+n2.format(sumU(sel)/sumP(sel)):'—'],
      ['Participação',(v/tot*100).toFixed(1).replace('.',',')+'%']]);});
  box.appendChild(s);
  let h='<div class="tw"><table><thead><tr><th>Parceira</th><th class="num">Valor</th>'+
    '<th class="num">PAX</th><th class="num">Custo/acesso</th><th class="num">Part.</th>'+
    '</tr></thead><tbody>';
  ord.forEach(e=>{const sel=rs.filter(r=>r.parc===e[0]);
    h+='<tr><td class="nm"><span class="sw" style="background:'+cv(PC[e[0]])+'"></span>'+
      pn(e[0])+'</td><td class="num">'+n2.format(sumU(sel))+'</td><td class="num">'+
      n0.format(sumP(sel))+'</td><td class="num">'+
      (sumP(sel)?n2.format(sumU(sel)/sumP(sel)):'—')+'</td><td class="num">'+
      (e[1]/tot*100).toFixed(1).replace('.',',')+'%</td></tr>';});
  document.getElementById('tv-rank').innerHTML=h+'</tbody></table></div>';
}

/* ========== contribuição YoY ========== */
function cContrib(){
  const box=document.getElementById('c-contrib'); box.innerHTML='';
  const p=ytd();
  if(!p){box.innerHTML='<p class="note-s">Selecione ao menos dois anos.</p>';return;}
  const ag={}; [...S.parc].forEach(x=>ag[x]=0);
  p.A.forEach(r=>ag[r.parc]+=val(r)); p.B.forEach(r=>ag[r.parc]-=val(r));
  const ct=Object.entries(ag).filter(e=>Math.abs(e[1])>0.005)
    .sort((a,b)=>Math.abs(b[1])-Math.abs(a[1]));
  if(!ct.length){box.innerHTML='<p class="note-s">Sem variação relevante.</p>';return;}
  const mx=Math.max(...ct.map(e=>Math.abs(e[1])))||1;
  const rh=29, lw=90, vw=84;
  const [s,W]=canvas(box,ct.length*rh+30);
  const eixo=lw+(W-lw-vw)/2, half=(W-lw-vw)/2-4;
  ct.forEach((e,i)=>{
    const [pa,v]=e, y=i*rh+4, pos=v>0;
    T(s,0,y+16,cut(pn(pa),12),{class:'tick',style:'font-size:11.5px'});
    const w=Math.max(2,(Math.abs(v)/mx)*half);
    const rc=mk('rect',{x:pos?eixo:eixo-w,y:y+5,width:w,height:15,rx:3.5,
      fill:pos?cv('--up'):cv('--dn'),class:pos?'bx':'bxr'});
    rc.style.animationDelay=(i*50)+'ms'; s.appendChild(rc);
    T(s,W,y+16,(pos?'+':'−')+unit()+kort(Math.abs(v)),
      {class:'dlab','text-anchor':'end',fill:pos?cv('--up'):cv('--dn')});
    hov(rc,pn(pa),[['Variação',(pos?'+':'−')+unit()+fmt(Math.abs(v),S.met)],
      [p.pre,unit()+fmt(sum(p.B.filter(r=>r.parc===pa)),S.met)],
      [p.cur,unit()+fmt(sum(p.A.filter(r=>r.parc===pa)),S.met)]]);});
  s.appendChild(mk('line',{x1:eixo,x2:eixo,y1:0,y2:ct.length*rh+2,class:'axisline'}));
  const ini=sum(p.B),fim=sum(p.A),d=pct(fim,ini);

  box.appendChild(s);
  let h='<div class="tw"><table><thead><tr><th>Parceira</th><th class="num">'+p.pre+
    '</th><th class="num">'+p.cur+'</th><th class="num">Variação</th>'+
    '<th class="num">%</th></tr></thead><tbody>';
  ct.forEach(function(e){
    const pa=e[0], v=e[1];
    const b=sum(p.B.filter(r=>r.parc===pa)), a=sum(p.A.filter(r=>r.parc===pa));
    const pc=pct(a,b);
    h+='<tr><td class="nm"><span class="sw" style="background:'+cv(PC[pa])+'"></span>'+
      pn(pa)+'</td><td class="num">'+fmt(b,S.met)+'</td><td class="num">'+
      fmt(a,S.met)+'</td><td class="num '+(v>0?'up':'dn')+'">'+
      (v>0?'+':'−')+fmt(Math.abs(v),S.met)+'</td><td class="num '+cls(pc)+'">'+
      pctS(pc)+'</td></tr>';});
  h+='</tbody><tfoot><tr><td>Total</td><td class="num">'+fmt(ini,S.met)+
    '</td><td class="num">'+fmt(fim,S.met)+'</td><td class="num '+
    (fim-ini>0?'up':'dn')+'">'+((fim-ini)>0?'+':'−')+fmt(Math.abs(fim-ini),S.met)+
    '</td><td class="num '+cls(d)+'">'+pctS(d)+'</td></tr></tfoot></table></div>';
  document.getElementById('tv-contrib').innerHTML=h;
}


/* ================= participação por sala ================= */
function cSplit(){
  const box=document.getElementById('c-split'); box.innerHTML='';
  const rs=filtra(), tot=sum(rs)||1;
  const ag=META.salas.map(sa=>[sa,sum(rs.filter(r=>r.sala===sa))])
    .filter(e=>e[1]>0).sort((a,b)=>b[1]-a[1]);
  if(!ag.length){box.innerHTML='<p class="note-s">Sem dados.</p>';return;}
  const [s,W]=canvas(box,52);
  let x=0;
  ag.forEach((e,i)=>{
    const w=(e[1]/tot)*W, c=cv(SEQ[Math.max(0,5-i)]);
    const rc=mk('rect',{x,y:0,width:Math.max(1,w-2),height:44,rx:5,fill:c,class:'bx'});
    rc.style.transformOrigin=x+'px 22px';
    rc.style.animationDelay=(i*75)+'ms'; s.appendChild(rc);
    if(w>48) T(s,x+w/2-1,27,(e[1]/tot*100).toFixed(0)+'%',
      {class:'dlab','text-anchor':'middle',fill:i<2?'#fff':cv('--text'),style:'font-size:12px'});
    const sel=rs.filter(r=>r.sala===e[0]);
    hov(rc,ln(e[0]),[['Valor',money(sumU(sel))],['PAX',n0.format(sumP(sel))],
      ['Participação',(e[1]/tot*100).toFixed(1).replace('.',',')+'%']]);
    x+=w;});
  box.appendChild(s);
  let h='<div class="tw" style="margin-top:12px"><table><tbody>';
  ag.forEach((e,i)=>{h+='<tr><td class="nm"><span class="sw" style="background:'+
    cv(SEQ[Math.max(0,5-i)])+'"></span>'+ln(e[0])+'</td><td class="num">'+
    unit()+fmt(e[1],S.met)+'</td><td class="num">'+
    (e[1]/tot*100).toFixed(1).replace('.',',')+'%</td></tr>';});
  box.insertAdjacentHTML('beforeend',h+'</tbody></table></div>');
}

/* ============ penetração pax pagantes ============ */
function payingPax(ano,meses,salas){
  let d=0, faltando=[];
  salas.forEach(sa=>meses.forEach(m=>{
    const v=META.pagantes[sa+'|'+ano+'|'+m];
    if(v) d+=v; else faltando.push(ln(sa)+' '+mn(m));}));
  return {den:d,faltando};
}
function cPaying(){
  const box=document.getElementById('c-paying'); box.innerHTML='';
  const p=ytd();
  if(!p){box.innerHTML='<p class="note-s">Selecione ao menos dois anos.</p>';return;}
  const salas=[...S.sala], w=janelaL3M(), linhas=[];
  [{ano:p.pre},{ano:p.cur}].forEach(o=>{
    const num=w?sum(ROWS.filter(r=>r.ano===o.ano&&w.jan.includes(r.mes)
      &&S.sala.has(r.sala)&&S.parc.has(r.parc)),'pax'):0;
    const yp=payingPax(o.ano,p.meses,salas);
    const lp=w?payingPax(o.ano,w.jan,salas):{den:0,faltando:[]};
    linhas.push({ano:o.ano, deck:yp.den?num/yp.den*100:null,
      like:lp.den?num/lp.den*100:null, falta:yp.faltando.length+lp.faltando.length});});
  const mx=Math.max(...linhas.flatMap(l=>[l.deck||0,l.like||0]),0.1);
  const [s,W]=canvas(box,linhas.length*72+38);
  const lw=44, vw=68, bw=Math.max(40,W-lw-vw);
  T(s,lw,11,'Base do slide · acessos L3M ÷ pax pagantes do acumulado',{class:'tick'});
  linhas.forEach((l,i)=>{
    const y=i*72+20;
    T(s,0,y+16,String(l.ano),{class:'tk2'});
    [[l.deck,cv('--primary')],[l.like,cv('--blue')]].forEach((rw,k)=>{
      const [v,c]=rw; if(v===null) return;
      const yy=y+3+k*26;
      const rc=mk('rect',{x:lw,y:yy,width:Math.max(2,(v/mx)*bw),height:19,rx:4,
        fill:c,class:'bx'});
      rc.style.animationDelay=((i*2+k)*70)+'ms'; s.appendChild(rc);
      T(s,W,yy+13.5,v.toFixed(2).replace('.',',')+'%',
        {class:'dlab','text-anchor':'end',fill:cv('--text')});
      hov(rc,l.ano+(k?' · like-for-like':' · base do slide'),
        [['Penetração',v.toFixed(2).replace('.',',')+'%']]);});});
  T(s,lw,linhas.length*72+32,'Like-for-like · acessos L3M ÷ pax pagantes L3M',{class:'tick'});
  box.appendChild(s);
  const falta=linhas.reduce((a,l)=>a+l.falta,0);
  box.insertAdjacentHTML('beforeend','<div class="legend">'+
    '<div class="it"><span class="sw" style="background:'+cv('--primary')+'"></span>Base do slide</div>'+
    '<div class="it"><span class="sw" style="background:'+cv('--blue')+'"></span>Like-for-like</div></div>'+
    (falta?'<div class="pill-w">⚠ '+falta+' célula(s) de mês/sala sem preenchimento em PAX PAGANTES</div>':''));
}

/* ================= L3M ================= */
/* L3M são sempre TRÊS meses fechados até o último mês do filtro.
   O início da janela ignora S.m1 de propósito: filtrar só julho deve mostrar
   MAI-JUN-JUL, não um L3M de um mês só. */
function janelaL3M(){
  const anos=[...S.ano].sort(); if(!anos.length) return null;
  const cur=anos[anos.length-1];
  const disp=[...new Set(ROWS.map(r=>r.ano))].sort();
  const pre=disp.filter(a=>a<cur).pop();
  if(pre===undefined) return null;
  const mc=ROWS.filter(r=>r.ano===cur).map(r=>mi(r.mes));
  if(!mc.length) return null;
  const fim=Math.min(S.m2,Math.max(...mc));
  const jan=[]; for(let i=Math.max(0,fim-2);i<=fim;i++) jan.push(MK[i]);
  return {cur,pre,jan};
}
function cL3M(){
  const box=document.getElementById('c-l3m'); box.innerHTML='';
  const lg=document.getElementById('l-l3m'); lg.innerHTML='';
  const w=janelaL3M();
  if(!w){box.innerHTML='<p class="note-s">Selecione ao menos dois anos.</p>';
    document.getElementById('c-l3msalas').innerHTML='';return;}
  document.getElementById('t-l3m').textContent='L3M '+w.cur+' vs '+w.pre;
  const csl=document.getElementById('cs-l3m');
  if(csl) csl.textContent='('+w.jan.map(mn).join(', ')+')';
  const parcs=META.parcs.filter(p=>S.parc.has(p)&&(META.contratos[p]||{}).pax_hist>0);
  const A={},B={};
  parcs.forEach(p=>{const f=a=>sum(ROWS.filter(r=>r.parc===p&&r.ano===a
    &&w.jan.includes(r.mes)&&S.sala.has(r.sala)));
    A[p]=f(w.cur);B[p]=f(w.pre);});
  const ord=parcs.filter(p=>A[p]>0||B[p]>0).sort((a,b)=>A[b]-A[a]);
  if(!ord.length){box.innerHTML='<p class="note-s">Sem dados na janela.</p>';return;}
  const max=Math.max(...ord.map(p=>Math.max(A[p],B[p])),1);
  const H=250,base=H-52,top=30,ml=54;
  const [s,W]=canvas(box,H);
  const gp=(W-ml)/ord.length, bw=Math.min(26,gp/2-4);
  /* A altura DEVE usar o mesmo teto das linhas de grade. Usando 'max' aqui e
     'tk' no eixo, a maior barra encostava no topo enquanto a grade dizia outro
     valor — era o gráfico "cortando" o número real. */
  const tk=ticks(max,3), topo=tk[tk.length-1]||1;
  tk.forEach(t=>{const y=base-(base-top)*(t/topo);
    s.appendChild(mk('line',{x1:ml,x2:W,y1:y,y2:y,class:'gridline'}));
    T(s,ml-8,y+3.5,unit()+kort(t),{class:'tick','text-anchor':'end'});});
  s.appendChild(mk('line',{x1:ml,x2:W,y1:base,y2:base,class:'axisline'}));
  const G=clip(s,ml,top-16,W-ml,base-top+20);
  const doisRot=gp>=88;   /* dois rótulos por par só cabem em coluna larga */
  ord.forEach((p,i)=>{
    const cx=ml+i*gp+gp/2;
    /* laranja no ano atual, cinza no anterior — igual à legenda do cartão.
       Antes cada par usava a cor da parceira e a legenda não fazia sentido. */
    [[B[p],cv('--text-muted'),w.pre,-1,.45],[A[p],cv('--primary'),w.cur,1,1]].forEach((rw,k)=>{
      const [v,c,ano,side,op]=rw, h=(v/topo)*(base-top);
      const rc=mk('rect',{x:cx+(side<0?-bw-1:1),y:base-h,width:bw,
        height:Math.max(1,h),rx:3,fill:c,opacity:op,class:'by'});
      rc.style.animationDelay=((i*2+k)*45)+'ms'; G.appendChild(rc);
      /* Só o ano atual recebe rótulo, e acima da barra. Mostrar os dois valores
         enchia o gráfico e eles colidiam; o ano anterior fica no tooltip. */
      if(v>0&&S.rot&&k===1)
        T(s,cx+bw/2+1,Math.max(top-3,base-h-6),kort(v),
          {class:'dlab','text-anchor':'middle',fill:cv('--text'),style:'font-size:9px'});
      hov(rc,pn(p)+' · '+ano,[[metName(),unit()+fmt(v,S.met)],
        ['Janela',w.jan.map(mn).join(', ')]]);});
    T(s,cx,base+15,cut(pn(p),9),{class:'tick','text-anchor':'middle'});
    const d=pct(A[p],B[p]);
    T(s,cx,base+30,pctS(d),{class:'tick','text-anchor':'middle',
      fill:d>0?cv('--up'):d<0?cv('--dn'):cv('--text-muted'),style:'font-weight:700'});});
  box.appendChild(s);
  lg.innerHTML='<div class="it"><span class="sw" style="background:'+cv('--text-muted')+
    ';opacity:.45"></span>'+w.pre+'</div><div class="it"><span class="sw" style="background:'+
    cv('--p1')+'"></span>'+w.cur+'</div>';
  let h='<div class="tw"><table><thead><tr><th>Parceira</th><th class="num">'+w.pre+
    '</th><th class="num">'+w.cur+'</th><th class="num">Δ</th></tr></thead><tbody>';
  ord.forEach(p=>{const d=pct(A[p],B[p]);
    h+='<tr><td class="nm">'+pn(p)+'</td><td class="num">'+fmt(B[p],S.met)+
      '</td><td class="num">'+fmt(A[p],S.met)+'</td><td class="num '+cls(d)+'">'+
      pctS(d)+'</td></tr>';});
  document.getElementById('tv-l3m').innerHTML=h+'</tbody></table></div>';
  cL3MSalas(w);
}
function cL3MSalas(w){
  const host=document.getElementById('c-l3msalas'); host.innerHTML='';
  const salas=META.salas.filter(x=>S.sala.has(x));
  const tot={};
  salas.forEach(sa=>tot[sa]=[w.pre,w.cur].map(a=>w.jan.map(m=>
    sum(ROWS.filter(r=>r.sala===sa&&r.ano===a&&r.mes===m&&S.parc.has(r.parc))))));
  /* Escala própria por sala. Com escala compartilhada, GIG DOM fica com ~4%
     da altura de GRU INTER e some. O rótulo do teto do eixo vai no cartão para
     ninguém comparar altura entre salas — a comparação certa é o % e o total. */
  salas.forEach((sa,k)=>{
    let max=0; tot[sa].forEach(ar=>ar.forEach(v=>{if(v>max)max=v;}));
    const card=document.createElement('div');
    card.className='card rise'; card.style.setProperty('--i',k);
    const a=tot[sa][1].reduce((x,y)=>x+y,0), b=tot[sa][0].reduce((x,y)=>x+y,0);
    const d=pct(a,b);
    card.innerHTML='<h3>'+ln(sa)+'</h3><p class="cs">'+unit()+kort(a)+' · '+
      '<strong class="'+cls(d)+'">'+pctS(d)+'</strong> contra '+w.pre+
      ' <span style="opacity:.6">· escala até '+unit()+kort(max)+'</span></p>'+
      '<div class="hc"></div>';
    host.appendChild(card);
    const box=card.querySelector('.hc');
    const H=152,base=H-26,top=34;   /* folga no topo para o rótulo de valor */
    const [s,W]=canvas(box,H);
    const gp=W/w.jan.length, bw=Math.min(19,gp/2-4);
    /* Dois rótulos por mês só cabem se o grupo for largo. Apertado, mostra só
       o do ano atual — antes os dois se escreviam por cima um do outro. */
    const doisRotulos=gp>=72;
    s.appendChild(mk('line',{x1:0,x2:W,y1:base,y2:base,class:'axisline'}));
    const G=clip(s,0,top-16,W,base-top+20);
    w.jan.forEach((m,i)=>{
      const cx=i*gp+gp/2;
      [[tot[sa][0][i],cv('--text-muted'),w.pre,-1,.45],
       [tot[sa][1][i],cv('--primary'),w.cur,1,1]].forEach((rw,j)=>{
        const [v,c,ano,side,op]=rw, h=max?(v/max)*(base-top)*0.9:0;
        const rc=mk('rect',{x:cx+(side<0?-bw-1:1),y:base-h,width:bw,
          height:Math.max(1,h),rx:2.5,fill:c,opacity:op,class:'by'});
        rc.style.animationDelay=((i*2+j)*45)+'ms'; G.appendChild(rc);
        /* Só o ano atual, acima da barra. O do ano anterior sai no tooltip. */
        if(v>0&&j===1)
          T(s,cx+bw/2+1,Math.max(9,base-h-5),kort(v),
            {class:'dlab','text-anchor':'middle',fill:cv('--text'),
             style:'font-size:8.5px'});
        hov(rc,ln(sa)+' · '+mlg(m)+' '+ano,[[metName(),unit()+fmt(v,S.met)]]);});
      T(s,cx,base+14,mn(m),{class:'tick','text-anchor':'middle'});});
    box.appendChild(s);});
}

/* ========== composição mensal empilhada por parceira ========== */
function cStack(){
  const box=document.getElementById('c-stack'); box.innerHTML='';
  const lg=document.getElementById('l-stack'); lg.innerHTML='';
  const anos=[...S.ano].sort(), ano=anos[anos.length-1];
  document.getElementById('t-stk').textContent='· '+ano;
  const parcs=META.parcs.filter(p=>S.parc.has(p)&&(META.contratos[p]||{}).pax_hist>0);
  const meses=[]; for(let i=S.m1;i<=S.m2;i++){
    if(ROWS.some(r=>r.ano===ano&&r.mes===MK[i])) meses.push(i);}
  if(!meses.length){box.innerHTML='<p class="note-s">Sem dados no recorte.</p>';return;}
  const dat=meses.map(i=>parcs.map(p=>sum(ROWS.filter(r=>r.ano===ano&&r.mes===MK[i]
    &&r.parc===p&&S.sala.has(r.sala)))));
  const totais=dat.map(col=>col.reduce((a,b)=>a+b,0));
  const max=Math.max(...totais,1);
  const H=250,base=H-40,top=20,ml=52;
  const [s,W]=canvas(box,H);
  const gp=(W-ml)/meses.length, bw=Math.min(46,gp-12);
  const tk=ticks(max,4);
  tk.forEach(t=>{const y=base-(base-top)*(t/tk[tk.length-1]);
    s.appendChild(mk('line',{x1:ml,x2:W,y1:y,y2:y,class:'gridline'}));
    T(s,ml-8,y+3.5,unit()+kort(t),{class:'tick','text-anchor':'end'});});
  s.appendChild(mk('line',{x1:ml,x2:W,y1:base,y2:base,class:'axisline'}));
  const G=clip(s,ml,top-16,W-ml,base-top+20);
  meses.forEach((mIdx,i)=>{
    const cx=ml+i*gp+gp/2, x=cx-bw/2;
    let acc=0;
    parcs.forEach((p,k)=>{
      const v=dat[i][k]; if(v<=0) return;
      const h=(v/tk[tk.length-1])*(base-top);
      const y=base-((acc+v)/tk[tk.length-1])*(base-top);
      const rc=mk('rect',{x,y,width:bw,height:Math.max(1,h-2),rx:2,
        fill:ciaCor(p),class:'by'});
      rc.style.animationDelay=(i*40+k*18)+'ms'; G.appendChild(rc);
      hov(rc,pn(p)+' · '+MPTL[mIdx]+' '+ano,
        [[metName(),unit()+fmt(v,S.met)],
         ['Do mês',(v/(totais[i]||1)*100).toFixed(1).replace('.',',')+'%']]);
      acc+=v;});
    T(s,cx,base-((totais[i])/tk[tk.length-1])*(base-top)-7,kort(totais[i]),
      {class:'dlab','text-anchor':'middle',fill:cv('--text'),style:'font-size:9.5px'});
    T(s,cx,base+15,MPT[mIdx],{class:'tick','text-anchor':'middle'});});
  box.appendChild(s);
  parcs.forEach(p=>{const el=document.createElement('div'); el.className='it';
    el.innerHTML='<span class=\"sw\" style=\"background:'+ciaCor(p)+'\"></span>'+pn(p);
    lg.appendChild(el);});
}

/* ========== índices complementares ========== */
function cIndices(){
  const box=document.getElementById('c-idx'); box.innerHTML='';
  const rs=filtra();
  if(!rs.length){box.innerHTML='<p class="note-s">Sem dados no recorte.</p>';return;}
  const tU=sumU(rs), tP=sumP(rs);
  const porMes={};
  rs.forEach(r=>{const k=r.mes+'/'+r.ano;
    porMes[k]=porMes[k]||[0,0]; porMes[k][0]+=r['usd_'+S.fx]; porMes[k][1]+=r.pax;});
  const meses=Object.entries(porMes);
  const maior=meses.reduce((a,b)=>b[1][0]>a[1][0]?b:a);
  const menor=meses.reduce((a,b)=>b[1][0]<a[1][0]?b:a);
  const eur=rs.filter(r=>r.moeda==='EUR'), expo=tU?sumU(eur)/tU*100:0;
  const salas={};
  rs.filter(r=>r.pax>0).forEach(r=>{salas[r.sala]=salas[r.sala]||[0,0];
    salas[r.sala][0]+=r['usd_'+S.fx]; salas[r.sala][1]+=r.pax;});
  const cps=Object.entries(salas).map(([k,v])=>[k,v[0]/v[1]]).sort((a,b)=>b[1]-a[1]);
  const rotulo=k=>k.split('/')[0]+'/'+k.split('/')[1];
  const cards=[
    ['Mês de maior custo',rotulo(maior[0]),
     'US$ '+n0.format(maior[1][0])+' · '+n0.format(maior[1][1])+' acessos'],
    ['Mês de menor custo',rotulo(menor[0]),
     'US$ '+n0.format(menor[1][0])+' · '+n0.format(menor[1][1])+' acessos'],
    ['Sala mais cara por acesso',cps.length?ln(cps[0][0]):'—',
     cps.length?'US$ '+n2.format(cps[0][1])+' por acesso':'—'],
    ['Sala mais barata por acesso',cps.length?ln(cps[cps.length-1][0]):'—',
     cps.length?'US$ '+n2.format(cps[cps.length-1][1])+' por acesso':'—'],
  ];
  box.innerHTML=cards.map((c,i)=>
    '<div class="card kpi rise" style="--i:'+i+'"><div class="lbl">'+c[0]+
    '</div><div class="val" style="font-size:19px">'+c[1]+'</div>'+
    '<div class="note" title="'+c[2]+'">'+c[2]+'</div></div>').join('');
}

/* ================= movimentos L3M =================
   Mesmo desenho da Evolução ano a ano — barras divergentes de uma linha zero,
   que mostram de cara quem melhorou e quem piorou. Um seletor no canto permite
   olhar sala a sala (ex.: só Avianca em GRU International). */
function cMov(){
  const box=document.getElementById('c-mov'); box.innerHTML='';
  const w=janelaL3M();
  const selBox=document.getElementById('mov-sala');
  if(!w){box.innerHTML='<p class="note-s">Sem ano anterior para comparar.</p>';
    if(selBox) selBox.innerHTML=''; return;}

  /* seletor de sala do cartão */
  if(selBox){
    selBox.innerHTML='';
    const sel=document.createElement('select');
    sel.style.cssText='font-size:11px;padding:3px 7px;border-radius:7px';
    const opts=[['todas','Todas as salas']].concat(
      META.salas.filter(x=>S.sala.has(x)).map(x=>[x,ln(x)]));
    opts.forEach(o=>{const op=document.createElement('option');
      op.value=o[0]; op.textContent=o[1]; sel.appendChild(op);});
    if(!opts.some(o=>o[0]===S.movSala)) S.movSala='todas';
    sel.value=S.movSala;
    sel.onchange=()=>{S.movSala=sel.value; cMov();};
    selBox.appendChild(sel);
  }
  const filtroSala=r=>S.movSala==='todas'?S.sala.has(r.sala):r.sala===S.movSala;
  const naJanela=(a)=>ROWS.filter(r=>r.ano===a&&w.jan.includes(r.mes)
    &&filtroSala(r)&&S.parc.has(r.parc));
  const A=naJanela(w.cur), B=naJanela(w.pre);

  const ag={}; [...S.parc].forEach(x=>ag[x]=0);
  A.forEach(r=>ag[r.parc]+=val(r)); B.forEach(r=>ag[r.parc]-=val(r));
  const ct=Object.entries(ag).filter(e=>Math.abs(e[1])>0.005)
    .sort((a,b)=>Math.abs(b[1])-Math.abs(a[1]));
  if(!ct.length){box.innerHTML='<p class="note-s">Sem variação na janela '+
    w.jan.map(mn).join(', ')+'.</p>';return;}
  const mx=Math.max(...ct.map(e=>Math.abs(e[1])))||1;
  const rh=29, lw=92;
  /* faixa da direita medida pelo maior rótulo, para a barra não passar por baixo */
  const larg=Math.max(...ct.map(e=>(unit()+kort(Math.abs(e[1]))).length));
  const vw=Math.round(larg*6.4)+16;
  const [s,W]=canvas(box,ct.length*rh+30);
  const eixo=lw+(W-lw-vw)/2, half=(W-lw-vw)/2-4;
  ct.forEach((e,i)=>{
    const [pa,v]=e, y=i*rh+4, pos=v>0;
    T(s,0,y+16,cut(pn(pa),12),{class:'tick',style:'font-size:11.5px'});
    const wid=Math.max(2,(Math.abs(v)/mx)*half);
    const rc=mk('rect',{x:pos?eixo:eixo-wid,y:y+5,width:wid,height:15,rx:3.5,
      fill:pos?cv('--up'):cv('--dn'),class:pos?'bx':'bxr'});
    rc.style.animationDelay=(i*50)+'ms'; s.appendChild(rc);
    T(s,W,y+16,(pos?'+':'−')+unit()+kort(Math.abs(v)),
      {class:'dlab','text-anchor':'end',fill:pos?cv('--up'):cv('--dn')});
    const a=sum(A.filter(r=>r.parc===pa)), b=sum(B.filter(r=>r.parc===pa));
    hov(rc,pn(pa)+(S.movSala==='todas'?'':' · '+ln(S.movSala)),
      [['Variação',(pos?'+':'−')+unit()+fmt(Math.abs(v),S.met)],
       [w.pre,unit()+fmt(b,S.met)],[w.cur,unit()+fmt(a,S.met)],
       ['Janela',w.jan.map(mn).join(', ')]]);});
  s.appendChild(mk('line',{x1:eixo,x2:eixo,y1:0,y2:ct.length*rh+2,class:'axisline'}));
  const ini=sum(B), fim=sum(A), d=pct(fim,ini);
  box.appendChild(s);

  let tb='<div class="tw"><table><thead><tr><th>Parceira</th>'+
    (S.movSala==='todas'?'<th>Sala VIP</th>':'')+
    '<th class="num">'+w.pre+'</th><th class="num">'+w.cur+'</th>'+
    '<th class="num">Variação</th><th class="num">%</th></tr></thead><tbody>';
  ct.forEach(function(e){
    const pa=e[0], v=e[1];
    const a=sum(A.filter(r=>r.parc===pa)), b=sum(B.filter(r=>r.parc===pa));
    const pc=pct(a,b);
    tb+='<tr><td class="nm"><span class="sw" style="background:'+cv(PC[pa])+'"></span>'+
      pn(pa)+'</td>'+(S.movSala==='todas'?'<td>todas</td>':'')+
      '<td class="num">'+fmt(b,S.met)+'</td><td class="num">'+fmt(a,S.met)+
      '</td><td class="num '+(v>0?'up':'dn')+'">'+(v>0?'+':'−')+
      fmt(Math.abs(v),S.met)+'</td><td class="num '+cls(pc)+'">'+pctS(pc)+
      '</td></tr>';});
  tb+='</tbody><tfoot><tr><td'+(S.movSala==='todas'?' colspan="2"':'')+'>Total</td>'+
    '<td class="num">'+fmt(ini,S.met)+'</td><td class="num">'+fmt(fim,S.met)+
    '</td><td class="num '+(fim-ini>0?'up':'dn')+'">'+((fim-ini)>0?'+':'−')+
    fmt(Math.abs(fim-ini),S.met)+'</td><td class="num '+cls(d)+'">'+pctS(d)+
    '</td></tr></tfoot></table></div>';
  document.getElementById('tv-mov').innerHTML=tb;
}

/* ================= leitura executiva (3 visões) ================= */
const RD=[{k:'',f:rdVariacao},{k:'',f:rdCusto},{k:'',f:rdConcentracao}];

/* Só descreve o escopo quando há filtro ativo. Sem filtro, a frase
   "todos os anos · todas as salas · JAN–DEZ" não informa nada e polui. */
function escopo(){
  const partes=[];
  const anos=[...S.ano].sort();
  if(anos.length!==META.anos.length) partes.push(anos.join(' e '));
  if(S.sala.size!==META.salas.length) partes.push([...S.sala].map(ln).join(', '));
  if(S.parc.size!==META.parcs.length) partes.push([...S.parc].map(pn).join(', '));
  if(S.modo==='unico') partes.push(MPTL[S.m1]);
  else if(S.m1!==0||S.m2!==11) partes.push(MPT[S.m1]+'–'+MPT[S.m2]);
  return partes.length? partes.join(' · ') : '';
}
const cap=s=>s?s.charAt(0).toUpperCase()+s.slice(1):s;

/* Como a leitura se refere ao recorte: um mês, um intervalo, uma sala. */
function contexto(){
  const per = S.modo==='unico' ? MPTL[S.m1]
            : (S.m1===0&&S.m2===11) ? 'o acumulado'
            : MPT[S.m1]+'–'+MPT[S.m2];
  const sala = S.sala.size===1 ? ln([...S.sala][0]) : null;
  return {per, sala, alvo: sala ? (per+' em '+sala) : per};
}

/* Cada leitura devolve {tom, titulo, corpo}. Cenário geral, frase curta: o
   título entrega o fato e o corpo mostra os números que o sustentam. */
/* junta nomes numa frase: 'Avianca, KLM e Ethiopian' */
function lista(a){ if(!a.length) return '';
  if(a.length===1) return a[0];
  return a.slice(0,-1).join(', ')+' e '+a[a.length-1]; }

function rdConcentracao(rs,tot){
  const out=[], ag={}, cx=contexto();
  rs.forEach(r=>ag[r.parc]=(ag[r.parc]||0)+val(r));
  const ord=Object.entries(ag).filter(e=>e[1]>0).sort((a,b)=>b[1]-a[1]);
  if(ord.length){
    const p=ord[0][0], v=ord[0][1], sel=rs.filter(r=>r.parc===p);
    out.push({tom:'alta',titulo:pn(p)+' concentra '+(v/tot*100).toFixed(0)+'% do custo',
      corpo:'São <strong>'+unit()+fmt(v,S.met)+'</strong> de '+unit()+fmt(tot,S.met)+
        ' em '+cx.alvo+', com '+n0.format(sumP(sel))+' acessos a US$ '+
        n2.format(sumP(sel)?sumU(sel)/sumP(sel):0)+' cada.'});
  }
  if(ord.length>=3){
    const t3=ord.slice(0,3).reduce((a,e)=>a+e[1],0), sh=t3/tot;
    out.push({tom:sh>.7?'atencao':'info',
      titulo:'Top 3 vale '+(sh*100).toFixed(0)+'% do total',
      corpo:ord.slice(0,3).map(e=>pn(e[0])).join(', ')+' somam '+unit()+fmt(t3,S.met)+'. '+
        (sh>.7?'Quase todo o custo sai dessas três.'
              :'O custo está bem dividido entre as parceiras.')});
  }
  if(cx.sala){
    const todas=ROWS.filter(r=>S.ano.has(r.ano)&&S.parc.has(r.parc)
      &&mi(r.mes)>=S.m1&&mi(r.mes)<=S.m2);
    const tg=sum(todas)||1, pxS=sumP(rs), pxG=sumP(todas);
    const cS=pxS?sumU(rs)/pxS:0, cG=pxG?sumU(todas)/pxG:0;
    out.push({tom:'info',titulo:cx.sala+' representa '+(tot/tg*100).toFixed(0)+'% das salas',
      corpo:'De '+unit()+fmt(tg,S.met)+' movimentados em '+cx.per+
        ' nas quatro salas, <strong>'+unit()+fmt(tot,S.met)+'</strong> saíram desta.'+
        (cS&&cG?' O acesso custa US$ '+n2.format(cS)+' aqui, contra US$ '+
          n2.format(cG)+' na média geral.':'')});
  }else{
    const sag={}; rs.forEach(r=>sag[r.sala]=(sag[r.sala]||0)+val(r));
    const so=Object.entries(sag).filter(e=>e[1]>0).sort((a,b)=>b[1]-a[1]);
    if(so.length) out.push({tom:'info',
      titulo:ln(so[0][0])+' puxa '+(so[0][1]/tot*100).toFixed(0)+'% do custo',
      corpo:'<strong>'+unit()+fmt(so[0][1],S.met)+'</strong> em '+cx.per+
        (so.length>1?', contra '+unit()+fmt(so[1][1],S.met)+' na '+ln(so[1][0])+'.':'.')+
        ' As internacionais pesam mais porque reúnem os contratos de fee mais alto.'});
  }
  if(ord.length>1){
    const u=ord[ord.length-1], sel=rs.filter(r=>r.parc===u[0]);
    out.push({tom:'info',titulo:pn(u[0])+' tem presença mínima',
      corpo:unit()+fmt(u[1],S.met)+' em '+cx.per+', ou '+
        (u[1]/tot*100).toFixed(1).replace('.',',')+'% do total, com '+
        n0.format(sumP(sel))+' acessos. O contrato segue ativo.'});
  }
  return out;
}

function rdVariacao(rs,tot){
  const out=[], p=ytd(), cx=contexto();
  if(!p) return [{tom:'atencao',titulo:'Sem base de comparação',
    corpo:'Não há ano anterior na base para comparar este recorte.'}];
  const ac={}; p.A.forEach(r=>ac[r.parc]=(ac[r.parc]||0)+val(r));
  p.B.forEach(r=>ac[r.parc]=(ac[r.parc]||0)-val(r));
  const cs=Object.entries(ac).sort((a,b)=>b[1]-a[1]);
  const va=sum(p.A), vb=sum(p.B), dif=va-vb, d=pct(va,vb);
  out.push({tom:d>0?'alta':'queda',
    titulo:cap(cx.per)+' '+(d>0?'subiu':'caiu')+' '+pctS(d)+' contra '+p.pre,
    corpo:'<strong>'+unit()+fmt(va,S.met)+'</strong> em '+p.cur+', contra '+
      unit()+fmt(vb,S.met)+' em '+p.pre+'. Diferença de '+
      (dif>0?'':'−')+unit()+fmt(Math.abs(dif),S.met)+'.'+
      (cx.sala?' Recorte: '+cx.sala+'.':'')});
  if(cs.length&&cs[0][1]>0){
    const p0=cs[0][0];
    const a0=sum(p.A.filter(r=>r.parc===p0)), b0=sum(p.B.filter(r=>r.parc===p0));
    const share=dif>0?cs[0][1]/dif*100:null;
    out.push({tom:'alta',titulo:'Maior alta: '+pn(p0),
      corpo:'De '+unit()+fmt(b0,S.met)+' para <strong>'+unit()+fmt(a0,S.met)+
        '</strong>, ou '+pctS(pct(a0,b0))+'. '+
        (share===null?'Ainda assim o período fechou abaixo do ano anterior.'
         :share>150?'Sozinha explica todo o aumento do período.'
         :'Representa '+share.toFixed(0)+'% do aumento do período.')});
  }
  const ult=cs[cs.length-1];
  if(ult&&ult[1]<0){
    const a1=sum(p.A.filter(r=>r.parc===ult[0])), b1=sum(p.B.filter(r=>r.parc===ult[0]));
    const ct=META.contratos[ult[0]]||{};
    out.push({tom:'queda',titulo:'Maior queda: '+pn(ult[0]),
      corpo:'De '+unit()+fmt(b1,S.met)+' para <strong>'+unit()+fmt(a1,S.met)+
        '</strong>, ou '+pctS(pct(a1,b1))+'.'+
        (ct.fee?' O fee segue em '+(ct.moeda==='EUR'?'€ ':'US$ ')+n0.format(ct.fee)+
          ', então a queda é de movimento e não de preço.':'')});
  }
  const sobeL=cs.filter(e=>e[1]>0.005).map(e=>pn(e[0]));
  const caiL=cs.filter(e=>e[1]<-0.005).map(e=>pn(e[0])).reverse();
  if(sobeL.length+caiL.length>1) out.push({tom:sobeL.length>=caiL.length?'alta':'queda',
    titulo:sobeL.length+' em alta, '+caiL.length+' em queda',
    corpo:(sobeL.length?'<strong>'+lista(sobeL)+'</strong> '+
             (sobeL.length>1?'estão':'está')+' em alta.':'')+
          (sobeL.length&&caiL.length?' ':'')+
          (caiL.length?'<strong>'+lista(caiL)+'</strong> '+
             (caiL.length>1?'caíram':'caiu')+' contra '+p.pre+'.':'')});
  if(cx.sala){
    const pa=sumP(p.A), pb=sumP(p.B);
    const ca=pa?sumU(p.A)/pa:0, cb=pb?sumU(p.B)/pb:0;
    out.push({tom:pct(ca,cb)<0?'alta':'atencao',
      titulo:'Custo por acesso: US$ '+n2.format(ca),
      corpo:'<strong>'+n0.format(pa)+'</strong> acessos ('+pctS(pct(pa,pb))+
        ' contra '+p.pre+') a US$ '+n2.format(ca)+' cada, ante US$ '+n2.format(cb)+
        ' no ano anterior.'});
  }else{
    const sa={}; p.A.forEach(r=>sa[r.sala]=(sa[r.sala]||0)+val(r));
    p.B.forEach(r=>sa[r.sala]=(sa[r.sala]||0)-val(r));
    const so=Object.entries(sa).sort((a,b)=>Math.abs(b[1])-Math.abs(a[1]));
    if(so.length) out.push({tom:so[0][1]>0?'alta':'queda',
      titulo:'Sala que mais variou: '+ln(so[0][0]),
      corpo:(so[0][1]>0?'+':'−')+unit()+fmt(Math.abs(so[0][1]),S.met)+
        ' contra '+p.pre+'. Nenhuma outra sala mudou tanto.'});
  }
  return out;
}

function rdCusto(rs,tot){
  const out=[], cx=contexto(), pax=sumP(rs), tU=sumU(rs);
  if(pax) out.push({tom:'info',titulo:'Custo médio de US$ '+n2.format(tU/pax)+' por acesso',
    corpo:'<strong>'+n0.format(pax)+'</strong> acessos e US$ '+n2.format(tU)+
      ' em '+cx.alvo+'. Contratos em euro entram já convertidos pelo câmbio do mês.'});
  const ag={};
  rs.filter(r=>r.pax>0).forEach(r=>{
    ag[r.parc]=ag[r.parc]||[0,0]; ag[r.parc][0]+=r['usd_'+S.fx]; ag[r.parc][1]+=r.pax;});
  const caro=Object.entries(ag).map(e=>[e[0],e[1][0]/e[1][1],e[1][1]])
    .sort((a,b)=>b[1]-a[1]);
  if(caro.length>1){
    const c0=caro[0], c1=caro[caro.length-1];
    out.push({tom:'atencao',titulo:'Acesso mais caro: '+pn(c0[0])+' a US$ '+n2.format(c0[1]),
      corpo:'Contra US$ '+n2.format(c1[1])+' da '+pn(c1[0])+', diferença de '+
        ((c0[1]/c1[1]-1)*100).toFixed(0)+'%. São '+n0.format(c0[2])+
        ' acessos nesse preço.'});
  }
  META.parcs.filter(x=>(META.contratos[x]||{}).moeda==='EUR'&&S.parc.has(x)).forEach(x=>{
    const sel=rs.filter(r=>r.parc===x&&r.pax>0), px=sumP(sel);
    if(!px) return;
    const real=sumU(sel)/px, fee=META.contratos[x].fee;
    if(fee) out.push({tom:'info',titulo:pn(x)+': contrato em euro',
      corpo:'Fee de € '+n0.format(fee)+' fechou em <strong>US$ '+n2.format(real)+
        '</strong> por acesso. A diferença de '+pctS((real-fee)/fee*100)+
        ' é câmbio, não reajuste.'});
  });
  const z=META.parcs.filter(x=>(META.contratos[x]||{}).pax_hist===0);
  if(z.length) out.push({tom:'atencao',titulo:z.map(pn).join(' e ')+' sem acesso na base',
    corpo:'Contrato ativo, mas nenhum acesso registrado em toda a série. '+
      'Vale confirmar se ainda vale manter.'});
  return out;
}

function cIns(){
  const box=document.getElementById('c-ins'); box.innerHTML='';
  const rs=filtra(), tot=sum(rs);
  if(!rs.length){box.innerHTML='<p class="note-s">Nenhum registro no recorte.</p>';return;}
  RD[S.rd].f(rs,tot||1).slice(0,6).forEach((o,i)=>{
    const d=document.createElement('div');
    d.className='ins '+(o.tom||'info')+' rise'; d.style.setProperty('--i',i);
    d.innerHTML='<div class="ins-title">'+o.titulo+'</div>'+
                '<div class="ins-body">'+o.corpo+'</div>';
    box.appendChild(d);});
}

/* ================= kpis ================= */
function kpis(){
  const rs=filtra(), p=ytd(), tU=sumU(rs), tP=sumP(rs), c=[];
  c.push(['Valor total','<span class="cur">US$</span> '+n0.format(tU),
    p?pct(sumU(p.A),sumU(p.B)):null,
    '']);
  c.push(['PAX Ex-GOL',n0.format(tP),p?pct(sumP(p.A),sumP(p.B)):null,
    '']);
  const cA=p&&sumP(p.A)?sumU(p.A)/sumP(p.A):null, cB=p&&sumP(p.B)?sumU(p.B)/sumP(p.B):null;
  /* deixa explícito que é média, e de quê: do recorte todo ou de uma sala só */
  const umaSala=S.sala.size===1?[...S.sala][0]:null;
  c.push(['Custo médio por acesso','<span class="cur">US$</span> '+n2.format(tP?tU/tP:0),
    (cA&&cB)?pct(cA,cB):null,
    umaSala?ln(umaSala):'']);
  /* o ranking dos dois últimos cartões segue a métrica escolhida no filtro,
     para o botão Valor/PAX ter efeito visível também aqui em cima */
  const ag={}; rs.forEach(r=>ag[r.parc]=(ag[r.parc]||0)+val(r));
  const base=S.met==='pax'?tP:tU;
  const ord=Object.entries(ag).filter(e=>e[1]>0).sort((a,b)=>b[1]-a[1]);
  c.push(['Maior parceira'+(S.met==='pax'?' em PAX':' em valor'),
    ord.length?pn(ord[0][0]):'—',null,
    ord.length?unit()+fmt(ord[0][1],S.met)+' · '+(ord[0][1]/base*100).toFixed(0)+'%':'—']);
  const t3=ord.slice(0,3).reduce((a,e)=>a+e[1],0);
  c.push(['Concentração top 3',base?(t3/base*100).toFixed(0)+'%':'—',null,
    ord.slice(0,3).map(e=>pn(e[0])).join(' · ')||'—']);
  document.getElementById('kpis').innerHTML=c.map((x,i)=>{
    const [l,v,d,n]=x;
    return '<div class="card kpi rise" style="--i:'+i+'"><div class="lbl">'+l+
      '</div><div class="val">'+v+'</div><div class="dlt '+cls(d)+'">'+
      (d!==null&&d!==undefined?pctS(d)+' <em>ano contra ano</em>':'')+
      '</div><div class="note" title="'+n+'">'+n+'</div></div>';}).join('');
}

function csv(){
  const rs=filtra();
  const L=[['Sala VIP','Parceira','Ano','Mes','PAX','Valor de contrato','Moeda',
    'Custo unitario USD','Valor USD','Base de cambio'].join(';')];
  rs.forEach(r=>L.push([ln(r.sala),pn(r.parc),r.ano,mn(r.mes),r.pax,r.fee,r.moeda,
    String(r['unit_'+S.fx].toFixed(4)).replace('.',','),
    String(r['usd_'+S.fx].toFixed(2)).replace('.',','),S.fx].join(';')));
  L.push(['TOTAL','','','',sumP(rs),'','','',
    String(sumU(rs).toFixed(2)).replace('.',','),''].join(';'));
  const nome='salas_vip_'+new Date().toISOString().slice(0,10)+'.csv';
  const b=new Blob(['\ufeff'+L.join('\r\n')],{type:'text/csv;charset=utf-8;'});
  /* IE/Edge legado tem API pr\u00f3pria */
  if(navigator.msSaveBlob){ navigator.msSaveBlob(b,nome); return; }
  const url=URL.createObjectURL(b);
  const a=document.createElement('a');
  a.href=url; a.download=nome; a.rel='noopener'; a.style.display='none';
  /* a \u00e2ncora PRECISA estar no documento: click() em elemento solto n\u00e3o baixa */
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{ document.body.removeChild(a); URL.revokeObjectURL(url); },1500);
}

/* ================= filtros ================= */
function dropdown(host,label,items,set,lab,cores){
  const el=document.getElementById(host); el.innerHTML='';
  const todos=set.size===items.length;
  const txt=todos?'Todos':set.size===1?lab([...set][0]):set.size+' selecionados';
  const b=document.createElement('button');
  b.innerHTML='<span class="lb">'+label+'</span><span class="vl">'+txt+
    '</span><span class="cv">▾</span>';
  el.appendChild(b);
  const pan=document.createElement('div'); pan.className='pan';
  items.forEach(it=>{
    const o=document.createElement('div');
    o.className='opt'+(set.has(it)?' on':'');
    o.innerHTML='<span class="bx">'+(set.has(it)?'✓':'')+'</span>'+
      (cores?'<span class="sw" style="background:'+cv(cores(it))+'"></span>':'')+
      '<span>'+lab(it)+'</span>';
    o.onclick=e=>{e.stopPropagation();
      set.has(it)?set.delete(it):set.add(it);
      paint(); render();};
    pan.appendChild(o);});
  const ft=document.createElement('div'); ft.className='panfoot';
  const ba=document.createElement('button'); ba.textContent='Todos';
  ba.onclick=e=>{e.stopPropagation();set.clear();items.forEach(x=>set.add(x));paint();render();};
  const bn=document.createElement('button'); bn.textContent='Limpar';
  bn.onclick=e=>{e.stopPropagation();set.clear();paint();render();};
  ft.append(ba,bn); pan.appendChild(ft); el.appendChild(pan);
  b.onclick=e=>{e.stopPropagation();
    const ja=el.classList.contains('open');
    document.querySelectorAll('.dd').forEach(x=>x.classList.remove('open'));
    if(!ja) el.classList.add('open');};
}
/* Período em dois modos: um mês só, ou um intervalo. */
function periodo(){
  const el=document.getElementById('dd-per'); el.innerHTML='';
  const unico=S.modo==='unico';
  const b=document.createElement('button');
  b.innerHTML='<span class="lb">Período</span><span class="vl">'+
    (unico?MPTL[S.m1]:MPT[S.m1]+'–'+MPT[S.m2])+'</span><span class="cv">▾</span>';
  el.appendChild(b);
  const pan=document.createElement('div'); pan.className='pan'; pan.style.padding='11px';

  const seg=document.createElement('div');
  seg.className='seg'; seg.style.marginBottom='10px';
  [['unico','Mês único'],['intervalo','Entre meses']].forEach(o=>{
    const t=document.createElement('button');
    t.className=S.modo===o[0]?'on':''; t.textContent=o[1];
    t.onclick=e=>{e.stopPropagation();
      S.modo=o[0];
      if(o[0]==='unico') S.m1=S.m2;            /* colapsa no mês final */
      paint(); render();};
    seg.appendChild(t);});
  pan.appendChild(seg);

  const validos=mesesValidos();
  function selMes(rot, valor, onSet){
    const w=document.createElement('div'); w.style.marginBottom='8px';
    w.innerHTML='<div style="font-size:9.5px;font-weight:700;letter-spacing:.09em;'+
      'text-transform:uppercase;color:var(--text-muted);margin-bottom:3px">'+rot+'</div>';
    const sel=document.createElement('select'); sel.style.width='100%';
    MPTL.forEach((m,i)=>{
      const o=document.createElement('option');
      o.value=i; o.textContent=m;
      /* mês sem dado no ano selecionado fica visível mas desabilitado */
      if(validos.length && validos.indexOf(i)<0){o.disabled=true; o.textContent=m;}
      sel.appendChild(o);});
    sel.value=valor;
    sel.onclick=e=>e.stopPropagation();
    sel.onchange=e=>{e.stopPropagation(); onSet(+sel.value); paint(); render();};
    w.appendChild(sel); pan.appendChild(w);
  }
  if(unico){
    selMes('Mês', S.m1, v=>{S.m1=v; S.m2=v;});
  }else{
    selMes('De',  S.m1, v=>{S.m1=v; if(S.m2<S.m1) S.m2=S.m1;});
    selMes('Até', S.m2, v=>{S.m2=v; if(S.m2<S.m1) S.m1=S.m2;});
  }

  const ir=document.createElement('button');
  ir.textContent='Ir para o mês mais recente';
  ir.style.cssText='width:100%;font-size:11px;margin-top:2px';
  ir.onclick=e=>{e.stopPropagation();
    S.ano=new Set([ULT_ANO]); S.modo='unico'; S.m1=ULT_MES; S.m2=ULT_MES;
    paint(); render();};
  pan.appendChild(ir);

  el.appendChild(pan);
  b.onclick=e=>{e.stopPropagation();
    const ja=el.classList.contains('open');
    document.querySelectorAll('.dd').forEach(x=>x.classList.remove('open'));
    if(!ja) el.classList.add('open');};
}
function segmento(host,opts,atual,cb){
  const el=document.getElementById(host); el.innerHTML='';
  opts.forEach(o=>{const b=document.createElement('button');
    b.className=atual===o[0]?'on':''; b.textContent=o[1]; b.title=o[2]||'';
    b.onclick=()=>{cb(o[0]);paint();render();}; el.appendChild(b);});
}
function paint(){
  dropdown('dd-ano','Ano',META.anos,S.ano,String);
  dropdown('dd-sala','Sala VIP',META.salas,S.sala,ln);
  dropdown('dd-parc','Parceira',META.parcs,S.parc,pn,p=>PC[p]);
  periodo();
  segmento('seg-met',[['usd','Valor'],['pax','PAX']],S.met,v=>S.met=v);
  document.getElementById('tgl-tab').className='tgl'+(S.tables?' on':'');
}

/* Cada bloco roda isolado. Se um gráfico falhar, os demais continuam e o erro
   aparece no próprio cartão — antes, uma exceção no meio do render() parava
   todos os blocos seguintes e dava a impressão de que o filtro não funcionava. */
const BLOCOS=[['kpis',kpis],['c-ins',cIns],['c-line',cLine],['c-stack',cStack],
  ['c-ytd',cYTD],['c-rank',cRank],['c-contrib',cContrib],
  ['c-split',cSplit],['c-idx',cIndices],['c-detparc',cDetalheParceira],['c-l3m',cL3M],
  ['c-mov',cMov]];
/* Filtro sem resultado: em vez de desenhar tudo zerado, esconde as seções e
   explica. Os filtros continuam visíveis para a pessoa desfazer o recorte. */
function semDados(){
  const rs=filtra();
  const vazio=rs.length===0 || (sumU(rs)===0 && sumP(rs)===0);
  document.body.classList.toggle('sem-dados',vazio);
  if(vazio){
    const e=escopo();
    const semFiltro=!S.ano.size||!S.sala.size||!S.parc.size;
    document.getElementById('vazio-sub').textContent=
      semFiltro ? 'Selecione ano, sala VIP e parceira no painel à esquerda.'
      : e ? 'Recorte atual: '+e+'.'
          : '';
  }
  return vazio;
}
/* Índices só fazem sentido com mais de um mês: "mês de maior custo" num mês
   só é o próprio mês. */
function mostraIndices(){
  const n=S.m2-S.m1+1, mostra=S.modo!=='unico'&&n>1;
  const sec=document.getElementById('c-idx');
  const bloco=sec?sec.closest('section'):null;
  if(bloco) bloco.style.display=mostra?'':'none';
  return mostra;
}
function render(){
  const vazio=semDados();
  mostraIndices();
  /* botão Limpar só aparece quando existe algo filtrado */
  const btn=document.getElementById('clearbtn');
  if(btn) btn.style.display=(escopo()||vazio)?'':'none';
  if(vazio){ paintFiltrosOnly(); return; }
  BLOCOS.forEach(function(par){
    if(par[0]==='c-idx'&&!mostraIndices()) return;
    try{ par[1](); }
    catch(e){
      console.error('falha ao desenhar '+par[0],e);
      const el=document.getElementById(par[0]);
      if(el) el.innerHTML='<p class="note-s" style="color:var(--dn)">'+
        'Não foi possível desenhar este bloco: '+e.message+'</p>';
    }});
  document.querySelectorAll('.tv').forEach(e=>e.classList.toggle('show',S.tables));
}
function paintFiltrosOnly(){}

/* Uma parceira filtrada sozinha: abre o detalhamento completo dela, linha a
   linha, com fórmula e total — serve de prova de que o número fecha. */
function cDetalheParceira(){
  const host=document.getElementById('c-detparc');
  const sec=host?host.closest('section'):null;
  if(!host) return;
  if(S.parc.size!==1){ if(sec) sec.style.display='none'; return; }
  if(sec) sec.style.display='';
  const p=[...S.parc][0], rs=filtra().filter(r=>r.parc===p);
  document.getElementById('t-detparc').textContent=pn(p);
  if(!rs.length){host.innerHTML='<p class="note-s">Sem lançamentos no recorte.</p>';return;}
  const ct=META.contratos[p]||{};
  const sim=ct.moeda==='EUR'?'€ ':'US$ ';
  const tU=sumU(rs), tP=sumP(rs);
  let h='<div class="tw" style="max-height:420px"><table><thead><tr>'+
    '<th>Sala VIP</th><th>Ano</th><th>Mês</th><th class="num">PAX</th>'+
    '<th class="num">Contrato</th><th>Moeda</th><th class="num">Câmbio</th>'+
    '<th class="num">Custo unit. US$</th><th class="num">Valor US$</th>'+
    '</tr></thead><tbody>';
  rs.slice().sort((a,b)=>a.ano-b.ano||mi(a.mes)-mi(b.mes)||
      a.sala.localeCompare(b.sala)).forEach(r=>{
    const u=r['unit_'+S.fx];
    const cam=r.moeda==='EUR'&&r.fee?(u/r.fee):1;
    h+='<tr><td>'+ln(r.sala)+'</td><td>'+r.ano+'</td><td>'+mn(r.mes)+
      '</td><td class="num">'+n0.format(r.pax)+'</td><td class="num">'+
      (r.fee?sim+n0.format(r.fee):'—')+'</td><td>'+r.moeda+
      '</td><td class="num">'+(r.moeda==='EUR'?n2.format(cam).replace(',','.'):'—')+
      '</td><td class="num">'+n2.format(u)+'</td><td class="num">'+
      n2.format(r['usd_'+S.fx])+'</td></tr>';});
  h+='</tbody><tfoot><tr><td colspan="3">Total do recorte</td><td class="num">'+
    n0.format(tP)+'</td><td colspan="4"></td><td class="num">'+n2.format(tU)+
    '</td></tr></tfoot></table></div>';
  h+='<p class="note-s">Contrato de '+sim+n0.format(ct.fee||0)+' por acesso'+
    (ct.moeda==='EUR'?'. O custo unitário em dólar é o contrato multiplicado pela '+
      'média de câmbio do mês, por isso varia mês a mês.'
     :'. Contrato em dólar, então o custo unitário não varia com câmbio.')+
    ' Custo médio no recorte: <strong>US$ '+n2.format(tP?tU/tP:0)+
    '</strong> por acesso, sobre '+n0.format(tP)+' acessos e '+
    n0.format(rs.length)+' lançamentos.</p>';
  host.innerHTML=h;
}
/* ---------- impressão / PDF ----------
   Os SVG são desenhados na largura real do container. Ao imprimir, o layout
   muda (menos colunas, página A4) e os gráficos ficariam deformados se não
   fossem redesenhados. Por isso o redesenho acontece quando as regras de
   impressão passam a valer (matchMedia('print')), e não no clique. */
let _antesTab=null, _imprimindo=false;
function preparaImpressao(){
  if(_imprimindo) return; _imprimindo=true;
  document.querySelectorAll('.dd').forEach(x=>x.classList.remove('open'));
  _antesTab=S.tables; S.tables=true;
  render();
}
function restauraImpressao(){
  if(!_imprimindo) return; _imprimindo=false;
  if(_antesTab!==null) S.tables=_antesTab;
  _antesTab=null; paint(); render();
}
addEventListener('beforeprint',preparaImpressao);
addEventListener('afterprint',restauraImpressao);
if(window.matchMedia){
  const mq=matchMedia('print');
  const onMq=e=>{ e.matches?preparaImpressao():restauraImpressao();
    if(e.matches) requestAnimationFrame(render); };  /* já na largura da página */
  mq.addEventListener?mq.addEventListener('change',onMq):mq.addListener(onMq);
}
document.getElementById('rd-next').onclick=e=>{
  S.rd=(S.rd+1)%RD.length;
  const b=e.currentTarget; b.classList.add('spin');
  setTimeout(()=>b.classList.remove('spin'),520);
  cIns();};
document.getElementById('tgl-tab').onclick=()=>{S.tables=!S.tables;paint();
  document.querySelectorAll('.tv').forEach(e=>e.classList.toggle('show',S.tables));};
/* Reset devolve o estado de abertura: mês mais recente, não o ano inteiro. */
/* Limpar zera mesmo: sem ano, sala ou parceira marcados o relatorio cai no
   estado vazio, que e a resposta honesta para "nenhum filtro selecionado". */
document.getElementById('clearbtn').onclick=()=>{
  S.ano=new Set();S.sala=new Set();S.parc=new Set();
  S.met='usd';S.fx='pub';S.hide.clear();S.rot=true;
  paint();render();};
function rotuloTema(){
  const c=document.documentElement.getAttribute('data-theme');
  const dark=c?c==='dark':matchMedia('(prefers-color-scheme:dark)').matches;
  document.getElementById('themebtn').textContent=dark?'Light':'Dark';
}
document.getElementById('themebtn').onclick=()=>{
  const c=document.documentElement.getAttribute('data-theme');
  const dark=c?c==='dark':matchMedia('(prefers-color-scheme:dark)').matches;
  document.documentElement.setAttribute('data-theme',dark?'light':'dark');
  rotuloTema(); requestAnimationFrame(render);};
rotuloTema();
document.addEventListener('click',()=>document.querySelectorAll('.dd')
  .forEach(x=>x.classList.remove('open')));
addEventListener('resize',()=>{clearTimeout(window._rz);window._rz=setTimeout(render,200);});
document.getElementById('fgen').textContent='Gerado em '+META.gerado;
document.getElementById('ffx').textContent='Câmbio: '+(META.fonte_fx||'—');

paint(); render();
requestAnimationFrame(()=>requestAnimationFrame(render));
</script></body></html>
"""

if __name__ == "__main__":
    main()





























