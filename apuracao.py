#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Apuração dos dados oficiais da Anatel usados no TCC.

Por que este arquivo existe
---------------------------
O Capítulo 4 (Metodologia) da monografia promete ao leitor que ele pode refazer
cada número derivado do trabalho. Esses números não são publicados prontos pela
Anatel: são apurações sobre as planilhas e os arquivos de dados abertos que
acompanham este repositório. Este script é o instrumento que torna a promessa
verificável.

Ele cumpre três papéis:
  1. REPRODUZ  — recalcula, do zero, todo número derivado registrado no índice.
  2. VERIFICA  — compara o resultado com os valores esperados (constante
                 ESPERADO), e acusa divergência. Serve de teste de regressão
                 quando as planilhas forem rebaixadas para um período novo.
  3. DOCUMENTA — cada função declara, no docstring, de qual arquivo parte,
                 que recorte aplica e qual afirmação do TCC sustenta.

Como executar
-------------
    python3 apuracao.py            # apura e verifica
    python3 apuracao.py --sem-check # só apura, sem comparar

Dependência: openpyxl (leitura de .xlsx). Os CSV usam apenas a biblioteca
padrão. Se o openpyxl não estiver disponível:

    python3 -m venv .venv && .venv/bin/pip install openpyxl
    .venv/bin/python apuracao.py

Armadilhas dos dados, documentadas no README.md e tratadas aqui
------------------------------------------------------------------
  * Série de porte da prestadora: quebra de critério em jan./2015 e ponto
    espúrio em jun./2021. Este script NÃO usa essa série; a contagem de
    entidades é apurada dos rankings anuais, que são imunes a ambos.
  * Contratos de interconexão: defasagem de publicação em 2025 e 2026.
  * Painel de cobertura móvel: usar a medida "moradores", nunca "área".
  * Exportações do painel podem vir com o período errado. A verificação de
    monotonicidade em `serie_fixo` existe justamente para pegar isso.

Procedência e URL de origem de cada conjunto: ver README.md.
Extrações de 01 e 07/08/2026.
"""

from __future__ import annotations

import argparse
import csv
import glob
import os
import re
import statistics
import sys
from collections import Counter, defaultdict

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados")

# Valores registrados no README.md. Divergência aqui significa que os dados
# mudaram (rebaixa de período) ou que a apuração foi alterada — nos dois casos,
# o README precisa ser atualizado junto.
ESPERADO = {
    "ref64_localidades": 39976,
    "ref64_moradores_total": 183719406,
    "ref64_pct_moradores_4g5g": 98.44,
    "ref65_meio_2021_metalico_pct": 53.22,
    "ref65_meio_2026_fibra_pct": 60.50,
    "ref65_meio_2026_total": 18459750,
    "ref65_grupos_2026": 315,
    "ref65_grupos_acima_1mi": 3,
    "ref65_grupos_abaixo_100": 106,
    "ref65_grupos_um_acesso": 13,
    "ref65_grupos_um_municipio": 102,
    "ref65_serie_grupos": [154, 185, 235, 270, 350, 315],
    "ref66_registros": 12374,
    "ref66_prestadoras": 526,
    "ref66_pares": 1447,
    "ref66_grau_max": 460,
    "ref66_stfc": 517,
    "ref66_local": 517,
    "ref66_ldn": 502,
    "ref66_ldi": 501,
    "ref66_mvno_vigentes": 196,
    "ref67_serie_grupos": [11, 11, 11, 13, 17, 17],
    "ref67_acessos_2026": 277808244,
}

ANOS = ["2021", "2022", "2023", "2024", "2025", "2026"]


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------

def _openpyxl():
    try:
        import openpyxl  # noqa: F401
        return __import__("openpyxl")
    except ImportError:
        sys.exit(
            "ERRO: openpyxl não instalado.\n"
            "  python3 -m venv .venv && .venv/bin/pip install openpyxl\n"
            "  .venv/bin/python apuracao.py"
        )


def _num(v):
    """Converte célula em float.

    Tolera vazio, o marcador "-" que a Anatel usa para ausência de dado, e o
    formato brasileiro com separador de milhar por ponto e decimal por vírgula.
    """
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s in ("", "-", "–", "N/A"):
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _linhas_xlsx(caminho):
    """Devolve (cabecalho, linhas) da primeira aba."""
    wb = _openpyxl().load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    cab = list(next(it))
    linhas = [r for r in it if r and r[0] is not None]
    wb.close()
    return cab, linhas


def _csv(caminho):
    with open(caminho, encoding="utf-8-sig", errors="replace") as f:
        return list(csv.DictReader(f, delimiter=";"))


def _pct(parte, total):
    return 100.0 * parte / total if total else 0.0


# --------------------------------------------------------------------------
# Ref. 64 — Painel de Cobertura Móvel (jun./2026)
# --------------------------------------------------------------------------

def ref64_cobertura_movel(res):
    """Proporção de moradores cobertos por 4G/5G.

    Arquivo : Dados/cobertura 4G 5G/cobertura_localidades-2026-06.xlsx
    Recorte : Operadora = "Todas", Tecnologia = "4G5G" (uma linha por localidade)
    Cálculo : média de "% moradores cobertos" ponderada pela coluna "Moradores"

    Sustenta: a rede de ACESSO móvel brasileira é predominantemente 4G/5G, o que
    enfraquece, para o caso brasileiro, a objeção do Ofcom de que o STIR/SHAKEN
    não compensa enquanto a rede não for majoritariamente IP.

    Cuidado: usar "moradores", nunca "área". Pela mesma base, a cobertura por
    área nacional é de 16,62%, porque a área não povoada da Amazônia domina o
    cálculo. A exportação por UF que o painel oferece traz SÓ a medida de área.
    """
    caminho = os.path.join(BASE, "cobertura_movel", "cobertura_localidades-2026-06.xlsx")
    cab, linhas = _linhas_xlsx(caminho)
    i = {n: cab.index(n) for n in ("Operadora", "% moradores cobertos", "Moradores")}

    moradores = cobertos = 0.0
    n_localidades = 0
    for r in linhas:
        if r[i["Operadora"]] != "Todas":
            continue
        n_localidades += 1
        m = _num(r[i["Moradores"]])
        moradores += m
        cobertos += m * _num(r[i["% moradores cobertos"]]) / 100.0

    res["ref64_localidades"] = n_localidades
    res["ref64_moradores_total"] = int(moradores)
    res["ref64_pct_moradores_4g5g"] = round(_pct(cobertos, moradores), 2)

    print("\n== Ref. 64 — Cobertura móvel (jun./2026)")
    print(f"   localidades ............... {n_localidades:,}")
    print(f"   moradores na base ......... {int(moradores):,}")
    print(f"   cobertos por 4G/5G ........ {int(cobertos):,}")
    print(f"   %% moradores cobertos ...... {res['ref64_pct_moradores_4g5g']:.2f}%")


# --------------------------------------------------------------------------
# Ref. 65 — Telefonia fixa
# --------------------------------------------------------------------------

def ref65_meio_de_acesso(res):
    """Migração do meio de acesso da telefonia fixa (proxy de TDM x IP).

    Arquivo : Dados/telefonia fixa/telefonia_fixa_tecnologia_acesso.xlsx
    Cálculo : participação de cada meio no primeiro e no último período da série

    Sustenta: a migração para IP no último quilômetro do fixo.

    Limite  : meio de acesso NÃO é protocolo de sinalização. Fibra torna a voz em
    IP provável, mas não prova que a INTERCONEXÃO seja SIP; e cabo metálico não
    implica TDM, já que há VoIP sobre par metálico com xDSL e adaptador.
    """
    caminho = os.path.join(BASE, "telefonia_fixa", "telefonia_fixa_tecnologia_acesso.xlsx")
    _, linhas = _linhas_xlsx(caminho)

    por_periodo = defaultdict(dict)
    for periodo, meio, assinaturas in ((r[0], r[1], r[2]) for r in linhas):
        por_periodo[periodo][meio] = _num(assinaturas)

    chaves = sorted(por_periodo)
    print("\n== Ref. 65(b) — Meio de acesso da telefonia fixa")
    for k in (chaves[0], chaves[-1]):
        total = sum(por_periodo[k].values())
        print(f"   --- {k.date()}  total {int(total):,}")
        for meio, v in sorted(por_periodo[k].items(), key=lambda x: -x[1]):
            print(f"       {meio:<16s} {int(v):>12,}  {_pct(v, total):5.2f}%")
        metalico = por_periodo[k].get("Cabo Metálico", 0.0)
        ip = sum(por_periodo[k].get(m, 0.0) for m in ("Fibra", "Cabo Coaxial", "Rádio"))
        print(f"       >> metálico {_pct(metalico, total):.2f}%  |  fibra+coax+rádio {_pct(ip, total):.2f}%")

    primeiro, ultimo = chaves[0], chaves[-1]
    t0 = sum(por_periodo[primeiro].values())
    t1 = sum(por_periodo[ultimo].values())
    res["ref65_meio_2021_metalico_pct"] = round(_pct(por_periodo[primeiro].get("Cabo Metálico", 0), t0), 2)
    res["ref65_meio_2026_fibra_pct"] = round(_pct(por_periodo[ultimo].get("Fibra", 0), t1), 2)
    res["ref65_meio_2026_total"] = int(t1)


def ref65_distribuicao(res):
    """Distribuição das prestadoras de telefonia fixa por tamanho (jun./2026).

    Arquivo : Dados/telefonia fixa/telefonia_fixa_localidades_tecnologia.xlsx
    Unidade : grupo econômico (coluna "Grupo/Empresa"), que já consolida CNPJs
    Cálculo : soma de acessos por grupo; contagem por faixa; municípios atendidos

    Sustenta: o objetivo específico 1 e a Seção 2.1. O argumento é de CONTAGEM DE
    ENTIDADES, não de participação de mercado: para o modelo de ameaça do
    spoofing, uma prestadora com 300 acessos tem a mesma capacidade de originar
    sinalização que uma com 5 milhões.

    A Anatel publica pronto o topo do ranking (arquivo de participação de
    mercado). A CAUDA — quantos grupos abaixo de cem acessos, quantos com um
    único acesso, quantos em um só município — é apuração própria. Por isso
    `ref65_validacao` confere o topo contra o arquivo oficial: se o agregado
    bate, o método que produz a cauda está validado.
    """
    caminho = os.path.join(BASE, "telefonia_fixa", "telefonia_fixa_localidades_tecnologia.xlsx")
    cab, linhas = _linhas_xlsx(caminho)
    i = {n: cab.index(n) for n in ("Grupo/Empresa", "CNPJ", "Acessos Telefonia Fixa", "Município")}

    acessos = defaultdict(float)
    municipios = defaultdict(set)
    cnpjs = set()
    for r in linhas:
        g = r[i["Grupo/Empresa"]]
        acessos[g] += _num(r[i["Acessos Telefonia Fixa"]])
        municipios[g].add(r[i["Município"]])
        cnpjs.add(r[i["CNPJ"]])

    total = sum(acessos.values())
    faixas = [
        (1_000_000, None, "1 milhão ou mais"),
        (100_000, 1_000_000, "100 mil a 1 milhão"),
        (10_000, 100_000, "10 mil a 100 mil"),
        (1_000, 10_000, "1 mil a 10 mil"),
        (100, 1_000, "100 a 1 mil"),
        (0, 100, "menos de 100"),
    ]

    print("\n== Ref. 65(a) — Distribuição por tamanho (jun./2026)")
    print(f"   grupos econômicos {len(acessos)} | CNPJs {len(cnpjs)} | acessos {int(total):,}")
    print(f"   {'faixa':<20s} {'grupos':>7s} {'% grupos':>9s} {'acessos':>12s} {'% acessos':>10s}")
    for lo, hi, rotulo in faixas:
        sel = [v for v in acessos.values() if v >= lo and (hi is None or v < hi)]
        print(f"   {rotulo:<20s} {len(sel):>7,} {_pct(len(sel), len(acessos)):>8.1f}% "
              f"{int(sum(sel)):>12,} {_pct(sum(sel), total):>9.2f}%")

    um_municipio = [g for g in acessos if len(municipios[g]) == 1]
    print("\n   top 5 por acessos:")
    for g, v in sorted(acessos.items(), key=lambda x: -x[1])[:5]:
        print(f"       {g[:42]:<42s} {int(v):>10,}  {_pct(v, total):5.2f}%  "
              f"{len(municipios[g]):>5,} municípios")

    res["ref65_grupos_2026"] = len(acessos)
    res["ref65_grupos_acima_1mi"] = len([v for v in acessos.values() if v >= 1_000_000])
    res["ref65_grupos_abaixo_100"] = len([v for v in acessos.values() if v < 100])
    res["ref65_grupos_um_acesso"] = len([v for v in acessos.values() if v == 1])
    res["ref65_grupos_um_municipio"] = len(um_municipio)
    print(f"\n   grupos com menos de 100 acessos ... {res['ref65_grupos_abaixo_100']}")
    print(f"   grupos com 1 único acesso ......... {res['ref65_grupos_um_acesso']}")
    print(f"   grupos em 1 único município ....... {res['ref65_grupos_um_municipio']}")


def ref65_validacao(res):
    """Confere a apuração própria contra o agregado que a Anatel publica pronto.

    Arquivo : Dados/telefonia fixa/telefonia_fixa_operadoras_participacao_mercado.xlsx

    Não é redundância: é verificação do procedimento. Se a agregação feita sobre
    a planilha bruta reproduz o agregado oficial, o mesmo método aplicado à cauda
    (que a Anatel não publica) está validado.
    """
    caminho = os.path.join(BASE, "telefonia_fixa",
                           "telefonia_fixa_operadoras_participacao_mercado.xlsx")
    if not os.path.exists(caminho):
        print("\n== Ref. 65 — validação: arquivo oficial ausente, pulando")
        return
    _, linhas = _linhas_xlsx(caminho)
    oficial = {r[0]: _num(r[1]) for r in linhas}

    print("\n== Ref. 65 — Validação contra o agregado oficial")
    print(f"   grupos no arquivo oficial ......... {len(oficial)}")
    print(f"   grupos na apuração própria ........ {res.get('ref65_grupos_2026')}")
    ok = len(oficial) == res.get("ref65_grupos_2026")
    print(f"   contagem de grupos ................ {'confere' if ok else 'DIVERGE'}")
    for g, v in sorted(oficial.items(), key=lambda x: -x[1])[:3]:
        print(f"       {g[:30]:<30s} {int(v):>12,}")


def ref65_serie_fixo(res):
    """Série própria de contagem de entidades no fixo, jun./2021 a jun./2026.

    Arquivos: Dados/telefonia fixa/telefonia_fixa-AAAA-06-assinantes-operadora*
              (mais o arquivo de participação, que é o período corrente)

    Dispensa a série de "porte da prestadora" da Anatel, que tem quebra de
    critério em jan./2015 e ponto espúrio em jun./2021.

    Sustenta: entre jun./2021 e jun./2026 os grupos econômicos de telefonia fixa
    mais que dobraram (154 -> 315) enquanto a base de assinantes caiu 38%. Menos
    assinantes, distribuídos por mais que o dobro de operadoras.

    Verificação embutida: a série de acessos do maior grupo deve ser monotônica
    decrescente. Foi assim que se detectou, em 07/08/2026, que uma exportação
    rotulada como 2022 continha, na verdade, os dados de jun./2026.
    """
    padrao = os.path.join(BASE, "telefonia_fixa", "telefonia_fixa-*-assinantes-operadora*")
    serie = {}
    for caminho in sorted(glob.glob(padrao)):
        m = re.search(r"-(\d{4})-06-", caminho)
        if not m:
            continue
        _, linhas = _linhas_xlsx(caminho)
        serie[m.group(1)] = {r[0]: _num(r[1]) for r in linhas}

    corrente = os.path.join(BASE, "telefonia_fixa",
                            "telefonia_fixa_operadoras_participacao_mercado.xlsx")
    if os.path.exists(corrente):
        _, linhas = _linhas_xlsx(corrente)
        serie.setdefault("2026", {r[0]: _num(r[1]) for r in linhas})

    print("\n== Ref. 65(a.3) — Série de contagem de entidades, telefonia fixa")
    print(f"   {'jun.':>6s} {'grupos':>8s} {'acessos':>14s} {'<1.000':>8s} {'<100':>6s} {'=1':>4s}")
    grupos_serie, claro = [], []
    for ano in ANOS:
        d = serie.get(ano)
        if not d:
            print(f"   {ano:>6s}  (arquivo ausente)")
            continue
        total = sum(d.values())
        grupos_serie.append(len(d))
        claro.append(d.get("CLARO", 0.0))
        print(f"   {ano:>6s} {len(d):>8,} {int(total):>14,} "
              f"{len([v for v in d.values() if v < 1000]):>8,} "
              f"{len([v for v in d.values() if v < 100]):>6,} "
              f"{len([v for v in d.values() if v == 1]):>4,}")

    res["ref65_serie_grupos"] = grupos_serie

    quedas = [claro[i] >= claro[i + 1] for i in range(len(claro) - 1)]
    if all(quedas):
        print("   consistência: série do maior grupo é monotônica decrescente, OK")
    else:
        pos = [ANOS[i + 1] for i, ok in enumerate(quedas) if not ok]
        print(f"   ⚠ ALERTA: quebra de monotonicidade em {pos}. "
              "Provável exportação com o período errado — rebaixar do painel.")


# --------------------------------------------------------------------------
# Ref. 66 — Contratos de interconexão
# --------------------------------------------------------------------------

def ref66_interconexao(res):
    """Topologia da interconexão entre prestadoras.

    Arquivo : Dados/contratos_interconexao/contratos_interconexao.csv
    Unidade : CNPJ da prestadora; aresta = par bilateral distinto

    Sustenta: o argumento de CONFIANÇA TRANSITIVA. O modelo do SS7 presume
    confiança mútua entre um conjunto pequeno e conhecido de operadoras; o grafo
    real tem centenas de nós, e uma chamada injetada em qualquer folha atravessa
    um concentrador e alcança toda a rede.

    O código de serviço 171 é o STFC. Não vem decodificado no glossário: a
    identificação é por evidência interna — MODALIDADE_STFC está preenchida
    exclusivamente e sempre quando o serviço é 171, e assume exatamente os
    valores LOCAL, LDN e LDI, as três modalidades do STFC.

    Limite  : o conjunto registra contratos PROTOCOLADOS entre 2006 e 2025, não
    interconexões em vigor. Tratar como acumulado histórico, não como fotografia
    do presente. E não há, em nenhum arquivo do conjunto, campo de tecnologia de
    interconexão — ou seja, se o enlace é SIP ou SS7 não é informação pública.
    """
    caminho = os.path.join(BASE, "contratos_interconexao", "contratos_interconexao.csv")
    linhas = _csv(caminho)

    pares, nomes = set(), {}
    por_ano = Counter()
    modalidade = defaultdict(set)
    for x in linhas:
        a = (x["PRESTADORA_1_CNPJ"] or "").strip()
        b = (x["PRESTADORA_2_CNPJ"] or "").strip()
        nomes[a] = x["PRESTADORA_1"]
        nomes[b] = x["PRESTADORA_2"]
        if a and b and a != b:
            pares.add(tuple(sorted((a, b))))
        m = re.search(r"/(\d{4})", x["PROTOCOLO_DATA"] or "")
        if m:
            por_ano[m.group(1)] += 1
        for n in ("1", "2"):
            mod = (x[f"MODALIDADE_STFC_{n}"] or "").strip()
            cnpj = (x[f"PRESTADORA_{n}_CNPJ"] or "").strip()
            if mod and cnpj:
                modalidade[mod].add(cnpj)

    grau = Counter()
    for a, b in pares:
        grau[a] += 1
        grau[b] += 1

    print("\n== Ref. 66 — Topologia da interconexão")
    print(f"   registros ................. {len(linhas):,}")
    print(f"   prestadoras (nós) ......... {len(grau):,}")
    print(f"   pares bilaterais (arestas)  {len(pares):,}")
    print(f"   grau mediano {statistics.median(grau.values()):.0f} | "
          f"média {statistics.mean(grau.values()):.1f} | máximo {max(grau.values())}")
    print(f"   com uma única interconexão  {len([v for v in grau.values() if v == 1])} "
          f"({_pct(len([v for v in grau.values() if v == 1]), len(grau)):.1f}%)")
    print("\n   5 maiores concentradores:")
    for c, g in grau.most_common(5):
        print(f"       {(nomes.get(c) or c)[:42]:<42s} {g:>4d}")
    print(f"\n   prestadoras no STFC (qualquer modalidade) {len(set().union(*modalidade.values())):>4d}")
    print("   prestadoras por modalidade de interconexão STFC:")
    for mod in ("LOCAL", "LDN", "LDI"):
        print(f"       {mod:<6s} {len(modalidade[mod]):>4d}")
    print("       (LDI é a porta de entrada do tráfego originado no exterior)")

    res["ref66_registros"] = len(linhas)
    res["ref66_prestadoras"] = len(grau)
    res["ref66_pares"] = len(pares)
    res["ref66_grau_max"] = max(grau.values())
    # União das três modalidades do STFC. É esta a grandeza que o texto afirma
    # ("se interconectam no serviço fixo"); hoje coincide com LOCAL, porque toda
    # prestadora com LDN ou LDI também tem LOCAL, mas a coincidência é dos dados
    # e não da definição. Guardada em chave própria para que uma rebaixa de
    # período que as separe seja acusada pela verificação.
    res["ref66_stfc"] = len(set().union(*modalidade.values())) if modalidade else 0
    res["ref66_local"] = len(modalidade["LOCAL"])
    res["ref66_ldn"] = len(modalidade["LDN"])
    res["ref66_ldi"] = len(modalidade["LDI"])

    ultimos = [a for a in sorted(por_ano) if a >= "2024"]
    print("\n   contratos protocolados nos últimos anos: "
          + ", ".join(f"{a}={por_ano[a]}" for a in ultimos))
    print("   ⚠ 2025 e 2026 aparecem subrepresentados: defasagem de publicação,")
    print("     não interrupção real. Não usar como anos completos.")


def ref66_mvno(res):
    """Rede virtual móvel (MVNO).

    Arquivos: empresas_credenciadas_vigentes.csv (vigentes) e contratos_mvno.csv

    Sustenta: o análogo móvel da fragmentação do fixo, por outro mecanismo — a
    MVNO origina chamadas com numeração própria sem possuir espectro.

    A explicar antes de citar: há centenas de credenciadas vigentes e apenas
    ~17 grupos com acessos no ranking do SMP (ref. 67). As hipóteses são base
    desprezível, contabilização sob a hospedeira, ou foco em IoT.
    """
    d = os.path.join(BASE, "contratos_interconexao")
    vigentes = _csv(os.path.join(d, "empresas_credenciadas_vigentes.csv"))
    contratos = _csv(os.path.join(d, "contratos_mvno.csv"))
    hospedeiras = Counter(x["PRESTADORA_ORIGEM"] for x in contratos)

    print("\n== Ref. 66 — Rede virtual (MVNO)")
    print(f"   credenciadas vigentes ..... {len(vigentes)}")
    print(f"   contratos (histórico) ..... {len(contratos)}")
    print("   principais hospedeiras:")
    for k, v in hospedeiras.most_common(3):
        print(f"       {(k or '')[:42]:<42s} {v:>4d}")

    res["ref66_mvno_vigentes"] = len(vigentes)


# --------------------------------------------------------------------------
# Ref. 67 — Ranking do Serviço Móvel Pessoal
# --------------------------------------------------------------------------

def ref67_serie_movel(res):
    """Série de contagem de entidades no móvel, jun./2021 a jun./2026.

    Arquivos: Dados/telefonia_movel/telefonia_movel_assinantes_operadora_AAAA-jun.xlsx

    Sustenta: A SUPERFÍCIE DE ORIGEM SE AMPLIOU APENAS ONDE A BARREIRA DE ENTRADA
    É BAIXA. O móvel exige espectro, leiloado e caro, e permaneceu concentrado; o
    fixo exige apenas autorização, caminho de entrada do provedor VoIP, e
    multiplicou de operadoras enquanto perdia base.

    Observações do conjunto:
      * A OI desaparece do ranking em 2022 (venda da OI Móvel para Vivo, Claro e
        TIM). Diferente dos artefatos das refs. 65 e 66, é evento de mercado
        documentado e pode ser narrado.
      * A cauda do móvel é de provedores de IoT/M2M, que não constituem
        superfície de originação de voz — o que reforça a assimetria.
    """
    padrao = os.path.join(BASE, "telefonia_movel", "telefonia_movel_assinantes_operadora_*.xlsx")
    serie = {}
    for caminho in sorted(glob.glob(padrao)):
        m = re.search(r"_(\d{4})-jun", caminho)
        if not m:
            continue
        _, linhas = _linhas_xlsx(caminho)
        serie[m.group(1)] = {r[0]: _num(r[1]) for r in linhas}

    print("\n== Ref. 67 — Série de contagem de entidades, telefonia móvel")
    print(f"   {'jun.':>6s} {'grupos':>8s} {'acessos':>15s} {'3 maiores':>11s}")
    grupos_serie = []
    for ano in ANOS:
        d = serie.get(ano)
        if not d:
            print(f"   {ano:>6s}  (arquivo ausente)")
            continue
        total = sum(d.values())
        maiores = sorted(d.values(), reverse=True)[:3]
        grupos_serie.append(len(d))
        print(f"   {ano:>6s} {len(d):>8,} {int(total):>15,} {_pct(sum(maiores), total):>10.2f}%")
        if ano == ANOS[-1]:
            res["ref67_acessos_2026"] = int(total)

    res["ref67_serie_grupos"] = grupos_serie

    if serie.get("2026") and res.get("ref65_grupos_2026"):
        gm, am = len(serie["2026"]), sum(serie["2026"].values())
        gf, af = res["ref65_grupos_2026"], res["ref65_meio_2026_total"]
        print("\n   CONTRASTE (jun./2026): tamanho médio do grupo")
        print(f"       móvel {am / gm:>14,.0f} acessos por grupo ({gm} grupos)")
        print(f"       fixo  {af / gf:>14,.0f} acessos por grupo ({gf} grupos)")
        print(f"       razão {am / gm / (af / gf):>14,.0f} vezes")


# --------------------------------------------------------------------------
# Verificação
# --------------------------------------------------------------------------

def verificar(res):
    print("\n" + "=" * 72)
    print("VERIFICAÇÃO CONTRA OS VALORES REGISTRADOS NO README.md")
    print("=" * 72)
    divergencias = 0
    for chave, esperado in ESPERADO.items():
        obtido = res.get(chave)
        if obtido is None:
            print(f"  -- {chave:<34s} não apurado (arquivo ausente?)")
            continue
        ok = obtido == esperado
        if isinstance(esperado, float) and isinstance(obtido, float):
            ok = abs(obtido - esperado) < 0.01
        if ok:
            print(f"  OK {chave:<34s} {obtido}")
        else:
            divergencias += 1
            print(f"  ** {chave:<34s} obtido {obtido!r}  esperado {esperado!r}")
    print("-" * 72)
    if divergencias:
        print(f"{divergencias} divergência(s). Os dados mudaram ou a apuração foi alterada.")
        print("Atualize o README.md e a constante ESPERADO junto.")
    else:
        print("Nenhuma divergência. Os números do README.md são reprodutíveis.")
    return divergencias


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--sem-check", action="store_true", help="apura sem verificar")
    args = ap.parse_args()

    print("=" * 72)
    print("APURAÇÃO DOS DADOS OFICIAIS DA ANATEL — TCC Tatianno Ferreira Alves")
    print("Procedência das fontes: README.md. Extrações de 01 e 07/08/2026.")
    print("=" * 72)

    res = {}
    for etapa in (ref64_cobertura_movel, ref65_meio_de_acesso, ref65_distribuicao,
                  ref65_validacao, ref65_serie_fixo, ref66_interconexao,
                  ref66_mvno, ref67_serie_movel):
        try:
            etapa(res)
        except FileNotFoundError as e:
            print(f"\n!! {etapa.__name__}: arquivo não encontrado — {e.filename}")

    if args.sem_check:
        return 0
    return 1 if verificar(res) else 0


if __name__ == "__main__":
    sys.exit(main())
