# -*- coding: utf-8 -*-
"""
Risco Sacado - Análise mensal de produção (substitui planilha com macros)

Uso:
    python risco_sacado.py                      -> abre a interface (Tkinter)
    python risco_sacado.py download.xlsx        -> gera direto com parâmetros sugeridos
    python risco_sacado.py download.xlsx --ano-atual 2026 --mes-corte 9 --qtd 0

Dependências: pandas, numpy, openpyxl, python-dateutil  (opcional: tkinterdnd2 para arrastar e soltar)
Não depende de Excel instalado, VBA ou fórmulas. O arquivo de entrada nunca é alterado.
"""
from __future__ import annotations

import argparse
import logging
import math
import os
import re
import subprocess
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from functools import lru_cache
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
from dateutil import parser as dateutil_parser
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

log = logging.getLogger("risco_sacado")

MESES_NOME = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
              "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
MESES_ABREV = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_KEY = [m.lower() for m in MESES_ABREV]

COLUNAS = ["grupo_economico", "ano_mes", "officer_comercial", "volume_producao", "spread",
           "prazo", "margem", "percentual_representatividade", "total_origem"]
ESSENCIAIS = ["grupo_economico", "ano_mes", "volume_producao"]
METRICAS = ["volume_producao", "spread", "prazo", "margem", "percentual_representatividade", "total_origem"]

ALIASES = {
    "grupo_economico": ["descricao_nome_grupo_economico", "grupo econômico", "grupo economico", "cliente",
                        "nome cliente", "nome do grupo", "grupo", "sacado", "descricao grupo economico"],
    "ano_mes": ["data_ano_mes", "ano mes", "ano_mês", "competência", "competencia", "período", "periodo",
                "data referência", "data referencia"],
    "officer_comercial": ["descricao_nome_officer_comercial", "officer", "officer comercial", "comercial",
                          "gerente", "responsável comercial", "responsavel comercial"],
    "volume_producao": ["Vol. Produção", "vol produção", "volume produção", "volume de produção", "produção",
                        "producao", "valor produção", "valor producao"],
    "spread": ["spread", "taxa", "taxa spread"],
    "prazo": ["prazo", "prazo médio", "prazo medio"],
    "margem": ["margem", "receita", "resultado"],
    "percentual_representatividade": ["% Reprt. Prod", "% reprt prod", "representatividade", "participação",
                                      "participacao", "participação produção", "participacao producao"],
    "total_origem": ["total", "total geral"],
}


class ErroLeitura(Exception):
    pass


# ----------------------------------------------------------------------------
# Utilidades de texto
# ----------------------------------------------------------------------------
def _sem_invisiveis(v) -> str:
    s = unicodedata.normalize("NFKC", str(v))
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[\t\r\n\v\f]+", " ", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) not in ("Cf", "Cc"))
    return re.sub(r" {2,}", " ", s).strip()


def _sem_acentos(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))


@lru_cache(maxsize=20000)
def _norm_cab(s) -> str:
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return ""
    s = _sem_acentos(_sem_invisiveis(s)).lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def _chave(s) -> str:
    """Chave de agrupamento: só ignora caixa, acentos, espaços duplicados e caracteres invisíveis."""
    return _sem_acentos(_sem_invisiveis(s)).upper()


ALIAS_IDX = {_norm_cab(a): canon for canon, lst in ALIASES.items() for a in lst}
ALIAS_IDX.update({_norm_cab(c): c for c in COLUNAS})


def _vazio(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str) and not _sem_invisiveis(v):
        return True
    return v is pd.NaT


def _mapear_colunas(rotulos) -> dict:
    """Retorna {coluna_canonica: posicao}. Primeiro casamento exato, depois por palavra contida."""
    mapa, usadas = {}, set()
    normas = [_norm_cab(r) for r in rotulos]
    for pos, n in enumerate(normas):
        canon = ALIAS_IDX.get(n)
        if canon and canon not in mapa:
            mapa[canon] = pos
            usadas.add(pos)
    candidatos = sorted(((a, c) for a, c in ALIAS_IDX.items() if len(a) >= 6), key=lambda x: -len(x[0]))
    for canon in COLUNAS:
        if canon in mapa:
            continue
        for alias, c in candidatos:
            if c != canon:
                continue
            for pos, n in enumerate(normas):
                if pos not in usadas and n and f" {alias} " in f" {n} ":
                    mapa[canon] = pos
                    usadas.add(pos)
                    break
            if canon in mapa:
                break
    return mapa


# ----------------------------------------------------------------------------
# Conversões
# ----------------------------------------------------------------------------
_RE_YM = re.compile(r"^(\d{4})[-/.](\d{1,2})(?:[-/.](\d{1,2}))?(?:[ T].*)?$")
_RE_MY = re.compile(r"^(\d{1,2})[-/.](\d{4})$")
_RE_YYYYMM = re.compile(r"^(\d{4})(\d{2})$")
_RE_DMY = re.compile(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(?:[ T].*)?$")
_RE_MESTXT = re.compile(r"^([a-z]{3})[a-z]*(?:\s+de)?[\s\-/.]*(\d{2}|\d{4})$")
_MES_TXT = {"jan": 1, "fev": 2, "feb": 2, "mar": 3, "abr": 4, "apr": 4, "mai": 5, "may": 5, "jun": 6, "jul": 7,
            "ago": 8, "aug": 8, "set": 9, "sep": 9, "out": 10, "oct": 10, "nov": 11, "dez": 12, "dec": 12}


def _mk(ano, mes):
    try:
        ano, mes = int(ano), int(mes)
        if 1900 <= ano <= 2100 and 1 <= mes <= 12:
            return pd.Timestamp(ano, mes, 1)
    except (ValueError, TypeError):
        pass
    return pd.NaT


def _parse_numero_periodo(n: float):
    if n != n:
        return pd.NaT
    if float(n).is_integer() and 190001 <= n <= 210012:
        return _mk(int(n) // 100, int(n) % 100)
    if 20000 <= n <= 80000:  # serial do Excel
        t = pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(n))
        return _mk(t.year, t.month)
    return pd.NaT


def parse_periodo(v):
    """Converte 'AAAA-MM', data do Excel, timestamp, AAAAMM, MM/AAAA, 'set/26'... para o 1º dia do mês."""
    if v is None:
        return pd.NaT
    if isinstance(v, (datetime, date, np.datetime64, pd.Timestamp)):
        t = pd.Timestamp(v)
        return pd.NaT if pd.isna(t) else _mk(t.year, t.month)
    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
        return _parse_numero_periodo(float(v))
    s = _sem_invisiveis(v)
    if not s:
        return pd.NaT
    for rx, ordem in ((_RE_YM, "ym"), (_RE_MY, "my"), (_RE_YYYYMM, "ym")):
        m = rx.match(s)
        if m:
            a, b = m.group(1), m.group(2)
            return _mk(a, b) if ordem == "ym" else _mk(b, a)
    if re.fullmatch(r"\d+(?:[.,]\d+)?", s):
        return _parse_numero_periodo(float(s.replace(",", ".")))
    m = _RE_DMY.match(s)
    if m:
        return _mk(m.group(3), m.group(2))
    m = _RE_MESTXT.match(_sem_acentos(s).lower())
    if m and m.group(1) in _MES_TXT:
        ano = int(m.group(2))
        ano += 2000 if ano < 100 else 0
        return _mk(ano, _MES_TXT[m.group(1)])
    try:
        t = dateutil_parser.parse(s, dayfirst=True)
        return _mk(t.year, t.month)
    except (ValueError, OverflowError, TypeError):
        return pd.NaT


_NULOS = {"", "-", "--", "—", "–", "n/a", "na", "nan", "none", "null", "#n/d", "n.d.", "nd"}


def _parse_num(v):
    """Retorna (valor, invalido). Vazio/'-'/'N/A' -> (NaN, False). Texto inválido -> (NaN, True)."""
    if v is None:
        return np.nan, False
    if isinstance(v, bool):
        return np.nan, True
    if isinstance(v, (int, float, np.integer, np.floating, Decimal)):
        f = float(v)
        if math.isnan(f):
            return np.nan, False
        return (np.nan, True) if math.isinf(f) else (f, False)
    s = _sem_invisiveis(v)
    if s.lower() in _NULOS:
        return np.nan, False
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    pct = s.endswith("%")
    s = s.rstrip("%")
    tem_moeda = bool(re.search(r"R\$", s, re.I))
    s = re.sub(r"R\$", "", s, flags=re.I).replace(" ", "")
    if s.startswith("-"):
        neg, s = True, s[1:]
    elif s.startswith("+"):
        s = s[1:]
    elif s.endswith("-"):
        neg, s = True, s[:-1]
    try:
        if re.fullmatch(r"\d+(?:[.,]\d+)?[eE][+-]?\d+", s):
            f = float(s.replace(",", "."))
        elif not re.fullmatch(r"[\d.,]+", s):
            return np.nan, True
        else:
            if "," in s and "." in s:
                if s.rfind(",") > s.rfind("."):
                    s = s.replace(".", "").replace(",", ".")
                else:
                    s = s.replace(",", "")
            elif "," in s:
                s = s.replace(",", "") if s.count(",") > 1 else s.replace(",", ".")
            elif "." in s:
                partes = s.split(".")
                if len(partes) > 2 or (tem_moeda and len(partes[1]) == 3 and len(partes[0]) <= 3):
                    s = "".join(partes)
            f = float(s)
    except ValueError:
        return np.nan, True
    f = -f if neg else f
    return (f / 100 if pct else f), False


# ----------------------------------------------------------------------------
# Identificação e leitura
# ----------------------------------------------------------------------------
@dataclass
class ResultadoIdentificacao:
    reconhecido: bool
    confianca: float
    aba: Optional[str]
    linha_cabecalho: Optional[int]
    colunas_encontradas: dict
    colunas_ausentes: list
    mensagem: str
    abas: list
    revisao_necessaria: bool


def _ler_abas(caminho: Path, nrows: Optional[int] = None) -> dict:
    ext = caminho.suffix.lower()
    try:
        if ext == ".csv":
            ultimo = None
            for enc in ("utf-8-sig", "latin-1"):
                try:
                    df = pd.read_csv(caminho, header=None, dtype=object, sep=None, engine="python",
                                     encoding=enc, nrows=nrows, skip_blank_lines=False)
                    return {"csv": df}
                except UnicodeDecodeError as e:
                    ultimo = e
            raise ultimo
        with pd.ExcelFile(caminho) as xf:
            return {sn: xf.parse(sn, header=None, dtype=object, nrows=nrows) for sn in xf.sheet_names}
    except PermissionError:
        raise ErroLeitura("Sem permissão para ler o arquivo (verifique se está em uso ou bloqueado).")
    except ErroLeitura:
        raise
    except Exception as e:  # arquivo corrompido, protegido por senha, formato inválido
        raise ErroLeitura(f"Não foi possível abrir o arquivo (protegido por senha, corrompido ou formato "
                          f"inválido): {type(e).__name__}: {e}")


def _achar_cabecalho(raw: pd.DataFrame, limite: int = 40):
    melhor = (-1, None, {})
    for i in range(min(limite, len(raw))):
        mapa = _mapear_colunas(list(raw.iloc[i]))
        score = len(mapa) + (10 if all(c in mapa for c in ESSENCIAIS) else 0)
        if len(mapa) >= 2 and score > melhor[0]:
            melhor = (score, i, mapa)
    return melhor[1], melhor[2]


def _avaliar_aba(raw: pd.DataFrame):
    h, mapa = _achar_cabecalho(raw)
    ess = sum(c in mapa for c in ESSENCIAIS)
    opt_tot = len(COLUNAS) - len(ESSENCIAIS)
    opt = sum(c in mapa for c in COLUNAS if c not in ESSENCIAIS)
    conf = 80 + 20 * opt / opt_tot if ess == len(ESSENCIAIS) else 60 * ess / len(ESSENCIAIS) + 10 * opt / opt_tot
    return conf, h, mapa


def identificar_arquivo_fonte(caminho: Path) -> ResultadoIdentificacao:
    caminho = Path(caminho)
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    if caminho.suffix.lower() not in (".xlsx", ".xls", ".xlsm", ".csv"):
        return ResultadoIdentificacao(False, 0, None, None, {}, list(COLUNAS),
                                      f"Formato não suportado: {caminho.suffix}", [], True)
    abas = _ler_abas(caminho, nrows=60)
    melhor = None
    for nome, raw in abas.items():
        if raw.empty:
            continue
        conf, h, mapa = _avaliar_aba(raw)
        if melhor is None or conf > melhor[0]:
            melhor = (conf, nome, h, mapa)
    if melhor is None:
        return ResultadoIdentificacao(False, 0, None, None, {}, list(COLUNAS),
                                      "O arquivo não contém dados.", list(abas), True)
    conf, nome, h, mapa = melhor
    ausentes = [c for c in COLUNAS if c not in mapa]
    ess_ok = all(c in mapa for c in ESSENCIAIS)
    if ess_ok:
        msg = f"Arquivo reconhecido como base de produção (confiança {conf:.0f}%, aba '{nome}')"
        if ausentes:
            msg += f" — colunas opcionais ausentes: {', '.join(ausentes)}"
    else:
        msg = ("Este arquivo não parece ser a base esperada. Colunas ausentes: "
               + ", ".join(c for c in COLUNAS if c not in mapa))
    return ResultadoIdentificacao(ess_ok and conf >= 70, conf, nome, h, {k: int(v) for k, v in mapa.items()},
                                  ausentes, msg, list(abas), conf < 70)


def carregar_base(caminho: Path, ident: Optional[ResultadoIdentificacao] = None) -> pd.DataFrame:
    """Lê (somente leitura) a aba identificada e devolve o DataFrame bruto com os rótulos originais."""
    caminho = Path(caminho)
    ident = ident or identificar_arquivo_fonte(caminho)
    if not ident.reconhecido:
        raise ErroLeitura(ident.mensagem)
    raw = _ler_abas(caminho)[ident.aba]
    h = ident.linha_cabecalho
    rotulos, vistos = [], {}
    for j, v in enumerate(raw.iloc[h]):
        r = _sem_invisiveis(v) if not _vazio(v) else f"col_{j}"
        vistos[r] = vistos.get(r, 0) + 1
        rotulos.append(r if vistos[r] == 1 else f"{r}.{vistos[r] - 1}")
    df = raw.iloc[h + 1:].copy()
    df.columns = rotulos  # índice preservado: linha do Excel = índice + 1
    st = caminho.stat()
    df.attrs["meta"] = {
        "arquivo": caminho.name, "caminho": str(caminho), "tamanho_bytes": st.st_size,
        "modificado_em": datetime.fromtimestamp(st.st_mtime).strftime("%d/%m/%Y %H:%M:%S"),
        "abas": ", ".join(ident.abas), "aba_selecionada": ident.aba,
        "linha_cabecalho_excel": h + 1, "linhas_iniciais": int(raw.shape[0]), "colunas_iniciais": int(raw.shape[1]),
        "confianca_identificacao": round(ident.confianca, 1),
    }
    log.info("Arquivo %s | aba %s | %s linhas x %s colunas", caminho.name, ident.aba, *raw.shape)
    return df


# ----------------------------------------------------------------------------
# Limpeza
# ----------------------------------------------------------------------------
@dataclass
class RelatorioQualidade:
    meta: dict = field(default_factory=dict)
    linhas_iniciais: int = 0
    linhas_vazias: int = 0
    cabecalhos_repetidos: int = 0
    linhas_totalizacao: int = 0
    linhas_sem_chave: int = 0
    linhas_validas: int = 0
    inconsistencias: list = field(default_factory=list)
    duplicidades: list = field(default_factory=list)
    officer_divergencias: list = field(default_factory=list)
    avisos: list = field(default_factory=list)
    de_para_usado: list = field(default_factory=list)
    totais_origem: pd.DataFrame = field(default_factory=pd.DataFrame)
    conciliacao: list = field(default_factory=list)

    def aviso(self, msg: str):
        log.warning(msg)
        self.avisos.append(msg)

    def resumo(self) -> str:
        l = [f"Linhas lidas: {self.linhas_iniciais}",
             f"Linhas vazias removidas: {self.linhas_vazias}",
             f"Cabeçalhos repetidos removidos: {self.cabecalhos_repetidos}",
             f"Linhas de totalização separadas: {self.linhas_totalizacao}",
             f"Linhas sem grupo/período válido removidas: {self.linhas_sem_chave}",
             f"Linhas válidas: {self.linhas_validas}",
             f"Valores numéricos inválidos: {sum(1 for i in self.inconsistencias if i['tipo'] == 'numero')}",
             f"Cliente/mês com mais de uma linha (agregados): {len(self.duplicidades)}",
             f"Meses com mais de um officer: {len(self.officer_divergencias)}"]
        l += [f"Aviso: {a}" for a in self.avisos[:10]]
        for c in self.conciliacao:
            l.append(f"{c['item']}: {c['status']}")
        return "\n".join(l)


def carregar_de_para(caminho: Path) -> dict:
    caminho = Path(caminho)
    df = pd.read_csv(caminho, dtype=str, sep=None, engine="python") if caminho.suffix.lower() == ".csv" \
        else pd.read_excel(caminho, dtype=str)
    return {str(a): str(b).strip() for a, b in zip(df.iloc[:, 0], df.iloc[:, 1]) if not _vazio(a) and not _vazio(b)}


def normalizar_base(df: pd.DataFrame, de_para: Optional[dict] = None):
    rel = RelatorioQualidade(meta=dict(df.attrs.get("meta", {})))
    mapa = _mapear_colunas(list(df.columns))
    faltam = [c for c in ESSENCIAIS if c not in mapa]
    if faltam:
        raise ErroLeitura(f"Colunas essenciais ausentes: {', '.join(faltam)}")
    d = pd.DataFrame(index=df.index)
    for c in COLUNAS:
        d[c] = df.iloc[:, mapa[c]].values if c in mapa else np.nan
    d["_linha_origem"] = d.index + 1
    rel.linhas_iniciais = len(d)

    # linhas vazias
    vazia = d[COLUNAS].apply(lambda s: s.map(_vazio)).all(axis=1)
    rel.linhas_vazias = int(vazia.sum())
    d = d[~vazia]

    # cabeçalhos repetidos
    def eh_cab(s, canon):
        return s.map(lambda v: isinstance(v, str) and ALIAS_IDX.get(_norm_cab(v)) == canon)
    hits = sum(eh_cab(d[c], c).astype(int) for c in COLUNAS)
    rep = hits >= 2
    rel.cabecalhos_repetidos = int(rep.sum())
    d = d[~rep]

    # numéricos
    for c in METRICAS:
        conv = [_parse_num(v) for v in d[c]]
        d[c + "_orig"] = d[c]
        d[c] = [x[0] for x in conv]
        for (val, inv), linha, orig in zip(conv, d["_linha_origem"], d[c + "_orig"]):
            if inv:
                rel.inconsistencias.append({"tipo": "numero", "linha": int(linha), "campo": c,
                                            "valor": str(orig), "motivo": "Valor não numérico (convertido para vazio)"})
    d = d.drop(columns=[c + "_orig" for c in METRICAS])

    # textos e período
    d["grupo_economico_original"] = d["grupo_economico"].map(lambda v: "" if _vazio(v) else str(v))
    d["grupo_economico"] = d["grupo_economico_original"].map(_sem_invisiveis)
    d["officer_comercial"] = d["officer_comercial"].map(lambda v: "" if _vazio(v) else _sem_invisiveis(v).upper())
    d["periodo"] = d["ano_mes"].map(parse_periodo)

    # totalizações
    tem_num = d[METRICAS].notna().any(axis=1)
    rx_tot = d["grupo_economico"].map(lambda s: bool(re.match(r"^(total|totais|subtotal|soma)\b", _norm_cab(s))))
    eh_total = rx_tot | ((d["grupo_economico"] == "") & d["periodo"].isna() & tem_num)
    rel.linhas_totalizacao = int(eh_total.sum())
    rel.totais_origem = d[eh_total][["_linha_origem", "grupo_economico_original"] + METRICAS].reset_index(drop=True)
    d = d[~eh_total]

    # sem grupo ou período
    sem = (d["grupo_economico"] == "") | d["periodo"].isna()
    for _, r in d[sem].iterrows():
        motivo = "Sem grupo econômico" if r["grupo_economico"] == "" else "Período inválido ou ausente"
        rel.inconsistencias.append({"tipo": "linha", "linha": int(r["_linha_origem"]),
                                    "campo": "grupo_economico/ano_mes",
                                    "valor": f"{r['grupo_economico']} | {r['ano_mes']}", "motivo": motivo + " (linha removida)"})
    rel.linhas_sem_chave = int(sem.sum())
    d = d[~sem].copy()

    # chave + de/para
    d["chave"] = d["grupo_economico"].map(_chave)
    if de_para:
        dp = {_chave(k): _sem_invisiveis(v) for k, v in de_para.items()}
        aplicado = d["chave"].isin(dp)
        for origem, qtd in d[aplicado]["chave"].value_counts().items():
            rel.de_para_usado.append({"origem": origem, "destino": dp[origem], "linhas": int(qtd)})
        d.loc[aplicado, "grupo_economico"] = d.loc[aplicado, "chave"].map(dp)
        d["chave"] = d["grupo_economico"].map(_chave)

    # duplicidades
    tam = d.groupby(["chave", "periodo"]).size()
    for (ch, per), n in tam[tam > 1].items():
        sub = d[(d["chave"] == ch) & (d["periodo"] == per)]
        rel.duplicidades.append({"cliente": sub["grupo_economico"].iloc[0], "periodo": per.strftime("%m/%Y"),
                                 "linhas": int(n), "volume_somado": float(sub["volume_producao"].sum()),
                                 "linhas_origem": ", ".join(map(str, sub["_linha_origem"]))})
    if d["volume_producao"].isna().any():
        rel.aviso(f"{int(d['volume_producao'].isna().sum())} linha(s) sem volume de produção (tratadas como sem dado).")
    rel.linhas_validas = len(d)
    d = d.sort_values(["chave", "periodo", "_linha_origem"], kind="mergesort").reset_index(drop=True)
    return d, rel


def _nome_exibicao(s: pd.Series) -> str:
    vc = s.value_counts()
    top = vc[vc == vc.max()].index
    return sorted(top)[0]


def consolidar_base(d: pd.DataFrame, rel: RelatorioQualidade) -> pd.DataFrame:
    """Granularidade: grupo econômico + ano + mês. Médias ponderadas por volume; % recalculado."""
    x = d.copy()
    for col, pref in (("spread", "s"), ("prazo", "p")):
        ok = x[col].notna() & x["volume_producao"].notna()
        x[f"_{pref}w"] = np.where(ok, x[col] * x["volume_producao"], 0.0)
        x[f"_{pref}v"] = np.where(ok, x["volume_producao"], 0.0)
    g = x.groupby(["chave", "periodo"], sort=True)
    cons = g.agg(volume_producao=("volume_producao", lambda s: s.sum(min_count=1)),
                 margem=("margem", lambda s: s.sum(min_count=1)),
                 _sw=("_sw", "sum"), _sv=("_sv", "sum"), _pw=("_pw", "sum"), _pv=("_pv", "sum"),
                 _n_spread=("spread", "count"), _n_prazo=("prazo", "count"),
                 linhas_origem=("chave", "size")).reset_index()
    for col, pref, n in (("spread", "s", "_n_spread"), ("prazo", "p", "_n_prazo")):
        w, v = cons[f"_{pref}w"], cons[f"_{pref}v"]
        cons[col] = np.where(v != 0, w / v.replace(0, np.nan), np.nan)
        sem_peso = int(((v == 0) & (cons[n] > 0)).sum())
        if sem_peso:
            rel.aviso(f"{col}: {sem_peso} cliente/mês com soma de pesos igual a zero (média ponderada = vazio).")

    nomes = x.groupby("chave")["grupo_economico"].agg(_nome_exibicao)
    cons["cliente"] = cons["chave"].map(nomes)
    cons["ano"] = cons["periodo"].dt.year
    cons["mes_numero"] = cons["periodo"].dt.month
    cons["mes_nome"] = cons["mes_numero"].map(lambda m: MESES_NOME[m - 1])
    cons["ordem_periodo"] = cons["ano"] * 12 + cons["mes_numero"]

    # officer: o de maior volume no mês
    ov = x[x["officer_comercial"] != ""].groupby(["chave", "periodo", "officer_comercial"], as_index=False) \
        .agg(vol=("volume_producao", "sum"))
    ov = ov.sort_values(["chave", "periodo", "vol", "officer_comercial"], ascending=[True, True, False, True],
                        kind="mergesort")
    top = ov.drop_duplicates(["chave", "periodo"])[["chave", "periodo", "officer_comercial"]]
    lst = ov.groupby(["chave", "periodo"])["officer_comercial"].agg(lambda s: " | ".join(sorted(s))) \
        .rename("officers_encontrados").reset_index()
    qtd = ov.groupby(["chave", "periodo"]).size().rename("_n_off").reset_index()
    cons = cons.merge(top, on=["chave", "periodo"], how="left").merge(lst, on=["chave", "periodo"], how="left") \
        .merge(qtd, on=["chave", "periodo"], how="left")
    cons["officer_comercial"] = cons["officer_comercial"].fillna("")
    cons["officers_encontrados"] = cons["officers_encontrados"].fillna("")
    cons["officer_divergente"] = cons["_n_off"].fillna(0) > 1
    for _, r in cons[cons["officer_divergente"]].iterrows():
        rel.officer_divergencias.append({"cliente": r["cliente"], "periodo": r["periodo"].strftime("%m/%Y"),
                                         "officers": r["officers_encontrados"], "escolhido": r["officer_comercial"]})

    tot = cons.groupby("periodo")["volume_producao"].transform("sum")
    cons["percentual_representatividade"] = np.where(tot > 0, cons["volume_producao"] / tot.replace(0, np.nan), np.nan)
    cons = cons.drop(columns=["_n_spread", "_n_prazo", "_n_off"]).sort_values(["chave", "periodo"]).reset_index(drop=True)

    # conciliação
    det = float(d["volume_producao"].sum())
    con = float(cons["volume_producao"].sum())
    rel.conciliacao.append({"item": "Volume: soma das linhas válidas x soma consolidada", "origem": det, "calculado": con,
                            "diferenca": det - con, "status": "OK" if abs(det - con) < 0.01 else "DIVERGENTE"})
    for _, t in rel.totais_origem.iterrows():
        if pd.notna(t["volume_producao"]):
            dif = float(t["volume_producao"]) - det
            ok = abs(dif) <= max(1.0, 1e-9 * abs(det))
            rel.conciliacao.append({"item": f"Total informado pela origem (linha {int(t['_linha_origem'])}) x soma dos detalhes",
                                    "origem": float(t["volume_producao"]), "calculado": det, "diferenca": dif,
                                    "status": "OK" if ok else "DIVERGENTE"})
    return cons


def processar_arquivo(caminho, de_para: Optional[dict] = None):
    ident = identificar_arquivo_fonte(Path(caminho))
    if not ident.reconhecido:
        raise ErroLeitura(ident.mensagem)
    bruto = carregar_base(Path(caminho), ident)
    d, rel = normalizar_base(bruto, de_para)
    if d.empty:
        raise ErroLeitura("Nenhuma linha válida encontrada após a limpeza.")
    cons = consolidar_base(d, rel)
    return ident, cons, rel


# ----------------------------------------------------------------------------
# Análise
# ----------------------------------------------------------------------------
@dataclass
class Parametros:
    ano_atual: int
    ano_comparacao: int
    mes_corte: int
    qtd_bloco: int = 60                 # 0 = todos
    exibir_sem_comparacao: bool = False
    exibir_producao_zero: bool = False
    officer: Optional[str] = None
    cliente_filtro: str = ""
    tolerancia: float = 0.01
    visao: str = "gerencial"            # gerencial | estrita
    spread_unidade: str = "percentual"  # percentual: 0,04 = 0,04% | fracao: 0,04 = 4%
    pasta_saida: Optional[Path] = None


def sugerir_parametros(cons: pd.DataFrame) -> Parametros:
    ano = int(cons["ano"].max())
    return Parametros(ano_atual=ano, ano_comparacao=ano - 1,
                      mes_corte=int(cons.loc[cons["ano"] == ano, "mes_numero"].max()))


@dataclass
class Analise:
    clientes: pd.DataFrame
    clientes_todos: pd.DataFrame
    matriz: pd.DataFrame
    plataforma: pd.DataFrame
    indicadores: list
    params: Parametros
    kpi: dict = field(default_factory=dict)


PA, PT = "producao_acumulada_anterior", "producao_acumulada_atual"


def calcular_analise(cons: pd.DataFrame, p: Parametros) -> Analise:
    ac, comp, corte = p.ano_atual, p.ano_comparacao, p.mes_corte
    if ac == comp:
        raise ValueError("Ano atual e ano de comparação devem ser diferentes.")
    c = cons[cons["ano"].isin([ac, comp])].copy()
    if c.empty:
        raise ValueError("Não há dados nos anos selecionados.")
    idx = ["chave", "ano", "mes_numero"]
    cols = pd.MultiIndex.from_product([[comp, ac], range(1, 13)], names=["ano", "mes_numero"])
    vol = c.set_index(idx)["volume_producao"].unstack(["ano", "mes_numero"]).reindex(columns=cols)
    pres = c.assign(_x=1.0).set_index(idx)["_x"].unstack(["ano", "mes_numero"]).reindex(columns=cols)
    pres = pres.reindex(vol.index).notna()
    v0 = vol.fillna(0.0)
    ate = list(range(1, corte + 1))

    def soma(ano):
        return v0[[(ano, m) for m in ate]].sum(axis=1)

    def acum(ano):
        s = c[(c["ano"] == ano) & (c["mes_numero"] <= corte)].groupby("chave")[
            ["_sw", "_sv", "_pw", "_pv", "margem"]].sum()
        return s.reindex(vol.index).fillna(0.0)

    ant, atu = acum(comp), acum(ac)
    cl = pd.DataFrame(index=vol.index)
    cl["cliente"] = c.groupby("chave")["cliente"].first()
    # officer de referência: mais recente com produção
    o = c[c["officer_comercial"] != ""].assign(_pos=lambda t: (t["volume_producao"].fillna(0) > 0).astype(int)) \
        .sort_values(["chave", "_pos", "periodo"])
    cl["officer_referencia"] = o.groupby("chave")["officer_comercial"].last().reindex(cl.index).fillna("")
    cl["officers_encontrados"] = c.groupby("chave")["officers_encontrados"].agg(
        lambda s: " | ".join(sorted({x for t in s for x in t.split(" | ") if x}))).reindex(cl.index).fillna("")
    cl["officer_divergente"] = c.groupby("chave")["officer_divergente"].any().reindex(cl.index).fillna(False)

    cl[PA], cl[PT] = soma(comp), soma(ac)
    cl["variacao_absoluta"] = cl[PT] - cl[PA]
    cl["variacao_percentual"] = np.where(cl[PA] > 0, cl["variacao_absoluta"] / cl[PA].replace(0, np.nan), np.nan)
    total_atu = float(cl[PT].sum())
    cl["participacao_cliente"] = cl[PT] / total_atu if total_atu else np.nan
    for suf, a in (("anterior", ant), ("atual", atu)):
        cl[f"spread_ponderado_{suf}"] = np.where(a["_sv"] > 0, a["_sw"] / a["_sv"].replace(0, np.nan), np.nan)
        cl[f"prazo_ponderado_{suf}"] = np.where(a["_pv"] > 0, a["_pw"] / a["_pv"].replace(0, np.nan), np.nan)
        cl[f"margem_acumulada_{suf}"] = a["margem"]
        cl[f"_sw_{suf}"], cl[f"_sv_{suf}"], cl[f"_pw_{suf}"], cl[f"_pv_{suf}"] = a["_sw"], a["_sv"], a["_pw"], a["_pv"]
    fator = 100.0 if p.spread_unidade == "percentual" else 10000.0
    cl["variacao_spread_bps"] = (cl["spread_ponderado_atual"] - cl["spread_ponderado_anterior"]) * fator
    cl["variacao_spread_percentual"] = np.where(
        cl["spread_ponderado_anterior"] > 0,
        cl["spread_ponderado_atual"] / cl["spread_ponderado_anterior"].replace(0, np.nan) - 1, np.nan)
    tm = float(atu["margem"].sum())
    cl["participacao_margem"] = atu["margem"] / tm if tm else np.nan

    for m in range(1, 13):
        cl[f"diferenca_{MESES_KEY[m - 1]}"] = (v0[(ac, m)] - v0[(comp, m)]) if m <= corte else np.nan

    def meses_aus(ano):
        return pres[[(ano, m) for m in ate]].apply(
            lambda r: ", ".join(MESES_ABREV[m - 1] for m, ok in zip(ate, r.values) if not ok), axis=1)

    for suf, ano in (("ano_anterior", comp), ("ano_atual", ac)):
        cl[f"tem_dado_{suf}"] = pres[[(ano, m) for m in ate]].any(axis=1)
        cl[f"quantidade_meses_{suf}"] = pres[[(ano, m) for m in ate]].sum(axis=1).astype(int)
        cl[f"meses_ausentes_{suf}"] = meses_aus(ano)

    # classificação
    tol = p.tolerancia
    ap, tp = cl[PA].round(2) > 0, cl[PT].round(2) > 0
    var = cl["variacao_absoluta"]
    cond = [~ap & tp, ap & ~tp, ap & tp & (var > tol), ap & tp & (var < -tol), ap & tp, ~ap & ~tp]
    nomes = ["NOVO_CLIENTE", "SEM_PRODUCAO_ATUAL", "SUPERAVIT", "DEFICIT", "ESTAVEL", "SEM_BASE_COMPARAVEL"]
    cl["status_variacao"] = np.select(cond, nomes, default="SEM_BASE_COMPARAVEL")
    vr = var.round(2)
    cl["status_estrito"] = np.select([~ap & ~tp, vr > 0, vr < 0], ["SEM_BASE_COMPARAVEL", "SUPERAVIT", "DEFICIT"],
                                     default="ESTAVEL")

    # matriz mensal (todos os clientes dos dois anos)
    mz = pd.DataFrame({"cliente": cl["cliente"], "officer_referencia": cl["officer_referencia"]})
    for suf, ano in (("ano_anterior", comp), ("ano_atual", ac)):
        for m in range(1, 13):
            mz[f"{MESES_KEY[m - 1]}_{suf}"] = v0[(ano, m)]
    for k in [c_ for c_ in cl.columns if c_.startswith(("tem_dado", "quantidade_meses", "meses_ausentes"))]:
        mz[k] = cl[k]

    cl = cl.reset_index()
    mz = mz.reset_index()
    todos = cl.copy()

    # filtros
    vis = cl
    if p.officer:
        vis = vis[vis["officer_referencia"] == p.officer]
    if p.cliente_filtro.strip():
        termos = [_chave(t) for t in re.split(r"[;|]", p.cliente_filtro) if t.strip()]
        vis = vis[vis["chave"].map(lambda k: any(t in k for t in termos))]
    mz = mz[mz["chave"].isin(vis["chave"])]
    keep = vis["chave"]

    # plataforma (universo filtrado)
    vv = v0.reset_index().set_index("chave").loc[keep]
    plat = []
    for m in range(1, 13):
        a_, b_ = float(vv[(comp, m)].sum()), float(vv[(ac, m)].sum())
        plat.append({"mes": m, "anterior": a_ if m <= corte else np.nan, "atual": b_ if m <= corte else np.nan})
    plat = pd.DataFrame(plat)
    plat["variacao_absoluta"] = plat["atual"] - plat["anterior"]
    plat["variacao_percentual"] = np.where(plat["anterior"] > 0, plat["variacao_absoluta"] / plat["anterior"], np.nan)
    cf = c[c["chave"].isin(keep)]

    def spm(ano, m):
        sub = cf[(cf["ano"] == ano) & (cf["mes_numero"] == m)]
        v = float(sub["_sv"].sum())
        return float(sub["_sw"].sum()) / v if v > 0 else np.nan

    plat["spread_anterior"] = [spm(comp, m) if m <= corte else np.nan for m in range(1, 13)]
    plat["spread_atual"] = [spm(ac, m) if m <= corte else np.nan for m in range(1, 13)]
    plat["variacao_spread_bps"] = (plat["spread_atual"] - plat["spread_anterior"]) * fator
    plat["variacao_spread_percentual"] = np.where(plat["spread_anterior"] > 0,
                                                  plat["spread_atual"] / plat["spread_anterior"] - 1, np.nan)

    pa_, pt_ = float(vis[PA].sum()), float(vis[PT].sum())

    def pond(col_w, col_v):
        v = float(vis[col_v].sum())
        return float(vis[col_w].sum()) / v if v > 0 else np.nan

    sp_ant, sp_atu = pond("_sw_anterior", "_sv_anterior"), pond("_sw_atual", "_sv_atual")
    sp_bps = (sp_atu - sp_ant) * fator if sp_ant == sp_ant and sp_atu == sp_atu else np.nan
    sp_pct = sp_atu / sp_ant - 1 if sp_ant == sp_ant and sp_atu == sp_atu and sp_ant > 0 else np.nan
    col_st = "status_variacao" if p.visao == "gerencial" else "status_estrito"
    cont = vis[col_st].value_counts()
    ind = [
        (f"Produção acumulada {comp}", pa_, "moeda"), (f"Produção acumulada {ac}", pt_, "moeda"),
        ("Variação absoluta total", pt_ - pa_, "moeda"),
        ("Variação percentual total", (pt_ - pa_) / pa_ if pa_ > 0 else np.nan, "pct"),
        (f"Spread médio ponderado {comp}", sp_ant, "spread"),
        (f"Spread médio ponderado {ac}", sp_atu, "spread"),
        ("Variação do spread (bps)", sp_bps, "bps"),
        ("Variação do spread (%)", sp_pct, "pctsinal"),
        (f"Prazo ponderado {comp}", pond("_pw_anterior", "_pv_anterior"), "num"),
        (f"Prazo ponderado {ac}", pond("_pw_atual", "_pv_atual"), "num"),
        (f"Margem total {comp}", float(vis["margem_acumulada_anterior"].sum()), "moeda"),
        (f"Margem total {ac}", float(vis["margem_acumulada_atual"].sum()), "moeda"),
        ("Quantidade de clientes", len(vis), "int"),
        ("Clientes em superávit", int(cont.get("SUPERAVIT", 0)), "int"),
        ("Clientes em déficit", int(cont.get("DEFICIT", 0)), "int"),
        ("Clientes novos", int(cont.get("NOVO_CLIENTE", 0)), "int"),
        ("Clientes sem produção no ano atual", int(cont.get("SEM_PRODUCAO_ATUAL", 0)), "int"),
        ("Clientes sem variação", int(cont.get("ESTAVEL", 0)), "int"),
        ("Clientes sem base comparável", int(cont.get("SEM_BASE_COMPARAVEL", 0)), "int"),
    ]
    kpi = {"pa": pa_, "pt": pt_, "var": pt_ - pa_, "varpct": (pt_ - pa_) / pa_ if pa_ > 0 else np.nan,
           "sp_ant": sp_ant, "sp_atu": sp_atu, "sp_bps": sp_bps, "sp_pct": sp_pct, "n": len(vis),
           "sup": int(cont.get("SUPERAVIT", 0)), "def": int(cont.get("DEFICIT", 0)),
           "novos": int(cont.get("NOVO_CLIENTE", 0)), "enc": int(cont.get("SEM_PRODUCAO_ATUAL", 0))}
    return Analise(vis.reset_index(drop=True), todos, mz.reset_index(drop=True), plat, ind, p, kpi)


# ----------------------------------------------------------------------------
# Exportação Excel
# ----------------------------------------------------------------------------
FONTE = "Arial"
AZUL, LARANJA = "0B2A4A", "EC7000"
COR_BLOCO = {"SUPERAVIT": "1E7B34", "DEFICIT": "B3261E", "NOVO_CLIENTE": "1F5FA8", "SEM_PRODUCAO_ATUAL": "6B5B4B",
             "ESTAVEL": "5F6B7A", "SEM_BASE_COMPARAVEL": "8A8D91"}
TITULO_BLOCO = {"SUPERAVIT": "SUPERÁVIT DE PRODUÇÃO", "DEFICIT": "QUEDA DE PRODUÇÃO", "NOVO_CLIENTE": "CLIENTES NOVOS",
                "SEM_PRODUCAO_ATUAL": "SEM PRODUÇÃO NO ANO ATUAL", "ESTAVEL": "SEM VARIAÇÃO",
                "SEM_BASE_COMPARAVEL": "SEM BASE COMPARÁVEL"}
FMT_MOEDA = '#,##0;-#,##0;"-"'
FMT_MOEDA2 = "#,##0.00"
FMT_PCT = "0.00%"
FINO = Side(style="thin", color="BFC5CE")


def _f(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONTE, bold=bold, color=color, size=size, italic=italic)


def _fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def _py(v):
    if v is None or v is pd.NaT:
        return None
    if isinstance(v, (np.floating, float)):
        return None if math.isnan(v) else float(v)
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, np.bool_):
        return bool(v)
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, str) and v[:1] == "=":
        return "'" + v
    return v


def _cel(ws, r, c, v, fonte=None, fmt=None, fill=None, al=None, borda=False, raw=False):
    cell = ws.cell(r, c, v if raw else _py(v))
    cell.font = fonte or _f()
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if al:
        cell.alignment = al
    if borda:
        cell.border = Border(top=FINO, bottom=FINO)
    return cell


def _cor_valor(v):
    if v is None or (isinstance(v, float) and math.isnan(v)) or abs(v) < 0.005:
        return "7F7F7F"
    return "1E7B34" if v > 0 else "B3261E"


FMT_BPS = '+0.0;-0.0;0.0'
FMT_SETA = '"▲ "#,##0;"▼ "#,##0;"-"'
FMT_SETA_PCT = '"▲ "0.0%;"▼ "0.0%;"-"'
FMT_PCT_SINAL = '+0.0%;-0.0%;0.0%'
NC = 11  # colunas do painel


def fmt_spread(p: Parametros) -> str:
    return '0.0000"%"' if p.spread_unidade == "percentual" else "0.00%"


def _fator(p: Parametros) -> float:
    return 100.0 if p.spread_unidade == "percentual" else 10000.0


def _pond(sub, suf):
    v = float(sub[f"_sv_{suf}"].sum())
    return float(sub[f"_sw_{suf}"].sum()) / v if v > 0 else None


def _barra(ws, r, texto, cor, altura=24):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
    for c in range(1, NC + 1):
        ws.cell(r, c).fill = _fill(cor)
    _cel(ws, r, 1, texto, _f(True, "FFFFFF", 12), fill=_fill(cor), al=Alignment(vertical="center", indent=1))
    ws.row_dimensions[r].height = altura


def _cabecalho(ws, r, rotulos):
    for j, t in enumerate(rotulos, 1):
        _cel(ws, r, j, t, _f(True, AZUL, 9), fill=_fill("E7E9EE"),
             al=Alignment(horizontal="left" if j == 3 else "center", vertical="center", wrap_text=True))
    ws.row_dimensions[r].height = 30


def _linha_painel(ws, rr, p, conf, part, nome, pa, pt, var, varp, sa, st, estilo="normal"):
    """Uma linha do painel: produção ano a ano, variação e spread médio ano a ano."""
    bps = (st - sa) * _fator(p) if sa is not None and st is not None else None
    spp = st / sa - 1 if sa is not None and st is not None and sa > 0 else None
    total, resto = estilo == "total", estilo == "resto"
    fill = _fill("DDE3EC") if total else _fill("F6F7F9") if estilo == "zebra" else None
    cor_base = AZUL if total else "595959" if resto else "000000"
    f = _f(total, cor_base, italic=resto)
    if conf is None:
        _cel(ws, rr, 1, None, fill=fill)
    else:
        _cel(ws, rr, 1, "✔" if conf else "✖", _f(True, "1E7B34" if conf else "B3261E"), fill=fill,
             al=Alignment(horizontal="center"), borda=True)
    _cel(ws, rr, 2, part, f, FMT_PCT, fill, Alignment(horizontal="center"), True)
    _cel(ws, rr, 3, nome, f, fill=fill, borda=True)
    _cel(ws, rr, 4, pa, f, FMT_MOEDA, fill, borda=True)
    _cel(ws, rr, 5, pt, f, FMT_MOEDA, fill, borda=True)
    _cel(ws, rr, 6, var, _f(True, _cor_valor(var)), FMT_SETA, fill, borda=True)
    _cel(ws, rr, 7, varp, _f(color=_cor_valor(varp)), FMT_SETA_PCT, fill, borda=True)
    _cel(ws, rr, 8, sa, f, fmt_spread(p), fill, Alignment(horizontal="right"), True)
    _cel(ws, rr, 9, st, f, fmt_spread(p), fill, Alignment(horizontal="right"), True)
    _cel(ws, rr, 10, bps, _f(True, _cor_valor(bps)), FMT_BPS, fill, borda=True)
    _cel(ws, rr, 11, spp, _f(color=_cor_valor(spp)), FMT_PCT_SINAL, fill, borda=True)


def _rotulos(p):
    return ["Conf.", "Part. do cliente", "Cliente", f"Produção {p.ano_comparacao}", f"Produção {p.ano_atual}",
            "Variação (R$)", "Variação (%)", f"Spread médio {p.ano_comparacao}", f"Spread médio {p.ano_atual}",
            "Var. spread (bps)", "Var. spread (%)"]


def _bloco_clientes(ws, r, status, df, p: Parametros, limite):
    n = len(df)
    _barra(ws, r, f"{TITULO_BLOCO[status]}  —  {n} cliente(s)", COR_BLOCO[status])
    r += 1
    _cabecalho(ws, r, _rotulos(p))
    r += 1
    exib = df if (limite <= 0 or n <= limite) else df.iloc[:limite]
    resto = df.iloc[len(exib):]
    mes_cols = [f"diferenca_{k}" for k in MESES_KEY[:p.mes_corte]]

    def n_(v):
        return None if v is None or (isinstance(v, float) and math.isnan(v)) else float(v)

    for i, (_, row) in enumerate(exib.iterrows()):
        ok = abs(float(np.nansum([row[c] for c in mes_cols])) - row["variacao_absoluta"]) <= 0.01
        _linha_painel(ws, r, p, ok, row["participacao_cliente"], row["cliente"], row[PA], row[PT],
                      row["variacao_absoluta"], n_(row["variacao_percentual"]),
                      n_(row["spread_ponderado_anterior"]), n_(row["spread_ponderado_atual"]),
                      "zebra" if i % 2 else "normal")
        r += 1

    def agregada(sub, rotulo, estilo):
        pa, pt = float(sub[PA].sum()), float(sub[PT].sum())
        _linha_painel(ws, r, p, None, float(sub["participacao_cliente"].sum()), rotulo, pa, pt, pt - pa,
                      (pt - pa) / pa if pa > 0 else None, _pond(sub, "anterior"), _pond(sub, "atual"), estilo)

    if len(resto):
        agregada(resto, f"Demais {len(resto)} cliente(s) não exibido(s) — limite de {limite} por bloco", "resto")
        r += 1
    agregada(df, f"TOTAL DO BLOCO ({n} clientes)", "total")
    return r + 2


def _bloco_plataforma(ws, r, a: Analise):
    p, pl, k = a.params, a.plataforma, a.kpi
    _barra(ws, r, "PRODUÇÃO TOTAL MENSAL DA PLATAFORMA  (mês a mês, ano x ano)", AZUL)
    r += 1
    rot = _rotulos(p)
    rot[0], rot[1], rot[2] = "", "", "Mês"
    _cabecalho(ws, r, rot)
    r += 1
    for m in range(p.mes_corte):
        x = pl.iloc[m]
        sa = None if pd.isna(x["spread_anterior"]) else float(x["spread_anterior"])
        st = None if pd.isna(x["spread_atual"]) else float(x["spread_atual"])
        vp = None if pd.isna(x["variacao_percentual"]) else float(x["variacao_percentual"])
        _linha_painel(ws, r, p, None, None, MESES_NOME[m], x["anterior"], x["atual"], x["variacao_absoluta"], vp,
                      sa, st, "zebra" if m % 2 else "normal")
        ws.cell(r, 2).value = None
        r += 1
    sa = None if pd.isna(k["sp_ant"]) else k["sp_ant"]
    st = None if pd.isna(k["sp_atu"]) else k["sp_atu"]
    _linha_painel(ws, r, p, None, None, f"ACUMULADO até {MESES_NOME[p.mes_corte - 1]}", k["pa"], k["pt"], k["var"],
                  None if pd.isna(k["varpct"]) else k["varpct"], sa, st, "total")
    ws.cell(r, 2).value = None
    return r + 2


def _bloco_indicadores(ws, r, a: Analise):
    p = a.params
    _barra(ws, r, "INDICADORES", LARANJA)
    r += 1
    fmts = {"moeda": FMT_MOEDA2, "pct": FMT_PCT, "pctsinal": FMT_PCT_SINAL, "num": "0.00", "int": "#,##0",
            "spread": fmt_spread(p), "bps": FMT_BPS}
    for i, (rot, val, tipo) in enumerate(a.indicadores):
        fill = _fill("F6F7F9") if i % 2 else None
        _cel(ws, r, 3, rot, _f(True), fill=fill, borda=True)
        cor = _cor_valor(val) if (rot.startswith("Variação") and val == val) else "000000"
        _cel(ws, r, 4, None if val != val else val, _f(True, cor), fmts[tipo], fill, borda=True)
        r += 1
    return r + 1


def _card(ws, r, c1, c2, rotulo, valor, fmt, sub, cor):
    for rr in range(r, r + 3):
        ws.merge_cells(start_row=rr, start_column=c1, end_row=rr, end_column=c2)
        for cc in range(c1, c2 + 1):
            ws.cell(rr, cc).fill = _fill("F4F6FA")
            ws.cell(rr, cc).border = Border(bottom=Side(style="thick", color=cor) if rr == r + 2 else None)
    al = Alignment(horizontal="center", vertical="center")
    _cel(ws, r, c1, rotulo, _f(True, "6B7280", 9), fill=_fill("F4F6FA"), al=al)
    _cel(ws, r + 1, c1, valor, _f(True, cor, 16), fmt, _fill("F4F6FA"), al)
    _cel(ws, r + 2, c1, sub, _f(False, "404040", 9), fill=_fill("F4F6FA"), al=al)
    ws.cell(r + 2, c1).border = Border(bottom=Side(style="thick", color=cor))
    ws.row_dimensions[r].height = 18
    ws.row_dimensions[r + 1].height = 28
    ws.row_dimensions[r + 2].height = 18


def _escrever_apresentacao(ws, a: Analise, meta: dict):
    p, k = a.params, a.kpi
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = LARANJA
    larg = {1: 7, 2: 11, 3: 44, 4: 17, 5: 17, 6: 17, 7: 12, 8: 15, 9: 15, 10: 13, 11: 13}
    for c in range(1, NC + 1):
        ws.column_dimensions[get_column_letter(c)].width = larg[c]
    _barra(ws, 1, "RISCO SACADO — SUPERÁVIT E QUEDA DE PRODUÇÃO", AZUL, 36)
    ws.cell(1, 1).font = _f(True, "FFFFFF", 16)
    mes = MESES_NOME[p.mes_corte - 1]
    filtros = []
    if p.officer:
        filtros.append(f"Officer: {p.officer}")
    if p.cliente_filtro.strip():
        filtros.append(f"Cliente: {p.cliente_filtro}")
    textos = [
        f"Período analisado: Janeiro a {mes}/{p.ano_atual}  versus  Janeiro a {mes}/{p.ano_comparacao}",
        f"Atualizado em {datetime.now():%d/%m/%Y %H:%M}   |   Origem: {meta.get('arquivo', '')}   |   "
        f"Ano atual: {p.ano_atual}   |   Comparação: {p.ano_comparacao}   |   Corte: {mes}",
        f"Visão: {'Gerencial' if p.visao == 'gerencial' else 'Estrita'}   |   Tolerância: R$ {p.tolerancia:,.2f}   |   "
        f"Filtros: {'; '.join(filtros) if filtros else 'nenhum'}   |   "
        f"Limite por bloco: {'todos' if p.qtd_bloco <= 0 else p.qtd_bloco}   |   "
        f"Spread: {'valor em % (0,04 = 0,04%)' if p.spread_unidade == 'percentual' else 'valor em fração (0,04 = 4%)'}",
    ]
    for i, t in enumerate(textos, 2):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=NC)
        _cel(ws, i, 1, t, _f(i == 2, "404040", 11 if i == 2 else 9), al=Alignment(indent=1))

    # cartões de resumo
    cv = _cor_valor(k["var"])
    spd = "" if pd.isna(k["sp_bps"]) else f"{k['sp_bps']:+.1f} bps  ({k['sp_pct']:+.1%})" if not pd.isna(k["sp_pct"]) else f"{k['sp_bps']:+.1f} bps"
    ws.row_dimensions[5].height = 8
    _card(ws, 6, 1, 3, f"PRODUÇÃO ACUMULADA {p.ano_atual}", k["pt"], '"R$ "#,##0',
          f"{p.ano_comparacao}: R$ {k['pa']:,.0f}", AZUL)
    _card(ws, 6, 4, 5, "VARIAÇÃO ANO x ANO", k["var"], '"▲ R$ "#,##0;"▼ R$ "#,##0;"-"',
          "" if pd.isna(k["varpct"]) else f"{k['varpct']:+.1%}", cv)
    _card(ws, 6, 6, 7, f"SPREAD MÉDIO {p.ano_atual}", None if pd.isna(k["sp_atu"]) else k["sp_atu"], fmt_spread(p),
          spd, _cor_valor(k["sp_bps"]) if not pd.isna(k["sp_bps"]) else AZUL)
    _card(ws, 6, 8, 9, "CLIENTES", k["n"], "#,##0", f"▲ {k['sup']}   ▼ {k['def']}", LARANJA)
    _card(ws, 6, 10, 11, "NOVOS | ENCERRADOS", f"{k['novos']} | {k['enc']}", None, "sem produção atual", "6B5B4B")
    r = 10

    cl = a.clientes
    col = "status_variacao" if p.visao == "gerencial" else "status_estrito"
    ordem = {"SUPERAVIT": [("variacao_absoluta", False), (PT, False), ("chave", True)],
             "DEFICIT": [("variacao_absoluta", True), (PT, False), ("chave", True)],
             "NOVO_CLIENTE": [(PT, False), ("chave", True)],
             "SEM_PRODUCAO_ATUAL": [(PA, False), ("chave", True)],
             "ESTAVEL": [(PT, False), ("chave", True)],
             "SEM_BASE_COMPARAVEL": [("chave", True)]}
    sequencia = ["SUPERAVIT", "DEFICIT"] + (["NOVO_CLIENTE", "SEM_PRODUCAO_ATUAL"] if p.visao == "gerencial" else []) \
        + ["ESTAVEL"] + (["SEM_BASE_COMPARAVEL"] if p.exibir_sem_comparacao else [])
    for st in sequencia:
        sub = cl[cl[col] == st]
        if not p.exibir_producao_zero:
            sub = sub[(sub[PA].round(2) != 0) | (sub[PT].round(2) != 0)]
        if sub.empty:
            continue
        ks = ordem[st]
        sub = sub.sort_values([k_ for k_, _ in ks], ascending=[a_ for _, a_ in ks], kind="mergesort")
        r = _bloco_clientes(ws, r, st, sub, p, p.qtd_bloco)
    r = _bloco_plataforma(ws, r, a)
    r = _bloco_indicadores(ws, r, a)
    notas = ["Conf. ✔: a soma das diferenças mensais do cliente é igual à variação acumulada.",
             "Part. do cliente = produção acumulada atual do cliente ÷ produção total da plataforma no mesmo período.",
             "Spread médio = média ponderada pelo volume de produção, de janeiro até o mês de corte. "
             "Var. spread (bps) = (spread atual − spread anterior) em pontos-base; (%) = variação relativa.",
             "Detalhe mês a mês de qualquer cliente: aba Cliente_Mensal."]
    for t in notas:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
        _cel(ws, r, 1, t, _f(italic=True, color="595959", size=9), al=Alignment(indent=1))
        r += 1
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _aba_mensal(ws, a: Analise, base_cols: list, n_base: int, n_cli: int, cliente_padrao: str):
    """Drill-down vertical: escolha o cliente na lista e veja mês a mês (ano x ano, spread)."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.worksheet.datavalidation import DataValidation
    p = a.params
    cel = lambda *args, **kw: _cel(*args, raw=True, **kw)  # fórmulas não podem passar por _py
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = AZUL
    for c, w in zip("ABCDEFGHIJ", (6, 16, 18, 18, 18, 12, 15, 15, 15, 4)):
        ws.column_dimensions[c].width = w
    ws.merge_cells("A1:I1")
    for c in range(1, 10):
        ws.cell(1, c).fill = _fill(AZUL)
    cel(ws, 1, 1, "ANÁLISE MENSAL POR CLIENTE", _f(True, "FFFFFF", 14), fill=_fill(AZUL), al=Alignment(vertical="center", indent=1))
    ws.row_dimensions[1].height = 30

    def rng(nome):
        col = get_column_letter(base_cols.index(nome) + 1)
        return f"Base_Consolidada!${col}$2:${col}${n_base}"
    cli, ano, mes = rng("cliente"), rng("ano"), rng("mes_numero")
    vol, sw, sv = rng("volume_producao"), rng("spread_x_volume"), rng("volume_com_spread")

    cel(ws, 3, 2, "Cliente ▸", _f(True, AZUL, 11))
    ws.merge_cells("C3:F3")
    for c in range(3, 7):
        ws.cell(3, c).fill = _fill("FFF4E5")
        ws.cell(3, c).border = Border(top=Side(style="medium", color=LARANJA), bottom=Side(style="medium", color=LARANJA))
    cel(ws, 3, 3, cliente_padrao, _f(True, "000000", 11), fill=_fill("FFF4E5"))
    dv = DataValidation(type="list", formula1=f"=Clientes!$A$2:$A${n_cli + 1}", allow_blank=False,
                        showErrorMessage=True, errorTitle="Cliente", error="Escolha um cliente da lista.")
    ws.add_data_validation(dv)
    dv.add("C3")
    cel(ws, 4, 2, "Officer", _f(True, "595959"))
    ws.merge_cells("C4:F4")
    cel(ws, 4, 3, '=IFERROR(INDEX(Clientes!$B$2:$B$%d,MATCH($C$3,Clientes!$A$2:$A$%d,0)),"")' % (n_cli + 1, n_cli + 1))
    cel(ws, 5, 2, "Situação", _f(True, "595959"))
    ws.merge_cells("C5:F5")
    cel(ws, 5, 3, '=IFERROR(INDEX(Clientes!$C$2:$C$%d,MATCH($C$3,Clientes!$A$2:$A$%d,0)),"")' % (n_cli + 1, n_cli + 1))
    for i in (4, 5):
        ws.cell(i, 3).font = _f(True)
    # parâmetros (fixos, vindos do relatório)
    for i, (rot, val) in enumerate([("Ano atual", p.ano_atual), ("Ano comparação", p.ano_comparacao),
                                    ("Mês de corte", p.mes_corte), ("Fator bps", _fator(p))]):
        cel(ws, 3 + i, 8, rot, _f(False, "7F7F7F", 8))
        cel(ws, 3 + i, 9, val, _f(False, "7F7F7F", 8))
    A_AT, A_AN, CORTE, FAT = "$I$3", "$I$4", "$I$5", "$I$6"

    cab = ["#", "Mês", f"Produção {p.ano_comparacao}", f"Produção {p.ano_atual}", "Variação (R$)", "Variação (%)",
           f"Spread {p.ano_comparacao}", f"Spread {p.ano_atual}", "Var. spread (bps)"]
    for j, t in enumerate(cab, 1):
        cel(ws, 8, j, t, _f(True, "FFFFFF", 9), fill=_fill(AZUL), al=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[8].height = 30
    sp = fmt_spread(p)

    def spread_f(an, r_):
        return (f'=IF($A{r_}>{CORTE},"",IFERROR(SUMIFS({sw},{cli},$C$3,{ano},{an},{mes},$A{r_})/'
                f'SUMIFS({sv},{cli},$C$3,{ano},{an},{mes},$A{r_}),""))')

    for m in range(1, 13):
        r_ = 8 + m
        fill = _fill("F6F7F9") if m % 2 == 0 else None
        cel(ws, r_, 1, m, _f(False, "A0A0A0", 8), fill=fill, al=Alignment(horizontal="center"), borda=True)
        cel(ws, r_, 2, MESES_NOME[m - 1], _f(True), fill=fill, borda=True)
        for col, an in ((3, A_AN), (4, A_AT)):
            cel(ws, r_, col, f'=IF($A{r_}>{CORTE},"",SUMIFS({vol},{cli},$C$3,{ano},{an},{mes},$A{r_}))', _f(), FMT_MOEDA, fill, borda=True)
        cel(ws, r_, 5, f'=IF($A{r_}>{CORTE},"",D{r_}-C{r_})', _f(True), FMT_SETA, fill, borda=True)
        cel(ws, r_, 6, f'=IF($A{r_}>{CORTE},"",IF(C{r_}>0,E{r_}/C{r_},""))', _f(), FMT_SETA_PCT, fill, borda=True)
        cel(ws, r_, 7, spread_f(A_AN, r_), _f(), sp, fill, borda=True)
        cel(ws, r_, 8, spread_f(A_AT, r_), _f(), sp, fill, borda=True)
        cel(ws, r_, 9, f'=IF(OR(G{r_}="",H{r_}=""),"",(H{r_}-G{r_})*{FAT})', _f(True), FMT_BPS, fill, borda=True)
    t = 21
    tf = _fill("DDE3EC")
    cel(ws, t, 1, None, fill=tf)
    cel(ws, t, 2, "ACUMULADO", _f(True, AZUL), fill=tf)
    cel(ws, t, 3, "=SUM(C9:C20)", _f(True, AZUL), FMT_MOEDA, tf)
    cel(ws, t, 4, "=SUM(D9:D20)", _f(True, AZUL), FMT_MOEDA, tf)
    cel(ws, t, 5, f"=D{t}-C{t}", _f(True, AZUL), FMT_SETA, tf)
    cel(ws, t, 6, f'=IF(C{t}>0,E{t}/C{t},"")', _f(True, AZUL), FMT_SETA_PCT, tf)
    for col, an in ((7, A_AN), (8, A_AT)):
        cel(ws, t, col, f'=IFERROR(SUMIFS({sw},{cli},$C$3,{ano},{an},{mes},"<="&{CORTE})/'
                         f'SUMIFS({sv},{cli},$C$3,{ano},{an},{mes},"<="&{CORTE}),"")', _f(True, AZUL), sp, tf)
    cel(ws, t, 9, f'=IF(OR(G{t}="",H{t}=""),"",(H{t}-G{t})*{FAT})', _f(True, AZUL), FMT_BPS, tf)
    verde, vermelho = Font(name=FONTE, color="1E7B34", bold=True), Font(name=FONTE, color="B3261E", bold=True)
    for faixa in ("E9:F21", "I9:I21"):
        ws.conditional_formatting.add(faixa, CellIsRule(operator="greaterThan", formula=["0"], font=verde))
        ws.conditional_formatting.add(faixa, CellIsRule(operator="lessThan", formula=["0"], font=vermelho))
    cel(ws, 23, 2, "Escolha o cliente na célula laranja. Valores até o mês de corte; spread = média ponderada pelo volume.",
         _f(italic=True, color="595959", size=9))

    ch = BarChart()
    ch.type, ch.grouping = "col", "clustered"
    ch.title = "Produção mensal — ano anterior x ano atual"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=8, max_row=20), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=9, max_row=20))
    ch.series[0].graphicalProperties.solidFill = "A6ADB8"
    ch.series[1].graphicalProperties.solidFill = LARANJA
    ch.y_axis.numFmt = "#,##0"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.height, ch.width = 8.5, 22
    ws.add_chart(ch, "B25")


def _fmt_coluna(nome: str, serie: pd.Series) -> Optional[str]:
    if pd.api.types.is_datetime64_any_dtype(serie):
        return "mm/yyyy"
    if nome.startswith(("participacao", "variacao_percentual", "percentual", "spread")):
        return FMT_PCT
    if nome.startswith("prazo"):
        return "0.00"
    if pd.api.types.is_bool_dtype(serie):
        return None
    if pd.api.types.is_integer_dtype(serie):
        return "0"
    if pd.api.types.is_float_dtype(serie):
        return FMT_MOEDA2
    return None


def _df_para_aba(ws, df: pd.DataFrame, destaque_status: Optional[str] = None, extras: Optional[dict] = None):
    df = df[[c for c in df.columns if not c.startswith("_")]]
    for j, c in enumerate(df.columns, 1):
        _cel(ws, 1, j, c, _f(True, "FFFFFF"), fill=_fill(AZUL), al=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[1].height = 34
    fmts = [(extras or {}).get(c, _fmt_coluna(c, df[c])) for c in df.columns]
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            cell = _cel(ws, i, j, v, _f(), fmts[j - 1])
            if destaque_status and df.columns[j - 1] == destaque_status and isinstance(v, str) and v in COR_BLOCO:
                cell.fill = _fill(COR_BLOCO[v])
                cell.font = _f(True, "FFFFFF")
    for j, c in enumerate(df.columns, 1):
        amostra = [len(str(x)) for x in df[c].head(200)]
        ws.column_dimensions[get_column_letter(j)].width = min(48, max(12, len(str(c)) * 0.9, (max(amostra) if amostra else 0) + 2))
    ws.freeze_panes = ws.cell(2, 3 if "cliente" in df.columns else 1)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{max(2, len(df) + 1)}"


def _secao(ws, r, titulo, cab, linhas, fmts=None):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max(4, len(cab)))
    _cel(ws, r, 1, titulo, _f(True, "FFFFFF", 11), fill=_fill(LARANJA))
    r += 1
    for j, t in enumerate(cab, 1):
        _cel(ws, r, j, t, _f(True, AZUL), fill=_fill("E7E9EE"))
    r += 1
    if not linhas:
        _cel(ws, r, 1, "Nenhum registro.", _f(italic=True, color="7F7F7F"))
        r += 1
    for ln in linhas:
        for j, v in enumerate(ln, 1):
            fmt = (fmts or {}).get(j)
            cell = _cel(ws, r, j, v, _f(), fmt if isinstance(v, (int, float)) else None)
            if v == "OK":
                cell.font = _f(True, "1E7B34")
            elif v == "DIVERGENTE":
                cell.font = _f(True, "B3261E")
        r += 1
    return r + 1


def _escrever_validacao(ws, rel: RelatorioQualidade, cons: pd.DataFrame, a: Analise):
    ws.sheet_view.showGridLines = False
    for c, w in zip("ABCDEFG", (58, 26, 26, 24, 40, 30, 30)):
        ws.column_dimensions[c].width = w
    p = a.params
    r = 1
    m = rel.meta
    r = _secao(ws, r, "ARQUIVO DE ORIGEM", ["Item", "Valor"], [
        ["Arquivo", m.get("arquivo")], ["Tamanho (bytes)", m.get("tamanho_bytes")],
        ["Modificado em", m.get("modificado_em")], ["Abas encontradas", m.get("abas")],
        ["Aba utilizada", m.get("aba_selecionada")], ["Linha do cabeçalho (Excel)", m.get("linha_cabecalho_excel")],
        ["Linhas iniciais / colunas iniciais", f"{m.get('linhas_iniciais')} / {m.get('colunas_iniciais')}"],
        ["Confiança da identificação", f"{m.get('confianca_identificacao')}%"]])
    r = _secao(ws, r, "RESUMO DA LIMPEZA", ["Item", "Quantidade"], [
        ["Linhas lidas (abaixo do cabeçalho)", rel.linhas_iniciais], ["Linhas vazias removidas", rel.linhas_vazias],
        ["Cabeçalhos repetidos removidos", rel.cabecalhos_repetidos],
        ["Linhas de totalização separadas", rel.linhas_totalizacao],
        ["Linhas sem grupo/período válido removidas", rel.linhas_sem_chave], ["Linhas válidas", rel.linhas_validas],
        ["Valores numéricos inválidos", sum(1 for i in rel.inconsistencias if i["tipo"] == "numero")],
        ["Cliente/mês com mais de uma linha (agregados)", len(rel.duplicidades)],
        ["Cliente/mês com mais de um officer", len(rel.officer_divergencias)],
        ["Clientes distintos na base", cons["chave"].nunique()]])
    r = _secao(ws, r, "CONCILIAÇÃO DE TOTAIS", ["Item", "Origem", "Calculado", "Diferença", "Status"],
               [[c["item"], c["origem"], c["calculado"], c["diferenca"], c["status"]] for c in rel.conciliacao]
               + [[f"Clientes (soma) x plataforma (mensal) — {p.ano_atual} acumulado",
                   float(a.clientes_todos[PT].sum()),
                   float(a.clientes_todos[[f'diferenca_{k}' for k in MESES_KEY]].sum().sum() + a.clientes_todos[PA].sum()),
                   None, "OK"]],
               {2: FMT_MOEDA2, 3: FMT_MOEDA2, 4: FMT_MOEDA2})
    # corrige a diferença/estatus da última linha de conciliação interna
    ult = r - 3
    cel_o, cel_c = ws.cell(ult, 2).value, ws.cell(ult, 3).value
    if cel_o is not None and cel_c is not None:
        dif = cel_o - cel_c
        _cel(ws, ult, 4, dif, _f(), FMT_MOEDA2)
        _cel(ws, ult, 5, "OK" if abs(dif) < 0.01 else "DIVERGENTE", _f(True, "1E7B34" if abs(dif) < 0.01 else "B3261E"))
    if len(rel.totais_origem):
        t = rel.totais_origem
        r = _secao(ws, r, "LINHAS DE TOTALIZAÇÃO INFORMADAS PELA ORIGEM (descartadas dos cálculos)",
                   ["Linha no arquivo", "Volume produção", "Spread", "Prazo", "Margem"],
                   [[int(x["_linha_origem"]), _py(x["volume_producao"]), _py(x["spread"]), _py(x["prazo"]), _py(x["margem"])]
                    for _, x in t.iterrows()], {2: FMT_MOEDA2, 3: "0.0000", 4: "0.00", 5: FMT_MOEDA2})
    # meses ausentes
    meses_base = sorted(cons["periodo"].unique())
    if meses_base:
        todos = pd.period_range(meses_base[0], meses_base[-1], freq="M")
        presentes = {pd.Timestamp(x).to_period("M") for x in meses_base}
        faltam = [x.strftime("%m/%Y") for x in todos if x not in presentes]
        r = _secao(ws, r, "MESES SEM NENHUM REGISTRO NA BASE (dentro do intervalo coberto)", ["Mês"], [[x] for x in faltam])
    cl = a.clientes_todos
    r = _secao(ws, r, f"CLIENTES COM MESES AUSENTES ATÉ O CORTE (detalhe na aba Matriz)", ["Item", "Quantidade"], [
        [f"Clientes com lacuna em {p.ano_comparacao}", int(((cl["tem_dado_ano_anterior"]) & (cl["meses_ausentes_ano_anterior"] != "")).sum())],
        [f"Clientes com lacuna em {p.ano_atual}", int(((cl["tem_dado_ano_atual"]) & (cl["meses_ausentes_ano_atual"] != "")).sum())],
        ["Clientes só no ano anterior", int((cl["tem_dado_ano_anterior"] & ~cl["tem_dado_ano_atual"]).sum())],
        ["Clientes só no ano atual", int((~cl["tem_dado_ano_anterior"] & cl["tem_dado_ano_atual"]).sum())]])
    r = _secao(ws, r, "AVISOS", ["Aviso"], [[x] for x in rel.avisos])
    r = _secao(ws, r, "DE/PARA DE NOMES APLICADO", ["Origem (chave)", "Destino", "Linhas"],
               [[x["origem"], x["destino"], x["linhas"]] for x in rel.de_para_usado])
    r = _secao(ws, r, "DUPLICIDADES (mesmo cliente e mês em várias linhas — somadas)",
               ["Cliente", "Período", "Linhas", "Volume somado", "Linhas no arquivo"],
               [[x["cliente"], x["periodo"], x["linhas"], x["volume_somado"], x["linhas_origem"]] for x in rel.duplicidades[:3000]],
               {4: FMT_MOEDA2})
    r = _secao(ws, r, "OFFICERS DIVERGENTES NO MESMO PERÍODO (adotado o de maior volume)",
               ["Cliente", "Período", "Officers encontrados", "Officer adotado"],
               [[x["cliente"], x["periodo"], x["officers"], x["escolhido"]] for x in rel.officer_divergencias[:3000]])
    inc = rel.inconsistencias
    r = _secao(ws, r, f"INCONSISTÊNCIAS ({len(inc)})" + (" — exibindo as 5000 primeiras" if len(inc) > 5000 else ""),
               ["Linha no arquivo", "Campo", "Valor original", "Motivo"],
               [[x["linha"], x["campo"], x["valor"], x["motivo"]] for x in inc[:5000]])


def exportar_excel(a: Analise, cons: pd.DataFrame, rel: RelatorioQualidade, caminho: Path) -> Path:
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Apresentação"
    _escrever_apresentacao(ws, a, rel.meta)

    p = a.params
    cl_cols = ["cliente", "officer_referencia", "status_variacao", "status_estrito", PA, PT, "variacao_absoluta",
               "variacao_percentual", "participacao_cliente", "spread_ponderado_anterior", "spread_ponderado_atual",
               "variacao_spread_bps", "variacao_spread_percentual",
               "prazo_ponderado_anterior", "prazo_ponderado_atual", "margem_acumulada_anterior", "margem_acumulada_atual",
               "participacao_margem"] + [f"diferenca_{k}" for k in MESES_KEY] + \
        ["officers_encontrados", "officer_divergente"] + [c for c in a.clientes.columns if c.startswith(("tem_dado", "quantidade_meses", "meses_ausentes"))]
    cl = a.clientes.sort_values(["variacao_absoluta", PT, "chave"], ascending=[False, False, True], kind="mergesort")
    ex_cl = {"spread_ponderado_anterior": fmt_spread(p), "spread_ponderado_atual": fmt_spread(p),
             "variacao_spread_bps": "0.0", "variacao_spread_percentual": "0.0%"}
    _df_para_aba(wb.create_sheet("Clientes"), cl[cl_cols], destaque_status="status_variacao", extras=ex_cl)
    _df_para_aba(wb.create_sheet("Matriz"), a.matriz.sort_values("cliente"))
    base = cons.copy()
    base["spread_x_volume"], base["volume_com_spread"] = base["_sw"], base["_sv"]
    base_cols = ["cliente", "ano", "mes_numero", "mes_nome", "periodo", "ordem_periodo", "officer_comercial",
                 "officers_encontrados", "officer_divergente", "volume_producao", "spread", "prazo", "margem",
                 "percentual_representatividade", "linhas_origem", "spread_x_volume", "volume_com_spread"]
    _df_para_aba(wb.create_sheet("Base_Consolidada"), base[base_cols],
                 extras={"spread": fmt_spread(p), "spread_x_volume": FMT_MOEDA2, "volume_com_spread": FMT_MOEDA2})
    padrao = str(cl["cliente"].iloc[0]) if len(cl) else ""
    _aba_mensal(wb.create_sheet("Cliente_Mensal", 1), a, base_cols, len(base) + 1, max(1, len(cl)), padrao)
    _escrever_validacao(wb.create_sheet("Validação"), rel, cons, a)
    wb.calculation.fullCalcOnLoad = True
    wb.save(caminho)
    log.info("Relatório salvo em %s", caminho)
    return caminho


def nome_saida(p: Parametros, pasta: Path) -> Path:
    return Path(pasta) / f"Apresentacao_RiscoSacado_{p.ano_atual}-{p.mes_corte:02d}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"


def abrir_arquivo(caminho: Path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(caminho))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(caminho)])
        else:
            subprocess.Popen(["xdg-open", str(caminho)])
    except Exception as e:
        log.warning("Não foi possível abrir automaticamente: %s", e)


# ----------------------------------------------------------------------------
# Interface (Tkinter)
# ----------------------------------------------------------------------------
def abrir_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD
        root = TkinterDnD.Tk()
        dnd = True
    except Exception:
        root, dnd = tk.Tk(), False
    root.title("Risco Sacado — Análise de Produção")
    root.geometry("860x700")
    estado = {"cons": None, "rel": None, "de_para": None}

    v_arq, v_saida, v_depara = tk.StringVar(), tk.StringVar(), tk.StringVar()
    v_ano, v_comp, v_mes = tk.StringVar(), tk.StringVar(), tk.StringVar()
    v_visao, v_qtd, v_todos = tk.StringVar(value="Gerencial"), tk.StringVar(value="60"), tk.BooleanVar(value=False)
    v_off, v_cli, v_tol = tk.StringVar(value="(Todos)"), tk.StringVar(), tk.StringVar(value="0,01")
    v_semcomp, v_zero = tk.BooleanVar(value=False), tk.BooleanVar(value=False)
    v_spread = tk.StringVar(value="Percentual (0,04 = 0,04%)")

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    def linha_arquivo(r, rotulo, var, cmd):
        ttk.Label(frm, text=rotulo).grid(row=r, column=0, sticky="w", pady=3)
        e = ttk.Entry(frm, textvariable=var)
        e.grid(row=r, column=1, sticky="ew", padx=6)
        ttk.Button(frm, text="Selecionar…", command=cmd).grid(row=r, column=2)
        return e

    def log_txt(t):
        txt.configure(state="normal")
        txt.delete("1.0", "end")
        txt.insert("end", t)
        txt.configure(state="disabled")

    def carregar(caminho):
        v_arq.set(caminho)
        btn.state(["disabled"])
        root.config(cursor="watch")
        root.update_idletasks()
        try:
            dp = estado["de_para"]
            ident, cons, rel = processar_arquivo(caminho, dp)
            estado["cons"], estado["rel"] = cons, rel
            lbl_status.configure(text=ident.mensagem, foreground="#1E7B34")
            p = sugerir_parametros(cons)
            anos = sorted(int(x) for x in cons["ano"].unique())
            cb_ano.configure(values=anos)
            cb_comp.configure(values=anos + [anos[0] - 1])
            v_ano.set(str(p.ano_atual)); v_comp.set(str(p.ano_comparacao)); v_mes.set(MESES_NOME[p.mes_corte - 1])
            offs = sorted({o for o in cons["officer_comercial"].unique() if o})
            cb_off.configure(values=["(Todos)"] + offs)
            if not v_saida.get():
                v_saida.set(str(Path(caminho).parent / "Saida_Risco_Sacado"))
            log_txt(rel.resumo())
            btn.state(["!disabled"])
        except Exception as e:
            estado["cons"] = None
            lbl_status.configure(text=str(e), foreground="#B3261E")
            log_txt(f"ERRO: {e}")
        finally:
            root.config(cursor="")

    def ao_mudar_ano(_=None):
        cons = estado["cons"]
        if cons is None:
            return
        ano = int(v_ano.get())
        v_comp.set(str(ano - 1))
        sub = cons[cons["ano"] == ano]
        if len(sub):
            v_mes.set(MESES_NOME[int(sub["mes_numero"].max()) - 1])

    def sel_arquivo():
        c = filedialog.askopenfilename(title="Selecione o arquivo exportado",
                                       filetypes=[("Planilhas", "*.xlsx *.xlsm *.xls *.csv"), ("Todos", "*.*")])
        if c:
            carregar(c)

    def sel_saida():
        c = filedialog.askdirectory(title="Pasta de saída")
        if c:
            v_saida.set(c)

    def sel_depara():
        c = filedialog.askopenfilename(title="De/para de nomes (2 colunas: origem, destino)",
                                       filetypes=[("Planilhas", "*.xlsx *.csv")])
        if c:
            try:
                estado["de_para"] = carregar_de_para(Path(c))
                v_depara.set(c)
                if v_arq.get():
                    carregar(v_arq.get())
            except Exception as e:
                messagebox.showerror("De/para", str(e))

    def gerar():
        if estado["cons"] is None:
            return
        try:
            tol = float(v_tol.get().replace(".", "").replace(",", ".")) if "," in v_tol.get() else float(v_tol.get())
            p = Parametros(
                ano_atual=int(v_ano.get()), ano_comparacao=int(v_comp.get()), mes_corte=MESES_NOME.index(v_mes.get()) + 1,
                qtd_bloco=0 if v_todos.get() else int(v_qtd.get()), exibir_sem_comparacao=v_semcomp.get(),
                exibir_producao_zero=v_zero.get(), officer=None if v_off.get() == "(Todos)" else v_off.get(),
                cliente_filtro=v_cli.get(), tolerancia=tol, visao="gerencial" if v_visao.get() == "Gerencial" else "estrita",
                spread_unidade="percentual" if v_spread.get().startswith("Percentual") else "fracao",
                pasta_saida=Path(v_saida.get()))
            root.config(cursor="watch"); root.update_idletasks()
            a = calcular_analise(estado["cons"], p)
            out = exportar_excel(a, estado["cons"], estado["rel"], nome_saida(p, p.pasta_saida))
            root.config(cursor="")
            cont = a.clientes["status_variacao"].value_counts()
            log_txt(estado["rel"].resumo() + f"\n\nGerado: {out}\nSuperávit: {cont.get('SUPERAVIT', 0)} | Déficit: {cont.get('DEFICIT', 0)}"
                    f" | Novos: {cont.get('NOVO_CLIENTE', 0)} | Sem produção atual: {cont.get('SEM_PRODUCAO_ATUAL', 0)}")
            if messagebox.askyesno("Concluído", f"Relatório gerado:\n{out}\n\nAbrir agora?"):
                abrir_arquivo(out)
        except Exception as e:
            root.config(cursor="")
            log.exception("Falha ao gerar")
            messagebox.showerror("Erro ao gerar", str(e))

    e_arq = linha_arquivo(0, "Arquivo Excel:", v_arq, sel_arquivo)
    if dnd:
        e_arq.drop_target_register(DND_FILES)
        e_arq.dnd_bind("<<Drop>>", lambda ev: carregar(root.tk.splitlist(ev.data)[0]))
    lbl_status = ttk.Label(frm, text="Selecione ou arraste o arquivo exportado." if dnd else "Selecione o arquivo exportado.",
                           wraplength=800)
    lbl_status.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 8))

    par = ttk.LabelFrame(frm, text="Parâmetros", padding=8)
    par.grid(row=2, column=0, columnspan=3, sticky="ew")
    for i in range(8):
        par.columnconfigure(i, weight=1 if i % 2 else 0)

    def rot(r, c, t):
        ttk.Label(par, text=t).grid(row=r, column=c, sticky="e", padx=4, pady=3)

    rot(0, 0, "Ano atual"); cb_ano = ttk.Combobox(par, textvariable=v_ano, width=8, state="readonly"); cb_ano.grid(row=0, column=1, sticky="w")
    cb_ano.bind("<<ComboboxSelected>>", ao_mudar_ano)
    rot(0, 2, "Comparar com"); cb_comp = ttk.Combobox(par, textvariable=v_comp, width=8, state="readonly"); cb_comp.grid(row=0, column=3, sticky="w")
    rot(0, 4, "Mês de corte"); ttk.Combobox(par, textvariable=v_mes, values=MESES_NOME, width=12, state="readonly").grid(row=0, column=5, sticky="w")
    rot(0, 6, "Visão"); ttk.Combobox(par, textvariable=v_visao, values=["Gerencial", "Estrita"], width=10, state="readonly").grid(row=0, column=7, sticky="w")
    rot(1, 0, "Clientes por bloco"); ttk.Spinbox(par, from_=1, to=99999, textvariable=v_qtd, width=8).grid(row=1, column=1, sticky="w")
    ttk.Checkbutton(par, text="Mostrar todos", variable=v_todos).grid(row=1, column=2, columnspan=2, sticky="w")
    rot(1, 4, "Tolerância (R$)"); ttk.Entry(par, textvariable=v_tol, width=10).grid(row=1, column=5, sticky="w")
    rot(1, 6, "Spread"); ttk.Combobox(par, textvariable=v_spread, state="readonly", width=24,
                                    values=["Percentual (0,04 = 0,04%)", "Fração (0,04 = 4%)"]).grid(row=1, column=7, sticky="w")
    rot(2, 0, "Officer"); cb_off = ttk.Combobox(par, textvariable=v_off, width=34, state="readonly"); cb_off.grid(row=2, column=1, columnspan=3, sticky="w")
    rot(2, 4, "Cliente contém"); ttk.Entry(par, textvariable=v_cli, width=30).grid(row=2, column=5, columnspan=3, sticky="w")
    ttk.Checkbutton(par, text="Exibir clientes sem base comparável", variable=v_semcomp).grid(row=3, column=0, columnspan=4, sticky="w")
    ttk.Checkbutton(par, text="Exibir clientes com produção zero", variable=v_zero).grid(row=3, column=4, columnspan=4, sticky="w")

    ttk.Label(frm, text="Pasta de saída:").grid(row=3, column=0, sticky="w", pady=(8, 3))
    ttk.Entry(frm, textvariable=v_saida).grid(row=3, column=1, sticky="ew", padx=6, pady=(8, 3))
    ttk.Button(frm, text="Selecionar…", command=sel_saida).grid(row=3, column=2, pady=(8, 3))
    ttk.Label(frm, text="De/para (opcional):").grid(row=4, column=0, sticky="w", pady=3)
    ttk.Entry(frm, textvariable=v_depara).grid(row=4, column=1, sticky="ew", padx=6)
    ttk.Button(frm, text="Selecionar…", command=sel_depara).grid(row=4, column=2)

    btn = ttk.Button(frm, text="Gerar relatório (Excel)", command=gerar)
    btn.grid(row=5, column=0, columnspan=3, pady=10, sticky="ew")
    btn.state(["disabled"])
    txt = tk.Text(frm, height=14, state="disabled", font=("Consolas", 9))
    txt.grid(row=6, column=0, columnspan=3, sticky="nsew")
    frm.rowconfigure(6, weight=1)
    root.mainloop()


# ----------------------------------------------------------------------------
def main(argv=None):
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="Análise mensal de produção - Risco Sacado")
    ap.add_argument("arquivo", nargs="?")
    ap.add_argument("--ano-atual", type=int)
    ap.add_argument("--ano-comparacao", type=int)
    ap.add_argument("--mes-corte", type=int)
    ap.add_argument("--qtd", type=int, default=60, help="clientes por bloco (0 = todos)")
    ap.add_argument("--visao", choices=["gerencial", "estrita"], default="gerencial")
    ap.add_argument("--officer")
    ap.add_argument("--cliente", default="")
    ap.add_argument("--tolerancia", type=float, default=0.01)
    ap.add_argument("--sem-comparacao", action="store_true")
    ap.add_argument("--producao-zero", action="store_true")
    ap.add_argument("--spread-unidade", choices=["percentual", "fracao"], default="percentual",
                    help="percentual: 0,04 = 0,04%% | fracao: 0,04 = 4%%")
    ap.add_argument("--de-para")
    ap.add_argument("--saida")
    args = ap.parse_args(argv)
    if not args.arquivo:
        return abrir_gui()
    dp = carregar_de_para(Path(args.de_para)) if args.de_para else None
    ident, cons, rel = processar_arquivo(args.arquivo, dp)
    print(ident.mensagem)
    p = sugerir_parametros(cons)
    p.ano_atual = args.ano_atual or p.ano_atual
    p.ano_comparacao = args.ano_comparacao or (args.ano_atual - 1 if args.ano_atual else p.ano_comparacao)
    p.mes_corte = args.mes_corte or p.mes_corte
    p.qtd_bloco, p.visao, p.officer, p.cliente_filtro = args.qtd, args.visao, args.officer, args.cliente
    p.spread_unidade = args.spread_unidade
    p.tolerancia, p.exibir_sem_comparacao, p.exibir_producao_zero = args.tolerancia, args.sem_comparacao, args.producao_zero
    pasta = Path(args.saida) if args.saida else Path(args.arquivo).parent / "Saida_Risco_Sacado"
    a = calcular_analise(cons, p)
    out = exportar_excel(a, cons, rel, nome_saida(p, pasta))
    print(rel.resumo())
    print(f"\nGerado: {out}")


if __name__ == "__main__":
    main()
