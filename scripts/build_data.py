#!/usr/bin/env python3
"""
Gera docs/data.json e public/data.json a partir de um arquivo .xlsx.

Uso:
  python3 scripts/build_data.py [caminho.xlsx]

Se nenhum caminho for informado, procura o primeiro .xlsx em data/.
Basta jogar o novo arquivo dentro de data/ e rodar o script — o
dashboard automaticamente passa a refletir os dados novos.

Estrutura esperada do .xlsx (mesmos nomes de aba/coluna do modelo atual):
  - d_comercial:     RV, CONCATENAÇÃO RV+NOME, CONCATENAÇÃO SV+NOME,
                     CONCATENAÇÃO CV+NOME, ds_uf
  - f_vendas_total:  cd_vendedor, ds_uf, vl_financeiro, Plataforma, Store Channel
  - d_metas_fin:     vd, ds_uf, vl_objetivo, Plataforma, Store Channel
"""
import json
import sys
import glob
import os
from datetime import datetime, timezone, timedelta
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_PATHS = [os.path.join(ROOT, "docs", "data.json"),
             os.path.join(ROOT, "public", "data.json")]

FARMA = {"DRUG/PHARMACY", "PERFUMERIES"}


def s(v):
    if v is None:
        return ""
    return str(v).strip()


def n(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except (ValueError, TypeError):
        return 0.0


def build(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)

    # ---------- d_comercial ----------
    ws = wb["d_comercial"]
    rows = ws.iter_rows(values_only=True)
    header = [s(h) for h in next(rows)]
    idx = {h: i for i, h in enumerate(header)}
    comercial = []
    seen = set()
    for r in rows:
        if not r:
            continue
        rv = s(r[idx["RV"]])
        uf = s(r[idx["ds_uf"]])
        if not rv or not uf:
            continue
        key = rv + "|" + uf
        if key in seen:
            continue
        seen.add(key)
        comercial.append({
            "rv": rv,
            "uf": uf,
            "rvName": s(r[idx["CONCATENAÇÃO RV + NOME"]]),
            "sv": s(r[idx["CONCATENAÇÃO SV + NOME"]]),
            "cv": s(r[idx["CONCATENAÇÃO CV + NOME"]]),
        })

    # ---------- f_vendas_total ----------
    ws = wb["f_vendas_total"]
    rows = ws.iter_rows(values_only=True)
    header = [s(h) for h in next(rows)]
    idx = {h: i for i, h in enumerate(header)}
    # busca coluna vl_faturamento tolerante a maiúsculas/espaços
    fat_key = None
    for h in header:
        if h.strip().lower().replace(" ", "_") == "vl_faturamento":
            fat_key = h
            break
    has_fat_flag = fat_key is not None
    print(f"[build_data] f_vendas_total headers: {header}", file=sys.stderr)
    print(f"[build_data] vl_faturamento column: {fat_key!r} (found={has_fat_flag})", file=sys.stderr)
    vagg = {}
    # Positivação: por (rv,uf,cnpj) somamos vl_financeiro split total/ali/far
    # Depois contamos CNPJs únicos onde a soma > 0.
    cnpj_sums = {}  # {(rv,uf): {cnpj: {"t":..., "a":..., "f":...}}}
    fat_rows = 0
    total_rows = 0
    for r in rows:
        if not r:
            continue
        rv = s(r[idx["cd_vendedor"]])
        uf = s(r[idx["ds_uf"]])
        if not rv or not uf:
            continue
        val = n(r[idx["vl_financeiro"]])
        plat = s(r[idx["Plataforma"]])
        chan = s(r[idx["Store Channel"]])
        cnpj = s(r[idx["nr_cnpj_cpf"]]) if "nr_cnpj_cpf" in idx else ""
        # aceita qualquer valor não-zero (1, "1", 1.0, "Sim", etc.) como faturado
        raw_fat = r[idx[fat_key]] if has_fat_flag else None
        if has_fat_flag:
            if isinstance(raw_fat, (int, float)):
                is_fat = raw_fat != 0
            else:
                sv = s(raw_fat).lower()
                is_fat = sv not in ("", "0", "0.0", "nao", "não", "no", "false")
        else:
            is_fat = False
        total_rows += 1
        if is_fat:
            fat_rows += 1
        k = rv + "|" + uf
        b = vagg.setdefault(k, {
            "rv": rv, "uf": uf,
            "v": 0.0, "ec": 0.0, "sp": 0.0, "ali": 0.0, "far": 0.0,
            "vf": 0.0, "vf_ec": 0.0, "vf_sp": 0.0, "vf_ali": 0.0, "vf_far": 0.0,
            "p": 0, "p_ali": 0, "p_far": 0,
            "pf": 0, "pf_ali": 0, "pf_far": 0,
        })
        b["v"] += val
        if plat == "Escolha Certa":
            b["ec"] += val
        if plat == "Store Platform":
            b["sp"] += val
        if chan in FARMA:
            b["far"] += val
        else:
            b["ali"] += val
        if is_fat:
            b["vf"] += val
            if plat == "Escolha Certa":
                b["vf_ec"] += val
            if plat == "Store Platform":
                b["vf_sp"] += val
            if chan in FARMA:
                b["vf_far"] += val
            else:
                b["vf_ali"] += val
        if cnpj:
            cm = cnpj_sums.setdefault(k, {})
            cs = cm.setdefault(cnpj, {"t": 0.0, "a": 0.0, "f": 0.0,
                                      "tf": 0.0, "af": 0.0, "ff": 0.0})
            cs["t"] += val
            if chan in FARMA:
                cs["f"] += val
            else:
                cs["a"] += val
            if is_fat:
                cs["tf"] += val
                if chan in FARMA:
                    cs["ff"] += val
                else:
                    cs["af"] += val

    # Conta CNPJs únicos com somatória > 0 por (rv,uf)
    # p*  = positivados (total) ; pf* = positivados faturados
    pos_total_all = 0
    for k, cm in cnpj_sums.items():
        b = vagg.get(k)
        if not b:
            continue
        for cnpj, cs in cm.items():
            if cs["t"] > 0:
                b["p"] += 1
                pos_total_all += 1
            if cs["a"] > 0:
                b["p_ali"] += 1
            if cs["f"] > 0:
                b["p_far"] += 1
            if cs["tf"] > 0:
                b["pf"] += 1
            if cs["af"] > 0:
                b["pf_ali"] += 1
            if cs["ff"] > 0:
                b["pf_far"] += 1

    vendas = list(vagg.values())
    print(f"[build_data] linhas: {total_rows}, faturadas: {fat_rows}, CNPJs positivados: {pos_total_all}", file=sys.stderr)



    # ---------- d_metas_fin ----------
    ws = wb["d_metas_fin"]
    rows = ws.iter_rows(values_only=True)
    header = [s(h) for h in next(rows)]
    idx = {h: i for i, h in enumerate(header)}
    magg = {}
    has_chan = "Store Channel" in idx
    for r in rows:
        if not r:
            continue
        rv = s(r[idx["vd"]])
        uf = s(r[idx["ds_uf"]])
        if not rv or not uf:
            continue
        val = n(r[idx["vl_objetivo"]])
        plat = s(r[idx["Plataforma"]])
        chan = s(r[idx["Store Channel"]]) if has_chan else ""
        k = rv + "|" + uf
        b = magg.setdefault(k, {"rv": rv, "uf": uf, "total": 0.0, "ec": 0.0, "sp": 0.0, "ali": 0.0, "far": 0.0})
        b["total"] += val
        if plat == "Escolha Certa":
            b["ec"] += val
        if plat == "Store Platform":
            b["sp"] += val
        if chan:
            if chan in FARMA:
                b["far"] += val
            else:
                b["ali"] += val
    # inicializa campos de meta de positivação (preenchidos abaixo por d_metas)
    for b in magg.values():
        b["p_total"] = 0.0
        b["p_ali"] = 0.0
        b["p_far"] = 0.0

    # ---------- d_metas (positivação) ----------
    if "d_metas" in wb.sheetnames:
        ws = wb["d_metas"]
        rows = ws.iter_rows(values_only=True)
        header = [s(h) for h in next(rows)]
        # busca colunas tolerante a espaços/case
        def find_col(name):
            target = name.strip().lower()
            for h in header:
                if h.strip().lower() == target:
                    return header.index(h)
            return None
        i_rv = find_col("RV")
        i_uf = find_col("ds_uf")
        i_tt = find_col("TT Positivação")
        i_ali = find_col("OBJ PRODUTIVIDADE HFS")
        i_far = find_col("OBJ PRODUTIVIDADE FARMA")
        print(f"[build_data] d_metas cols: RV={i_rv} ds_uf={i_uf} TT={i_tt} HFS={i_ali} FARMA={i_far}", file=sys.stderr)
        if i_rv is not None and i_uf is not None:
            for r in rows:
                if not r:
                    continue
                rv = s(r[i_rv])
                uf = s(r[i_uf])
                if not rv or not uf:
                    continue
                k = rv + "|" + uf
                b = magg.setdefault(k, {"rv": rv, "uf": uf, "total": 0.0, "ec": 0.0, "sp": 0.0, "ali": 0.0, "far": 0.0,
                                        "p_total": 0.0, "p_ali": 0.0, "p_far": 0.0})
                if i_tt is not None:
                    b["p_total"] += n(r[i_tt])
                if i_ali is not None:
                    b["p_ali"] += n(r[i_ali])
                if i_far is not None:
                    b["p_far"] += n(r[i_far])

    metas = list(magg.values())

    # timestamp em horário de Brasília
    br = timezone(timedelta(hours=-3))
    generated_at = datetime.now(br).strftime("%Y-%m-%dT%H:%M:%S-03:00")

    data = {
        "generated_at": generated_at,
        "source_file": os.path.basename(xlsx_path),
        "comercial": comercial,
        "vendas": vendas,
        "metas": metas,
    }
    return data


def main():
    if len(sys.argv) > 1:
        xlsx = sys.argv[1]
    else:
        cands = sorted(glob.glob(os.path.join(ROOT, "data", "*.xlsx")))
        if not cands:
            print("ERRO: nenhum .xlsx encontrado em data/. Passe o caminho como argumento.")
            sys.exit(1)
        xlsx = cands[-1]
    print(f"Lendo {xlsx} ...")
    data = build(xlsx)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    for p in OUT_PATHS:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            f.write(payload)
        print(f"  -> {p} ({len(payload):,} bytes)")
    print(f"OK. generated_at = {data['generated_at']}")
    print(f"comercial: {len(data['comercial'])}  vendas: {len(data['vendas'])}  metas: {len(data['metas'])}")


if __name__ == "__main__":
    main()
