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
    has_fat_flag = "vl_faturamento" in idx
    vagg = {}
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
        is_fat = int(n(r[idx["vl_faturamento"]])) == 1 if has_fat_flag else False
        k = rv + "|" + uf
        b = vagg.setdefault(k, {
            "rv": rv, "uf": uf,
            "v": 0.0, "ec": 0.0, "sp": 0.0, "ali": 0.0, "far": 0.0,
            "vf": 0.0, "vf_ec": 0.0, "vf_sp": 0.0, "vf_ali": 0.0, "vf_far": 0.0,
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
    vendas = list(vagg.values())


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
