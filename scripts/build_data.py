#!/usr/bin/env python3
"""
Gera docs/data.json e public/data.json a partir de dois arquivos .xlsx em data/:

  - Dados.xlsx      -> aba f_venda_total (dados de venda)
  - Estrutura.xlsx  -> abas d_comercial (estrutura) e d_metas (metas)

Relacionamento: RV + ds_uf (comercial) <-> cd_vendedor + ds_uf (vendas)
                                       <-> RV + Uf (metas)
"""
import json
import sys
import os
import subprocess
from urllib.request import urlopen
from urllib.error import HTTPError
from datetime import datetime, timezone, timedelta
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")
OUT_PATHS = [os.path.join(ROOT, "docs", "data.json"),
             os.path.join(ROOT, "public", "data.json")]

# EANs válidos para o card "Positivação Always Noturno"
ALWAYS_EANS = {
    "7506309805498", "7500435214650", "7500435214667", "7500435248334",
    "7506339326055", "7506339394603", "7500435190640", "7506339326031",
    "7500435265263", "7500435233446", "7506339325263", "7500435190657",
    "7506339394535", "7506339325249",
}



def s(v):
    return "" if v is None else str(v).strip()


def n(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except (ValueError, TypeError):
        return 0.0


def rv_key(v):
    """Normaliza código de vendedor (448 / '448' / 448.0 -> '448')."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    t = str(v).strip()
    if t.endswith(".0"):
        t = t[:-2]
    return t


def header_map(ws):
    rows = ws.iter_rows(values_only=True)
    header = [s(h) for h in next(rows)]
    idx = {}
    for i, h in enumerate(header):
        if h and h not in idx:
            idx[h] = i
        low = h.strip().lower()
        idx.setdefault("~" + low, i)
    return rows, header, idx


def col(idx, *names):
    for nm in names:
        if nm in idx:
            return idx[nm]
        k = "~" + nm.strip().lower()
        if k in idx:
            return idx[k]
    return None


def build(dados_path, estrutura_path):
    # ---------- Estrutura: d_comercial ----------
    wbe = openpyxl.load_workbook(estrutura_path, read_only=True, data_only=True)
    ws = wbe["d_comercial"]
    rows, header, idx = header_map(ws)
    i_rv = col(idx, "RV")
    i_uf = col(idx, "ds_uf")
    i_rvn = col(idx, "CONCATENAÇÃO RV + NOME")
    i_svn = col(idx, "CONCATENAÇÃO SV + NOME")
    i_cvn = col(idx, "CONCATENAÇÃO CV + NOME")
    comercial, seen = [], set()
    for r in rows:
        if not r:
            continue
        rv = rv_key(r[i_rv])
        uf = s(r[i_uf])
        if not rv or not uf:
            continue
        k = rv + "|" + uf
        if k in seen:
            continue
        seen.add(k)
        comercial.append({
            "rv": rv, "uf": uf,
            "rvName": s(r[i_rvn]),
            "sv": s(r[i_svn]),
            "cv": s(r[i_cvn]),
        })

    # ---------- Estrutura: d_metas ----------
    ws = wbe["d_metas"]
    rows, header, idx = header_map(ws)
    m_rv = col(idx, "RV")
    m_uf = col(idx, "Uf", "ds_uf")
    m_tot = col(idx, "Meta Financeira Total")
    m_ec = col(idx, "Meta Financeira Escolha Certa")
    m_sp = col(idx, "Meta Financeira Store Platform")
    m_ali = col(idx, "Meta Financeira Alimentar")
    m_far = col(idx, "Meta Financeira Farma")
    p_tot = col(idx, "Objetivo Positivação Total")
    p_ali = col(idx, "Objetivo Positivação Alimentar")
    p_far = col(idx, "Objetivo Positivação Farma")
    r_hfs = col(idx, "OBJ PRODUTIVIDADE HFS")
    r_far = col(idx, "OBJ PRODUTIVIDADE FARMA", "OBJ PRODUTIVIDADE FARMA ")
    r_alw = col(idx, "Objetivo Marca 1")
    r_pmp = col(idx, "Objetivo Marca 2")
    e_ch2 = col(idx, "OBJ ESCOLHA CERTA_NOVO")
    e_pp = col(idx, "Objetivo Platinum Points")
    magg = {}
    for r in rows:
        if not r:
            continue
        rv = rv_key(r[m_rv])
        uf = s(r[m_uf])
        if not rv or not uf:
            continue
        k = rv + "|" + uf
        b = magg.setdefault(k, {"rv": rv, "uf": uf, "total": 0.0, "ec": 0.0,
                                "sp": 0.0, "ali": 0.0, "far": 0.0,
                                "p_total": 0.0, "p_ali": 0.0, "p_far": 0.0,
                                "r_hfs": 0.0, "r_far": 0.0,
                                "r_alw": 0.0, "r_pmp": 0.0,
                                "e_ch2": 0.0, "e_pp": 0.0})
        b["total"] += n(r[m_tot]) if m_tot is not None else 0.0
        b["ec"] += n(r[m_ec]) if m_ec is not None else 0.0
        b["sp"] += n(r[m_sp]) if m_sp is not None else 0.0
        b["ali"] += n(r[m_ali]) if m_ali is not None else 0.0
        b["far"] += n(r[m_far]) if m_far is not None else 0.0
        b["p_total"] += n(r[p_tot]) if p_tot is not None else 0.0
        b["p_ali"] += n(r[p_ali]) if p_ali is not None else 0.0
        b["p_far"] += n(r[p_far]) if p_far is not None else 0.0
        b["r_hfs"] += n(r[r_hfs]) if r_hfs is not None else 0.0
        b["r_far"] += n(r[r_far]) if r_far is not None else 0.0
        b["r_alw"] += n(r[r_alw]) if r_alw is not None else 0.0
        b["r_pmp"] += n(r[r_pmp]) if r_pmp is not None else 0.0
        b["e_ch2"] += n(r[e_ch2]) if e_ch2 is not None else 0.0
        b["e_pp"] += n(r[e_pp]) if e_pp is not None else 0.0
    metas = list(magg.values())

    # ---------- Estrutura: d_clientes_braveo (potencial) ----------
    # Contagem distinta de CNPJ por rv|uf, aplicando as mesmas regras de
    # negócio (Canal Ranking / Plataforma) de cada card de ranking.
    potencial = {}
    if "d_clientes_braveo" in wbe.sheetnames:
        ws = wbe["d_clientes_braveo"]
        rows, header, idx = header_map(ws)
        c_rv = col(idx, "cd_vendedor")
        c_uf = col(idx, "ds_uf")
        c_cnpj = col(idx, "nr_cnpj_cpf")
        c_can = col(idx, "Canal Ranking")
        c_plat = col(idx, "Plataforma")
        psets = {}
        for r in rows:
            if not r:
                continue
            rv = rv_key(r[c_rv])
            uf = s(r[c_uf])
            cnpj = s(r[c_cnpj]) if c_cnpj is not None else ""
            if not rv or not uf or not cnpj:
                continue
            canal = s(r[c_can]) if c_can is not None else ""
            plat = s(r[c_plat]) if c_plat is not None else ""
            k = rv + "|" + uf
            b = psets.setdefault(k, {"hfs": set(), "far": set(),
                                     "alw": set(), "pmp": set(), "ecp": set()})
            plat_ok = plat in ("Escolha Certa", "Store Platform")
            if canal == "HFS" and plat_ok:
                b["hfs"].add(cnpj)
            if canal in ("Farma Indep", "Farma Rede") and plat_ok:
                b["far"].add(cnpj)
            if canal in ("HFS", "Farma Indep") and plat == "Escolha Certa":
                b["alw"].add(cnpj)
            if canal in ("HFS", "Farma Indep") and plat_ok:
                b["pmp"].add(cnpj)
            if plat == "Escolha Certa":
                b["ecp"].add(cnpj)
        for k, b in psets.items():
            rv, uf = k.split("|", 1)
            potencial[k] = {"rv": rv, "uf": uf,
                            "hfs": len(b["hfs"]), "far": len(b["far"]),
                            "alw": len(b["alw"]), "pmp": len(b["pmp"]),
                            "ecp": len(b["ecp"])}


    # ---------- Dados: f_venda_total ----------
    wbd = openpyxl.load_workbook(dados_path, read_only=True, data_only=True)
    sheet = "f_venda_total" if "f_venda_total" in wbd.sheetnames else wbd.sheetnames[0]
    ws = wbd[sheet]
    rows, header, idx = header_map(ws)
    print(f"[build_data] {sheet} headers: {header}", file=sys.stderr)
    v_rv = col(idx, "cd_vendedor")
    v_uf = col(idx, "ds_uf")
    v_val = col(idx, "vl_financeiro")
    v_plat = col(idx, "Plataforma")
    v_can = col(idx, "Canal")
    v_fat = col(idx, "vl_faturamento")
    v_cnpj = col(idx, "nr_cnpj_cpf")
    v_ger = col(idx, "cd_gerente")
    v_sup = col(idx, "cd_vendedor_superior")
    v_crank = col(idx, "Canal Ranking")
    v_ean = col(idx, "ds_ean")
    v_grupo = col(idx, "nm_grupo")
    v_prod = col(idx, "nm_produto")

    vagg = {}
    cnpj_sums = {}
    rank_sums = {}
    total_rows = fat_rows = 0

    for r in rows:
        if not r:
            continue
        rv = rv_key(r[v_rv])
        uf = s(r[v_uf])
        if not rv or not uf:
            continue
        val = n(r[v_val])
        plat = s(r[v_plat]) if v_plat is not None else ""
        canal = s(r[v_can]).lower() if v_can is not None else ""
        is_far = canal == "farma"
        raw = r[v_fat] if v_fat is not None else None
        if isinstance(raw, (int, float)):
            is_fat = raw != 0
        else:
            sv = s(raw).lower()
            is_fat = sv not in ("", "0", "0.0", "nao", "não", "no", "false")
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
        if v_sup is not None and not b.get("cdSv"):
            b["cdSv"] = rv_key(r[v_sup])
        if v_ger is not None and not b.get("cdCv"):
            b["cdCv"] = rv_key(r[v_ger])
        b["v"] += val
        if plat == "Escolha Certa":
            b["ec"] += val
        elif plat == "Store Platform":
            b["sp"] += val
        if is_far:
            b["far"] += val
        else:
            b["ali"] += val
        if is_fat:
            b["vf"] += val
            if plat == "Escolha Certa":
                b["vf_ec"] += val
            elif plat == "Store Platform":
                b["vf_sp"] += val
            if is_far:
                b["vf_far"] += val
            else:
                b["vf_ali"] += val
        cnpj = s(r[v_cnpj]) if v_cnpj is not None else ""
        if cnpj:
            cs = cnpj_sums.setdefault(k, {}).setdefault(
                cnpj, {"t": 0.0, "a": 0.0, "f": 0.0, "tf": 0.0, "af": 0.0, "ff": 0.0})
            cs["t"] += val
            cs["f" if is_far else "a"] += val
            if is_fat:
                cs["tf"] += val
                cs["ff" if is_far else "af"] += val

            # ----- Ranking (Canal Ranking / Plataforma / EAN / grupo) -----
            crank = s(r[v_crank]) if v_crank is not None else ""
            plat_ok = plat in ("Escolha Certa", "Store Platform")
            ean = rv_key(r[v_ean]) if v_ean is not None else ""
            grupo = s(r[v_grupo]).upper() if v_grupo is not None else ""
            prod = s(r[v_prod]).upper() if v_prod is not None else ""
            rb = rank_sums.setdefault(k, {})

            def add_rank(metric):
                cc = rb.setdefault(metric, {}).setdefault(cnpj, {"t": 0.0, "f": 0.0})
                cc["t"] += val
                if is_fat:
                    cc["f"] += val

            if crank == "HFS" and plat_ok:
                add_rank("hfs")
            if crank in ("Farma Indep", "Farma Rede") and plat_ok:
                add_rank("far")
            if crank in ("HFS", "Farma Indep") and plat == "Escolha Certa" \
                    and ean in ALWAYS_EANS:
                add_rank("alw")
            if crank in ("HFS", "Farma Indep") and plat_ok \
                    and ("PAMPERS" in grupo or "FRALDA" in grupo) \
                    and "TOALHAS" not in prod:
                add_rank("pmp")
            if plat == "Escolha Certa":
                add_rank("ecp")


    pos_total_all = 0
    for k, cm in cnpj_sums.items():
        b = vagg.get(k)
        if not b:
            continue
        for cs in cm.values():
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

    for k, mm in rank_sums.items():
        b = vagg.get(k)
        if not b:
            continue
        for metric, cm in mm.items():
            tot = sum(1 for cs in cm.values() if cs["t"] > 0)
            fat = sum(1 for cs in cm.values() if cs["f"] > 0)
            b["r_" + metric] = tot
            b["rf_" + metric] = fat

    vendas = list(vagg.values())
    print(f"[build_data] linhas: {total_rows}, faturadas: {fat_rows}, "
          f"CNPJs positivados: {pos_total_all}", file=sys.stderr)

    # ---------- Dados: f_ec_oniz (Escolha Certa) ----------
    # chaves >= 2: contagem distinta de nr_doc com nr_chave >= 2
    # platinum points: contagem distinta de (nr_doc, ds_combo_sku_lista_ativacao)
    #                  com "Platinum Point?" = "Sim"
    ec = []
    if "f_ec_oniz" in wbd.sheetnames:
        ws = wbd["f_ec_oniz"]
        rows, header, idx = header_map(ws)
        o_rv = col(idx, "cd_vendedor")
        o_uf = col(idx, "ds_sigla", "ds_uf")
        o_doc = col(idx, "nr_doc")
        o_ch = col(idx, "nr_chave")
        o_combo = col(idx, "ds_combo_sku_lista_ativacao")
        o_pp = col(idx, "Platinum Point?")
        esets = {}
        for r in rows:
            if not r:
                continue
            rv = rv_key(r[o_rv])
            uf = s(r[o_uf])
            doc = rv_key(r[o_doc]) if o_doc is not None else ""
            if not rv or not uf or not doc:
                continue
            k = rv + "|" + uf
            b = esets.setdefault(k, {"ch2": set(), "pp": set()})
            if o_ch is not None and n(r[o_ch]) >= 2:
                b["ch2"].add(doc)
            if o_pp is not None and s(r[o_pp]).lower() in ("sim", "s", "yes", "1"):
                combo = s(r[o_combo]) if o_combo is not None else ""
                b["pp"].add(doc + "|" + combo)
        for k, b in esets.items():
            rv, uf = k.split("|", 1)
            ec.append({"rv": rv, "uf": uf,
                       "ch2": len(b["ch2"]), "pp": len(b["pp"])})
        print(f"[build_data] f_ec_oniz: chaves>=2 = "
              f"{sum(e['ch2'] for e in ec)}, platinum = "
              f"{sum(e['pp'] for e in ec)}", file=sys.stderr)

    br = timezone(timedelta(hours=-3))
    return {
        "generated_at": datetime.now(br).strftime("%Y-%m-%dT%H:%M:%S-03:00"),
        "source_file": os.path.basename(dados_path),
        "comercial": comercial,
        "vendas": vendas,
        "metas": metas,
        "potencial": list(potencial.values()),
        "ec": ec,
    }




def find(name):
    p = os.path.join(DATA_DIR, name)
    return p if os.path.exists(p) else None


def build_orfaos(data):
    """Combinações rv|uf presentes na base mas ausentes da d_comercial."""
    est = {c["rv"] + "|" + c["uf"] for c in data["comercial"]}
    itens = []
    for x in data["vendas"]:
        k = x["rv"] + "|" + x["uf"]
        if k in est:
            continue
        itens.append({
            "chave": k, "rv": x["rv"], "uf": x["uf"], "origem": "vendas",
            "vl_financeiro": round(x.get("v", 0.0), 2),
            "vl_faturado": round(x.get("vf", 0.0), 2),
            "positivados": x.get("p", 0),
            "cdSv": x.get("cdSv", ""), "cdCv": x.get("cdCv", ""),
        })
    vkeys = {x["rv"] + "|" + x["uf"] for x in data["vendas"]}
    for m in data["metas"]:
        k = m["rv"] + "|" + m["uf"]
        if k in est or k in vkeys:
            continue
        itens.append({
            "chave": k, "rv": m["rv"], "uf": m["uf"], "origem": "metas",
            "vl_financeiro": 0.0, "vl_faturado": 0.0, "positivados": 0,
            "cdSv": "", "cdCv": "",
            "meta_financeira": round(m.get("total", 0.0), 2),
            "meta_positivacao": round(m.get("p_total", 0.0), 2),
        })
    itens.sort(key=lambda i: -i.get("vl_financeiro", 0.0))
    return itens


def code_of(label):
    """Extrai o código de rótulos como '11 - NOME' ou '211-NOME'."""
    t = s(label)
    return rv_key(t.split("-")[0]) if "-" in t else rv_key(t)


def completar_estrutura(data, itens):
    """Adiciona à d_comercial as combinações órfãs, reaproveitando nomes
    já existentes de supervisor/gerente quando o código for conhecido."""
    known = {}
    for c in data["comercial"]:
        for field in ("sv", "cv"):
            lbl = c.get(field) or ""
            cd = code_of(lbl)
            if cd and "-" in lbl and lbl.split("-", 1)[1].strip() not in ("", "-"):
                known.setdefault(cd, lbl)

    def label(cd):
        if not cd:
            return "-"
        return known.get(cd, f"{cd} - -")

    for i in itens:
        data["comercial"].append({
            "rv": i["rv"], "uf": i["uf"],
            "rvName": f"{i['rv']} - -",
            "sv": label(i.get("cdSv")),
            "cv": label(i.get("cdCv")),
        })


def csv_escape(v):
    t = "" if v is None else str(v)
    return '"' + t.replace('"', '""') + '"' if (";" in t or '"' in t or "\n" in t) else t


def write_csv(name, rows, header):
    lines = [";".join(header)]
    for r in rows:
        lines.append(";".join(csv_escape(c) for c in r))
    payload = "\ufeff" + "\r\n".join(lines) + "\r\n"
    for base in ("docs", "public"):
        p = os.path.join(ROOT, base, name)
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            f.write(payload)
        print(f"  -> {p} ({len(payload):,} bytes)")


def write_all(name, obj, indent=None):
    payload = json.dumps(obj, ensure_ascii=False,
                         separators=(",", ":") if indent is None else None,
                         indent=indent)
    for base in ("docs", "public"):
        p = os.path.join(ROOT, base, name)
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            f.write(payload)
        print(f"  -> {p} ({len(payload):,} bytes)")


def main():
    dados = sys.argv[1] if len(sys.argv) > 1 else find("Dados.xlsx")
    estrutura = sys.argv[2] if len(sys.argv) > 2 else find("Estrutura.xlsx")
    if not dados or not estrutura:
        print("ERRO: coloque Dados.xlsx e Estrutura.xlsx na pasta data/.")
        sys.exit(1)
    print(f"Lendo {dados} + {estrutura} ...")
    data = build(dados, estrutura)
    orfaos = build_orfaos(data)
    completar_estrutura(data, orfaos)
    write_all("data.json", data)
    write_csv("sem-estrutura.csv",
              [[i["rv"], i["uf"], "-", i.get("cdSv", ""), i.get("cdCv", ""), i["origem"],
                f"{i.get('vl_financeiro', 0.0):.2f}".replace(".", ","),
                f"{i.get('vl_faturado', 0.0):.2f}".replace(".", ","),
                i.get("positivados", 0)] for i in orfaos],
              ["cd_vendedor", "ds_uf", "nome", "cd_vendedor_superior", "cd_gerente", "origem",
               "vl_financeiro", "vl_faturado", "positivados"])
    print(f"OK. generated_at = {data['generated_at']}")
    print(f"comercial: {len(data['comercial'])}  vendas: {len(data['vendas'])}  metas: {len(data['metas'])}")
    print(f"sem estrutura (adicionadas): {len(orfaos)} combinações, "
          f"R$ {sum(i.get('vl_financeiro', 0.0) for i in orfaos):,.2f}")




if __name__ == "__main__":
    main()
