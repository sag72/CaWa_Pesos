# -*- coding: utf-8 -*-
"""
Interface para acrescentar lojas nos XML do CAWA (MFT_NCR_FTPPUBLICO_SCOT_EANS).

Le o Excel da pasta e, para cada insignia (sheet CNT / CBD / MDL), acrescenta no
XML correspondente as lojas que estao na sheet e que ainda NAO existem no XML:
  - cria o bloco <unix_job> (clonado a partir de um job ja existente no proprio XML)
  - acrescenta a dependencia correspondente na lista do link "@Inicial"

E idempotente (nao duplica lojas ja presentes) e cria backup .bak antes de gravar.

Uso:
    python adicionar_lojas_xml.py            # aplica as alteracoes
    python adicionar_lojas_xml.py --dry-run  # so mostra o que faria, sem gravar
"""

import os
import re
import sys
import glob
import shutil
import argparse

import openpyxl

# Pasta onde estao o Excel e os XML (a mesma do script)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Mapeamento sheet (insignia) -> ficheiro XML
INSIGNIAS = {
    "CNT": "MFT_NCR_FTPPUBLICO_SCOT_EANS_CNT.xml",
    "CBD": "MFT_NCR_FTPPUBLICO_SCOT_EANS_CBD.xml",
    "MDL": "MFT_NCR_FTPPUBLICO_SCOT_EANS_MDL.xml",
}

# Colunas esperadas na sheet: APPL | JOB NAME | PARAMETROS
COL_APPL, COL_JOB, COL_PARAM = 0, 1, 2


def encontrar_excel():
    xlsx = [f for f in glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
            if not os.path.basename(f).startswith("~$")]
    if not xlsx:
        sys.exit("ERRO: nao encontrei nenhum .xlsx na pasta.")
    # se houver mais que um, usa o que tem 'ficha' no nome, senao o primeiro
    for f in xlsx:
        if "ficha" in os.path.basename(f).lower():
            return f
    return xlsx[0]


def ler_lojas_sheet(wb, sheet):
    """Devolve lista de (job_name, parametros) da sheet, ignorando linhas vazias."""
    ws = wb[sheet]
    lojas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) <= COL_PARAM:
            continue
        job = row[COL_JOB]
        param = row[COL_PARAM]
        if not job or not param:
            continue
        lojas.append((str(job).strip(), str(param).strip()))
    return lojas


def jobs_existentes(xml):
    return set(re.findall(r'<unix_job name="([^"]+)"', xml))


def extrair_template(xml):
    """Usa o primeiro <unix_job>...</unix_job> do proprio XML como molde."""
    m = re.search(r'<unix_job name="[^"]+" qualifier="0001">.*?</unix_job>',
                  xml, re.DOTALL)
    if not m:
        return None
    return m.group(0)


def construir_job(template, job_name, parametros):
    """Clona o template, trocando o nome do job e o conteudo de <args>."""
    novo = re.sub(r'(<unix_job name=")[^"]+(")',
                  lambda mm: mm.group(1) + job_name + mm.group(2),
                  template, count=1)
    novo = re.sub(r'<args>.*?</args>',
                  lambda mm: "<args>" + parametros + "</args>",
                  novo, count=1, flags=re.DOTALL)
    return novo


def inserir_jobs(xml, blocos):
    """Insere os novos blocos <unix_job> imediatamente antes de </appl>."""
    idx = xml.rfind("</appl>")
    return xml[:idx] + "".join(blocos) + xml[idx:]


def inserir_relconditions(xml, appl, job_names):
    """Acrescenta as dependencias das novas lojas na relconditionlist do @Inicial."""
    marca = 'name="%s@Inicial"' % appl
    pos = xml.find(marca)
    if pos == -1:
        raise RuntimeError("Nao encontrei o link @Inicial para %s" % appl)
    fim = xml.find("</relconditionlist>", pos)
    if fim == -1:
        raise RuntimeError("Nao encontrei </relconditionlist> no @Inicial de %s" % appl)
    novas = "".join(
        "<relcondition><successorname>%s.0001</successorname>"
        "<condition>NORMAL</condition></relcondition>" % jn
        for jn in job_names
    )
    return xml[:fim] + novas + xml[fim:]


def processar(sheet, xml_file, wb, dry_run):
    caminho = os.path.join(BASE_DIR, xml_file)
    if not os.path.exists(caminho):
        print("  [%s] AVISO: XML nao encontrado (%s) -- ignorado." % (sheet, xml_file))
        return
    appl = os.path.splitext(xml_file)[0]

    with open(caminho, "r", encoding="utf-8") as fh:
        xml = fh.read()

    template = extrair_template(xml)
    if not template:
        print("  [%s] ERRO: sem <unix_job> no XML para servir de molde -- ignorado." % sheet)
        return

    existentes = jobs_existentes(xml)
    lojas = ler_lojas_sheet(wb, sheet)

    novos, vistos = [], set()
    for job_name, param in lojas:
        if job_name in existentes or job_name in vistos:
            continue
        vistos.add(job_name)
        novos.append((job_name, param))

    if not novos:
        print("  [%s] nada a acrescentar (todas as %d lojas ja existem)." % (sheet, len(lojas)))
        return

    print("  [%s] %d lojas na sheet | %d ja existiam | %d a acrescentar"
          % (sheet, len(lojas), len(existentes), len(novos)))
    for jn, _ in novos:
        print("        + %s" % jn)

    if dry_run:
        return

    blocos = [construir_job(template, jn, pr) for jn, pr in novos]
    xml = inserir_jobs(xml, blocos)
    xml = inserir_relconditions(xml, appl, [jn for jn, _ in novos])

    shutil.copy2(caminho, caminho + ".bak")
    with open(caminho, "w", encoding="utf-8") as fh:
        fh.write(xml)
    print("  [%s] gravado (backup em %s.bak)" % (sheet, xml_file))


def main():
    ap = argparse.ArgumentParser(description="Acrescenta lojas do Excel nos XML do CAWA.")
    ap.add_argument("--dry-run", action="store_true",
                    help="mostra o que faria sem gravar nada")
    args = ap.parse_args()

    excel = encontrar_excel()
    print("Excel : %s" % os.path.basename(excel))
    print("Modo  : %s\n" % ("DRY-RUN (sem gravar)" if args.dry_run else "APLICAR"))

    wb = openpyxl.load_workbook(excel, data_only=True)
    for sheet, xml_file in INSIGNIAS.items():
        if sheet not in wb.sheetnames:
            print("  [%s] AVISO: sheet inexistente no Excel -- ignorada." % sheet)
            continue
        processar(sheet, xml_file, wb, args.dry_run)

    print("\nConcluido.")


if __name__ == "__main__":
    main()
