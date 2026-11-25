import os
import re
import shutil
from pathlib import Path
from datetime import datetime, timezone
import openpyxl
from cryptography.hazmat.primitives.serialization import pkcs12

PASTA_CERTIFICADOS = Path(r"T:\Geral\Publica\CERTIFICADOS DIGITAIS\CERTIFICADOS e-CNPJ")
PASTA_CERTIFICADOS_PF = Path(r"T:\Geral\Publica\CERTIFICADOS DIGITAIS\CERTIFICADOS PF")
PLANILHA_ORIGINAL = Path(r"C:\Users\supervisor\Downloads\CERTIFICADOS VENCIMENTOS.xlsx")

def extrair_codigo_empresa(pasta_nome):
    """
    Extrai o código numérico da empresa a partir do nome da pasta.
    Exemplos:
      '001 - AGECON' -> '001'
      '010 - ECP ENGENHARIA' -> '010'
      '3304 - CODAIBA FILIAL 06' -> '3304'
      '746 - RECANTO DAS BRAUNAS' -> '746'
    """
    if not pasta_nome:
        return ""
    m = re.match(r'^(\d+)\s*[-_]', pasta_nome.strip())
    if m:
        return m.group(1)
    # Se começar apenas com números
    m2 = re.match(r'^(\d+)', pasta_nome.strip())
    if m2:
        return m2.group(1)
    return ""

def extrair_senhas_candidatas(pasta, pfx_path):
    senhas = []
    nome_pfx = pfx_path.stem
    if ' - ' in nome_pfx:
        primeira_parte = nome_pfx.split(' - ')[0].strip()
        if primeira_parte and len(primeira_parte) <= 30:
            senhas.append(primeira_parte)
    elif nome_pfx.startswith('@'):
        m = re.match(r'(@[a-zA-Z0-9_\-]+)', nome_pfx)
        if m:
            senhas.append(m.group(1))

    try:
        for f in pasta.iterdir():
            if not f.is_file() or f == pfx_path or f.suffix.lower() in ['.pfx', '.p12', '.url', '.lnk', '.ini']:
                continue
            nome_sem_ext = f.stem.strip()
            if nome_sem_ext and len(nome_sem_ext) <= 40:
                senhas.append(nome_sem_ext)
            if not f.suffix and len(f.name.strip()) <= 40:
                senhas.append(f.name.strip())
            try:
                if f.stat().st_size <= 200:
                    c1 = f.read_text(encoding='utf-8', errors='ignore').strip()
                    if c1 and len(c1) <= 40 and '\n' not in c1:
                        senhas.append(c1)
                    else:
                        c2 = f.read_text(encoding='latin-1', errors='ignore').strip()
                        if c2 and len(c2) <= 40 and '\n' not in c2:
                            senhas.append(c2)
            except Exception:
                pass
    except Exception:
        pass

    senhas.extend(['1234', '123456', '12345678', '123456789', 'admin', 'cert', ''])
    vistas = set()
    unicas = []
    for s in senhas:
        if s not in vistas:
            vistas.add(s)
            unicas.append(s)
    return unicas

def ler_dados_pfx(pfx_path):
    try:
        data = pfx_path.read_bytes()
    except Exception as e:
        return None, f"Erro ao ler arquivo: {e}"

    pasta = pfx_path.parent
    cod_empresa = extrair_codigo_empresa(pasta.name)
    senhas = extrair_senhas_candidatas(pasta, pfx_path)
    cert = None
    senha_sucesso = None

    for pwd in senhas:
        for enc in ['utf-8', 'latin-1']:
            try:
                pwd_bytes = pwd.encode(enc) if pwd else None
                _, cert_obj, _ = pkcs12.load_key_and_certificates(data, pwd_bytes)
                if cert_obj:
                    cert = cert_obj
                    senha_sucesso = pwd
                    break
            except Exception:
                pass
        if cert:
            break

    if not cert:
        return None, "Senha não encontrada ou arquivo corrompido"

    common_name = ""
    for attr in cert.subject:
        if attr.oid.dotted_string == "2.5.4.3":
            common_name = str(attr.value)
            break

    razao = common_name
    doc = ""
    cnpj_match = re.search(r':(\d{14})$', common_name)
    cpf_match = re.search(r':(\d{11})$', common_name)

    if cnpj_match:
        doc_raw = cnpj_match.group(1)
        doc = f"{doc_raw[:2]}.{doc_raw[2:5]}.{doc_raw[5:8]}/{doc_raw[8:12]}-{doc_raw[12:14]}"
        razao = common_name[:cnpj_match.start()].strip()
    elif cpf_match:
        doc_raw = cpf_match.group(1)
        doc = f"{doc_raw[:3]}.{doc_raw[3:6]}.{doc_raw[6:9]}-{doc_raw[9:11]}"
        razao = common_name[:cpf_match.start()].strip()
    else:
        m_doc = re.search(r'(\d{14})', pfx_path.name)
        if m_doc:
            doc_raw = m_doc.group(1)
            doc = f"{doc_raw[:2]}.{doc_raw[2:5]}.{doc_raw[5:8]}/{doc_raw[8:12]}-{doc_raw[12:14]}"

    vencimento = cert.not_valid_after_utc if hasattr(cert, 'not_valid_after_utc') else cert.not_valid_after
    if vencimento.tzinfo is not None:
        vencimento = vencimento.astimezone(timezone.utc).replace(tzinfo=None)

    serial = str(cert.serial_number)

    return {
        "arquivo": pfx_path.name,
        "pasta": pasta.name,
        "codigo_empresa": cod_empresa,
        "nome": razao,
        "cnpj": doc,
        "vencimento": vencimento,
        "serial": serial
    }, None

def formatar_doc_chave(doc_str):
    if not doc_str:
        return ""
    return re.sub(r'\D', '', str(doc_str)).strip()

def atualizar_planilha_com_codigo(resultados):
    if not PLANILHA_ORIGINAL.exists():
        print("Planilha original não encontrada.")
        return

    wb = openpyxl.load_workbook(str(PLANILHA_ORIGINAL))
    sheet_name = "Planilha3" if "Planilha3" in wb.sheetnames else wb.sheetnames[0]
    sheet = wb[sheet_name]

    # Atualiza o cabeçalho da Coluna 1: de 'Série' para 'Nº Empresa'
    sheet.cell(row=3, column=1, value="Nº Empresa")

    # Mapeia linhas existentes pelo CNPJ e pelo Nome
    mapa_existentes = {} # {doc_limpo: row_idx}
    mapa_nomes = {}      # {nome_upper: row_idx}

    for row_idx in range(4, sheet.max_row + 1):
        doc_val = sheet.cell(row=row_idx, column=2).value
        nome_val = sheet.cell(row=row_idx, column=3).value
        
        doc_limpo = formatar_doc_chave(doc_val)
        if doc_limpo:
            mapa_existentes[doc_limpo] = row_idx
        if nome_val:
            mapa_nomes[str(nome_val).strip().upper()] = row_idx

    # Atualiza ou insere com o código da empresa
    atualizados = 0
    inseridos = 0

    for r in resultados:
        doc_limpo = formatar_doc_chave(r.get("cnpj"))
        nome_upper = str(r.get("nome", "")).strip().upper()
        cod = r.get("codigo_empresa", "")
        row_target = None

        if doc_limpo and doc_limpo in mapa_existentes:
            row_target = mapa_existentes[doc_limpo]
        elif nome_upper and nome_upper in mapa_nomes:
            row_target = mapa_nomes[nome_upper]

        if row_target:
            # Se a linha já tem código existente (ex: 3304, 746), preserva caso a pasta não tenha número
            cod_atual = sheet.cell(row=row_target, column=1).value
            novo_cod = cod if cod else cod_atual

            sheet.cell(row=row_target, column=1, value=novo_cod)
            sheet.cell(row=row_target, column=2, value=r.get("cnpj") or sheet.cell(row=row_target, column=2).value)
            sheet.cell(row=row_target, column=3, value=r.get("nome") or sheet.cell(row=row_target, column=3).value)
            sheet.cell(row=row_target, column=4, value=r.get("vencimento"))
            atualizados += 1
        else:
            next_row = sheet.max_row + 1
            sheet.cell(row=next_row, column=1, value=cod)
            sheet.cell(row=next_row, column=2, value=r.get("cnpj"))
            sheet.cell(row=next_row, column=3, value=r.get("nome"))
            sheet.cell(row=next_row, column=4, value=r.get("vencimento"))
            if doc_limpo:
                mapa_existentes[doc_limpo] = next_row
            if nome_upper:
                mapa_nomes[nome_upper] = next_row
            inseridos += 1

    wb.save(str(PLANILHA_ORIGINAL))
    print(f"Planilha atualizada com 'Nº Empresa' com sucesso!")
    print(f"Total atualizados: {atualizados}, novos: {inseridos}")

def executar():
    arquivos_pfx = list(PASTA_CERTIFICADOS.rglob("*.pfx")) + list(PASTA_CERTIFICADOS.rglob("*.p12"))
    if PASTA_CERTIFICADOS_PF.exists():
        arquivos_pfx.extend(list(PASTA_CERTIFICADOS_PF.rglob("*.pfx")) + list(PASTA_CERTIFICADOS_PF.rglob("*.p12")))
    
    print(f"Processando {len(arquivos_pfx)} arquivos...")
    resultados = []
    for pfx in arquivos_pfx:
        d, err = ler_dados_pfx(pfx)
        if d:
            resultados.append(d)

    print(f"Sucesso em {len(resultados)} certificados.")
    atualizar_planilha_com_codigo(resultados)

if __name__ == "__main__":
    executar()
