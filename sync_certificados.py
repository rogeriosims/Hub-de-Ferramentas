import os
import re
import json
import time
import shutil
from pathlib import Path
from datetime import datetime, timezone
import openpyxl
from cryptography.hazmat.primitives.serialization import pkcs12

PASTA_CERTIFICADOS = Path(r"T:\Geral\Publica\CERTIFICADOS DIGITAIS\CERTIFICADOS e-CNPJ")
PASTA_CERTIFICADOS_PF = Path(r"T:\Geral\Publica\CERTIFICADOS DIGITAIS\CERTIFICADOS PF")
PLANILHA_ORIGINAL = Path(r"C:\Users\supervisor\Downloads\CERTIFICADOS VENCIMENTOS.xlsx")
CACHE_FILE = Path(r"C:\Users\supervisor\Downloads\pdf-optimizer-dashboard\certificados_scan_cache.json")

def carregar_cache():
    if CACHE_FILE.exists():
        try:
            return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def salvar_cache(cache):
    try:
        CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"Erro ao salvar cache: {e}")

def extrair_codigo_empresa(pasta_nome):
    if not pasta_nome:
        return ""
    m = re.match(r'^(\d+)\s*[-_]', pasta_nome.strip())
    if m:
        return m.group(1)
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
                    conteudo = f.read_text(encoding='utf-8', errors='ignore').strip()
                    if conteudo and len(conteudo) <= 40 and '\n' not in conteudo:
                        senhas.append(conteudo)
                    else:
                        conteudo_lat = f.read_text(encoding='latin-1', errors='ignore').strip()
                        if conteudo_lat and len(conteudo_lat) <= 40 and '\n' not in conteudo_lat:
                            senhas.append(conteudo_lat)
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

    senhas = extrair_senhas_candidatas(pfx_path.parent, pfx_path)
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
        doc = f"{doc_raw[:2]}.{doc_raw[2:5]}.{doc_raw[5:8]}/{doc_raw[8:12]}-{doc_raw[12:]}"
        razao = common_name[:cnpj_match.start()].strip()
    elif cpf_match:
        doc_raw = cpf_match.group(1)
        doc = f"{doc_raw[:3]}.{doc_raw[3:6]}.{doc_raw[6:9]}-{doc_raw[9:]}"
        razao = common_name[:cpf_match.start()].strip()

    num_empresa = extrair_codigo_empresa(pfx_path.parent.name)

    try:
        vencimento = cert.not_valid_after_utc.replace(tzinfo=None)
    except AttributeError:
        vencimento = cert.not_valid_after

    return {
        "empresa_num": num_empresa,
        "nome": razao,
        "cnpj": doc,
        "vencimento": vencimento,
        "vencimento_str": vencimento.strftime("%d/%m/%Y %H:%M"),
        "vencimento_iso": vencimento.strftime("%Y-%m-%d"),
        "senha_utilizada": senha_sucesso,
        "arquivo": pfx_path.name,
        "pasta": pfx_path.parent.name
    }, None

def formatar_doc_chave(doc_str):
    if not doc_str:
        return ""
    return re.sub(r'\D', '', str(doc_str)).strip()

def atualizar_planilha(resultados):
    if not PLANILHA_ORIGINAL.exists():
        print("Planilha original não encontrada.")
        return 0, 0

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bkp_path = PLANILHA_ORIGINAL.parent / f"CERTIFICADOS VENCIMENTOS_bkp_{timestamp}.xlsx"
    try:
        shutil.copy2(PLANILHA_ORIGINAL, bkp_path)
    except Exception as e:
        print(f"Aviso ao criar backup: {e}")

    wb = openpyxl.load_workbook(str(PLANILHA_ORIGINAL))
    sheet_name = "Planilha3" if "Planilha3" in wb.sheetnames else wb.sheetnames[0]
    sheet = wb[sheet_name]

    # Linha 3 = Cabeçalho: Col 1='Nº Empresa', Col 2='NCPJ', Col 3='Nome', Col 4='Vencimento'
    sheet.cell(row=3, column=1, value="Nº Empresa")

    mapa_existentes = {} # {doc_limpo: row_index}
    mapa_nomes = {} # {nome_upper: row_index}

    for row_idx in range(4, sheet.max_row + 1):
        cnpj_val = sheet.cell(row=row_idx, column=2).value
        nome_val = sheet.cell(row=row_idx, column=3).value

        doc_limpo = formatar_doc_chave(cnpj_val)
        if doc_limpo:
            mapa_existentes[doc_limpo] = row_idx
        if nome_val:
            mapa_nomes[str(nome_val).strip().upper()] = row_idx

    atualizados = 0
    inseridos = 0

    for r in resultados:
        doc_limpo = formatar_doc_chave(r.get("cnpj"))
        nome_upper = str(r.get("nome", "")).strip().upper()
        row_target = None

        if doc_limpo and doc_limpo in mapa_existentes:
            row_target = mapa_existentes[doc_limpo]
        elif nome_upper and nome_upper in mapa_nomes:
            row_target = mapa_nomes[nome_upper]

        novo_venc = r.get("vencimento")
        num_emp = r.get("empresa_num")

        if row_target:
            val_emp_atual = sheet.cell(row=row_target, column=1).value
            if not val_emp_atual and num_emp:
                sheet.cell(row=row_target, column=1, value=num_emp)
            
            # Atualiza apenas se vencimento novo for maior ou igual
            venc_atual = sheet.cell(row=row_target, column=4).value
            deve_atualizar = True
            if isinstance(venc_atual, datetime) and isinstance(novo_venc, datetime):
                if novo_venc < venc_atual:
                    deve_atualizar = False

            if deve_atualizar:
                if num_emp:
                    sheet.cell(row=row_target, column=1, value=num_emp)
                if r.get("cnpj"):
                    sheet.cell(row=row_target, column=2, value=r.get("cnpj"))
                if r.get("nome"):
                    sheet.cell(row=row_target, column=3, value=r.get("nome"))
                sheet.cell(row=row_target, column=4, value=novo_venc)
                atualizados += 1
        else:
            next_row = sheet.max_row + 1
            sheet.cell(row=next_row, column=1, value=num_emp)
            sheet.cell(row=next_row, column=2, value=r.get("cnpj"))
            sheet.cell(row=next_row, column=3, value=r.get("nome"))
            sheet.cell(row=next_row, column=4, value=novo_venc)
            if doc_limpo:
                mapa_existentes[doc_limpo] = next_row
            if nome_upper:
                mapa_nomes[nome_upper] = next_row
            inseridos += 1

    wb.save(str(PLANILHA_ORIGINAL))
    return atualizados, inseridos

def executar_varredura(callback_progresso=None, forcar_tudo=False):
    """
    Executa a verificação incremental de certificados na pasta de rede.
    Utiliza cache baseado em mtime e tamanho para continuar de onde parou.
    """
    if not PASTA_CERTIFICADOS.exists():
        return {"sucesso": False, "erro": f"Pasta de rede não acessível: {PASTA_CERTIFICADOS}"}

    cache = {} if forcar_tudo else carregar_cache()
    
    arquivos_pfx = list(PASTA_CERTIFICADOS.rglob("*.pfx")) + list(PASTA_CERTIFICADOS.rglob("*.p12"))
    if PASTA_CERTIFICADOS_PF.exists():
        arquivos_pfx.extend(list(PASTA_CERTIFICADOS_PF.rglob("*.pfx")) + list(PASTA_CERTIFICADOS_PF.rglob("*.p12")))

    total_arquivos = len(arquivos_pfx)
    processados_novos = 0
    pulados_cache = 0
    falhas = 0
    certificados_alterados_ou_novos = []

    for idx, pfx in enumerate(arquivos_pfx, 1):
        try:
            stat = pfx.stat()
            file_key = str(pfx.resolve())
            mtime = stat.st_mtime
            size = stat.st_size
        except Exception:
            continue

        cached_entry = cache.get(file_key)
        # Se arquivo já foi lido e não foi modificado
        if cached_entry and not forcar_tudo:
            if cached_entry.get("mtime") == mtime and cached_entry.get("size") == size:
                pulados_cache += 1
                if callback_progresso and idx % 25 == 0:
                    callback_progresso(idx, total_arquivos, f"Verificando {pfx.parent.name} (já em cache)...")
                continue

        # Ler arquivo novo ou modificado
        dados, erro = ler_dados_pfx(pfx)
        if dados:
            processados_novos += 1
            certificados_alterados_ou_novos.append(dados)
            # Salvar no cache
            cache[file_key] = {
                "mtime": mtime,
                "size": size,
                "empresa_num": dados["empresa_num"],
                "nome": dados["nome"],
                "cnpj": dados["cnpj"],
                "vencimento_iso": dados["vencimento_iso"],
                "status": "ok"
            }
        else:
            falhas += 1
            cache[file_key] = {
                "mtime": mtime,
                "size": size,
                "status": "falha",
                "erro": erro
            }

        if callback_progresso:
            callback_progresso(idx, total_arquivos, f"Processando {pfx.parent.name}...")

    # Salva cache atualizado
    salvar_cache(cache)

    # Se houver novos ou modificados, atualiza a planilha Excel
    atualizados = 0
    inseridos = 0
    if certificados_alterados_ou_novos:
        atualizados, inseridos = atualizar_planilha(certificados_alterados_ou_novos)

    return {
        "sucesso": True,
        "total_encontrados": total_arquivos,
        "lidos_novos_ou_modificados": processados_novos,
        "inalterados_em_cache": pulados_cache,
        "falhas_leitura": falhas,
        "planilha_atualizados": atualizados,
        "planilha_inseridos": inseridos
    }

if __name__ == "__main__":
    def print_prog(curr, tot, msg):
        if curr % 20 == 0 or curr == tot:
            print(f"[{curr}/{tot}] {msg}")
    res = executar_varredura(callback_progresso=print_prog)
    print("\nResultado:")
    print(json.dumps(res, indent=2, ensure_ascii=False))
