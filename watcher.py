import os
import sys
import glob
import time
import shutil
import subprocess
from pathlib import Path

# Configuração da Pasta Monitorada
WATCH_BASE_DIR = Path(r"C:\Users\supervisor\Downloads\pasta_monitorada")
WATCH_IN_DIR = WATCH_BASE_DIR / "1_Entrada_Para_Comprimir"
WATCH_OUT_DIR = WATCH_BASE_DIR / "2_Otimizados"
WATCH_ORIG_DIR = WATCH_BASE_DIR / "3_Originais_Arquivados"

for folder in [WATCH_BASE_DIR, WATCH_IN_DIR, WATCH_OUT_DIR, WATCH_ORIG_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

def find_ghostscript():
    for cmd in ["gswin64c", "gswin32c", "gs"]:
        path = shutil.which(cmd)
        if path:
            return path
    patterns = [
        r"C:\Program Files\gs\gs*\bin\gswin64c.exe",
        r"C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe",
    ]
    for pattern in patterns:
        matches = glob.glob(pattern)
        if matches:
            matches.sort(reverse=True)
            return matches[0]
    return None

GS = find_ghostscript()

def format_bytes(b):
    if b < 1024:
        return f"{b} B"
    elif b < 1024 * 1024:
        return f"{b / 1024:.1f} KB"
    else:
        return f"{b / (1024 * 1024):.2f} MB"

def main():
    if not GS:
        print("[ERRO] Ghostscript não encontrado no sistema!")
        sys.exit(1)

    print("=" * 60)
    print(" GHOSTPDF HOT FOLDER - MONITORAMENTO EM TEMPO REAL")
    print("=" * 60)
    print(f" [Motor GS]    : {GS}")
    print(f" [Monitorando] : {WATCH_IN_DIR}")
    print(f" [Otimizados]  : {WATCH_OUT_DIR}")
    print(f" [Originais]   : {WATCH_ORIG_DIR}")
    print("=" * 60)
    print(" Aguardando arquivos PDF serem colocados na pasta de entrada...")
    print(" Pressione Ctrl+C para encerrar.\n")

    while True:
        try:
            pdf_files = list(WATCH_IN_DIR.glob("*.pdf")) + list(WATCH_IN_DIR.glob("*.PDF"))
            for pdf_file in pdf_files:
                if pdf_file.name.startswith("~") or pdf_file.name.endswith(".tmp") or pdf_file.name.endswith(".crdownload"):
                    continue

                try:
                    sz1 = pdf_file.stat().st_size
                    time.sleep(1)
                    if not pdf_file.exists() or pdf_file.stat().st_size != sz1:
                        continue
                    with open(pdf_file, "rb"):
                        pass
                except (IOError, OSError):
                    continue

                orig_size = pdf_file.stat().st_size
                base_name = pdf_file.stem
                out_filename = f"{base_name}_otimizado.pdf"
                target_out = WATCH_OUT_DIR / out_filename
                temp_out = WATCH_OUT_DIR / f"temp_{out_filename}"

                print(f"[{time.strftime('%H:%M:%S')}] Detectado: {pdf_file.name} ({format_bytes(orig_size)})")
                print(" -> Otimizando com Ghostscript...")

                cmd = [
                    GS,
                    "-sDEVICE=pdfwrite",
                    "-dCompatibilityLevel=1.4",
                    "-dPDFSETTINGS=/ebook",
                    "-dNOPAUSE",
                    "-dQUIET",
                    "-dBATCH",
                    f"-sOutputFile={temp_out}",
                    str(pdf_file)
                ]

                start_t = time.time()
                res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180, check=False)
                dur = round(time.time() - start_t, 2)

                if res.returncode == 0 and temp_out.exists() and temp_out.stat().st_size > 0:
                    new_size = temp_out.stat().st_size
                    if new_size >= orig_size:
                        shutil.copy2(str(pdf_file), str(target_out))
                        temp_out.unlink(missing_ok=True)
                        new_size = orig_size
                        pct = 0.0
                        print(f" -> Concluído ({dur}s): Arquivo já ideal ({format_bytes(new_size)}).")
                    else:
                        shutil.move(str(temp_out), str(target_out))
                        pct = round((1 - (new_size / orig_size)) * 100, 1)
                        saved = orig_size - new_size
                        print(f" -> Sucesso ({dur}s): {format_bytes(orig_size)} -> {format_bytes(new_size)} (-{pct}% / Economizou {format_bytes(saved)})")

                    # Arquivar original
                    dest_orig = WATCH_ORIG_DIR / pdf_file.name
                    if dest_orig.exists():
                        dest_orig = WATCH_ORIG_DIR / f"{base_name}_{int(time.time())}.pdf"
                    shutil.move(str(pdf_file), str(dest_orig))
                    print(f" -> Entregue em: {target_out.name}\n")
                else:
                    temp_out.unlink(missing_ok=True)
                    print(f" -> [FALHA] Não foi possível otimizar {pdf_file.name}\n")

        except KeyboardInterrupt:
            print("\nMonitoramento encerrado pelo usuário.")
            break
        except Exception as e:
            print(f"[Aviso] Erro no ciclo de monitoramento: {e}")
        time.sleep(2)

if __name__ == "__main__":
    main()
