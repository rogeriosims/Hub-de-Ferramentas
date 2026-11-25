import os
import sys
import glob
import time
import uuid
import shutil
import zipfile
import tempfile
import threading
import subprocess
from io import BytesIO
import openpyxl
from datetime import datetime, timezone
from pathlib import Path
from PIL import Image, ImageOps
import pytesseract
from pypdf import PdfReader, PdfWriter
from flask import Flask, render_template, request, jsonify, send_file

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # Max 500 MB per request

# Pastas de armazenamento temporário
STORAGE_DIR = Path(tempfile.gettempdir()) / "ghostpdf_dashboard"
STORAGE_DIR.mkdir(parents=True, exist_ok=True)

# Registro de arquivos em memória {file_id: {"path": ..., "orig_name": ..., "created_at": ...}}
FILE_REGISTRY = {}

# Configuração da Pasta Monitorada (Feature 7)
WATCH_BASE_DIR = Path(r"C:\Users\supervisor\Downloads\pasta_monitorada")
WATCH_IN_DIR = WATCH_BASE_DIR / "1_Entrada_Para_Comprimir"
WATCH_OUT_DIR = WATCH_BASE_DIR / "2_Otimizados"
WATCH_ORIG_DIR = WATCH_BASE_DIR / "3_Originais_Arquivados"

for folder in [WATCH_BASE_DIR, WATCH_IN_DIR, WATCH_OUT_DIR, WATCH_ORIG_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

WATCHER_STATE = {
    "enabled": True,
    "processed_count": 0,
    "saved_bytes_total": 0,
    "logs": []  # List of { timestamp, filename, orig_size, new_size, reduction, status }
}

# Detecção do Ghostscript no Windows / Linux
def find_ghostscript():
    for cmd in ["gswin64c", "gswin32c", "gs"]:
        path = shutil.which(cmd)
        if path:
            return path

    common_patterns = [
        r"C:\Program Files\gs\gs*\bin\gswin64c.exe",
        r"C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe",
        r"C:\Program Files\gs\*\bin\gswin64c.exe",
    ]
    for pattern in common_patterns:
        matches = glob.glob(pattern)
        if matches:
            matches.sort(reverse=True)
            return matches[0]

    return None

GS_EXECUTABLE = find_ghostscript()

# Detecção do Tesseract OCR
def find_tesseract():
    cmd_path = shutil.which("tesseract")
    if cmd_path:
        return cmd_path
    common_tess = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Users\supervisor\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
    ]
    for p in common_tess:
        if os.path.exists(p):
            return p
    return None

TESSERACT_EXECUTABLE = find_tesseract()
if TESSERACT_EXECUTABLE:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXECUTABLE

PRESETS = {
    "screen": {
        "id": "screen",
        "name": "Super Compacto (Screen - 72 DPI)",
        "desc": "Máxima redução de MB. Recomendado para portais com limite rígido (ex: SEI, e-mail).",
        "gs_param": "/screen"
    },
    "ebook": {
        "id": "ebook",
        "name": "Equilibrado (eBook - 150 DPI) [Recomendado]",
        "desc": "Excelente nitidez para leitura em telas e tamanho reduzido em até 90%.",
        "gs_param": "/ebook"
    },
    "printer": {
        "id": "printer",
        "name": "Impressão (Printer - 300 DPI)",
        "desc": "Mantém alta resolução adequada para impressão física.",
        "gs_param": "/printer"
    },
    "prepress": {
        "id": "prepress",
        "name": "Alta Fidelidade (Prepress)",
        "desc": "Preservação máxima de cores e elementos vetoriais.",
        "gs_param": "/prepress"
    }
}

def format_bytes_pt(size_in_bytes):
    if size_in_bytes < 1024:
        return f"{size_in_bytes} B"
    elif size_in_bytes < 1024 * 1024:
        return f"{size_in_bytes / 1024:.1f} KB"
    else:
        return f"{size_in_bytes / (1024 * 1024):.2f} MB"

def clean_old_files():
    now = time.time()
    to_delete = []
    for fid, meta in FILE_REGISTRY.items():
        if now - meta.get("created_at", 0) > 7200:
            to_delete.append(fid)
            try:
                if os.path.exists(meta["path"]):
                    os.remove(meta["path"])
            except Exception:
                pass
    for fid in to_delete:
        FILE_REGISTRY.pop(fid, None)

def parse_page_ranges(range_str, max_pages):
    pages = []
    parts = [p.strip() for p in range_str.split(',') if p.strip()]
    for p in parts:
        if '-' in p:
            start_str, end_str = p.split('-', 1)
            try:
                start = max(1, int(start_str.strip()))
                end = min(max_pages, int(end_str.strip()))
                for i in range(start, end + 1):
                    if (i - 1) not in pages:
                        pages.append(i - 1)
            except ValueError:
                continue
        else:
            try:
                pg = int(p.strip())
                if 1 <= pg <= max_pages and (pg - 1) not in pages:
                    pages.append(pg - 1)
            except ValueError:
                continue
    pages.sort()
    return pages

# ==========================================
# THREAD DO FOLDER WATCHER (FEATURE 7)
# ==========================================
def watcher_worker():
    while True:
        try:
            if WATCHER_STATE["enabled"] and GS_EXECUTABLE:
                # Procura PDFs na pasta de entrada
                pdf_files = list(WATCH_IN_DIR.glob("*.pdf")) + list(WATCH_IN_DIR.glob("*.PDF"))
                for pdf_file in pdf_files:
                    # Ignora arquivos temporários
                    if pdf_file.name.startswith("~") or pdf_file.name.endswith(".tmp") or pdf_file.name.endswith(".crdownload"):
                        continue
                    
                    # Aguarda o término da cópia/download (tamanho estável)
                    try:
                        initial_size = pdf_file.stat().st_size
                        time.sleep(1)
                        if not pdf_file.exists():
                            continue
                        if pdf_file.stat().st_size != initial_size:
                            continue  # Arquivo ainda sendo gravado
                        # Tenta abrir para verificar se não está travado por outro processo
                        with open(pdf_file, "rb"):
                            pass
                    except (IOError, OSError):
                        continue

                    # Processa a compressão com Ghostscript
                    orig_size = pdf_file.stat().st_size
                    base_name = pdf_file.stem
                    out_filename = f"{base_name}_otimizado.pdf"
                    target_out = WATCH_OUT_DIR / out_filename
                    temp_out = WATCH_OUT_DIR / f"temp_{out_filename}"

                    cmd = [
                        GS_EXECUTABLE,
                        "-sDEVICE=pdfwrite",
                        "-dCompatibilityLevel=1.4",
                        "-dPDFSETTINGS=/ebook",
                        "-dNOPAUSE",
                        "-dQUIET",
                        "-dBATCH",
                        f"-sOutputFile={temp_out}",
                        str(pdf_file)
                    ]

                    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180, check=False)
                    
                    if res.returncode == 0 and temp_out.exists() and temp_out.stat().st_size > 0:
                        new_size = temp_out.stat().st_size
                        if new_size >= orig_size:
                            # Se o Ghostscript gerou arquivo maior, preserva o original
                            shutil.copy2(str(pdf_file), str(target_out))
                            temp_out.unlink(missing_ok=True)
                            new_size = orig_size
                            reduction = 0.0
                        else:
                            shutil.move(str(temp_out), str(target_out))
                            reduction = round((1 - (new_size / orig_size)) * 100, 1)

                        # Move o original para a pasta de arquivamento seguro
                        dest_orig = WATCH_ORIG_DIR / pdf_file.name
                        if dest_orig.exists():
                            dest_orig = WATCH_ORIG_DIR / f"{base_name}_{int(time.time())}.pdf"
                        shutil.move(str(pdf_file), str(dest_orig))

                        saved_bytes = max(0, orig_size - new_size)
                        WATCHER_STATE["processed_count"] += 1
                        WATCHER_STATE["saved_bytes_total"] += saved_bytes
                        
                        log_entry = {
                            "timestamp": time.strftime("%H:%M:%S"),
                            "filename": pdf_file.name,
                            "orig_str": format_bytes_pt(orig_size),
                            "new_str": format_bytes_pt(new_size),
                            "saved_str": format_bytes_pt(saved_bytes),
                            "reduction": reduction,
                            "status": "Otimizado" if reduction > 0 else "Já Ideal"
                        }
                        WATCHER_STATE["logs"].insert(0, log_entry)
                        if len(WATCHER_STATE["logs"]) > 50:
                            WATCHER_STATE["logs"].pop()
                    else:
                        temp_out.unlink(missing_ok=True)

        except Exception as e:
            pass
        time.sleep(2)

# Inicia o Folder Watcher em segundo plano
watcher_thread = threading.Thread(target=watcher_worker, daemon=True)
watcher_thread.start()


# ==========================================
# ROTAS PRINCIPAIS
# ==========================================
@app.route("/")
def home():
    return render_template("home.html")

@app.route("/pdf-studio")
def index():
    return render_template("index.html")

@app.route("/api/status")
def status():
    return jsonify({
        "ghostscript_available": GS_EXECUTABLE is not None,
        "gs_path": GS_EXECUTABLE,
        "tesseract_available": TESSERACT_EXECUTABLE is not None,
        "tess_path": TESSERACT_EXECUTABLE,
        "presets": list(PRESETS.values()),
        "default_preset": "ebook"
    })

# ==========================================
# ROTA 1: COMPRESSÃO DE PDFS
# ==========================================
@app.route("/api/compress", methods=["POST"])
def compress_pdf():
    clean_old_files()

    if not GS_EXECUTABLE:
        return jsonify({"error": "Ghostscript não foi encontrado no sistema."}), 500

    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado."}), 400

    file = request.files["file"]
    if not file or not file.filename:
        return jsonify({"error": "Arquivo inválido."}), 400

    orig_filename = file.filename
    if not orig_filename.lower().endswith(".pdf"):
        return jsonify({"error": "Apenas arquivos PDF são aceitos."}), 400

    preset_key = request.form.get("preset", "ebook")
    preset_info = PRESETS.get(preset_key, PRESETS["ebook"])
    gs_setting = preset_info["gs_param"]

    unique_id = uuid.uuid4().hex
    input_path = STORAGE_DIR / f"in_{unique_id}.pdf"
    output_path = STORAGE_DIR / f"out_{unique_id}.pdf"

    try:
        file.save(str(input_path))
        orig_size = os.path.getsize(str(input_path))
        start_time = time.time()

        cmd = [
            GS_EXECUTABLE,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS={gs_setting}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={output_path}",
            str(input_path)
        ]

        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120, check=False)
        duration = round(time.time() - start_time, 2)

        if result.returncode != 0 or not output_path.exists():
            err_msg = result.stderr.decode("latin1", errors="ignore")
            return jsonify({"error": f"Falha na compressão do Ghostscript: {err_msg}"}), 500

        new_size = os.path.getsize(str(output_path))
        
        if new_size >= orig_size:
            shutil.copy(str(input_path), str(output_path))
            new_size = orig_size
            saved_bytes = 0
            reduction_percent = 0.0
            status_tag = "already_optimized"
        else:
            saved_bytes = orig_size - new_size
            reduction_percent = round((saved_bytes / orig_size) * 100, 1) if orig_size > 0 else 0.0
            status_tag = "reduced"

        base_name, _ = os.path.splitext(orig_filename)
        clean_out_name = f"{base_name}_otimizado.pdf"

        FILE_REGISTRY[unique_id] = {
            "path": str(output_path),
            "orig_name": clean_out_name,
            "created_at": time.time()
        }

        return jsonify({
            "success": True,
            "file_id": unique_id,
            "filename": orig_filename,
            "output_filename": clean_out_name,
            "orig_size": orig_size,
            "new_size": new_size,
            "saved_bytes": saved_bytes,
            "orig_size_str": format_bytes_pt(orig_size),
            "new_size_str": format_bytes_pt(new_size),
            "saved_str": format_bytes_pt(saved_bytes),
            "orig_size_mb": round(orig_size / (1024 * 1024), 2),
            "new_size_mb": round(new_size / (1024 * 1024), 2),
            "reduction_percent": reduction_percent,
            "status_tag": status_tag,
            "duration": duration,
            "preset": preset_info["name"],
            "download_url": f"/api/download/{unique_id}",
            "preview_url": f"/api/view/{unique_id}"
        })

    except subprocess.TimeoutExpired:
        return jsonify({"error": "Tempo limite de compressão excedido (120s)."}), 504
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if input_path.exists():
            try:
                os.remove(str(input_path))
            except Exception:
                pass

# ==========================================
# ROTA 1 (FEATURE 1): MESCLAR / UNIR PDFS
# ==========================================
@app.route("/api/merge-pdfs", methods=["POST"])
def merge_pdfs():
    clean_old_files()

    if not GS_EXECUTABLE:
        return jsonify({"error": "Ghostscript não foi encontrado no sistema."}), 500

    files = request.files.getlist("pdfs")
    if not files or len(files) < 2:
        return jsonify({"error": "Selecione pelo menos 2 arquivos PDF para mesclar."}), 400

    preset_key = request.form.get("preset", "ebook")
    preset_info = PRESETS.get(preset_key, PRESETS["ebook"])
    gs_setting = preset_info["gs_param"]

    custom_title = request.form.get("pdf_title", "documento_mesclado.pdf").strip()
    if not custom_title.lower().endswith(".pdf"):
        custom_title += ".pdf"

    unique_id = uuid.uuid4().hex
    saved_inputs = []
    total_orig_size = 0
    start_time = time.time()

    try:
        total_pages = 0
        for i, f in enumerate(files):
            temp_in = STORAGE_DIR / f"merge_in_{unique_id}_{i}.pdf"
            f.save(str(temp_in))
            total_orig_size += os.path.getsize(str(temp_in))
            saved_inputs.append(str(temp_in))
            try:
                r = PdfReader(str(temp_in))
                total_pages += len(r.pages)
            except Exception:
                pass

        merged_out = STORAGE_DIR / f"merged_{unique_id}.pdf"

        # Ghostscript faz a fusão e a otimização de uma só vez!
        cmd = [
            GS_EXECUTABLE,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS={gs_setting}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={merged_out}",
            *saved_inputs
        ]

        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180, check=False)
        duration = round(time.time() - start_time, 2)

        if result.returncode != 0 or not merged_out.exists():
            err_msg = result.stderr.decode("latin1", errors="ignore")
            return jsonify({"error": f"Falha ao mesclar PDFs: {err_msg}"}), 500

        final_size = os.path.getsize(str(merged_out))
        saved_bytes = max(0, total_orig_size - final_size)
        reduction_percent = round((saved_bytes / total_orig_size) * 100, 1) if total_orig_size > 0 else 0.0

        FILE_REGISTRY[unique_id] = {
            "path": str(merged_out),
            "orig_name": custom_title,
            "created_at": time.time()
        }

        return jsonify({
            "success": True,
            "file_id": unique_id,
            "output_filename": custom_title,
            "page_count": total_pages,
            "file_count": len(saved_inputs),
            "orig_size": total_orig_size,
            "new_size": final_size,
            "saved_bytes": saved_bytes,
            "orig_size_str": format_bytes_pt(total_orig_size),
            "new_size_str": format_bytes_pt(final_size),
            "saved_str": format_bytes_pt(saved_bytes),
            "reduction_percent": reduction_percent,
            "duration": duration,
            "preset": preset_info["name"],
            "download_url": f"/api/download/{unique_id}",
            "preview_url": f"/api/view/{unique_id}"
        })

    except Exception as e:
        return jsonify({"error": f"Erro durante a mesclagem: {str(e)}"}), 500
    finally:
        for p in saved_inputs:
            try:
                os.remove(p)
            except Exception:
                pass

# ==========================================
# ROTA 2 (FEATURE 2): DIVIDIR / EXTRAIR PÁGINAS
# ==========================================
@app.route("/api/pdf-info", methods=["POST"])
def get_pdf_info():
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado."}), 400
    f = request.files["file"]
    try:
        reader = PdfReader(f.stream)
        num_pages = len(reader.pages)
        f.seek(0, os.SEEK_END)
        size = f.tell()
        return jsonify({
            "success": True,
            "filename": f.filename,
            "page_count": num_pages,
            "size": size,
            "size_str": format_bytes_pt(size)
        })
    except Exception as e:
        return jsonify({"error": f"Não foi possível ler o PDF: {str(e)}"}), 400

@app.route("/api/split-pdf", methods=["POST"])
def split_pdf():
    clean_old_files()

    if not GS_EXECUTABLE:
        return jsonify({"error": "Ghostscript não foi encontrado no sistema."}), 500

    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado."}), 400

    file = request.files["file"]
    mode = request.form.get("mode", "range")  # "range" ou "all"
    range_str = request.form.get("range", "").strip()
    preset_key = request.form.get("preset", "ebook")
    preset_info = PRESETS.get(preset_key, PRESETS["ebook"])
    gs_setting = preset_info["gs_param"]

    unique_id = uuid.uuid4().hex
    temp_input = STORAGE_DIR / f"split_in_{unique_id}.pdf"
    temp_extracted = STORAGE_DIR / f"split_raw_{unique_id}.pdf"
    opt_output = STORAGE_DIR / f"split_out_{unique_id}.pdf"

    try:
        file.save(str(temp_input))
        orig_size = os.path.getsize(str(temp_input))
        reader = PdfReader(str(temp_input))
        total_pages = len(reader.pages)
        start_time = time.time()

        if mode == "range":
            selected_indices = parse_page_ranges(range_str, total_pages)
            if not selected_indices:
                return jsonify({"error": "Intervalo de páginas inválido."}), 400

            writer = PdfWriter()
            for idx in selected_indices:
                writer.add_page(reader.pages[idx])

            with open(temp_extracted, "wb") as f_out:
                writer.write(f_out)

            # Otimiza o PDF extraído com Ghostscript
            cmd = [
                GS_EXECUTABLE,
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS={gs_setting}",
                "-dNOPAUSE",
                "-dQUIET",
                "-dBATCH",
                f"-sOutputFile={opt_output}",
                str(temp_extracted)
            ]
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120, check=False)

            final_path = opt_output if (opt_output.exists() and opt_output.stat().st_size > 0) else temp_extracted
            final_size = os.path.getsize(str(final_path))
            duration = round(time.time() - start_time, 2)

            base_name = Path(file.filename).stem
            clean_out_name = f"{base_name}_paginas_{range_str.replace(' ', '').replace(',', '_')}.pdf"

            FILE_REGISTRY[unique_id] = {
                "path": str(final_path),
                "orig_name": clean_out_name,
                "created_at": time.time()
            }

            return jsonify({
                "success": True,
                "file_id": unique_id,
                "output_filename": clean_out_name,
                "page_count": len(selected_indices),
                "orig_size": orig_size,
                "new_size": final_size,
                "orig_size_str": format_bytes_pt(orig_size),
                "new_size_str": format_bytes_pt(final_size),
                "duration": duration,
                "download_url": f"/api/download/{unique_id}",
                "preview_url": f"/api/view/{unique_id}"
            })

        elif mode == "all":
            # Separa cada página como um PDF e empacota em um ZIP
            zip_buffer = BytesIO()
            base_name = Path(file.filename).stem

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx in range(total_pages):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[idx])
                    page_buf = BytesIO()
                    writer.write(page_buf)
                    page_bytes = page_buf.getvalue()
                    page_filename = f"{base_name}_pag_{idx+1:02d}.pdf"
                    zf.writestr(page_filename, page_bytes)

            zip_buffer.seek(0)
            zip_filename = f"{base_name}_todas_as_paginas.zip"

            zip_path = STORAGE_DIR / f"zip_{unique_id}.zip"
            with open(zip_path, "wb") as fz:
                fz.write(zip_buffer.getvalue())

            FILE_REGISTRY[unique_id] = {
                "path": str(zip_path),
                "orig_name": zip_filename,
                "created_at": time.time()
            }

            return jsonify({
                "success": True,
                "file_id": unique_id,
                "output_filename": zip_filename,
                "page_count": total_pages,
                "orig_size": orig_size,
                "new_size": len(zip_buffer.getvalue()),
                "orig_size_str": format_bytes_pt(orig_size),
                "new_size_str": format_bytes_pt(len(zip_buffer.getvalue())),
                "duration": round(time.time() - start_time, 2),
                "download_url": f"/api/download/{unique_id}",
                "is_zip": True
            })

    except Exception as e:
        return jsonify({"error": f"Erro ao dividir PDF: {str(e)}"}), 500
    finally:
        for p in [temp_input, temp_extracted]:
            if p.exists():
                try:
                    os.remove(p)
                except Exception:
                    pass

# ==========================================
# ROTA 3: IMAGENS PARA PDF (FEATURE 4)
# ==========================================
def process_image_to_page(img, page_size="a4", orientation="auto", margin=30):
    img = ImageOps.exif_transpose(img)
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        bg = Image.new("RGB", img.size, (255, 255, 255))
        if img.mode != "RGBA":
            img = img.convert("RGBA")
        bg.paste(img, mask=img.split()[3])
        img = bg
    elif img.mode != "RGB":
        img = img.convert("RGB")
        
    if page_size != "a4":
        return img
        
    w, h = img.size
    is_landscape = (w > h) if orientation == "auto" else (orientation == "landscape")
    page_w, page_h = (1754, 1240) if is_landscape else (1240, 1754)
    canvas = Image.new("RGB", (page_w, page_h), (255, 255, 255))
    
    max_w = max(10, page_w - (margin * 2))
    max_h = max(10, page_h - (margin * 2))
    
    scale = min(max_w / w, max_h / h)
    new_w = max(1, int(w * scale))
    new_h = max(1, int(h * scale))
    
    resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    pos_x = (page_w - new_w) // 2
    pos_y = (page_h - new_h) // 2
    canvas.paste(resized, (pos_x, pos_y))
    return canvas

@app.route("/api/images-to-pdf", methods=["POST"])
def images_to_pdf():
    clean_old_files()
    
    if not GS_EXECUTABLE:
        return jsonify({"error": "Ghostscript não foi encontrado no sistema."}), 500
        
    files = request.files.getlist("images")
    if not files or len(files) == 0:
        return jsonify({"error": "Nenhuma imagem selecionada."}), 400
        
    preset_key = request.form.get("preset", "ebook")
    preset_info = PRESETS.get(preset_key, PRESETS["ebook"])
    gs_setting = preset_info["gs_param"]
    
    output_mode = request.form.get("output_mode", "single") # "single" ou "individual"
    page_size = request.form.get("page_size", "a4")
    orientation = request.form.get("orientation", "auto")
    margin_str = request.form.get("margin", "small")
    margin_map = {"none": 0, "small": 30, "normal": 60}
    margin = margin_map.get(margin_str, 30)
    
    custom_title = request.form.get("pdf_title", "documento_fotos.pdf").strip()
    if not custom_title:
        custom_title = "documento_fotos.pdf"
    if not custom_title.lower().endswith(".pdf"):
        custom_title += ".pdf"
        
    enable_ocr = request.form.get("enable_ocr") in ["true", "1", "on"]
    ocr_lang = request.form.get("ocr_lang", "por")
    
    start_time = time.time()
    
    if output_mode == "individual":
        # Processa cada imagem para um PDF individual e empacota em ZIP
        unique_id = uuid.uuid4().hex
        zip_buffer = BytesIO()
        total_orig_size = 0
        total_final_size = 0
        converted_count = 0
        temp_files = []
        
        try:
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx, f in enumerate(files):
                    if not f or not f.filename:
                        continue
                    
                    f.seek(0, os.SEEK_END)
                    file_size = f.tell()
                    total_orig_size += file_size
                    f.seek(0)
                    
                    img = Image.open(f.stream)
                    page = process_image_to_page(img, page_size=page_size, orientation=orientation, margin=margin)
                    
                    stem = Path(f.filename).stem or f"imagem_{idx+1:02d}"
                    raw_single = STORAGE_DIR / f"raw_single_{unique_id}_{idx}.pdf"
                    opt_single = STORAGE_DIR / f"opt_single_{unique_id}_{idx}.pdf"
                    temp_files.extend([raw_single, opt_single])
                    
                    if enable_ocr and TESSERACT_EXECUTABLE:
                        try:
                            pdf_bytes = pytesseract.image_to_pdf_or_hocr(page, extension="pdf", lang=ocr_lang)
                            with open(raw_single, "wb") as f_ocr:
                                f_ocr.write(pdf_bytes)
                        except Exception:
                            page.save(str(raw_single), format="PDF", resolution=150.0)
                    else:
                        page.save(str(raw_single), format="PDF", resolution=150.0)
                    
                    cmd = [
                        GS_EXECUTABLE,
                        "-sDEVICE=pdfwrite",
                        "-dCompatibilityLevel=1.4",
                        f"-dPDFSETTINGS={gs_setting}",
                        "-dNOPAUSE",
                        "-dQUIET",
                        "-dBATCH",
                        f"-sOutputFile={opt_single}",
                        str(raw_single)
                    ]
                    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60, check=False)
                    
                    final_single = opt_single if (opt_single.exists() and opt_single.stat().st_size > 0) else raw_single
                    total_final_size += os.path.getsize(str(final_single))
                    
                    zf.write(str(final_single), arcname=f"{stem}.pdf")
                    converted_count += 1
                    
            if converted_count == 0:
                return jsonify({"error": "Nenhuma imagem válida pôde ser convertida."}), 400
                
            zip_buffer.seek(0)
            zip_bytes = zip_buffer.getvalue()
            zip_size = len(zip_bytes)
            base_zip_name = Path(custom_title).stem + "_individuais.zip"
            zip_path = STORAGE_DIR / f"zip_img_{unique_id}.zip"
            
            with open(zip_path, "wb") as fz:
                fz.write(zip_bytes)
                
            FILE_REGISTRY[unique_id] = {
                "path": str(zip_path),
                "orig_name": base_zip_name,
                "created_at": time.time()
            }
            
            duration = round(time.time() - start_time, 2)
            saved_bytes = max(0, total_orig_size - zip_size)
            reduction_percent = round((saved_bytes / total_orig_size) * 100, 1) if total_orig_size > 0 else 0.0
            
            return jsonify({
                "success": True,
                "file_id": unique_id,
                "output_filename": base_zip_name,
                "page_count": converted_count,
                "orig_size": total_orig_size,
                "new_size": zip_size,
                "saved_bytes": saved_bytes,
                "orig_size_str": format_bytes_pt(total_orig_size),
                "new_size_str": format_bytes_pt(zip_size),
                "saved_str": format_bytes_pt(saved_bytes),
                "reduction_percent": reduction_percent,
                "duration": duration,
                "preset": preset_info["name"],
                "download_url": f"/api/download/{unique_id}",
                "is_zip": True
            })
        except Exception as e:
            return jsonify({"error": f"Erro na conversão individual: {str(e)}"}), 500
        finally:
            for p in temp_files:
                if p.exists():
                    try:
                        os.remove(str(p))
                    except Exception:
                        pass

    # Modo padrão: PDF Único
    unique_id = uuid.uuid4().hex
    raw_pdf_path = STORAGE_DIR / f"raw_img_{unique_id}.pdf"
    compressed_pdf_path = STORAGE_DIR / f"out_{unique_id}.pdf"
    
    total_orig_size = 0
    processed_pages = []
    
    try:
        for f in files:
            if not f or not f.filename:
                continue
            
            f.seek(0, os.SEEK_END)
            total_orig_size += f.tell()
            f.seek(0)
            
            img = Image.open(f.stream)
            page = process_image_to_page(img, page_size=page_size, orientation=orientation, margin=margin)
            processed_pages.append(page)
            
        if not processed_pages:
            return jsonify({"error": "Nenhuma imagem válida pôde ser lida."}), 400
            
        if enable_ocr and TESSERACT_EXECUTABLE:
            writer = PdfWriter()
            temp_ocr_pages = []
            for p_img in processed_pages:
                try:
                    ocr_pdf_bytes = pytesseract.image_to_pdf_or_hocr(p_img, extension="pdf", lang=ocr_lang)
                    r_ocr = PdfReader(BytesIO(ocr_pdf_bytes))
                    writer.add_page(r_ocr.pages[0])
                except Exception:
                    b_fallback = BytesIO()
                    p_img.save(b_fallback, format="PDF", resolution=150.0)
                    b_fallback.seek(0)
                    r_fb = PdfReader(b_fallback)
                    writer.add_page(r_fb.pages[0])
            with open(raw_pdf_path, "wb") as f_out:
                writer.write(f_out)
        else:
            first_page = processed_pages[0]
            other_pages = processed_pages[1:] if len(processed_pages) > 1 else []
            first_page.save(
                str(raw_pdf_path),
                format="PDF",
                save_all=True,
                append_images=other_pages,
                resolution=150.0
            )
        
        cmd = [
            GS_EXECUTABLE,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS={gs_setting}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={compressed_pdf_path}",
            str(raw_pdf_path)
        ]
        
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120, check=False)
        duration = round(time.time() - start_time, 2)
        
        final_path = compressed_pdf_path if (compressed_pdf_path.exists() and os.path.getsize(str(compressed_pdf_path)) > 0) else raw_pdf_path
        final_size = os.path.getsize(str(final_path))
        
        saved_bytes = max(0, total_orig_size - final_size)
        reduction_percent = round((saved_bytes / total_orig_size) * 100, 1) if total_orig_size > 0 else 0.0
        
        FILE_REGISTRY[unique_id] = {
            "path": str(final_path),
            "orig_name": custom_title,
            "created_at": time.time()
        }
        
        return jsonify({
            "success": True,
            "file_id": unique_id,
            "output_filename": custom_title,
            "page_count": len(processed_pages),
            "orig_size": total_orig_size,
            "new_size": final_size,
            "saved_bytes": saved_bytes,
            "orig_size_str": format_bytes_pt(total_orig_size),
            "new_size_str": format_bytes_pt(final_size),
            "saved_str": format_bytes_pt(saved_bytes),
            "reduction_percent": reduction_percent,
            "duration": duration,
            "preset": preset_info["name"],
            "download_url": f"/api/download/{unique_id}",
            "preview_url": f"/api/view/{unique_id}",
            "is_zip": False
        })
        
    except Exception as e:
        return jsonify({"error": f"Erro na conversão: {str(e)}"}), 500
    finally:
        if raw_pdf_path.exists() and final_path != raw_pdf_path:
            try:
                os.remove(str(raw_pdf_path))
            except Exception:
                pass

# ==========================================
# ROTA 4 (FEATURE 7): CONTROLE DO FOLDER WATCHER
# ==========================================
@app.route("/api/watcher/status")
def watcher_status():
    return jsonify({
        "enabled": WATCHER_STATE["enabled"],
        "processed_count": WATCHER_STATE["processed_count"],
        "saved_bytes_total": WATCHER_STATE["saved_bytes_total"],
        "saved_total_str": format_bytes_pt(WATCHER_STATE["saved_bytes_total"]),
        "logs": WATCHER_STATE["logs"],
        "base_dir": str(WATCH_BASE_DIR),
        "in_dir": str(WATCH_IN_DIR),
        "out_dir": str(WATCH_OUT_DIR),
        "orig_dir": str(WATCH_ORIG_DIR)
    })

@app.route("/api/watcher/toggle", methods=["POST"])
def watcher_toggle():
    WATCHER_STATE["enabled"] = not WATCHER_STATE["enabled"]
    return jsonify({"success": True, "enabled": WATCHER_STATE["enabled"]})

@app.route("/api/watcher/open-folder", methods=["POST"])
def watcher_open_folder():
    data = request.get_json() or {}
    folder_type = data.get("type", "base")
    path_map = {
        "base": WATCH_BASE_DIR,
        "in": WATCH_IN_DIR,
        "out": WATCH_OUT_DIR,
        "orig": WATCH_ORIG_DIR
    }
    target = path_map.get(folder_type, WATCH_BASE_DIR)
    try:
        os.startfile(str(target))
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ==========================================
# ROTAS DE DOWNLOAD E VISUALIZAÇÃO
# ==========================================
@app.route("/api/view/<file_id>")
def view_file(file_id):
    meta = FILE_REGISTRY.get(file_id)
    if not meta or not os.path.exists(meta["path"]):
        return "Arquivo expirado ou não encontrado.", 404

    return send_file(
        meta["path"],
        as_attachment=False,
        download_name=meta["orig_name"],
        mimetype="application/pdf"
    )

@app.route("/api/download/<file_id>")
def download_file(file_id):
    meta = FILE_REGISTRY.get(file_id)
    if not meta or not os.path.exists(meta["path"]):
        return "Arquivo expirado ou não encontrado.", 404

    is_zip = meta["orig_name"].lower().endswith(".zip")
    mimetype = "application/zip" if is_zip else "application/pdf"

    return send_file(
        meta["path"],
        as_attachment=True,
        download_name=meta["orig_name"],
        mimetype=mimetype
    )

@app.route("/api/download-batch", methods=["POST"])
def download_batch():
    data = request.get_json() or {}
    file_ids = data.get("file_ids", [])

    if not file_ids:
        return jsonify({"error": "Nenhum arquivo selecionado."}), 400

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        added = 0
        for fid in file_ids:
            meta = FILE_REGISTRY.get(fid)
            if meta and os.path.exists(meta["path"]):
                zip_file.write(meta["path"], arcname=meta["orig_name"])
                added += 1

    if added == 0:
        return jsonify({"error": "Nenhum arquivo válido encontrado para o pacote."}), 404

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name="pdfs_otimizados.zip",
        mimetype="application/zip"
    )

# ==========================================
# ROTAS DO MONITOR DE CERTIFICADOS DIGITAIS
# ==========================================
EXCEL_CERT_PATH = Path(r"C:\Users\supervisor\Downloads\CERTIFICADOS VENCIMENTOS.xlsx")

@app.route("/conciliacao")
def pagina_conciliacao():
    return render_template("conciliacao.html")

@app.route("/certificados")
def pagina_certificados():
    return render_template("certificados.html")

@app.route("/api/certificados/listar")
def listar_certificados():
    if not EXCEL_CERT_PATH.exists():
        return jsonify({"error": "Planilha de certificados não encontrada em Downloads."}), 404
        
    try:
        wb = openpyxl.load_workbook(str(EXCEL_CERT_PATH), data_only=True)
        sheet = wb["Planilha3"] if "Planilha3" in wb.sheetnames else wb.active
        
        certificados = []
        now = datetime.now()
        
        # Pula as linhas de título/cabeçalho
        for row in list(sheet.iter_rows(values_only=True))[3:]:
            num_empresa = str(row[0]).strip() if row[0] is not None else ""
            cnpj = str(row[1]).strip() if row[1] is not None else ""
            nome = str(row[2]).strip() if row[2] is not None else ""
            venc_raw = row[3]
            
            if not nome and not cnpj:
                continue
                
            venc_str = ""
            venc_iso = ""
            status = "desconhecido"
            dias_restantes = None
            
            if isinstance(venc_raw, datetime):
                venc_dt = venc_raw
                venc_str = venc_dt.strftime("%d/%m/%Y")
                venc_iso = venc_dt.strftime("%Y-%m-%d")
                diff = (venc_dt.date() - now.date()).days
                dias_restantes = diff
                if diff < 0:
                    status = "vencido"
                elif diff <= 30:
                    status = "urgente"
                else:
                    status = "valido"
            elif venc_raw:
                venc_str = str(venc_raw)[:10]
                venc_iso = venc_str
                
            certificados.append({
                "empresa_num": num_empresa,
                "cnpj": cnpj,
                "nome": nome,
                "vencimento": venc_str,
                "vencimento_iso": venc_iso,
                "dias_restantes": dias_restantes,
                "status": status
            })
            
        return jsonify({
            "success": True,
            "total": len(certificados),
            "certificados": certificados
        })
    except Exception as e:
        return jsonify({"error": f"Erro ao ler certificados: {str(e)}"}), 500

# Estado global da sincronização incremental de certificados
SYNC_STATE = {
    "running": False,
    "progresso_atual": 0,
    "total": 0,
    "mensagem": "",
    "ultimo_resultado": None,
    "ultima_execucao": None
}

def worker_sync_certificados():
    global SYNC_STATE
    try:
        from sync_certificados import executar_varredura
        
        def prog_callback(curr, tot, msg):
            SYNC_STATE["progresso_atual"] = curr
            SYNC_STATE["total"] = tot
            SYNC_STATE["mensagem"] = msg

        res = executar_varredura(callback_progresso=prog_callback, forcar_tudo=False)
        SYNC_STATE["ultimo_resultado"] = res
        SYNC_STATE["ultima_execucao"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        SYNC_STATE["mensagem"] = "Varredura concluída com sucesso!"
    except Exception as ex:
        SYNC_STATE["ultimo_resultado"] = {"sucesso": False, "erro": str(ex)}
        SYNC_STATE["mensagem"] = f"Erro na varredura: {str(ex)}"
    finally:
        SYNC_STATE["running"] = False

@app.route("/api/certificados/sincronizar", methods=["POST"])
def sincronizar_certificados_api():
    global SYNC_STATE
    if SYNC_STATE["running"]:
        return jsonify({"success": False, "message": "A varredura já está em andamento no momento."}), 400

    SYNC_STATE["running"] = True
    SYNC_STATE["progresso_atual"] = 0
    SYNC_STATE["total"] = 0
    SYNC_STATE["mensagem"] = "Iniciando varredura incremental..."
    
    t = threading.Thread(target=worker_sync_certificados, daemon=True)
    t.start()
    
    return jsonify({"success": True, "message": "Varredura iniciada em segundo plano."})

@app.route("/api/certificados/status-sync")
def status_sync_certificados_api():
    return jsonify(SYNC_STATE)

if __name__ == "__main__":
    print(f"==================================================")
    print(f" GhostPDF Studio iniciado com sucesso!")
    print(f" Ghostscript detectado: {GS_EXECUTABLE}")
    print(f" Folder Watcher ativo em: {WATCH_BASE_DIR}")
    print(f" Acesse no navegador: http://127.0.0.1:5000")
    print(f"==================================================")
    app.run(host="0.0.0.0", port=5000, debug=False)
