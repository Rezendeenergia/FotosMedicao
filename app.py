import os
import io
import zipfile
import copy
import time
import logging
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image, ImageOps
import sys

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, static_folder=BASE_DIR)
CORS(app, expose_headers=["X-Photos-Used", "X-Slides-Filled", "X-Processing-Time"])

# Configurações para produção
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

# ── Dimensões alvo ──
TARGET_W_CM  = 7.62
TARGET_H_CM  = 10.16
CM_TO_EMU    = 360000
TARGET_W_EMU = int(TARGET_W_CM * CM_TO_EMU)
TARGET_H_EMU = int(TARGET_H_CM * CM_TO_EMU)

PHOTO_SLOTS = [
    {"name": "Retângulo 19", "x": 189290,  "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 41", "x": 3193696, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 15", "x": 6198102, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 8",  "x": 9202508, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
]

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'

# ── Rotas frontend ──
@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")

@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "message": "Backend Rezende Energia rodando!",
        "timestamp": time.time()
    })

# ── Utilitários de imagem ──
def resize_photo(img_bytes: bytes) -> bytes:
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        target_w_px = int(TARGET_W_CM / 2.54 * 150)
        target_h_px = int(TARGET_H_CM / 2.54 * 150)
        img = ImageOps.fit(img, (target_w_px, target_h_px), method=Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85, optimize=True)
        logger.info(f"Imagem redimensionada: {target_w_px}x{target_h_px}")
        return buf.getvalue()
    except Exception as e:
        logger.error(f"Erro ao redimensionar imagem: {e}")
        raise

def add_photo_to_slide(slide, slot: dict, img_bytes: bytes):
    """Remove o placeholder e insere a foto no slot correto."""
    sp_tree = slide.shapes._spTree

    for sp in list(sp_tree):
        if sp.tag.split("}")[-1] == "sp":
            cNvPr = sp.find(".//" + qn("p:cNvPr"))
            if cNvPr is not None and cNvPr.get("name") == slot["name"]:
                sp_tree.remove(sp)
                break

    img_resized = resize_photo(img_bytes)
    _, rId = slide.part.get_or_add_image_part(io.BytesIO(img_resized))

    pic_xml = f"""<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvPicPr>
    <p:cNvPr id="9999" name="{slot['name']}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="{rId}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="{slot['x']}" y="{slot['y']}"/>
      <a:ext cx="{slot['cx']}" cy="{slot['cy']}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>"""

    sp_tree.append(etree.fromstring(pic_xml))

# ── Gerenciamento dinâmico de slides ──
def duplicate_slide(prs: Presentation, source_index: int):
    """Clona o slide no source_index (incluindo fundo e shapes)."""
    source_slide = prs.slides[source_index]
    new_slide = prs.slides.add_slide(source_slide.slide_layout)

    sp_tree_src = copy.deepcopy(source_slide.shapes._spTree)
    sp_tree_new = new_slide.shapes._spTree
    for child in list(sp_tree_new):
        sp_tree_new.remove(child)
    for child in list(sp_tree_src):
        sp_tree_new.append(child)

    bg_src = source_slide._element.find(f'{{{P_NS}}}cSld/{{{P_NS}}}bg')
    if bg_src is not None:
        bg_copy = copy.deepcopy(bg_src)
        cSld = new_slide._element.find(f'{{{P_NS}}}cSld')
        existing_bg = cSld.find(f'{{{P_NS}}}bg')
        if existing_bg is not None:
            cSld.remove(existing_bg)
        cSld.insert(0, bg_copy)

    for rel in source_slide.part.rels.values():
        if 'image' in rel.reltype:
            new_slide.part.rels._rels[rel.rId] = rel

    return new_slide

def remove_last_slide(prs: Presentation):
    """Remove o último slide da apresentação."""
    sldIdLst = prs.slides._sldIdLst
    if len(sldIdLst) > 0:
        sldIdLst.remove(sldIdLst[-1])

def process_pptx(pptx_bytes: bytes, photos: list) -> bytes:
    """Lógica dinâmica com slides."""
    logger.info(f"Iniciando processamento com {len(photos)} fotos")
    prs = Presentation(io.BytesIO(pptx_bytes))

    n_photos = len(photos)
    n_slides_needed = (n_photos + 3) // 4
    n_slides_have = len(list(prs.slides)) - 1
    
    logger.info(f"Slides necessários: {n_slides_needed}, disponíveis: {n_slides_have}")

    if n_slides_needed > n_slides_have:
        logger.info(f"Duplicando {n_slides_needed - n_slides_have} slides")
        for _ in range(n_slides_needed - n_slides_have):
            duplicate_slide(prs, 1)

    all_slides = list(prs.slides)
    photo_idx = 0
    slides_used = 0

    for slide in all_slides[1:]:
        if photo_idx >= n_photos:
            break
        for slot in PHOTO_SLOTS:
            if photo_idx >= n_photos:
                break
            _, img_bytes = photos[photo_idx]
            add_photo_to_slide(slide, slot, img_bytes)
            photo_idx += 1
        slides_used += 1
        logger.info(f"Slide {slides_used} processado, {photo_idx}/{n_photos} fotos inseridas")

    total_slides_now = len(list(prs.slides))
    slides_to_remove = total_slides_now - 1 - slides_used
    logger.info(f"Removendo {slides_to_remove} slides vazios")
    for _ in range(slides_to_remove):
        remove_last_slide(prs)

    out = io.BytesIO()
    prs.save(out)
    logger.info("Processamento concluído")
    return out.getvalue()

def extract_photos_from_zip(zip_bytes: bytes) -> list:
    VALID_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tiff", ".tif"}
    photos = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = sorted([
            n for n in zf.namelist()
            if os.path.splitext(n.lower())[1] in VALID_EXT
            and not n.startswith("__MACOSX")
            and not os.path.basename(n).startswith(".")
            and not n.startswith(".")
        ])
        logger.info(f"Encontradas {len(names)} imagens no ZIP")
        for name in names[:100]:  # Limitar a 100 fotos
            photos.append((os.path.basename(name), zf.read(name)))
    return photos

# ── Rotas API ──
@app.route("/process", methods=["POST"])
def process():
    start_time = time.time()
    try:
        if "pptx" not in request.files:
            return jsonify({"error": "Arquivo .pptx não enviado"}), 400
        if "zip" not in request.files:
            return jsonify({"error": "Arquivo .zip não enviado"}), 400

        pptx_file = request.files["pptx"]
        zip_file = request.files["zip"]

        if not pptx_file.filename.lower().endswith(".pptx"):
            return jsonify({"error": "O arquivo deve ser .pptx"}), 400
        if not zip_file.filename.lower().endswith(".zip"):
            return jsonify({"error": "O pacote de fotos deve ser .zip"}), 400

        # Verificar tamanhos
        pptx_file.seek(0, 2)
        pptx_size = pptx_file.tell()
        pptx_file.seek(0)
        
        zip_file.seek(0, 2)
        zip_size = zip_file.tell()
        zip_file.seek(0)
        
        logger.info(f"Recebendo arquivos: PPTX={pptx_size} bytes, ZIP={zip_size} bytes")
        
        if pptx_size > 100 * 1024 * 1024:
            return jsonify({"error": "Arquivo PPTX muito grande (máx 100MB)"}), 400
        if zip_size > 100 * 1024 * 1024:
            return jsonify({"error": "Arquivo ZIP muito grande (máx 100MB)"}), 400

        pptx_bytes = pptx_file.read()
        zip_bytes = zip_file.read()

        photos = extract_photos_from_zip(zip_bytes)
        
        if not photos:
            return jsonify({"error": "Nenhuma imagem encontrada no ZIP"}), 400

        result_bytes = process_pptx(pptx_bytes, photos)

        n_photos = len(photos)
        n_slides = (n_photos + 3) // 4
        processing_time = time.time() - start_time

        logger.info(f"Processamento concluído em {processing_time:.2f} segundos")

        response = send_file(
            io.BytesIO(result_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="relatorio_preenchido.pptx",
        )
        response.headers["X-Photos-Used"] = str(n_photos)
        response.headers["X-Slides-Filled"] = str(n_slides)
        response.headers["X-Processing-Time"] = f"{processing_time:.2f}"
        return response
        
    except Exception as e:
        logger.error(f"Erro no processamento: {str(e)}", exc_info=True)
        return jsonify({"error": f"Erro ao processar: {str(e)}"}), 500

@app.route("/validate", methods=["POST"])
def validate():
    result = {}
    
    if "pptx" in request.files:
        f = request.files["pptx"]
        try:
            prs = Presentation(io.BytesIO(f.read()))
            result["pptx"] = {"ok": True, "slides": len(prs.slides), "name": f.filename}
        except Exception as e:
            result["pptx"] = {"ok": False, "error": str(e)}
    
    if "zip" in request.files:
        f = request.files["zip"]
        try:
            photos = extract_photos_from_zip(f.read())
            result["zip"] = {
                "ok": True,
                "photos": len(photos),
                "names": [p[0] for p in photos[:5]],
                "name": f.filename,
            }
        except Exception as e:
            result["zip"] = {"ok": False, "error": str(e)}
    
    return jsonify(result)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    logger.info(f"Iniciando servidor na porta {port}")
    app.run(debug=False, port=port, host="0.0.0.0")
