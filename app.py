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

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, static_folder=BASE_DIR)
CORS(app, expose_headers=["X-Photos-Used", "X-Slides-Filled", "X-Barramentos", "X-Processing-Time"])

# Configurações
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

# ── Dimensões ──
TARGET_W_CM = 7.62
TARGET_H_CM = 10.16
CM_TO_EMU = 360000
TARGET_W_EMU = int(TARGET_W_CM * CM_TO_EMU)
TARGET_H_EMU = int(TARGET_H_CM * CM_TO_EMU)

TARGET_W_PX = int(TARGET_W_CM / 2.54 * 150)
TARGET_H_PX = int(TARGET_H_CM / 2.54 * 150)

# Slots relatório fotográfico (4 fotos por slide)
PHOTO_SLOTS_4 = [
    {"name": "Retângulo 19", "x": 189290,  "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 41", "x": 3193696, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 15", "x": 6198102, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 8",  "x": 9202508, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
]

# Slots base concretada (3 fotos por slide)
PHOTO_SLOTS_3 = [
    {"name": "Retângulo 19", "x": 720435,  "y": 2299851, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 41", "x": 4724400, "y": 2299850, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 15", "x": 8728366, "y": 2299851, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
]

# Posição da caixa de texto do número do barramento
BARRAMENTO_TEXTBOX = {
    "x": 5186217, "y": 1547498, "cx": 1819564, "cy": 369332
}

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'

# ── Rotas ──
@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")

@app.route("/<path:filename>")
def static_files(filename):
    """Serve arquivos estáticos da raiz (como logo-rezende.png)"""
    return send_from_directory(BASE_DIR, filename)

@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "message": "Backend Rezende Energia rodando!",
        "timestamp": time.time()
    })

# ── Utilitários de imagem (sem corte) ──
def resize_photo(img_bytes: bytes, w_cm=TARGET_W_CM, h_cm=TARGET_H_CM) -> bytes:
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        target_w_px = int(w_cm / 2.54 * 150)
        target_h_px = int(h_cm / 2.54 * 150)

        img_w, img_h = img.size
        scale = min(target_w_px / img_w, target_h_px / img_h)
        new_w = int(img_w * scale)
        new_h = int(img_h * scale)

        img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

        canvas = Image.new('RGB', (target_w_px, target_h_px), (255, 255, 255))
        offset_x = (target_w_px - new_w) // 2
        offset_y = (target_h_px - new_h) // 2
        canvas.paste(img_resized, (offset_x, offset_y))

        buf = io.BytesIO()
        canvas.save(buf, format="JPEG", quality=85, optimize=True)
        return buf.getvalue()
    except Exception as e:
        logger.error(f"Erro ao redimensionar imagem: {e}")
        raise

def add_photo_to_slide(slide, slot: dict, img_bytes: bytes):
    sp_tree = slide.shapes._spTree
    for sp in list(sp_tree):
        if sp.tag.split("}")[-1] == "sp":
            cNvPr = sp.find(".//" + qn("p:cNvPr"))
            if cNvPr is not None and cNvPr.get("name") == slot["name"]:
                sp_tree.remove(sp)
                break

    img_resized = resize_photo(img_bytes)
    _, rId = slide.part.get_or_add_image_part(io.BytesIO(img_resized))

    pic_xml = f'''<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
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
</p:pic>'''
    sp_tree.append(etree.fromstring(pic_xml))

def set_barramento_number(slide, numero: str):
    sp_tree = slide.shapes._spTree
    for sp in list(sp_tree):
        if sp.tag.split("}")[-1] == "sp":
            cNvPr = sp.find(".//" + qn("p:cNvPr"))
            if cNvPr is not None and "CaixaDeTexto" in cNvPr.get("name", ""):
                for t in sp.iter():
                    if t.tag.endswith("}t"):
                        t.text = numero
                        return

    tb = BARRAMENTO_TEXTBOX
    sp_xml = f"""<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="9998" name="CaixaDeTexto 27"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{tb['x']}" y="{tb['y']}"/><a:ext cx="{tb['cx']}" cy="{tb['cy']}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:gradFill>
      <a:gsLst>
        <a:gs pos="63000"><a:schemeClr val="accent2"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr></a:gs>
        <a:gs pos="4000"><a:schemeClr val="bg1"/></a:gs>
        <a:gs pos="86000"><a:schemeClr val="accent2"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:gs>
      </a:gsLst>
      <a:lin ang="5400000" scaled="1"/>
    </a:gradFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r><a:rPr lang="pt-BR" b="1" dirty="0"/><a:t>{numero}</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
    slide.shapes._spTree.append(etree.fromstring(sp_xml))

def duplicate_slide(prs, source_index):
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

def remove_last_slide(prs):
    sldIdLst = prs.slides._sldIdLst
    if len(sldIdLst) > 0:
        sldIdLst.remove(sldIdLst[-1])

def process_pptx(pptx_bytes: bytes, photos: list) -> bytes:
    start_time = time.time()
    logger.info(f"Iniciando relatório fotográfico com {len(photos)} fotos")
    prs = Presentation(io.BytesIO(pptx_bytes))
    n_photos = len(photos)
    n_slides_needed = (n_photos + 3) // 4
    n_slides_have = len(list(prs.slides)) - 1

    if n_slides_needed > n_slides_have:
        for _ in range(n_slides_needed - n_slides_have):
            duplicate_slide(prs, 1)

    all_slides = list(prs.slides)
    photo_idx = 0
    slides_used = 0
    for slide in all_slides[1:]:
        if photo_idx >= n_photos:
            break
        for slot in PHOTO_SLOTS_4:
            if photo_idx >= n_photos:
                break
            _, img_bytes = photos[photo_idx]
            add_photo_to_slide(slide, slot, img_bytes)
            photo_idx += 1
        slides_used += 1

    total_slides_now = len(list(prs.slides))
    for _ in range(total_slides_now - 1 - slides_used):
        remove_last_slide(prs)

    out = io.BytesIO()
    prs.save(out)
    logger.info(f"Relatório concluído em {time.time()-start_time:.2f}s")
    return out.getvalue()

def process_base_concretada(pptx_bytes: bytes, barramentos: list) -> bytes:
    start_time = time.time()
    logger.info(f"Iniciando base concretada com {len(barramentos)} barramentos")
    prs = Presentation(io.BytesIO(pptx_bytes))
    n_barramentos = len(barramentos)
    n_slides_have = len(list(prs.slides)) - 1

    if n_barramentos > n_slides_have:
        for _ in range(n_barramentos - n_slides_have):
            duplicate_slide(prs, 1)

    all_slides = list(prs.slides)
    for i, barr in enumerate(barramentos):
        slide = all_slides[i + 1]
        set_barramento_number(slide, barr["numero"])
        fotos = [
            (PHOTO_SLOTS_3[0], barr["barramento"]),
            (PHOTO_SLOTS_3[1], barr["poste"]),
            (PHOTO_SLOTS_3[2], barr["base"]),
        ]
        for slot, img_bytes in fotos:
            add_photo_to_slide(slide, slot, img_bytes)

    total_slides_now = len(list(prs.slides))
    for _ in range(total_slides_now - 1 - n_barramentos):
        remove_last_slide(prs)

    out = io.BytesIO()
    prs.save(out)
    logger.info(f"Base concluída em {time.time()-start_time:.2f}s")
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
        ])
        logger.info(f"Encontradas {len(names)} imagens no ZIP")
        for name in names[:100]:
            photos.append((os.path.basename(name), zf.read(name)))
    return photos

# ── Rotas de processamento ──
@app.route("/process", methods=["POST"])
def process():
    try:
        if "pptx" not in request.files or "zip" not in request.files:
            return jsonify({"error": "Arquivos .pptx e .zip são obrigatórios"}), 400
        pptx_bytes = request.files["pptx"].read()
        zip_bytes = request.files["zip"].read()
        photos = extract_photos_from_zip(zip_bytes)
        if not photos:
            return jsonify({"error": "Nenhuma imagem encontrada no ZIP"}), 400
        result_bytes = process_pptx(pptx_bytes, photos)
        response = send_file(
            io.BytesIO(result_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="relatorio_preenchido.pptx",
        )
        response.headers["X-Photos-Used"] = str(len(photos))
        response.headers["X-Slides-Filled"] = str((len(photos)+3)//4)
        return response
    except Exception as e:
        logger.error(f"Erro em /process: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.route("/process-base", methods=["POST"])
def process_base():
    try:
        if "pptx" not in request.files:
            return jsonify({"error": "Arquivo .pptx não enviado"}), 400
        pptx_bytes = request.files["pptx"].read()
        numeros = request.form.getlist("numeros[]")
        if not numeros:
            return jsonify({"error": "Nenhum número de barramento enviado"}), 400
        barramentos = []
        for i, num in enumerate(numeros):
            barr_key = f"barramento_{i}"
            poste_key = f"poste_{i}"
            base_key = f"base_{i}"
            if barr_key not in request.files or poste_key not in request.files or base_key not in request.files:
                return jsonify({"error": f"Fotos incompletas para barramento {i+1}"}), 400
            barramentos.append({
                "numero": num,
                "barramento": request.files[barr_key].read(),
                "poste": request.files[poste_key].read(),
                "base": request.files[base_key].read(),
            })
        result_bytes = process_base_concretada(pptx_bytes, barramentos)
        response = send_file(
            io.BytesIO(result_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="base_concretada_preenchida.pptx",
        )
        response.headers["X-Barramentos"] = str(len(barramentos))
        return response
    except Exception as e:
        logger.error(f"Erro em /process-base: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.route("/validate", methods=["POST"])
def validate():
    result = {}
    if "pptx" in request.files:
        try:
            prs = Presentation(io.BytesIO(request.files["pptx"].read()))
            result["pptx"] = {"ok": True, "slides": len(prs.slides)}
        except Exception as e:
            result["pptx"] = {"ok": False, "error": str(e)}
    if "zip" in request.files:
        try:
            photos = extract_photos_from_zip(request.files["zip"].read())
            result["zip"] = {"ok": True, "photos": len(photos)}
        except Exception as e:
            result["zip"] = {"ok": False, "error": str(e)}
    return jsonify(result)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
