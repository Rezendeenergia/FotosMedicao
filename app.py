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

# Configurações para produção
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

# ── Dimensões ──
TARGET_W_CM = 7.62
TARGET_H_CM = 10.16
CM_TO_EMU = 360000
TARGET_W_EMU = int(TARGET_W_CM * CM_TO_EMU)
TARGET_H_EMU = int(TARGET_H_CM * CM_TO_EMU)

# Dimensões em pixels para redimensionamento (150 DPI)
TARGET_W_PX = int(TARGET_W_CM / 2.54 * 150)
TARGET_H_PX = int(TARGET_H_CM / 2.54 * 150)

# Slots relatório fotográfico (4 fotos por slide)
PHOTO_SLOTS_4 = [
    {"name": "Retângulo 19", "x": 189290, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 41", "x": 3193696, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 15", "x": 6198102, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 8",  "x": 9202508, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
]

# Slots base concretada (3 fotos por slide: barramento, poste, base)
PHOTO_SLOTS_3 = [
    {"name": "Retângulo 19", "x": 720435,  "y": 2299851, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 41", "x": 4724400, "y": 2299850, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retângulo 15", "x": 8728366, "y": 2299851, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
]

# Posição e tamanho da caixa de texto do número do barramento
BARRAMENTO_TEXTBOX = {
    "x": 5186217, "y": 1547498, "cx": 1819564, "cy": 369332
}

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

# ── Frontend ──
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

# ── Utilitários de imagem otimizados ──
def resize_photo(img_bytes: bytes, w_cm=TARGET_W_CM, h_cm=TARGET_H_CM) -> bytes:
    """Redimensiona imagem com crop central e compressão otimizada"""
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        target_w_px = int(w_cm / 2.54 * 150)
        target_h_px = int(h_cm / 2.54 * 150)
        
        # Redimensionar mantendo proporção
        img.thumbnail((target_w_px * 2, target_h_px * 2), Image.Resampling.LANCZOS)
        
        # Crop central para o tamanho exato
        width, height = img.size
        left = max(0, (width - target_w_px) / 2)
        top = max(0, (height - target_h_px) / 2)
        right = left + target_w_px
        bottom = top + target_h_px
        
        img = img.crop((int(left), int(top), int(right), int(bottom)))
        
        # Salvar com compressão otimizada
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85, optimize=True)
        return buf.getvalue()
        
    except Exception as e:
        logger.error(f"Erro ao redimensionar imagem: {e}")
        raise

def add_photo_to_slide(slide, slot: dict, img_bytes: bytes):
    """Remove o placeholder e insere a foto no slot correto"""
    sp_tree = slide.shapes._spTree
    
    # Encontrar e remover o placeholder
    for sp in list(sp_tree):
        if sp.tag.split("}")[-1] == "sp":
            cNvPr = sp.find(".//" + qn("p:cNvPr"))
            if cNvPr is not None and cNvPr.get("name") == slot["name"]:
                sp_tree.remove(sp)
                break
    
    # Redimensionar e adicionar imagem
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
    """Atualiza o número do barramento na CaixaDeTexto do slide."""
    sp_tree = slide.shapes._spTree
    for sp in list(sp_tree):
        if sp.tag.split("}")[-1] == "sp":
            cNvPr = sp.find(".//" + qn("p:cNvPr"))
            if cNvPr is not None and "CaixaDeTexto" in cNvPr.get("name", ""):
                # Atualizar o texto
                for t in sp.iter():
                    if t.tag.endswith("}t"):
                        t.text = numero
                        return
    
    # Se não encontrou, cria a caixa do zero (para slides clonados)
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

# ── Gerenciamento de slides ──
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

# ══════════════════════════════════════════════
# RELATÓRIO FOTOGRÁFICO (4 fotos/slide)
# ══════════════════════════════════════════════
def process_pptx(pptx_bytes: bytes, photos: list) -> bytes:
    """Processa o relatório fotográfico com 4 fotos por slide"""
    start_time = time.time()
    logger.info(f"Iniciando relatório fotográfico com {len(photos)} fotos")
    
    prs = Presentation(io.BytesIO(pptx_bytes))
    n_photos = len(photos)
    n_slides_needed = (n_photos + 3) // 4
    n_slides_have = len(list(prs.slides)) - 1

    logger.info(f"Slides necessários: {n_slides_needed}, disponíveis: {n_slides_have}")

    if n_slides_needed > n_slides_have:
        slides_to_add = n_slides_needed - n_slides_have
        logger.info(f"Duplicando {slides_to_add} slides")
        for i in range(slides_to_add):
            duplicate_slide(prs, 1)
            if (i + 1) % 5 == 0:
                logger.info(f"  {i + 1}/{slides_to_add} slides duplicados")

    all_slides = list(prs.slides)
    photo_idx = 0
    slides_used = 0
    
    for slide_idx, slide in enumerate(all_slides[1:], start=1):
        if photo_idx >= n_photos:
            break
        for slot in PHOTO_SLOTS_4:
            if photo_idx >= n_photos:
                break
            _, img_bytes = photos[photo_idx]
            add_photo_to_slide(slide, slot, img_bytes)
            photo_idx += 1
        slides_used += 1
        
        if slide_idx % 5 == 0:
            logger.info(f"Progresso: {photo_idx}/{n_photos} fotos inseridas em {slides_used} slides")

    total_slides_now = len(list(prs.slides))
    slides_to_remove = total_slides_now - 1 - slides_used
    if slides_to_remove > 0:
        logger.info(f"Removendo {slides_to_remove} slides vazios")
        for _ in range(slides_to_remove):
            remove_last_slide(prs)

    out = io.BytesIO()
    prs.save(out)
    
    total_time = time.time() - start_time
    logger.info(f"Relatório fotográfico concluído em {total_time:.2f} segundos")
    return out.getvalue()

# ══════════════════════════════════════════════
# BASE CONCRETADA (3 fotos/slide + nº barramento)
# ══════════════════════════════════════════════
def process_base_concretada(pptx_bytes: bytes, barramentos: list) -> bytes:
    """
    barramentos: lista de dicts {numero: str, barramento: bytes, poste: bytes, base: bytes}
    """
    start_time = time.time()
    logger.info(f"Iniciando base concretada com {len(barramentos)} barramentos")
    
    prs = Presentation(io.BytesIO(pptx_bytes))
    n_barramentos = len(barramentos)
    n_slides_have = len(list(prs.slides)) - 1

    # Expandir slides se necessário
    if n_barramentos > n_slides_have:
        slides_to_add = n_barramentos - n_slides_have
        logger.info(f"Duplicando {slides_to_add} slides")
        for i in range(slides_to_add):
            duplicate_slide(prs, 1)
            if (i + 1) % 5 == 0:
                logger.info(f"  {i + 1}/{slides_to_add} slides duplicados")

    all_slides = list(prs.slides)
    
    for i, barr in enumerate(barramentos):
        slide = all_slides[i + 1]  # pula capa
        
        logger.info(f"Processando barramento {i+1}/{n_barramentos}: {barr['numero']}")
        
        # Número do barramento
        set_barramento_number(slide, barr["numero"])

        # 3 fotos
        fotos = [
            (PHOTO_SLOTS_3[0], barr["barramento"]),
            (PHOTO_SLOTS_3[1], barr["poste"]),
            (PHOTO_SLOTS_3[2], barr["base"]),
        ]
        for slot, img_bytes in fotos:
            add_photo_to_slide(slide, slot, img_bytes)

    # Remover slides vazios
    total_slides_now = len(list(prs.slides))
    slides_to_remove = total_slides_now - 1 - n_barramentos
    if slides_to_remove > 0:
        logger.info(f"Removendo {slides_to_remove} slides vazios")
        for _ in range(slides_to_remove):
            remove_last_slide(prs)

    out = io.BytesIO()
    prs.save(out)
    
    total_time = time.time() - start_time
    logger.info(f"Base concretada concluída em {total_time:.2f} segundos")
    return out.getvalue()

# ── Utilitários ZIP ──
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
        for name in names[:100]:  # Limitar a 100 fotos
            photos.append((os.path.basename(name), zf.read(name)))
    return photos

# ══════════════════════════════════════════════
# ROTAS API
# ══════════════════════════════════════════════
@app.route("/process", methods=["POST"])
def process():
    """Endpoint para relatório fotográfico"""
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

        pptx_bytes = pptx_file.read()
        zip_bytes = zip_file.read()

        photos = extract_photos_from_zip(zip_bytes)
        
        if not photos:
            return jsonify({"error": "Nenhuma imagem encontrada no ZIP"}), 400

        result_bytes = process_pptx(pptx_bytes, photos)

        n_photos = len(photos)
        n_slides = (n_photos + 3) // 4
        processing_time = time.time() - start_time

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

@app.route("/process-base", methods=["POST"])
def process_base():
    """Endpoint para base concretada"""
    start_time = time.time()
    
    try:
        if "pptx" not in request.files:
            return jsonify({"error": "Arquivo .pptx não enviado"}), 400

        pptx_bytes = request.files["pptx"].read()

        # Coletar barramentos do form
        numeros = request.form.getlist("numeros[]")
        if not numeros:
            return jsonify({"error": "Nenhum número de barramento enviado"}), 400

        logger.info(f"Processando {len(numeros)} barramentos")
        
        barramentos = []
        for i, numero in enumerate(numeros):
            barr_key = f"barramento_{i}"
            poste_key = f"poste_{i}"
            base_key = f"base_{i}"

            if barr_key not in request.files or poste_key not in request.files or base_key not in request.files:
                return jsonify({"error": f"Fotos incompletas para barramento {i+1} (nº {numero})"}), 400

            barramentos.append({
                "numero": numero,
                "barramento": request.files[barr_key].read(),
                "poste": request.files[poste_key].read(),
                "base": request.files[base_key].read(),
            })

        result_bytes = process_base_concretada(pptx_bytes, barramentos)
        
        processing_time = time.time() - start_time
        logger.info(f"Base concretada processada em {processing_time:.2f} segundos")

        response = send_file(
            io.BytesIO(result_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="base_concretada_preenchida.pptx",
        )
        response.headers["X-Barramentos"] = str(len(barramentos))
        response.headers["X-Processing-Time"] = f"{processing_time:.2f}"
        return response
        
    except Exception as e:
        logger.error(f"Erro no processamento da base: {str(e)}", exc_info=True)
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
            result["zip"] = {"ok": True, "photos": len(photos), "names": [p[0] for p in photos[:5]], "name": f.filename}
        except Exception as e:
            result["zip"] = {"ok": False, "error": str(e)}
    return jsonify(result)

@app.route("/debug-template", methods=["POST"])
def debug_template():
    """Endpoint para debug - lista todos os shapes do template"""
    if "pptx" not in request.files:
        return jsonify({"error": "PPTX não enviado"}), 400
    
    pptx_file = request.files["pptx"]
    try:
        prs = Presentation(io.BytesIO(pptx_file.read()))
        
        shapes_info = []
        for i, slide in enumerate(prs.slides):
            slide_info = {"index": i, "shapes": []}
            for shape in slide.shapes:
                if hasattr(shape, "name"):
                    slide_info["shapes"].append(shape.name)
            shapes_info.append(slide_info)
        
        return jsonify({
            "total_slides": len(prs.slides),
            "slides": shapes_info,
            "photo_slots_4": PHOTO_SLOTS_4,
            "photo_slots_3": PHOTO_SLOTS_3
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    logger.info(f"Iniciando servidor na porta {port}")
    app.run(debug=False, port=port, host="0.0.0.0")
