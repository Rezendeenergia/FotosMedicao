import os
import io
import zipfile
import copy
import time
import logging
from concurrent.futures import ThreadPoolExecutor
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, static_folder=BASE_DIR)
CORS(app, expose_headers=["X-Photos-Used", "X-Slides-Filled", "X-Barramentos", "X-Processing-Time"])

app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

# Dimensões PADRÃO RETRATO: 7,62 x 10,16 cm
TARGET_W_CM = 7.62
TARGET_H_CM = 10.16
CM_TO_EMU = 360000
TARGET_W_EMU = int(TARGET_W_CM * CM_TO_EMU)
TARGET_H_EMU = int(TARGET_H_CM * CM_TO_EMU)

# Slots relatorio fotografico (4 fotos por slide)
PHOTO_SLOTS_4 = [
    {"name": "Retangulo 19", "x": 189290,  "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retangulo 41", "x": 3193696, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retangulo 15", "x": 6198102, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
    {"name": "Retangulo 8",  "x": 9202508, "y": 2478960, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},
]

# Slots base concretada — ordem: Poste | Barramento | Base
PHOTO_SLOTS_3 = [
    {"name": "Retangulo 41", "x": 4724400, "y": 2299850, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},  # Poste
    {"name": "Retangulo 19", "x": 720435,  "y": 2299851, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},  # Barramento
    {"name": "Retangulo 15", "x": 8728366, "y": 2299851, "cx": TARGET_W_EMU, "cy": TARGET_H_EMU},  # Base
]

BARRAMENTO_TEXTBOX = {
    "x": 5186217, "y": 1547498, "cx": 1819564, "cy": 369332
}

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'

@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")

@app.route("/<path:filename>")
def static_files(filename):
    return send_from_directory(BASE_DIR, filename)

@app.route("/health")
def health():
    return jsonify({"status": "ok", "message": "Backend Rezende Energia rodando!", "timestamp": time.time()})

def crop_and_resize(img_bytes, target_w_px, target_h_px):
    """Corte cover rapido: sem bordas brancas, preenche totalmente o espaco."""
    img = Image.open(io.BytesIO(img_bytes))
    if img.mode != "RGB":
        img = img.convert("RGB")
    img_w, img_h = img.size

    # Pre-reduz agressivamente para economizar RAM (max 1.5x o alvo)
    pre_w = int(target_w_px * 1.5)
    pre_h = int(target_h_px * 1.5)
    if img_w > pre_w or img_h > pre_h:
        img.thumbnail((pre_w, pre_h), Image.Resampling.BILINEAR)
        img_w, img_h = img.size

    scale = max(target_w_px / img_w, target_h_px / img_h)
    new_w = int(img_w * scale)
    new_h = int(img_h * scale)
    img_scaled = img.resize((new_w, new_h), Image.Resampling.BILINEAR)
    left = (new_w - target_w_px) // 2
    top  = (new_h - target_h_px) // 2
    img_cropped = img_scaled.crop((left, top, left + target_w_px, top + target_h_px))
    buf = io.BytesIO()
    img_cropped.save(buf, format="JPEG", quality=65, subsampling=2)
    del img, img_scaled  # libera RAM imediatamente
    return buf.getvalue()

def add_photo_to_slide(slide, slot, img_bytes, already_processed=False, is_landscape=None):
    """
    Insere foto no slot com orientacao inteligente:
    - Padrao: retrato 7,62 x 10,16 cm
    - Se a foto for paisagem (largura > altura): inverte slot para 10,16 x 7,62 cm
    Sem bordas brancas. Com efeito Predefinicao 1 do PowerPoint.
    """
    sp_tree = slide.shapes._spTree

    # Remove placeholder original
    for sp in list(sp_tree):
        if sp.tag.split("}")[-1] == "sp":
            cNvPr = sp.find(".//" + qn("p:cNvPr"))
            if cNvPr is not None and cNvPr.get("name") == slot["name"]:
                sp_tree.remove(sp)
                break

    # Verifica orientacao da foto
    if is_landscape is None:
        # Detecta da imagem original (nao foi pre-processada)
        try:
            probe = Image.open(io.BytesIO(img_bytes))
            is_landscape = probe.width > probe.height
        except Exception:
            is_landscape = False

    # Sempre usa as dimensoes exatas do slot (sem inverter, sem borda branca)
    # Foto esticada para preencher completamente o espaco definido
    final_cx = slot["cx"]
    final_cy = slot["cy"]
    final_x = slot["x"]
    final_y = slot["y"]

    if already_processed:
        img_resized = img_bytes  # ja foi pre-processado em paralelo
    else:
        target_w_px = int(final_cx / CM_TO_EMU * 2.54 * 96)
        target_h_px = int(final_cy / CM_TO_EMU * 2.54 * 96)
        img_resized = crop_and_resize(img_bytes, target_w_px, target_h_px)

    _, rId = slide.part.get_or_add_image_part(io.BytesIO(img_resized))

    # Efeito Predefinicao 1: sombra externa suave — sem linha de borda
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
      <a:off x="{final_x}" y="{final_y}"/>
      <a:ext cx="{final_cx}" cy="{final_cy}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:noFill/></a:ln>
    <a:effectLst>
      <a:outerShdw blurRad="63500" dist="38100" dir="2700000" algn="tl" rotWithShape="0">
        <a:prstClr val="black">
          <a:alpha val="40000"/>
        </a:prstClr>
      </a:outerShdw>
    </a:effectLst>
  </p:spPr>
</p:pic>'''

    sp_tree.append(etree.fromstring(pic_xml))


def set_barramento_number(slide, numero):
    for shape in slide.shapes:
        if shape.has_text_frame:
            tf = shape.text_frame
            for para in tf.paragraphs:
                for run in para.runs:
                    if any(c.isdigit() for c in run.text) or run.text.strip() == "":
                        run.text = numero
                        return
    for shape in slide.shapes:
        if shape.has_text_frame:
            tf = shape.text_frame
            if tf.paragraphs and tf.paragraphs[0].runs:
                tf.paragraphs[0].runs[0].text = numero
                return


def duplicate_slide(prs, slide_index):
    """Duplica slide corretamente usando SlidePart do python-pptx."""
    from pptx.opc.packuri import PackURI
    from pptx.parts.slide import SlidePart

    template = prs.slides[slide_index]

    # Cópia profunda do elemento XML do slide
    new_element = copy.deepcopy(template._element)

    # Determina partname único
    existing = [prs.slides[i].part.partname for i in range(len(prs.slides))]
    idx = len(prs.slides) + 1
    while PackURI(f'/ppt/slides/slide{idx}.xml') in existing:
        idx += 1
    new_partname = PackURI(f'/ppt/slides/slide{idx}.xml')

    # Cria SlidePart (não Part base) — necessário para ter o atributo .slide
    new_part = SlidePart(new_partname, template.part.content_type,
                         template.part.package, new_element)

    # Copia todas as relações (layout, imagens, etc.)
    for rel in template.part.rels.values():
        new_part.relate_to(rel.target_part, rel.reltype,
                           is_external=rel.is_external)

    # Registra o novo slide na apresentação
    rId = prs.slides.part.relate_to(
        new_part,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
    )

    # Adiciona à lista XML de slides
    sldIdLst = prs.slides._sldIdLst
    max_id = max((int(s.get('id')) for s in sldIdLst), default=255)
    sldId = etree.SubElement(
        sldIdLst,
        '{http://schemas.openxmlformats.org/presentationml/2006/main}sldId'
    )
    sldId.set('id', str(max_id + 1))
    sldId.set(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id',
        rId
    )
    return prs.slides[-1]


def remove_last_slide(prs):
    """Remove o último slide da apresentação."""
    sldIdLst = prs.slides._sldIdLst
    if len(sldIdLst) > 1:
        sldIdLst.remove(sldIdLst[-1])


def extract_photos_from_zip(zip_bytes):
    """Retorna lista de (nome, dados) — mantida para /validate e compatibilidade."""
    photos = []
    valid_ext = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = sorted([n for n in zf.namelist()
                        if not n.startswith('__MACOSX') and not os.path.basename(n).startswith('.')
                        and os.path.splitext(n.lower())[1] in valid_ext])
        for name in names:
            data = zf.read(name)
            if len(data) > 1000:
                photos.append((os.path.basename(name), data))
    return photos

def list_photo_names_in_zip(zip_bytes):
    """Retorna apenas os nomes das fotos validas, sem carregar dados na RAM."""
    valid_ext = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = sorted([n for n in zf.namelist()
                        if not n.startswith('__MACOSX') and not os.path.basename(n).startswith('.')
                        and os.path.splitext(n.lower())[1] in valid_ext
                        and zf.getinfo(n).file_size > 1000])
    return names


@app.route("/validate", methods=["POST"])
def validate():
    if 'zip' not in request.files:
        return jsonify({"error": "ZIP nao enviado"}), 400
    try:
        photos = extract_photos_from_zip(request.files['zip'].read())
        slides_needed = max(1, -(-len(photos) // 4))
        return jsonify({"photos": len(photos), "slides_needed": slides_needed, "names": [p[0] for p in photos[:20]]})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


# Dimensoes alvo em pixels para as fotos (pre-calculado uma vez)
_TARGET_W_PX = int(TARGET_W_EMU / CM_TO_EMU * 2.54 * 96)
_TARGET_H_PX = int(TARGET_H_EMU / CM_TO_EMU * 2.54 * 96)

def _preprocess_photo(args):
    """
    Processa uma foto (crop+resize) sequencial.
    Sempre usa dimensoes fixas do slot W x H, sem inverter orientacao.
    Retorna (idx, img_bytes_processado, is_landscape=False).
    """
    idx, img_bytes = args
    try:
        data = crop_and_resize(img_bytes, _TARGET_W_PX, _TARGET_H_PX)
        return idx, data, False
    except Exception as e:
        logger.warning(f"Foto {idx} ignorada: {e}")
        return idx, None, False

@app.route("/process", methods=["POST"])
def process_pptx():
    t0 = time.time()
    if 'pptx' not in request.files or 'zip' not in request.files:
        return jsonify({"error": "Arquivos PPTX e ZIP sao obrigatorios"}), 400
    pptx_bytes = request.files['pptx'].read()
    zip_bytes = request.files['zip'].read()
    try:
        # 1) Lista nomes sem carregar dados — O(1) de RAM
        photo_names = list_photo_names_in_zip(zip_bytes)
        if not photo_names:
            return jsonify({"error": "Nenhuma foto encontrada no ZIP"}), 400

        n_photos = min(len(photo_names), 100)
        photo_names = photo_names[:n_photos]
        slides_needed = max(1, -(-n_photos // 4))
        logger.info(f"Processando {n_photos} fotos em {slides_needed} slides...")

        # 2) Monta apresentacao com slides corretos
        prs = Presentation(io.BytesIO(pptx_bytes))
        del pptx_bytes  # libera RAM do PPTX original
        if len(prs.slides) == 0:
            return jsonify({"error": "Apresentacao sem slides"}), 400
        while len(prs.slides) < slides_needed:
            duplicate_slide(prs, 0)
        while len(prs.slides) > slides_needed:
            remove_last_slide(prs)

        # 3) Processa e insere foto a foto direto do ZIP — nunca acumula tudo na RAM
        photos_used = 0
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            del zip_bytes  # libera bytes brutos do ZIP, mantendo apenas o objeto ZipFile
            for photo_idx, name in enumerate(photo_names):
                slide_idx = photo_idx // 4
                slot_idx  = photo_idx % 4
                slide = prs.slides[slide_idx]
                slot  = PHOTO_SLOTS_4[slot_idx]
                try:
                    raw = zf.read(name)
                    _, data, lsc = _preprocess_photo((photo_idx, raw))
                    del raw  # libera original imediatamente
                    if data is not None:
                        add_photo_to_slide(slide, slot, data,
                                           already_processed=True, is_landscape=lsc)
                        del data
                        photos_used += 1
                except Exception as ex:
                    logger.warning(f"Foto {name} ignorada: {ex}")

        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        elapsed = time.time() - t0
        logger.info(f"Concluido em {elapsed:.1f}s — {photos_used} fotos inseridas")
        response = send_file(out, mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                             as_attachment=True, download_name='relatorio_preenchido.pptx')
        response.headers['X-Photos-Used'] = str(photos_used)
        response.headers['X-Slides-Filled'] = str(slides_needed)
        response.headers['X-Processing-Time'] = f"{elapsed:.2f}s"
        return response
    except Exception as e:
        logger.exception("Erro no processamento")
        return jsonify({"error": str(e)}), 500


@app.route("/process-base", methods=["POST"])
def process_base_concretada():
    """
    Base concretada: ordem Poste | Barramento | Base
    Dimensoes: 10,16 x 7,62 cm (paisagem, sem bordas brancas, com Predefinicao 1)
    """
    t0 = time.time()
    if 'pptx' not in request.files:
        return jsonify({"error": "Arquivo PPTX e obrigatorio"}), 400
    pptx_bytes = request.files['pptx'].read()
    numeros = request.form.getlist('numeros[]')
    if not numeros:
        return jsonify({"error": "Nenhum barramento enviado"}), 400
    try:
        prs = Presentation(io.BytesIO(pptx_bytes))
        if len(prs.slides) == 0:
            return jsonify({"error": "Apresentacao sem slides"}), 400
        n_barramentos = len(numeros)
        while len(prs.slides) < n_barramentos:
            duplicate_slide(prs, 0)
        while len(prs.slides) > n_barramentos:
            remove_last_slide(prs)
        for i, numero in enumerate(numeros):
            slide = prs.slides[i]
            # Ordem: poste_i → slot Poste, barramento_i → slot Barramento, base_i → slot Base
            fotos = [
                request.files.get(f'poste_{i}'),      # PHOTO_SLOTS_3[0] = Poste
                request.files.get(f'barramento_{i}'),  # PHOTO_SLOTS_3[1] = Barramento
                request.files.get(f'base_{i}'),        # PHOTO_SLOTS_3[2] = Base
            ]
            for slot, foto in zip(PHOTO_SLOTS_3, fotos):
                if foto:
                    add_photo_to_slide(slide, slot, foto.read())
            if numero:
                set_barramento_number(slide, numero)
        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        elapsed = time.time() - t0
        response = send_file(out, mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                             as_attachment=True, download_name='base_concretada_preenchida.pptx')
        response.headers['X-Barramentos'] = str(n_barramentos)
        response.headers['X-Processing-Time'] = f"{elapsed:.2f}s"
        return response
    except Exception as e:
        logger.exception("Erro no processamento de base")
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
