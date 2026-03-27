import os
import io
import zipfile
import copy
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image, ImageOps

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, static_folder=BASE_DIR)
CORS(app, expose_headers=["X-Photos-Used", "X-Slides-Filled"])

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
    return jsonify({"status": "ok", "message": "Backend Rezende Energia rodando!"})

# ── Utilitários de imagem ──
def resize_photo(img_bytes: bytes) -> bytes:
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    target_w_px = int(TARGET_W_CM / 2.54 * 150)
    target_h_px = int(TARGET_H_CM / 2.54 * 150)
    img = ImageOps.fit(img, (target_w_px, target_h_px), method=Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    return buf.getvalue()

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
    """
    Clona o slide no source_index (incluindo fundo e shapes).
    Adiciona o clone ao final da apresentação.
    """
    source_slide = prs.slides[source_index]
    new_slide = prs.slides.add_slide(source_slide.slide_layout)

    # Substituir shapes do novo slide pelos do template
    sp_tree_src = copy.deepcopy(source_slide.shapes._spTree)
    sp_tree_new = new_slide.shapes._spTree
    for child in list(sp_tree_new):
        sp_tree_new.remove(child)
    for child in list(sp_tree_src):
        sp_tree_new.append(child)

    # Copiar background (gradiente laranja do template)
    bg_src = source_slide._element.find(f'{{{P_NS}}}cSld/{{{P_NS}}}bg')
    if bg_src is not None:
        bg_copy = copy.deepcopy(bg_src)
        cSld = new_slide._element.find(f'{{{P_NS}}}cSld')
        existing_bg = cSld.find(f'{{{P_NS}}}bg')
        if existing_bg is not None:
            cSld.remove(existing_bg)
        cSld.insert(0, bg_copy)

    # Compartilhar relacionamentos de imagem (logos)
    for rel in source_slide.part.rels.values():
        if 'image' in rel.reltype:
            new_slide.part.rels._rels[rel.rId] = rel

    return new_slide


def remove_last_slide(prs: Presentation):
    """Remove o último slide da apresentação."""
    sldIdLst = prs.slides._sldIdLst
    sldIdLst.remove(sldIdLst[-1])


def process_pptx(pptx_bytes: bytes, photos: list) -> bytes:
    """
    Lógica dinâmica:
    - Menos de 100 fotos: remove slides de foto vazios
    - Mais de 100 fotos: cria novos slides clonando o slide 2
    """
    prs = Presentation(io.BytesIO(pptx_bytes))

    n_photos        = len(photos)
    n_slides_needed = (n_photos + 3) // 4        # slides de foto necessários
    n_slides_have   = len(list(prs.slides)) - 1  # slides de foto no template (excl. capa)

    # ── Expandir se precisar de mais slides ──
    if n_slides_needed > n_slides_have:
        for _ in range(n_slides_needed - n_slides_have):
            duplicate_slide(prs, 1)  # sempre clona o slide 2 (índice 1)

    # ── Preencher fotos ──
    all_slides  = list(prs.slides)
    photo_idx   = 0
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

    # ── Remover slides de foto vazios ──
    total_slides_now = len(list(prs.slides))
    slides_to_remove = total_slides_now - 1 - slides_used
    for _ in range(slides_to_remove):
        remove_last_slide(prs)

    out = io.BytesIO()
    prs.save(out)
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
        for name in names:
            photos.append((os.path.basename(name), zf.read(name)))
    return photos


# ── Rotas API ──
@app.route("/process", methods=["POST"])
def process():
    if "pptx" not in request.files:
        return jsonify({"error": "Arquivo .pptx não enviado"}), 400
    if "zip" not in request.files:
        return jsonify({"error": "Arquivo .zip não enviado"}), 400

    pptx_file = request.files["pptx"]
    zip_file  = request.files["zip"]

    if not pptx_file.filename.lower().endswith(".pptx"):
        return jsonify({"error": "O arquivo deve ser .pptx"}), 400
    if not zip_file.filename.lower().endswith(".zip"):
        return jsonify({"error": "O pacote de fotos deve ser .zip"}), 400

    pptx_bytes = pptx_file.read()
    zip_bytes  = zip_file.read()

    try:
        photos = extract_photos_from_zip(zip_bytes)
    except Exception as e:
        return jsonify({"error": f"Erro ao ler o ZIP: {str(e)}"}), 400

    if not photos:
        return jsonify({"error": "Nenhuma imagem encontrada no ZIP"}), 400

    try:
        result_bytes = process_pptx(pptx_bytes, photos)
    except Exception as e:
        import traceback
        return jsonify({"error": f"Erro ao processar PPTX: {str(e)}", "trace": traceback.format_exc()}), 500

    n_photos = len(photos)
    n_slides  = (n_photos + 3) // 4

    response = send_file(
        io.BytesIO(result_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name="relatorio_preenchido.pptx",
    )
    response.headers["X-Photos-Used"]   = str(n_photos)
    response.headers["X-Slides-Filled"] = str(n_slides)
    return response


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
    print("=" * 55)
    print("  Rezende Energia — Automação de Relatório Fotográfico")
    print("=" * 55)
    print("  Acesse no navegador: http://127.0.0.1:5050")
    print("  Para encerrar: Ctrl + C")
    print("=" * 55)
    app.run(debug=True, port=5050, host="127.0.0.1")
