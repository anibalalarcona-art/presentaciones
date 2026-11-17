#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Combina dos .pptx en uno solo: todas las láminas de A, seguidas de todas
las láminas de B. Edita el paquete OOXML directamente (como el resto de
tools/) — no usa python-pptx, que no sabe copiar gráficos nativos entre
presentaciones sin perder la fuente de datos.

Cómo evita colisiones: cada parte de B (slideN.xml, chartN.xml, colorsN.xml,
styleN.xml, slideLayoutN.xml, slideMasterN.xml, imageN.png, theme1.xml) se
renumera para no pisar la numeración que ya usa A, y cada referencia a esas
partes —en los .rels y en [Content_Types].xml— se reescribe para apuntar al
nuevo número. El tema propio de B (y sus "theme override" por gráfico) se
copian intactos junto a su propio slide master, así que cada mitad del
archivo combinado conserva exactamente su diseño original.

Qué NO se copia de B (no hace falta, nada los referencia si no hay notas):
  - su notesMaster (B no tiene notas en ninguna lámina)
  - su theme2.xml (sólo lo usaba ese notesMaster)

Uso:
    python3 tools/merge_pptx.py --a "Laminas 11112026.pptx" \
        --b "Láminas_TC.pptx" --out "Laminas combinadas.pptx"
"""
import argparse
import re
import zipfile
from pathlib import Path

RENAME_TYPES = ("chart", "colors", "style", "slideLayout", "slideMaster", "slide", "image", "theme")
FNAME_RE = re.compile(r"\b(" + "|".join(RENAME_TYPES) + r")(\d+)(\.\w+)\b")
RID_RE = re.compile(r"\brId(\d+)\b")


def read_zip(path):
    z = zipfile.ZipFile(path)
    return {n: z.read(n) for n in z.namelist()}


def build_rename_map(parts_a, parts_b):
    """type -> {old_num: new_num} para que B no choque con los números que A ya usa."""
    def max_num(parts, typ, dirpat):
        best = 0
        for n in parts:
            m = re.match(dirpat + typ + r"(\d+)\.\w+$", n)
            if m:
                best = max(best, int(m.group(1)))
        return best

    dirs = {
        "slide": r"ppt/slides/",
        "slideLayout": r"ppt/slideLayouts/",
        "slideMaster": r"ppt/slideMasters/",
        "chart": r"ppt/charts/",
        "colors": r"ppt/charts/",
        "style": r"ppt/charts/",
        "image": r"ppt/media/",
        "theme": r"ppt/theme/",
    }

    rename = {}
    for typ, dirpat in dirs.items():
        max_a = max_num(parts_a, typ, dirpat)
        b_nums = sorted({int(m.group(1)) for n in parts_b for m in [re.match(dirpat + typ + r"(\d+)\.\w+$", n)] if m})
        if typ == "theme":
            # sólo renumeramos el theme PRINCIPAL de B (theme1.xml); theme2.xml
            # (el de su notesMaster, que no copiamos) se descarta más abajo.
            b_nums = [n for n in b_nums if n == 1]
        rename[typ] = {old: max_a + i + 1 for i, old in enumerate(b_nums)}
    return rename


def rename_fname(fname, rename_map):
    # "slideMaster1.xml.rels" tiene doble extensión: separa el ".rels" antes
    # de reconocer el patrón "(tipo)(número).xml", si no, nunca hace match
    # (fullmatch exige consumir TODO el nombre) y el archivo .rels de B
    # termina escribiéndose en la MISMA ruta que el de A, pisándolo.
    suffix = ""
    base = fname
    if base.endswith(".rels"):
        suffix, base = ".rels", base[: -len(".rels")]
    m = re.fullmatch(r"(" + "|".join(RENAME_TYPES) + r")(\d+)(\.\w+)", base)
    if not m:
        return fname
    typ, num, ext = m.group(1), int(m.group(2)), m.group(3)
    new_num = rename_map.get(typ, {}).get(num)
    if new_num is None:
        return fname
    return f"{typ}{new_num}{ext}{suffix}"


def rename_path(path, rename_map):
    parts = path.split("/")
    parts[-1] = rename_fname(parts[-1], rename_map)
    return "/".join(parts)


def rewrite_text(text, rename_map):
    def repl(m):
        typ, num, ext = m.group(1), int(m.group(2)), m.group(3)
        new_num = rename_map.get(typ, {}).get(num)
        return f"{typ}{new_num}{ext}" if new_num is not None else m.group(0)
    return FNAME_RE.sub(repl, text)


SKIP_FROM_B = {
    "ppt/notesMasters/notesMaster1.xml",
    "ppt/notesMasters/_rels/notesMaster1.xml.rels",
    "ppt/theme/theme2.xml",
}
BINARY_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".emf", ".wmf", ".ico")


def merge(a_path, b_path, out_path):
    parts_a = read_zip(a_path)
    parts_b_all = read_zip(b_path)
    parts_b = {n: d for n, d in parts_b_all.items() if n not in SKIP_FROM_B}

    rename_map = build_rename_map(parts_a.keys(), parts_b.keys())

    # --- ids globales: sldMasterId, sldId y sldLayoutId (estos últimos dentro
    #     de cada slideMaster propio) comparten un solo espacio de ids en todo
    #     el paquete, así que cualquier id nuevo para B debe evitar los tres a
    #     la vez. sldId además está limitado por el esquema a < 2^31, mientras
    #     que A ya usa valores enormes (~2^31) para sldMasterId/sldLayoutId:
    #     por eso van en dos "pools" separados, no uno compartido.
    pres_a_text = parts_a["ppt/presentation.xml"].decode("utf-8")
    sld_master_ids_a = [int(x) for x in re.findall(r'<p:sldMasterId id="(\d+)"', pres_a_text)]
    sld_ids_a = [int(x) for x in re.findall(r'<p:sldId id="(\d+)"', pres_a_text)]
    sld_layout_ids_a = []
    for name, data in parts_a.items():
        if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", name):
            sld_layout_ids_a += [int(x) for x in re.findall(r'<p:sldLayoutId id="(\d+)"', data.decode("utf-8"))]

    next_master_id = max(sld_master_ids_a + sld_layout_ids_a) + 1
    next_slide_id = max(sld_ids_a) + 1

    def take_next_master_id():
        nonlocal next_master_id
        v = next_master_id
        next_master_id += 1
        return v

    def take_next_slide_id():
        nonlocal next_slide_id
        v = next_slide_id
        next_slide_id += 1
        return v

    def remap_slidemaster_layout_ids(xml_text):
        # el slideMaster de B trae sus propios ids de sldLayoutId, que por
        # convención arrancan igual que los de A (ej: 2147483664) y chocan
        # con ellos; se reasignan desde el pool de "master ids". El r:id no
        # cambia: es una relación local al .rels propio del slideMaster.
        def repl(m):
            return f'<p:sldLayoutId id="{take_next_master_id()}" r:id="{m.group(1)}"'
        return re.sub(r'<p:sldLayoutId id="\d+" r:id="(rId\d+)"', repl, xml_text)

    # --- copia las partes de B, renombradas, reescribiendo su propio texto ---
    merged = dict(parts_a)
    for name, data in parts_b.items():
        if name in ("[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels",
                     "_rels/.rels", "docProps/app.xml", "docProps/core.xml", "docProps/thumbnail.jpeg",
                     "docMetadata/LabelInfo.xml", "ppt/presProps.xml", "ppt/viewProps.xml", "ppt/tableStyles.xml"):
            continue  # se tratan aparte (o se descartan: B no aporta nada nuevo ahí)
        new_name = rename_path(name, rename_map)
        if new_name.endswith(BINARY_EXT):
            merged[new_name] = data
        else:
            text = rewrite_text(data.decode("utf-8"), rename_map)
            if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", name):
                text = remap_slidemaster_layout_ids(text)
            merged[new_name] = text.encode("utf-8")

    # --- [Content_Types].xml: agrega los Override de B (renombrados), evita duplicar Default ---
    ct_a = parts_a["[Content_Types].xml"].decode("utf-8")
    ct_b = parts_b_all["[Content_Types].xml"].decode("utf-8")
    existing_defaults = set(re.findall(r'<Default Extension="([^"]+)"', ct_a))
    new_defaults = [m for m in re.findall(r'<Default Extension="[^"]+" ContentType="[^"]+"/>', ct_b)
                     if re.search(r'Extension="([^"]+)"', m).group(1) not in existing_defaults]
    skip_overrides = {rename_path("/" + n, rename_map) for n in SKIP_FROM_B if n.startswith("ppt/")}
    overrides_b = re.findall(r'<Override PartName="[^"]+" ContentType="[^"]+"/>', ct_b)
    overrides_b_renamed = []
    for ov in overrides_b:
        part = re.search(r'PartName="([^"]+)"', ov).group(1)
        if part in skip_overrides or part.lstrip("/") in SKIP_FROM_B:
            continue
        new_part = "/" + rename_path(part.lstrip("/"), rename_map)
        overrides_b_renamed.append(ov.replace(f'PartName="{part}"', f'PartName="{new_part}"'))
    ct_merged = ct_a.replace("</Types>", "".join(new_defaults) + "".join(overrides_b_renamed) + "</Types>")
    merged["[Content_Types].xml"] = ct_merged.encode("utf-8")

    # --- ppt/_rels/presentation.xml.rels: agrega las relaciones de B con rIds nuevos ---
    rels_a = parts_a["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    rels_b = parts_b_all["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    max_rid_a = max(int(n) for n in RID_RE.findall(rels_a))

    rid_map = {}
    b_rel_entries = re.findall(r'<Relationship Id="rId(\d+)"[^>]*Target="([^"]+)"[^>]*/>', rels_b)
    keep_entries = []
    next_rid = max_rid_a + 1
    for old_rid, target in b_rel_entries:
        fname = target.split("/")[-1]
        if fname in ("notesMaster1.xml",):
            continue  # no copiamos el notesMaster de B
        rid_map[old_rid] = str(next_rid)
        next_rid += 1

    def bump_rid_rels(m):
        old = m.group(1)
        return f'rId{rid_map[old]}' if old in rid_map else m.group(0)

    rels_b_body = re.sub(r"<\?xml[^>]*\?>\s*<Relationships[^>]*>", "", rels_b)
    rels_b_body = rels_b_body.replace("</Relationships>", "")
    # descarta el Relationship del notesMaster de B
    rels_b_body = re.sub(r'<Relationship Id="rId\d+"[^>]*Target="notesMasters/notesMaster1\.xml"[^>]*/>', "", rels_b_body)
    rels_b_body = re.sub(r'rId(\d+)', lambda m: ("rId" + rid_map[m.group(1)]) if m.group(1) in rid_map else m.group(0), rels_b_body)
    rels_b_body = rewrite_text(rels_b_body, rename_map)
    merged["ppt/_rels/presentation.xml.rels"] = (rels_a.replace("</Relationships>", rels_b_body + "</Relationships>")).encode("utf-8")

    # --- ppt/presentation.xml: agrega sldMasterId y sldId de B, con rIds e ids nuevos ---
    pres_a = pres_a_text
    pres_b = parts_b_all["ppt/presentation.xml"].decode("utf-8")

    def take_block(xml, tag):
        m = re.search(rf"<p:{tag}>(.*?)</p:{tag}>", xml, re.S)
        return m.group(1) if m else ""

    master_items_b = re.findall(r'<p:sldMasterId[^/]*/>', take_block(pres_b, "sldMasterIdLst"))
    slide_items_b = re.findall(r'<p:sldId[^/]*/>', take_block(pres_b, "sldIdLst"))

    def retarget_id_block(item, take_id):
        rid = re.search(r'r:id="rId(\d+)"', item).group(1)
        new_rid = rid_map.get(rid, rid)
        new_id = take_id()
        item = re.sub(r'r:id="rId\d+"', f'r:id="rId{new_rid}"', item)
        item = re.sub(r'id="\d+"', f'id="{new_id}"', item)
        return item

    new_master_items = "".join(retarget_id_block(i, take_next_master_id) for i in master_items_b)
    new_slide_items = "".join(retarget_id_block(i, take_next_slide_id) for i in slide_items_b)

    pres_merged = pres_a
    pres_merged = pres_merged.replace("</p:sldMasterIdLst>", new_master_items + "</p:sldMasterIdLst>")
    pres_merged = pres_merged.replace("</p:sldIdLst>", new_slide_items + "</p:sldIdLst>")
    merged["ppt/presentation.xml"] = pres_merged.encode("utf-8")

    # --- escribe el paquete combinado ---
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in merged.items():
            zout.writestr(name, data)

    n_slides_a = len(re.findall(r"<p:sldId ", pres_a))
    n_slides_b = len(slide_items_b)
    return n_slides_a, n_slides_b


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--a", required=True, help="Primer .pptx (sus láminas van primero)")
    ap.add_argument("--b", required=True, help="Segundo .pptx (sus láminas se agregan al final)")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    na, nb = merge(Path(args.a), Path(args.b), Path(args.out))
    print(f"Listo: {na} + {nb} = {na + nb} láminas -> {args.out}")


if __name__ == "__main__":
    main()
