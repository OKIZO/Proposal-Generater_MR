from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import re
import json
from io import BytesIO

PLACEHOLDER_RE = re.compile(r"\{\{[^}]+\}\}")

def _safe_get(d, *keys, default=""):
    cur = d
    for k in keys:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur if cur is not None else default

def _arr_get(a, idx, default=""):
    if isinstance(a, list) and len(a) > idx and a[idx] is not None:
        return a[idx]
    return default

def build_placeholder_map(data: dict) -> dict:
    basic = data.get("basic", {})
    materials = data.get("materials", {})
    methods = data.get("methods", {})
    md = data.get("md_result", {})
    target_info = data.get("target_info", {})
    focus = data.get("focus", {})

    motifs = _safe_get(materials, "motifs", default=[])
    worlds = _safe_get(materials, "worldview_tags", default=[])
    tones = _safe_get(methods, "tone_tags", default=[])
    functional = _safe_get(md, "functional_elements", default=[])

    insight_lines = _safe_get(md, "target_insight_lines", default=[])
    if not isinstance(insight_lines, list):
        insight_lines = []

    branding = basic.get("branding", "")
    if branding == "branded":
        branding_label = "ブランデッド"
    elif branding == "unbranded":
        branding_label = "アンブランデッド"
    else:
        branding_label = str(branding)

    mapping = {
        "{{branding}}": branding_label,
        "{{itemName}}": str(basic.get("item_name", "")),
        
        # --- 不足していた変数を追加 ---
        "{{productName}}": str(basic.get("product_name", "")),
        "{{diseaseName}}": str(basic.get("disease_name", "")),
        "{{spec}}": str(basic.get("spec", "")),
        "{{objective}}": str(basic.get("objective", "")),
        "{{target}}": str(basic.get("target", "")),
        "{{demo}}": str(basic.get("demo", "")),
        # ------------------------------

        "{{brandPrimary}}": str(basic.get("brand_primary", "")),
        "{{brandSecondary}}": str(basic.get("brand_secondary", "")),
        "{{brandElement}}": str(basic.get("brand_elements", "")),

        "{{creativeConcept}}": str(md.get("creative_concept", "")),
        "{{direction1}}": _arr_get(md.get("direction_short", []), 0, ""),
        "{{direction2}}": _arr_get(md.get("direction_short", []), 1, ""),
        "{{direction3}}": _arr_get(md.get("direction_short", []), 2, ""),

        "{{focusCenter}}": str(target_info.get("focusCenter", "")),
        "{{treatmentPhase}}": str(target_info.get("treatmentPhase", "")),

        "{{targetInsightLine1}}": _arr_get(insight_lines, 0, ""),
        "{{targetInsightLine2}}": _arr_get(insight_lines, 1, ""),

        "{{emotion1}}": _arr_get(_safe_get(md, "reasons", "emotion_keywords", default=[]), 0, ""),
        "{{emotion2}}": _arr_get(_safe_get(md, "reasons", "emotion_keywords", default=[]), 1, ""),
        "{{emotion3}}": _arr_get(_safe_get(md, "reasons", "emotion_keywords", default=[]), 2, ""),

        "{{behavior1}}": _arr_get(_safe_get(md, "reasons", "behavior_keywords", default=[]), 0, ""),
        "{{behavior2}}": _arr_get(_safe_get(md, "reasons", "behavior_keywords", default=[]), 1, ""),
        "{{behavior3}}": _arr_get(_safe_get(md, "reasons", "behavior_keywords", default=[]), 2, ""),

        "{{info1}}": _arr_get(_safe_get(md, "reasons", "info_need_type", default=[]), 0, ""),
        "{{info2}}": _arr_get(_safe_get(md, "reasons", "info_need_type", default=[]), 1, ""),
        "{{info3}}": _arr_get(_safe_get(md, "reasons", "info_need_type", default=[]), 2, ""),

        "{{motif1}}": _arr_get(motifs, 0, ""),
        "{{motif2}}": _arr_get(motifs, 1, ""),
        "{{motif3}}": _arr_get(motifs, 2, ""),

        "{{world1}}": _arr_get(worlds, 0, ""),
        "{{world2}}": _arr_get(worlds, 1, ""),
        "{{world3}}": _arr_get(worlds, 2, ""),

        "{{tone1}}": _arr_get(tones, 0, ""),
        "{{tone2}}": _arr_get(tones, 1, ""),
        "{{tone3}}": _arr_get(tones, 2, ""),

        "{{focusSummary}}": str(methods.get("focus_summary", "")),

        "{{functional1}}": _arr_get(functional, 0, ""),
        "{{functional2}}": _arr_get(functional, 1, ""),
        "{{functional3}}": _arr_get(functional, 2, ""),
    }

    return mapping

def replace_text_in_shape(shape, mapping):
    """図形、グループ化された図形、表の中のテキストを再帰的に置換する関数"""
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            if not p.text or "{{" not in p.text:
                continue
            new_text = p.text
            for k, v in mapping.items():
                new_text = new_text.replace(k, str(v))
            if new_text != p.text:
                p.text = new_text

    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    for p in cell.text_frame.paragraphs:
                        if not p.text or "{{" not in p.text:
                            continue
                        new_text = p.text
                        for k, v in mapping.items():
                            new_text = new_text.replace(k, str(v))
                        if new_text != p.text:
                            p.text = new_text

    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child_shape in shape.shapes:
            replace_text_in_shape(child_shape, mapping)

def replace_placeholders_in_pptx(template_path: str, data: dict) -> bytes:
    prs = Presentation(template_path)
    mapping = build_placeholder_map(data)

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_text_in_shape(shape, mapping)

    bio = BytesIO()
    prs.save(bio)
    return bio.getvalue()
