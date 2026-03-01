import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import re
import json
import os
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
        "{{productName}}": str(basic.get("product_name", "")),
        "{{diseaseName}}": str(basic.get("disease_name", "")),
        "{{spec}}": str(basic.get("spec", "")),
        "{{objective}}": str(basic.get("objective", "")),
        "{{target}}": str(basic.get("target", "")),
        "{{demo}}": str(basic.get("demo", "")),
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

def replace_text_in_paragraph(paragraph, mapping):
    """段落内のプレースホルダーを置換し、フォント書式を維持する"""
    original_text = paragraph.text
    if not original_text or "{{" not in original_text:
        return

    new_text = original_text
    for k, v in mapping.items():
        new_text = new_text.replace(k, str(v))

    if new_text != original_text:
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for i in range(len(paragraph.runs) - 1, 0, -1):
                r = paragraph.runs[i]._r
                paragraph._p.remove(r)
        else:
            paragraph.text = new_text

def replace_text_in_shape(shape, mapping):
    """図形やグループ、表から再帰的にテキストを置換する"""
    if getattr(shape, "has_text_frame", False):
        for p in shape.text_frame.paragraphs:
            replace_text_in_paragraph(p, mapping)

    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                if getattr(cell, "text_frame", None):
                    for p in cell.text_frame.paragraphs:
                        replace_text_in_paragraph(p, mapping)

    if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
        for child_shape in shape.shapes:
            replace_text_in_shape(child_shape, mapping)

def replace_placeholders_in_pptx(template_file, data: dict) -> bytes:
    prs = Presentation(template_file)
    mapping = build_placeholder_map(data)

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_text_in_shape(shape, mapping)

    bio = BytesIO()
    prs.save(bio)
    return bio.getvalue()

# --- Streamlitの画面（UI）設定 ---
def main():
    st.title("企画書テンプレート自動生成アプリ")
    st.write("JSONデータを入力して、PPTXテンプレートに自動記入します。")

    # template.pptx が同じフォルダにあるか確認
    template_path = "template.pptx"
    template_file = None

    if os.path.exists(template_path):
        st.info(f"デフォルトのテンプレート `{template_path}` を使用します。")
        template_file = template_path
    else:
        st.warning("`template.pptx` が見つかりません。テンプレートファイルをアップロードしてください。")
        template_file = st.file_uploader("PPTXテンプレートをアップロード", type=["pptx"])

    json_input = st.text_area("JSONデータを貼り付けてください", height=300)

    if st.button("スライドを生成する"):
        if not template_file:
            st.error("テンプレートファイルがありません。")
            return
        if not json_input.strip():
            st.error("JSONデータが入力されていません。")
            return

        try:
            data = json.loads(json_input)
        except json.JSONDecodeError:
            st.error("JSONの形式が正しくありません。不要なカンマや括弧の抜け落ちがないか確認してください。")
            return

        with st.spinner("PPTXを生成中..."):
            try:
                output_pptx = replace_placeholders_in_pptx(template_file, data)
                st.success("生成が完了しました！")
                
                st.download_button(
                    label="📥 完成したPPTXをダウンロード",
                    data=output_pptx,
                    file_name="generated_proposal.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()
