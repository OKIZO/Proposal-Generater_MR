import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import re
import json
from io import BytesIO

PLACEHOLDER_RE = re.compile(r"\{\{[^}]+\}\}")

# --- PPTXの裏側処理 ---
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

    branding = basic.get("branding", "")
    branding_label = "ブランデッド" if branding == "branded" else "アンブランデッド" if branding == "unbranded" else str(branding)

    mapping = {
        "{{branding}}": branding_label,
        "{{itemName}}": str(basic.get("item_name", "")),
        "{{productName}}": str(basic.get("product_name", "")),
        "{{diseaseName}}": str(basic.get("disease_name", "")),
        "{{spec}}": str(basic.get("item_spec", "")),
        "{{objective}}": str(basic.get("objective", "")),
        "{{target}}": str(basic.get("target", "")),
        "{{demo}}": "", 
        "{{brandPrimary}}": str(basic.get("brand_primary", "")),
        "{{brandSecondary}}": str(basic.get("brand_secondary", "")),
        "{{brandElement}}": str(basic.get("brand_elements", "")),
        "{{creativeConcept}}": str(md.get("visual_concept", "")),
        "{{direction1}}": _arr_get(md.get("direction_short", []), 0, ""),
        "{{direction2}}": _arr_get(md.get("direction_short", []), 1, ""),
        "{{direction3}}": _arr_get(md.get("direction_short", []), 2, ""),
        "{{focusCenter}}": str(target_info.get("focusCenter", "")),
        "{{treatmentPhase}}": str(target_info.get("infoPhase", "")),
        "{{targetInsightLine1}}": str(_safe_get(md, "target_insight", "line1", default="")),
        "{{targetInsightLine2}}": str(_safe_get(md, "target_insight", "line2", default="")),
        "{{emotion1}}": _arr_get(_safe_get(md, "reasons", "emotion_keywords", default=[]), 0, ""),
        "{{emotion2}}": _arr_get(_safe_get(md, "reasons", "emotion_keywords", default=[]), 1, ""),
        "{{emotion3}}": _arr_get(_safe_get(md, "reasons", "emotion_keywords", default=[]), 2, ""),
        "{{behavior1}}": _arr_get(_safe_get(md, "reasons", "behavior_keywords", default=[]), 0, ""),
        "{{behavior2}}": _arr_get(_safe_get(md, "reasons", "behavior_keywords", default=[]), 1, ""),
        "{{behavior3}}": _arr_get(_safe_get(md, "reasons", "behavior_keywords", default=[]), 2, ""),
        "{{info1}}": _arr_get(_safe_get(md, "reasons", "info_need_type", default=[]), 0, ""),
        "{{info2}}": _arr_get(_safe_get(md, "reasons", "info_need_type", default=[]), 1, ""),
        "{{info3}}": _arr_get(_safe_get(md, "reasons", "info_need_type", default=[]), 2, ""),
        "{{motif1}}": _arr_get(_safe_get(materials, "motifs", default=[]), 0, ""),
        "{{motif2}}": _arr_get(_safe_get(materials, "motifs", default=[]), 1, ""),
        "{{motif3}}": _arr_get(_safe_get(materials, "motifs", default=[]), 2, ""),
        "{{world1}}": _arr_get(_safe_get(materials, "worldview_tags", default=[]), 0, ""),
        "{{world2}}": _arr_get(_safe_get(materials, "worldview_tags", default=[]), 1, ""),
        "{{world3}}": _arr_get(_safe_get(materials, "worldview_tags", default=[]), 2, ""),
        "{{tone1}}": _arr_get(_safe_get(methods, "tone_tags", default=[]), 0, ""),
        "{{tone2}}": _arr_get(_safe_get(methods, "tone_tags", default=[]), 1, ""),
        "{{tone3}}": _arr_get(_safe_get(methods, "tone_tags", default=[]), 2, ""),
        "{{focusSummary}}": str(methods.get("focus_summary", "")),
        "{{functional1}}": _arr_get(_safe_get(md, "functional_elements", default=[]), 0, ""),
        "{{functional2}}": _arr_get(_safe_get(md, "functional_elements", default=[]), 1, ""),
        "{{functional3}}": _arr_get(_safe_get(md, "functional_elements", default=[]), 2, ""),
    }
    return mapping

def replace_text_in_paragraph(paragraph, mapping):
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

# 画像をスライドに貼り付ける関数（最大6枚をグリッド配置）
def insert_images_to_slide(prs, slide_index, image_files):
    # スライドが存在しない、または画像がない場合はスキップ
    if not image_files or slide_index >= len(prs.slides):
        return
    
    slide = prs.slides[slide_index]
    image_files = image_files[:6] # 強制的に最大6枚までに制限
    
    # スライドのサイズ（EMU単位）を取得して配置位置を計算
    margin_x = prs.slide_width * 0.05
    margin_top = prs.slide_height * 0.25 # 上部のタイトルエリアを開ける
    usable_w = prs.slide_width - (margin_x * 2)
    usable_h = prs.slide_height - margin_top - (prs.slide_height * 0.05)
    
    col_w = usable_w / 3 # 3列
    row_h = usable_h / 2 # 2段
    
    for idx, img_file in enumerate(image_files):
        c = idx % 3  # 列番号 (0, 1, 2)
        r = idx // 3 # 行番号 (0, 1)
        
        # 各画像の左上座標を計算
        cell_x = int(margin_x + c * col_w)
        cell_y = int(margin_top + r * row_h)
        
        # セルの幅の90%を画像の幅に設定（少し余白をもたせる）
        img_width = int(col_w * 0.9)
        
        try:
            img_stream = BytesIO(img_file.getvalue())
            slide.shapes.add_picture(img_stream, cell_x, cell_y, width=img_width)
        except Exception as e:
            st.warning(f"画像の貼り付けに失敗しました（{img_file.name}）: {e}")

def replace_placeholders_in_pptx(template_path: str, data: dict, images: dict) -> bytes:
    prs = Presentation(template_path)
    mapping = build_placeholder_map(data)

    # 1. テキストの置換
    for slide in prs.slides:
        for shape in slide.shapes:
            replace_text_in_shape(shape, mapping)

    # 2. 画像の貼り付け (スライド番号は 0から始まるため、スライド6 = index5)
    insert_images_to_slide(prs, 5, images.get("A", [])) # スライド6へA案
    insert_images_to_slide(prs, 6, images.get("B", [])) # スライド7へB案
    insert_images_to_slide(prs, 7, images.get("C", [])) # スライド8へC案

    bio = BytesIO()
    prs.save(bio)
    return bio.getvalue()

# --- Streamlit UI（画面表示部分） ---
def main():
    st.set_page_config(page_title="企画書自動生成アプリ", layout="wide")
    st.title("企画書自動生成アプリ")

    st.subheader("1. 画像のアップロード（各案最大6枚まで）")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        # accept_multiple_files=True で複数選択・ドラッグ可能に
        img_a_list = st.file_uploader("A案の画像", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        if len(img_a_list) > 6:
            st.warning("A案の画像が6枚を超えています。最初の6枚のみが使用されます。")
            
    with col2:
        img_b_list = st.file_uploader("B案の画像", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        if len(img_b_list) > 6:
            st.warning("B案の画像が6枚を超えています。最初の6枚のみが使用されます。")
            
    with col3:
        img_c_list = st.file_uploader("C案の画像", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        if len(img_c_list) > 6:
            st.warning("C案の画像が6枚を超えています。最初の6枚のみが使用されます。")

    st.subheader("2. JSONデータの入力")
    json_input = st.text_area("JSONデータを貼り付けてください", height=300)

    if st.button("企画書を生成する", type="primary"):
        if not json_input.strip():
            st.error("JSONデータが入力されていません。")
            return
            
        try:
            data = json.loads(json_input)
        except json.JSONDecodeError:
            st.error("JSONの形式が正しくありません。")
            return

        # アップロードされた画像をまとめる辞書
        images_dict = {
            "A": img_a_list,
            "B": img_b_list,
            "C": img_c_list
        }

        with st.spinner("PPTXを生成中..."):
            try:
                # template.pptx が同じフォルダにある前提
                output_pptx = replace_placeholders_in_pptx("template.pptx", data, images_dict)
                st.success("テキストの置換と画像の挿入が完了しました！")
                
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
