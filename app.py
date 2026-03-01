import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.util import Pt
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

def get_unique_list(lists):
    """複数のリストを結合し、順番を保ったまま重複を削除する"""
    seen = set()
    res = []
    for lst in lists:
        for item in lst:
            if item and item not in seen:
                seen.add(item)
                res.append(item)
    return res

def build_placeholder_map(data: dict) -> dict:
    basic = data.get("basic", {})
    materials = data.get("materials", {})
    methods = data.get("methods", {})
    md = data.get("md_result", {})
    target_info = data.get("target_info", {})

    branding = basic.get("branding", "")
    branding_label = "ブランデッド" if branding == "branded" else "アンブランデッド" if branding == "unbranded" else str(branding)

    merged_tones = get_unique_list([
        _safe_get(md, "tone_manner", "keys", default=[]),
        _safe_get(methods, "tone_tags", default=[]),
        _safe_get(materials, "worldview_tags", default=[])
    ])
    
    merged_motifs = get_unique_list([
        _safe_get(md, "motif", "keys", default=[]),
        _safe_get(materials, "motifs", default=[])
    ])

    mapping = {
        "{{branding}}": branding_label,
        "{{itemName}}": str(basic.get("item_name", "")),
        "{{productName}}": str(basic.get("product_name", "")),
        "{{diseaseName}}": str(basic.get("disease_name", "")),
        "{{spec}}": str(basic.get("item_spec", "")),
        "{{objective}}": str(basic.get("objective", "")),
        "{{target}}": str(basic.get("target", "")),
        "{{demo}}": "", # 空欄にして消去
        
        "{{brandPrimary}}": "",
        "{{brandSecondary}}": "",
        "{{brandElement}}": str(basic.get("brand_elements", "")),
        "{{brandFont}}": str(basic.get("brand_font", "")),
        
        # テンプレートに合わせて元の名前に修正
        "{{creativeConcept}}": str(md.get("visual_concept", "")),
        "{{targetInsightLine1}}": str(_safe_get(md, "target_insight", "line1", default="")),
        "{{targetInsightLine2}}": str(_safe_get(md, "target_insight", "line2", default="")),
        
        "{{direction1}}": _arr_get(md.get("direction_short", []), 0, ""),
        "{{direction2}}": _arr_get(md.get("direction_short", []), 1, ""),
        "{{direction3}}": _arr_get(md.get("direction_short", []), 2, ""),
        
        "{{focusCenter}}": str(target_info.get("focusCenter", "")),
        "{{treatmentPhase}}": str(target_info.get("infoPhase", "")),
        "{{focusSummary}}": str(methods.get("focus_summary", "")),
        "{{toneSummary}}": str(_safe_get(md, "tone_manner", "summary", default="")),
        
        "{{note1}}": _arr_get(md.get("notes", []), 0, ""),
        "{{note2}}": _arr_get(md.get("notes", []), 1, ""),
    }

    # トーンのタグ (t1〜t10)
    for i in range(10):
        mapping[f"{{{{t{i+1}}}}}"] = _arr_get(merged_tones, i, "")

    # モチーフのタグ (m1〜m10)
    for i in range(10):
        mapping[f"{{{{m{i+1}}}}}"] = _arr_get(merged_motifs, i, "")

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

def insert_images_to_slide(prs, slide_index, image_files):
    if not image_files or slide_index >= len(prs.slides):
        return
    
    slide = prs.slides[slide_index]
    image_files = image_files[:6]
    
    margin_x = prs.slide_width * 0.05
    margin_top = prs.slide_height * 0.25 
    usable_w = prs.slide_width - (margin_x * 2)
    usable_h = prs.slide_height - margin_top - (prs.slide_height * 0.05)
    
    col_w = usable_w / 3 
    row_h = usable_h / 2 
    
    for idx, img_file in enumerate(image_files):
        c = idx % 3  
        r = idx // 3 
        
        cell_x = int(margin_x + c * col_w)
        cell_y = int(margin_top + r * row_h)
        img_width = int(col_w * 0.9)
        
        try:
            img_stream = BytesIO(img_file.getvalue())
            slide.shapes.add_picture(img_stream, cell_x, cell_y, width=img_width)
        except Exception as e:
            st.warning(f"画像の貼り付けに失敗しました（{img_file.name}）: {e}")

def replace_placeholders_in_pptx(template_path: str, data: dict, images: dict) -> bytes:
    prs = Presentation(template_path)
    mapping = build_placeholder_map(data)

    primary_hex = str(data.get("basic", {}).get("brand_primary", ""))
    secondary_hex = str(data.get("basic", {}).get("brand_secondary", ""))
    
    color_shapes_to_add = []

    def process_shape_for_colors(slide, shape):
        if getattr(shape, "has_text_frame", False):
            text = shape.text
            if "{{brandPrimary}}" in text and primary_hex:
                color_shapes_to_add.append((slide, shape.left, shape.top, primary_hex, True))
            if "{{brandSecondary}}" in text and secondary_hex:
                color_shapes_to_add.append((slide, shape.left, shape.top, secondary_hex, False))

        if getattr(shape, "has_table", False):
            table = shape.table
            current_top = shape.top
            for row in table.rows:
                current_left = shape.left
                for col_idx, cell in enumerate(row.cells):
                    if getattr(cell, "text_frame", None):
                        text = cell.text_frame.text
                        if "{{brandPrimary}}" in text and primary_hex:
                            color_shapes_to_add.append((slide, current_left + Pt(10), current_top + Pt(5), primary_hex, True))
                        if "{{brandSecondary}}" in text and secondary_hex:
                            color_shapes_to_add.append((slide, current_left + Pt(10), current_top + Pt(5), secondary_hex, False))
                    current_left += table.columns[col_idx].width
                current_top += row.height
                
        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
            for child in shape.shapes:
                process_shape_for_colors(slide, child)

    for slide in prs.slides:
        for shape in slide.shapes:
            process_shape_for_colors(slide, shape)
            replace_text_in_shape(shape, mapping)

    for slide, base_left, base_top, hex_color, is_primary in color_shapes_to_add:
        hex_color = hex_color.lstrip('#')
        if len(hex_color) == 6:
            try:
                r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
                box_width = Pt(50) 
                box_height = Pt(15) 
                offset_left = 0 if is_primary else box_width + Pt(10)
                
                new_shape = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, 
                    base_left + offset_left, 
                    base_top, 
                    box_width, 
                    box_height
                )
                new_shape.fill.solid()
                new_shape.fill.fore_color.rgb = RGBColor(r, g, b)
                new_shape.line.color.rgb = RGBColor(r, g, b)
            except ValueError:
                pass

    insert_images_to_slide(prs, 5, images.get("A", []))
    insert_images_to_slide(prs, 6, images.get("B", []))
    insert_images_to_slide(prs, 7, images.get("C", []))

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

        images_dict = {
            "A": img_a_list,
            "B": img_b_list,
            "C": img_c_list
        }

        with st.spinner("PPTXを生成中..."):
            try:
                output_pptx = replace_placeholders_in_pptx("template.pptx", data, images_dict)
                st.success("テキストの置換、カラー図形の配置、画像の挿入が完了しました！")
                
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
