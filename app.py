# -*- coding: utf-8 -*-
"""
🔮 사주 PDF 자동 생성기 v2
- 에러 방지 강화
- 폰트 문제 해결
"""

import streamlit as st
import os
import io
import re
import time

# 페이지 설정
st.set_page_config(
    page_title="사주 PDF 자동 생성기",
    page_icon="🔮",
    layout="wide"
)

# 필요한 라이브러리 체크
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    st.error("python-docx 라이브러리가 없습니다!")

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.utils import ImageReader
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False
    st.error("reportlab 라이브러리가 없습니다!")

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    st.error("Pillow 라이브러리가 없습니다!")

# ============================================
# 폰트 설정
# ============================================
def setup_fonts():
    """한글 폰트 설정"""
    
    font_paths = [
        # 사용자 폰트 (.otf 포함!)
        './fonts/나눔바른고딕2-R.otf',
        'fonts/나눔바른고딕2-R.otf',
        './fonts/나눔바른고딕2-R.ttf',
        'fonts/나눔바른고딕2-R.ttf',
        './fonts/NanumGothic.ttf',
        'fonts/NanumGothic.ttf',
        # 시스템 폰트 (Streamlit Cloud)
        '/usr/share/fonts/truetype/nanum/NanumGothic.ttf',
        '/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf',
        '/usr/share/fonts/truetype/nanum/NanumMyeongjo.ttf',
    ]
    
    bold_paths = [
        # Bold 폰트 (.otf 포함!)
        './fonts/나눔바른고딕1-B.otf',
        'fonts/나눔바른고딕1-B.otf',
        './fonts/나눔바른고딕2-B.otf',
        'fonts/나눔바른고딕2-B.otf',
        './fonts/나눔바른고딕2-B.ttf',
        'fonts/나눔바른고딕2-B.ttf',
        './fonts/NanumGothicBold.ttf',
        'fonts/NanumGothicBold.ttf',
        '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf',
    ]
    
    font_name = 'Helvetica'
    bold_name = 'Helvetica-Bold'
    
    # 일반 폰트
    for path in font_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('Korean', path))
                font_name = 'Korean'
                st.sidebar.success(f"✅ 폰트: {os.path.basename(path)}")
                break
            except:
                continue
    
    # Bold 폰트
    for path in bold_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('KoreanBold', path))
                bold_name = 'KoreanBold'
                break
            except:
                continue
    
    # Bold 없으면 일반으로 대체
    if bold_name == 'Helvetica-Bold' and font_name == 'Korean':
        bold_name = 'Korean'
    
    return font_name, bold_name

# ============================================
# DOCX 읽기
# ============================================
def read_docx(file):
    """docx 파일에서 텍스트 추출"""
    try:
        doc = Document(file)
        content = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                style = para.style.name if para.style else "Normal"
                content.append({"text": text, "style": style})
        return content
    except Exception as e:
        st.error(f"DOCX 읽기 실패: {e}")
        return []

# ============================================
# PDF 생성
# ============================================
def create_pdf(docx_files, images, customer_name, progress_bar, status_text):
    """PDF 생성"""
    
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    font_name, bold_name = setup_fonts()
    
    # 폰트 크기
    TITLE_SIZE = 30
    SUBTITLE_SIZE = 25
    BODY_SIZE = 17
    
    # 이미지 분류
    cover_img = None
    page_bg_img = None
    table_images = {}
    
    for img_file in images:
        name = img_file.name.lower()
        data = img_file.read()
        img_file.seek(0)
        
        if "표지" in name or "cover" in name:
            cover_img = data
        elif "내지" in name or "bg" in name or "page" in name:
            page_bg_img = data
        else:
            table_images[img_file.name] = data
    
    total = len(docx_files) + 2
    step = 0
    
    # ===== 1. 표지 =====
    status_text.text("📄 표지 생성 중...")
    progress_bar.progress(step / total)
    time.sleep(0.1)
    
    if cover_img:
        try:
            img_buffer = io.BytesIO(cover_img)
            c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
        except:
            pass
    
    c.setFont(bold_name, 28)
    name_text = f"{customer_name} 님"
    text_width = c.stringWidth(name_text, bold_name, 28)
    c.drawString((width - text_width) / 2, height * 0.2, name_text)
    c.showPage()
    step += 1
    
    # ===== 2. 목차 =====
    status_text.text("📋 목차 생성 중...")
    progress_bar.progress(step / total)
    time.sleep(0.1)
    
    if page_bg_img:
        try:
            img_buffer = io.BytesIO(page_bg_img)
            c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
        except:
            pass
    
    c.setFont(bold_name, 24)
    c.drawString(width/2 - 30, height - 80, "목 차")
    
    c.setFont(font_name, 14)
    y = height - 140
    
    for i, f in enumerate(docx_files):
        name = f.name.replace(".docx", "").replace("_", " ")
        c.drawString(80, y, f"{i+1}. {name}")
        y -= 30
    
    c.showPage()
    step += 1
    
    # ===== 3. 본문 =====
    page_num = 3
    
    for idx, docx_file in enumerate(docx_files):
        progress = (step + idx / len(docx_files)) / total
        progress_bar.progress(progress)
        status_text.text(f"📝 {idx+1}/{len(docx_files)} 장 처리 중...")
        time.sleep(0.05)
        
        docx_file.seek(0)
        content = read_docx(docx_file)
        
        if not content:
            continue
        
        # 새 페이지 시작
        if page_bg_img:
            try:
                img_buffer = io.BytesIO(page_bg_img)
                c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
            except:
                pass
        
        y = height - 80
        margin = 60
        max_width = width - margin * 2
        
        for item in content:
            text = item["text"]
            style = item["style"]
            
            # 스타일 설정
            if "Heading" in style or re.match(r'^제\d+장', text):
                c.setFont(bold_name, TITLE_SIZE)
                line_h = 40
                current_font = bold_name
                current_size = TITLE_SIZE
            elif re.match(r'^\d+\.', text):
                c.setFont(bold_name, SUBTITLE_SIZE)
                line_h = 35
                current_font = bold_name
                current_size = SUBTITLE_SIZE
            else:
                c.setFont(font_name, BODY_SIZE)
                line_h = 25
                current_font = font_name
                current_size = BODY_SIZE
            
            # 줄바꿈
            lines = []
            line = ""
            for char in text:
                if c.stringWidth(line + char, current_font, current_size) < max_width:
                    line += char
                else:
                    lines.append(line)
                    line = char
            if line:
                lines.append(line)
            
            # 그리기
            for ln in lines:
                if y < 80:
                    c.setFont(font_name, 10)
                    c.drawString(width/2 - 10, 40, str(page_num))
                    c.showPage()
                    page_num += 1
                    
                    if page_bg_img:
                        try:
                            img_buffer = io.BytesIO(page_bg_img)
                            c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
                        except:
                            pass
                    
                    y = height - 80
                    c.setFont(current_font, current_size)
                
                c.drawString(margin, y, ln)
                y -= line_h
            
            y -= 5
        
        c.setFont(font_name, 10)
        c.drawString(width/2 - 10, 40, str(page_num))
        c.showPage()
        page_num += 1
    
    progress_bar.progress(1.0)
    status_text.text("✅ PDF 생성 완료!")
    
    c.save()
    buffer.seek(0)
    return buffer

# ============================================
# 메인 UI
# ============================================
def main():
    st.title("🔮 사주 PDF 자동 생성기")
    
    if not (DOCX_AVAILABLE and REPORTLAB_AVAILABLE and PIL_AVAILABLE):
        st.error("필요한 라이브러리가 없습니다. requirements.txt를 확인하세요.")
        return
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 DOCX 파일")
        docx_files = st.file_uploader(
            "장별 docx 파일들",
            type=["docx"],
            accept_multiple_files=True
        )
        if docx_files:
            st.success(f"✅ {len(docx_files)}개 업로드됨")
    
    with col2:
        st.subheader("🖼️ 이미지 파일")
        image_files = st.file_uploader(
            "표지, 내지, 사주표 이미지들",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True
        )
        if image_files:
            st.success(f"✅ {len(image_files)}개 업로드됨")
    
    st.markdown("---")
    
    customer_name = st.text_input("고객명", value="홍길동")
    
    st.markdown("---")
    
    if st.button("🚀 PDF 생성하기", type="primary", use_container_width=True):
        if not docx_files:
            st.error("❌ docx 파일을 업로드해주세요!")
            return
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            sorted_docx = sorted(docx_files, key=lambda x: x.name)
            
            pdf_buffer = create_pdf(
                sorted_docx,
                image_files if image_files else [],
                customer_name,
                progress_bar,
                status_text
            )
            
            st.success("✅ 완료!")
            
            st.download_button(
                f"📥 {customer_name}_사주감정서.pdf 다운로드",
                pdf_buffer,
                f"{customer_name}_사주감정서.pdf",
                "application/pdf",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"❌ 오류: {e}")

if __name__ == "__main__":
    main()
