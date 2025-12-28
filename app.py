# -*- coding: utf-8 -*-
"""
🔮 사주 PDF 자동 생성기
- docx 14개 + 이미지 18개 업로드
- "~표 해설" 앞에 해당 이미지 자동 삽입
- PDF 자동 변환
"""

import streamlit as st
import os
import io
import re
import time
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from PIL import Image

# ============================================
# 페이지 설정
# ============================================
st.set_page_config(
    page_title="사주 PDF 자동 생성기",
    page_icon="🔮",
    layout="wide"
)

# ============================================
# 이미지-마커 매핑 테이블
# ============================================
IMAGE_MARKERS = {
    "01_원국표": "원국표 해설",
    "02_대운표": "대운표 해설",
    "03_세운표": "세운표 해설",
    "04_월운표": "월운표 해설",
    "05_오행분석": "오행분석 해설",
    "06_십성표": "십성표 해설",
    "07_신살표": "신살표 해설",
    "08_12운성표": "12운성표 해설",
    "09_지장간표": "지장간표 해설",
    "10_합충형파해표": "합충형파해표 해설",
    "11_궁성표": "궁성표 해설",
    "12_육친표": "육친표 해설",
    "13_납음오행표": "납음오행표 해설",
    "14_격국표": "격국표 해설",
    "15_공망표": "공망표 해설",
    "16_용신표": "용신표 해설",
}

# 대체 마커 (유연한 매칭)
ALT_MARKERS = {
    "01_원국표": ["원국표", "원국 표", "사주 원국"],
    "02_대운표": ["대운표", "대운 표", "대운 분석"],
    "03_세운표": ["세운표", "세운 표", "세운 분석"],
    "04_월운표": ["월운표", "월운 표", "월운 분석"],
    "05_오행분석": ["오행분석", "오행 분석", "오행표"],
    "06_십성표": ["십성표", "십성 표", "십성 분석"],
    "07_신살표": ["신살표", "신살 표", "신살 분석"],
    "08_12운성표": ["12운성표", "12운성 표", "운성표", "십이운성"],
    "09_지장간표": ["지장간표", "지장간 표", "지장간"],
    "10_합충형파해표": ["합충형파해표", "합충형파해", "충형파해"],
    "11_궁성표": ["궁성표", "궁성 표", "궁성 분석"],
    "12_육친표": ["육친표", "육친 표", "육친 분석"],
    "13_납음오행표": ["납음오행표", "납음오행", "납음 오행"],
    "14_격국표": ["격국표", "격국 표", "격국 분석"],
    "15_공망표": ["공망표", "공망 표", "공망 분석"],
    "16_용신표": ["용신표", "용신 표", "용신 분석"],
}

# ============================================
# 폰트 설정
# ============================================
def setup_fonts():
    """한글 폰트 설정 (일반 + Bold)"""
    
    # 일반 폰트 경로 (나눔바른고딕2-R 우선!)
    font_paths = [
        # 사용자 폰트 (최우선)
        './fonts/나눔바른고딕2-R.ttf',
        'fonts/나눔바른고딕2-R.ttf',
        './fonts/NanumBarunGothic2-R.ttf',
        'fonts/NanumBarunGothic2-R.ttf',
        # 시스템 폰트
        '/usr/share/fonts/truetype/nanum/NanumGothic.ttf',
        '/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf',
        './fonts/NanumGothic.ttf',
        'fonts/NanumGothic.ttf',
    ]
    
    # Bold 폰트 경로
    bold_font_paths = [
        # 사용자 Bold 폰트 (최우선)
        './fonts/나눔바른고딕2-B.ttf',
        'fonts/나눔바른고딕2-B.ttf',
        './fonts/NanumBarunGothic2-B.ttf',
        'fonts/NanumBarunGothic2-B.ttf',
        # 시스템 폰트
        '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf',
        '/usr/share/fonts/truetype/nanum/NanumBarunGothicBold.ttf',
        './fonts/NanumGothicBold.ttf',
        'fonts/NanumGothicBold.ttf',
    ]
    
    font_name = 'Helvetica'
    bold_font_name = 'Helvetica-Bold'
    
    # 일반 폰트 등록
    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont('Korean', font_path))
                font_name = 'Korean'
                st.sidebar.success(f"✅ 폰트 로드: {font_path}")
                break
            except Exception as e:
                continue
    
    # Bold 폰트 등록
    for font_path in bold_font_paths:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont('KoreanBold', font_path))
                bold_font_name = 'KoreanBold'
                st.sidebar.success(f"✅ Bold 폰트 로드: {font_path}")
                break
            except Exception as e:
                continue
    
    # Bold가 없으면 일반 폰트로 대체
    if bold_font_name == 'Helvetica-Bold' and font_name == 'Korean':
        try:
            pdfmetrics.registerFont(TTFont('KoreanBold', font_paths[0]))
            bold_font_name = 'KoreanBold'
        except:
            pass
    
    return font_name, bold_font_name

# ============================================
# DOCX 파일 읽기
# ============================================
def read_docx(file):
    """docx 파일에서 텍스트와 구조 추출"""
    doc = Document(file)
    
    content = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            # 스타일 확인 (제목인지 본문인지)
            style = para.style.name if para.style else "Normal"
            content.append({
                "text": text,
                "style": style
            })
    
    return content

def extract_chapter_title(content):
    """장 제목 추출"""
    for item in content:
        text = item["text"]
        # "제1장", "제2장" 등의 패턴 찾기
        if re.match(r'^제\d+장', text):
            return text
    return None

# ============================================
# 이미지 매칭
# ============================================
def find_matching_image(text, images_dict):
    """텍스트에 맞는 이미지 찾기"""
    text_lower = text.lower().replace(" ", "")
    
    for img_key, markers in ALT_MARKERS.items():
        for marker in markers:
            marker_clean = marker.lower().replace(" ", "")
            if marker_clean in text_lower:
                # 해당 이미지 파일 찾기
                for img_name, img_data in images_dict.items():
                    img_name_clean = img_name.lower().replace(" ", "")
                    if img_key.lower().replace("_", "") in img_name_clean.replace("_", ""):
                        return img_name, img_data
    
    return None, None

# ============================================
# PDF 생성
# ============================================
def create_pdf(docx_files, images, customer_name, progress_callback=None):
    """docx 파일들과 이미지로 PDF 생성"""
    
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    font_name, bold_font_name = setup_fonts()
    
    # 폰트 크기 설정
    TITLE_SIZE = 30      # 장제목 (제1장, 제2장...)
    SUBTITLE_SIZE = 25   # 소제목 (1., 2., 3...)
    BODY_SIZE = 17       # 본문
    
    # 이미지 딕셔너리 구성
    images_dict = {}
    cover_img = None
    page_bg_img = None
    
    for img_file in images:
        img_name = img_file.name.lower()
        img_data = img_file.read()
        img_file.seek(0)
        
        if "표지" in img_name or "cover" in img_name:
            cover_img = img_data
        elif "내지" in img_name or "page" in img_name or "bg" in img_name:
            page_bg_img = img_data
        else:
            images_dict[img_file.name] = img_data
    
    total_steps = len(docx_files) + 2  # 표지 + 목차 + 각 장
    current_step = 0
    
    # ========== 1. 표지 페이지 ==========
    if progress_callback:
        progress_callback(current_step / total_steps, "표지 생성 중...")
    time.sleep(0.1)
    
    if cover_img:
        try:
            img_buffer = io.BytesIO(cover_img)
            c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
        except Exception as e:
            st.warning(f"표지 이미지 로드 실패: {e}")
    
    # 고객명 추가
    c.setFont(font_name, 28)
    name_text = f"{customer_name} 님"
    text_width = c.stringWidth(name_text, font_name, 28)
    c.drawString((width - text_width) / 2, height * 0.2, name_text)
    
    c.showPage()
    current_step += 1
    
    # ========== 2. 목차 페이지 ==========
    if progress_callback:
        progress_callback(current_step / total_steps, "목차 생성 중...")
    time.sleep(0.1)
    
    # 배경
    if page_bg_img:
        try:
            img_buffer = io.BytesIO(page_bg_img)
            c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
        except:
            pass
    
    c.setFont(font_name, 24)
    c.drawString(width/2 - 30, height - 80, "목 차")
    
    c.setFont(font_name, 14)
    y_pos = height - 140
    
    # 장 제목들 수집
    chapter_titles = []
    for docx_file in docx_files:
        docx_file.seek(0)
        content = read_docx(docx_file)
        title = extract_chapter_title(content)
        if title:
            chapter_titles.append(title)
        else:
            # 파일명에서 추출
            name = docx_file.name.replace(".docx", "").replace("_", " ")
            chapter_titles.append(name)
    
    for i, title in enumerate(chapter_titles):
        c.drawString(80, y_pos, f"{title}")
        y_pos -= 30
        if y_pos < 100:
            c.showPage()
            if page_bg_img:
                try:
                    img_buffer = io.BytesIO(page_bg_img)
                    c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
                except:
                    pass
            y_pos = height - 80
    
    c.showPage()
    current_step += 1
    
    # ========== 3. 본문 페이지들 ==========
    page_number = 3
    
    for doc_idx, docx_file in enumerate(docx_files):
        if progress_callback:
            progress_callback(
                (current_step + doc_idx * 0.8 / len(docx_files)) / total_steps,
                f"'{chapter_titles[doc_idx]}' 처리 중... ({doc_idx + 1}/{len(docx_files)})"
            )
        time.sleep(0.05)
        
        docx_file.seek(0)
        content = read_docx(docx_file)
        
        # 배경 이미지
        if page_bg_img:
            try:
                img_buffer = io.BytesIO(page_bg_img)
                c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
            except:
                pass
        
        y_pos = height - 80
        margin_left = 60
        margin_right = 60
        line_height = 20
        max_width = width - margin_left - margin_right
        
        for item in content:
            text = item["text"]
            style = item["style"]
            
            # 이미지 삽입 체크 ("~표 해설" 앞에)
            img_name, img_data = find_matching_image(text, images_dict)
            if img_data:
                # 새 페이지로 이동
                if y_pos < height - 100:
                    c.showPage()
                    page_number += 1
                    if page_bg_img:
                        try:
                            img_buffer = io.BytesIO(page_bg_img)
                            c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
                        except:
                            pass
                    y_pos = height - 80
                
                # 이미지 삽입
                try:
                    img_buffer = io.BytesIO(img_data)
                    img = Image.open(img_buffer)
                    img_width, img_height = img.size
                    
                    # 이미지 크기 조정 (페이지에 맞게)
                    scale = min((width - 100) / img_width, (height - 200) / img_height, 1)
                    new_width = img_width * scale
                    new_height = img_height * scale
                    
                    img_buffer.seek(0)
                    c.drawImage(
                        ImageReader(img_buffer),
                        (width - new_width) / 2,
                        (height - new_height) / 2,
                        width=new_width,
                        height=new_height
                    )
                    
                    c.showPage()
                    page_number += 1
                    if page_bg_img:
                        try:
                            img_buffer = io.BytesIO(page_bg_img)
                            c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
                        except:
                            pass
                    y_pos = height - 80
                    
                    # 해당 이미지 제거 (중복 삽입 방지)
                    if img_name in images_dict:
                        del images_dict[img_name]
                        
                except Exception as e:
                    st.warning(f"이미지 삽입 실패 ({img_name}): {e}")
            
            # 텍스트 스타일 설정
            if "Heading" in style or re.match(r'^제\d+장', text):
                # 장제목: 30pt, Bold
                c.setFont(bold_font_name, TITLE_SIZE)
                line_height = 40
                current_font = bold_font_name
                current_size = TITLE_SIZE
            elif re.match(r'^\d+\.', text):  # 소제목 (1., 2., 3. 등)
                # 소제목: 25pt, Bold
                c.setFont(bold_font_name, SUBTITLE_SIZE)
                line_height = 35
                current_font = bold_font_name
                current_size = SUBTITLE_SIZE
            else:
                # 본문: 17pt
                c.setFont(font_name, BODY_SIZE)
                line_height = 25
                current_font = font_name
                current_size = BODY_SIZE
            
            # 텍스트 줄바꿈 처리
            words = text
            lines = []
            current_line = ""
            
            for char in words:
                test_line = current_line + char
                if c.stringWidth(test_line, current_font, current_size) < max_width:
                    current_line = test_line
                else:
                    if current_line:
                        lines.append(current_line)
                    current_line = char
            if current_line:
                lines.append(current_line)
            
            # 텍스트 그리기
            for line in lines:
                if y_pos < 80:
                    # 페이지 번호
                    c.setFont(font_name, 10)
                    c.drawString(width/2 - 10, 40, str(page_number))
                    
                    c.showPage()
                    page_number += 1
                    
                    if page_bg_img:
                        try:
                            img_buffer = io.BytesIO(page_bg_img)
                            c.drawImage(ImageReader(img_buffer), 0, 0, width=width, height=height)
                        except:
                            pass
                    
                    y_pos = height - 80
                    
                    # 현재 스타일에 맞게 폰트 재설정
                    c.setFont(current_font, current_size)
                
                c.drawString(margin_left, y_pos, line)
                y_pos -= line_height
            
            y_pos -= 5  # 문단 간격
        
        # 장 끝: 페이지 번호 추가 후 새 페이지
        c.setFont(font_name, 10)
        c.drawString(width/2 - 10, 40, str(page_number))
        c.showPage()
        page_number += 1
        
        current_step = 2 + doc_idx + 1
    
    if progress_callback:
        progress_callback(1.0, "✅ PDF 생성 완료!")
    
    c.save()
    buffer.seek(0)
    return buffer

# ============================================
# 메인 UI
# ============================================
def main():
    st.title("🔮 사주 PDF 자동 생성기")
    st.markdown("docx 파일과 이미지를 업로드하면 자동으로 PDF를 생성합니다.")
    
    st.markdown("---")
    
    # ========== 파일 업로드 영역 ==========
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 DOCX 파일 업로드")
        docx_files = st.file_uploader(
            "장별 docx 파일들 (14개)",
            type=["docx"],
            accept_multiple_files=True,
            key="docx_upload"
        )
        
        if docx_files:
            st.success(f"✅ {len(docx_files)}개 파일 업로드됨")
            with st.expander("업로드된 파일 목록"):
                for f in docx_files:
                    st.write(f"- {f.name}")
    
    with col2:
        st.subheader("🖼️ 이미지 파일 업로드")
        image_files = st.file_uploader(
            "표지, 내지, 사주표 이미지들 (18개)",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key="image_upload"
        )
        
        if image_files:
            st.success(f"✅ {len(image_files)}개 이미지 업로드됨")
            with st.expander("업로드된 이미지 목록"):
                for f in image_files:
                    st.write(f"- {f.name}")
    
    st.markdown("---")
    
    # ========== 설정 영역 ==========
    st.subheader("⚙️ 설정")
    
    customer_name = st.text_input(
        "고객명",
        value="홍길동",
        help="표지에 표시될 고객 이름"
    )
    
    st.markdown("---")
    
    # ========== PDF 생성 버튼 ==========
    if st.button("🚀 PDF 생성하기", type="primary", use_container_width=True):
        
        if not docx_files:
            st.error("❌ docx 파일을 업로드해주세요!")
            return
        
        if not image_files:
            st.warning("⚠️ 이미지 없이 진행합니다.")
        
        # 진행률 표시
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        def update_progress(progress, message):
            progress_bar.progress(progress)
            status_text.text(message)
        
        try:
            # docx 파일 정렬 (파일명 기준)
            sorted_docx = sorted(docx_files, key=lambda x: x.name)
            
            # PDF 생성
            pdf_buffer = create_pdf(
                sorted_docx,
                image_files if image_files else [],
                customer_name,
                update_progress
            )
            
            # 다운로드 버튼
            st.success("✅ PDF 생성 완료!")
            
            filename = f"{customer_name}_사주감정서.pdf"
            
            st.download_button(
                label=f"📥 {filename} 다운로드",
                data=pdf_buffer,
                file_name=filename,
                mime="application/pdf",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"❌ 오류 발생: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
    
    # ========== 사용 안내 ==========
    st.markdown("---")
    with st.expander("📖 사용 안내"):
        st.markdown("""
        ### 파일 준비
        
        **DOCX 파일 (14개)**
        - 제1장_리포트안내및해석기준.docx
        - 제2장_사주원국분석.docx
        - ... (총 14개 장)
        
        **이미지 파일 (18개)**
        - 표지.png (또는 cover.png)
        - 내지.png (또는 page_bg.png)
        - 01_원국표.png
        - 02_대운표.png
        - ... (총 16개 사주표)
        
        ### 자동 매칭 규칙
        
        docx 본문에서 "원국표 해설", "대운표 해설" 등의 텍스트를 찾으면
        해당 위치 앞에 자동으로 이미지가 삽입됩니다.
        
        ### 주의사항
        
        - docx 파일명은 순서대로 정렬됩니다 (제1장, 제2장...)
        - 이미지 파일명에 "표지" 또는 "cover"가 포함되면 표지로 인식
        - 이미지 파일명에 "내지" 또는 "bg"가 포함되면 본문 배경으로 인식
        """)

if __name__ == "__main__":
    main()
