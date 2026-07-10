# -*- coding: utf-8 -*-
import os
import zipfile
import tempfile
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

def get_image_files(src_path, temp_dir):
    """
    ZIP 파일이거나 이미지 디렉터리 경로에서 이미지 파일 목록을 확보하여 리스트로 반환합니다.
    """
    image_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tif', '.tiff')
    image_files = []
    
    # 만약 ZIP 파일인 경우 압축 해제
    if zipfile.is_zipfile(src_path):
        with zipfile.ZipFile(src_path, 'r') as z:
            z.extractall(temp_dir)
        # 압축 해제된 경로에서 이미지 파일 수집
        for root, _, files in os.walk(temp_dir):
            for file in files:
                if file.lower().endswith(image_extensions) and not file.startswith('._'):
                    image_files.append(os.path.join(root, file))
    elif os.path.isdir(src_path):
        for root, _, files in os.walk(src_path):
            for file in files:
                if file.lower().endswith(image_extensions):
                    image_files.append(os.path.join(root, file))
    elif os.path.isfile(src_path) and src_path.lower().endswith(image_extensions):
        image_files.append(src_path)
        
    # 알파벳/이름 순서대로 정렬
    image_files.sort()
    return image_files

def calculate_fit_dimensions(img_path, max_w, max_h):
    """
    이미지 원본의 가로세로 비율을 보존하면서 max_w, max_h 영역에 쏙 들어가도록 크기를 재계산합니다.
    """
    try:
        with Image.open(img_path) as img:
            img_w, img_h = img.size
    except Exception:
        # 비율 획득 실패 시 기본 정사각형 스케일
        return max_w, max_h
        
    img_aspect = img_w / img_h
    box_aspect = max_w / max_h
    
    if img_aspect > box_aspect:
        # 가로가 꽉 차는 경우
        w = max_w
        h = max_w / img_aspect
    else:
        # 세로가 꽉 차는 경우
        h = max_h
        w = max_h * img_aspect
        
    return w, h

def add_images_to_slide(slide, images, slide_w, slide_h, max_per_slide):
    """
    슬라이드 1장에 속한 이미지 목록을 장수에 맞는 최적의 배치 좌표로 렌더링합니다.
    """
    n = len(images)
    if n == 0:
        return
        
    # 여백 설정
    margin_top = Inches(0.8)
    margin_left = Inches(0.6)
    
    # 가용 영역
    avail_w = slide_w - (margin_left * 2)
    avail_h = slide_h - (margin_top * 1.5)
    
    # 1장~6장에 따른 좌표 구역 계산
    boxes = [] # (left, top, max_w, max_h)
    
    if n == 1:
        # 1장: 전체 화면 중앙 배치
        w_max = avail_w * 0.95
        h_max = avail_h * 0.9
        boxes.append((
            margin_left + (avail_w - w_max) / 2,
            margin_top + (avail_h - h_max) / 2,
            w_max,
            h_max
        ))
    elif n == 2:
        # 2장: 가로 2분할 배치
        gap = Inches(0.4)
        w_max = (avail_w - gap) / 2
        h_max = avail_h * 0.8
        y = margin_top + (avail_h - h_max) / 2
        boxes.append((margin_left, y, w_max, h_max))
        boxes.append((margin_left + w_max + gap, y, w_max, h_max))
    elif n == 3:
        # 3장: 가로 3분할 배치
        gap = Inches(0.3)
        w_max = (avail_w - gap * 2) / 3
        h_max = avail_h * 0.75
        y = margin_top + (avail_h - h_max) / 2
        for i in range(3):
            boxes.append((margin_left + (w_max + gap) * i, y, w_max, h_max))
    elif n == 4:
        # 4장: 2 x 2 격자 배치
        gap_w = Inches(0.4)
        gap_h = Inches(0.4)
        w_max = (avail_w - gap_w) / 2
        h_max = (avail_h - gap_h) / 2
        for r in range(2):
            for c in range(2):
                boxes.append((
                    margin_left + (w_max + gap_w) * c,
                    margin_top + (h_max + gap_h) * r,
                    w_max,
                    h_max
                ))
    elif n == 5:
        # 5장: 상단 3개, 하단 2개 정렬
        gap_w = Inches(0.3)
        gap_h = Inches(0.4)
        w_max = (avail_w - gap_w * 2) / 3
        h_max = (avail_h - gap_h) / 2
        
        # 1행 3개
        for c in range(3):
            boxes.append((
                margin_left + (w_max + gap_w) * c,
                margin_top,
                w_max,
                h_max
            ))
        # 2행 2개 (가로 가운데 정렬)
        indent = (avail_w - (w_max * 2 + gap_w)) / 2
        for c in range(2):
            boxes.append((
                margin_left + indent + (w_max + gap_w) * c,
                margin_top + h_max + gap_h,
                w_max,
                h_max
            ))
    else:
        # 6장: 3 x 2 격자 배치
        gap_w = Inches(0.3)
        gap_h = Inches(0.4)
        w_max = (avail_w - gap_w * 2) / 3
        h_max = (avail_h - gap_h) / 2
        for r in range(2):
            for c in range(3):
                boxes.append((
                    margin_left + (w_max + gap_w) * c,
                    margin_top + (h_max + gap_h) * r,
                    w_max,
                    h_max
                ))

    # 계산된 박스 좌표에 맞추어 이미지를 비율 보존 렌더링
    for i, img_path in enumerate(images):
        left, top, max_w, max_h = boxes[i]
        
        # 비율 유지 가로세로 계산
        w, h = calculate_fit_dimensions(img_path, max_w, max_h)
        
        # 박스 영역 안에서 중앙 정렬
        x = left + (max_w - w) / 2
        y = top + (max_h - h) / 2
        
        # 슬라이드에 이미지 추가
        slide.shapes.add_picture(img_path, x, y, w, h)

def generate_photo_slides(src_path, template_path, output_path, max_per_slide=6, log=print):
    """
    이미지 리소스(ZIP/폴더)를 파싱하여 슬라이드당 최적의 수로 배치한 사진 슬라이드 PPTX를 생성합니다.
    """
    log("[시작] 이미지 파일 파싱을 수행합니다...")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        image_files = get_image_files(src_path, temp_dir)
        n_images = len(image_files)
        
        if n_images == 0:
            raise FileNotFoundError("사진 파일이 발견되지 않았습니다. 올바른 ZIP 파일 또는 이미지 폴더를 지정해 주세요.")
            
        log(f"[정보] 총 {n_images}개의 이미지가 로드되었습니다.")
        
        # 프레젠테이션 초기화
        if template_path and os.path.exists(template_path):
            log(f"[템플릿] '{os.path.basename(template_path)}' 기반으로 새 슬라이드를 추가합니다.")
            prs = Presentation(template_path)
        else:
            log("[기본] 기본 빈 템플릿(16:9)을 사용하여 프레젠테이션을 생성합니다.")
            prs = Presentation()
            # 16:9 슬라이드 크기 표준으로 설정
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
            
        slide_w = prs.slide_width
        slide_h = prs.slide_height
        
        # 빈 슬라이드 레이아웃 (보통 인덱스 6이 완전히 빈 슬라이드 레이아웃임)
        blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
        
        # 이미지 분할 배치 루프
        index = 0
        slide_count = 0
        
        while index < n_images:
            # 한 장에 들어갈 이미지 묶음
            chunk_size = min(max_per_slide, n_images - index)
            chunk_images = image_files[index : index + chunk_size]
            
            # 새 슬라이드 생성
            slide = prs.slides.add_slide(blank_layout)
            slide_count += 1
            
            # 상단 제목 추가 (예: "교육 사진 스케치")
            title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.2), slide_w - Inches(1.2), Inches(0.5))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = f"■ 교육 활동 사진 스케치 (페이지 {slide_count})"
            p.font.name = "맑은 고딕"
            p.font.size = Pt(16)
            p.font.bold = True
            p.font.color.rgb = RGBColor(45, 108, 223) # KOOS 파란색 테마
            
            # 이미지 얹기
            add_images_to_slide(slide, chunk_images, slide_w, slide_h, max_per_slide)
            
            index += chunk_size
            log(f"[진행] 슬라이드 {slide_count} 생성 완료 (사진 {chunk_size}장 배치)")
            
        # 프레젠테이션 저장
        out_dir = os.path.dirname(os.path.abspath(output_path))
        os.makedirs(out_dir, exist_ok=True)
        prs.save(output_path)
        
        log(f"[성공] 총 {slide_count}장의 교육 사진 슬라이드가 저장되었습니다 -> {output_path}")
        return output_path
