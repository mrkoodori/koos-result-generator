# -*- coding: utf-8 -*-
import os

def export_slides_if_needed(pptx_path, output_dir):
    """
    윈도우 환경에서 win32com을 사용해 PPT 템플릿 슬라이드를 이미지로 내보냅니다.
    이미 파일이 존재한다면 불필요한 실행을 방지하기 위해 즉시 리턴합니다.
    """
    slide1_path = os.path.join(output_dir, "slide1_origin.png")
    slide4_path = os.path.join(output_dir, "slide4_origin.png")
    
    # 캐싱: 파일이 이미 존재하면 다시 변환하지 않음
    if os.path.exists(slide1_path) and os.path.exists(slide4_path):
        return True
        
    if not os.path.exists(pptx_path):
        return False
        
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("[Exporter] win32com is not installed. Fallback to virtual slides.")
        return False
        
    try:
        # COM 초기화 (멀티스레드 환경 대비)
        pythoncom.CoInitialize()
        
        os.makedirs(output_dir, exist_ok=True)
        
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True 
        
        abs_pptx = os.path.abspath(pptx_path)
        presentation = powerpoint.Presentations.Open(
            abs_pptx, 
            ReadOnly=True, 
            Untitled=False, 
            WithWindow=False
        )
        
        presentation.Slides(1).Export(slide1_path, "PNG")
        presentation.Slides(4).Export(slide4_path, "PNG")
        
        presentation.Close()
        powerpoint.Quit()
        
        print("[Exporter] Successfully exported slides to assets.")
        return True
    except Exception as e:
        print(f"[Exporter] Failed to export slides: {str(e)}")
        return False
