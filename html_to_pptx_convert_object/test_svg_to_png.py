#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SVG를 PNG로 변환하는 테스트
"""

import requests
from pathlib import Path
from PIL import Image
import io
import base64

def svg_to_png_with_pillow(svg_file, output_file, size=(64, 64)):
    """Pillow를 사용하여 SVG를 PNG로 변환"""
    try:
        # SVG 파일 읽기
        with open(svg_file, 'r', encoding='utf-8') as f:
            svg_content = f.read()
        
        print(f"SVG 파일 읽기: {svg_file}")
        print(f"SVG 크기: {len(svg_content)} bytes")
        
        # SVG를 base64로 인코딩
        svg_base64 = base64.b64encode(svg_content.encode('utf-8')).decode('utf-8')
        data_url = f"data:image/svg+xml;base64,{svg_base64}"
        
        # requests를 사용하여 이미지 다운로드 (Pillow는 SVG를 직접 지원하지 않음)
        # 대신 cairosvg나 다른 라이브러리가 필요하지만, 여기서는 간단한 방법 시도
        
        print(f"PNG 변환 시도: {output_file}")
        print(f"크기: {size}")
        
        # 실제로는 cairosvg나 다른 라이브러리가 필요
        # 여기서는 SVG 내용을 출력하여 확인
        print("SVG 내용 (처음 300자):")
        print(svg_content[:300])
        print("...")
        
        return True
        
    except Exception as e:
        print(f"변환 실패: {e}")
        return False

def test_svg_files():
    """다운로드된 SVG 파일들 테스트"""
    icons_dir = Path("icons")
    if not icons_dir.exists():
        print("icons 디렉토리가 없습니다.")
        return
    
    svg_files = list(icons_dir.glob("*.svg"))
    print(f"SVG 파일 {len(svg_files)}개 발견")
    print("=" * 50)
    
    for svg_file in svg_files:
        print(f"\n--- {svg_file.name} ---")
        png_file = svg_file.with_suffix('.png')
        
        # SVG 내용 확인
        with open(svg_file, 'r', encoding='utf-8') as f:
            svg_content = f.read()
        
        print(f"파일 크기: {svg_file.stat().st_size} bytes")
        print(f"SVG 내용 (처음 200자):")
        print(svg_content[:200])
        print("...")
        
        # viewBox 확인
        if 'viewBox=' in svg_content:
            import re
            viewbox_match = re.search(r'viewBox="([^"]*)"', svg_content)
            if viewbox_match:
                print(f"ViewBox: {viewbox_match.group(1)}")
        
        print("-" * 30)

def test_html_rendering():
    """HTML에서 SVG 렌더링 테스트"""
    icons_dir = Path("icons")
    if not icons_dir.exists():
        return
    
    # 간단한 HTML 생성
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body { margin: 20px; }
            .icon { width: 64px; height: 64px; margin: 10px; }
            .icon-container { display: flex; flex-wrap: wrap; }
        </style>
    </head>
    <body>
        <h1>FontAwesome 아이콘 테스트</h1>
        <div class="icon-container">
    """
    
    svg_files = list(icons_dir.glob("*.svg"))
    for svg_file in svg_files:
        with open(svg_file, 'r', encoding='utf-8') as f:
            svg_content = f.read()
        
        # SVG 내용을 HTML에 직접 삽입
        html_content += f"""
            <div style="margin: 10px; text-align: center;">
                <div class="icon">{svg_content}</div>
                <p>{svg_file.stem}</p>
            </div>
        """
    
    html_content += """
        </div>
    </body>
    </html>
    """
    
    # HTML 파일 저장
    html_file = Path("icon_test.html")
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"HTML 테스트 파일 생성: {html_file}")
    print("브라우저에서 열어서 아이콘들이 제대로 표시되는지 확인하세요.")

if __name__ == "__main__":
    print("SVG 파일들 테스트")
    print("=" * 50)
    test_svg_files()
    
    print("\nHTML 렌더링 테스트")
    print("=" * 50)
    test_html_rendering()
