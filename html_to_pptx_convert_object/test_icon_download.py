#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FontAwesome 아이콘 SVG 다운로드 테스트
"""

import requests
from pathlib import Path
from PIL import Image
import io
import base64

def download_fontawesome_svg(icon_name, output_dir="icons"):
    """FontAwesome 아이콘 SVG 다운로드"""
    try:
        # FontAwesome SVG URL
        svg_url = f"https://raw.githubusercontent.com/FortAwesome/Font-Awesome/6.x/svgs/solid/{icon_name}.svg"
        
        print(f"다운로드 시도: {icon_name}")
        print(f"URL: {svg_url}")
        
        # SVG 다운로드
        response = requests.get(svg_url, timeout=10)
        print(f"응답 상태: {response.status_code}")
        
        if response.status_code == 200:
            svg_content = response.text
            print(f"SVG 크기: {len(svg_content)} bytes")
            
            # 출력 디렉토리 생성
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)
            
            # SVG 파일 저장
            svg_file = output_path / f"{icon_name}.svg"
            with open(svg_file, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            
            print(f"SVG 저장 완료: {svg_file}")
            
            # SVG 내용 일부 출력
            print(f"SVG 내용 (처음 200자):")
            print(svg_content[:200])
            print("...")
            
            return svg_file
            
        else:
            print(f"다운로드 실패: HTTP {response.status_code}")
            return None
            
    except Exception as e:
        print(f"오류 발생: {e}")
        return None

def test_icons():
    """여러 아이콘 다운로드 테스트"""
    test_icons = [
        'react',  # fa-react
        'js',     # fa-js  
        'css3',   # fa-css3
        'database', # fa-database
        'server', # fa-server
        'github', # fa-github
        'globe',  # fa-globe
        'history', # fa-history
        'bullseye', # fa-bullseye
        'star',   # fa-star
        'users',  # fa-users
        'graduation-cap', # fa-graduation-cap
        'project-diagram', # fa-project-diagram
        'mobile-alt' # fa-mobile-alt
    ]
    
    print("FontAwesome 아이콘 다운로드 테스트 시작")
    print("=" * 50)
    
    success_count = 0
    for icon_name in test_icons:
        print(f"\n--- {icon_name} ---")
        result = download_fontawesome_svg(icon_name)
        if result:
            success_count += 1
        print("-" * 30)
    
    print(f"\n총 {len(test_icons)}개 아이콘 중 {success_count}개 성공")
    
    # 다운로드된 파일들 확인
    icons_dir = Path("icons")
    if icons_dir.exists():
        svg_files = list(icons_dir.glob("*.svg"))
        print(f"\n다운로드된 SVG 파일들:")
        for svg_file in svg_files:
            print(f"  - {svg_file.name} ({svg_file.stat().st_size} bytes)")

if __name__ == "__main__":
    test_icons()
