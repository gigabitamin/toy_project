#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FontAwesome 아이콘 SVG 다운로드 테스트 V2
다양한 경로와 브랜치 시도
"""

import requests
from pathlib import Path

def test_different_paths(icon_name):
    """다양한 경로로 아이콘 다운로드 시도"""
    paths = [
        f"https://raw.githubusercontent.com/FortAwesome/Font-Awesome/6.x/svgs/solid/{icon_name}.svg",
        f"https://raw.githubusercontent.com/FortAwesome/Font-Awesome/6.x/svgs/brands/{icon_name}.svg",
        f"https://raw.githubusercontent.com/FortAwesome/Font-Awesome/6.x/svgs/regular/{icon_name}.svg",
        f"https://raw.githubusercontent.com/FortAwesome/Font-Awesome/5.x/svgs/solid/{icon_name}.svg",
        f"https://raw.githubusercontent.com/FortAwesome/Font-Awesome/5.x/svgs/brands/{icon_name}.svg",
        f"https://raw.githubusercontent.com/FortAwesome/Font-Awesome/5.x/svgs/regular/{icon_name}.svg",
    ]
    
    for path in paths:
        try:
            print(f"  시도: {path}")
            response = requests.get(path, timeout=5)
            if response.status_code == 200:
                print(f"  ✅ 성공!")
                return path, response.text
            else:
                print(f"  ❌ 실패: {response.status_code}")
        except Exception as e:
            print(f"  ❌ 오류: {e}")
    
    return None, None

def test_missing_icons():
    """누락된 아이콘들 테스트"""
    missing_icons = [
        'react', 'js', 'css3', 'github', 'history', 
        'project-diagram', 'mobile-alt'
    ]
    
    print("누락된 아이콘들 다른 경로로 시도")
    print("=" * 50)
    
    for icon_name in missing_icons:
        print(f"\n--- {icon_name} ---")
        path, content = test_different_paths(icon_name)
        if path and content:
            # 성공한 경우 파일 저장
            output_dir = Path("icons")
            output_dir.mkdir(exist_ok=True)
            
            svg_file = output_dir / f"{icon_name}.svg"
            with open(svg_file, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"  저장: {svg_file}")
        print("-" * 30)

def test_alternative_names():
    """대체 이름으로 시도"""
    alternatives = {
        'react': ['react', 'reactjs'],
        'js': ['js', 'javascript', 'js-square'],
        'css3': ['css3', 'css3-alt'],
        'github': ['github', 'github-alt', 'github-square'],
        'history': ['history', 'clock'],
        'project-diagram': ['project-diagram', 'sitemap', 'diagram-project'],
        'mobile-alt': ['mobile-alt', 'mobile', 'phone']
    }
    
    print("\n대체 이름으로 시도")
    print("=" * 50)
    
    for original, alts in alternatives.items():
        print(f"\n--- {original} ---")
        for alt in alts:
            print(f"  시도: {alt}")
            path, content = test_different_paths(alt)
            if path and content:
                print(f"  ✅ 성공! ({alt})")
                # 파일 저장
                output_dir = Path("icons")
                output_dir.mkdir(exist_ok=True)
                
                svg_file = output_dir / f"{original}.svg"
                with open(svg_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                print(f"  저장: {svg_file}")
                break
            else:
                print(f"  ❌ 실패")
        print("-" * 30)

if __name__ == "__main__":
    test_missing_icons()
    test_alternative_names()
    
    # 최종 결과 확인
    icons_dir = Path("icons")
    if icons_dir.exists():
        svg_files = list(icons_dir.glob("*.svg"))
        print(f"\n최종 다운로드된 SVG 파일들 ({len(svg_files)}개):")
        for svg_file in svg_files:
            print(f"  - {svg_file.name} ({svg_file.stat().st_size} bytes)")
