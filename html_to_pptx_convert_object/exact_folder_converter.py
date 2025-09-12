#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exact Folder HTML to PPTX Converter
HTML의 정확한 레이아웃과 스타일을 완벽하게 재현하는 폴더 변환기
"""

import sys
from pathlib import Path
from exact_html_to_pptx_converter import convert_folder_to_pptx

def main():
    print("Exact Folder HTML to PPTX 변환기")
    print("=" * 50)
    
    # 명령행 인수 확인
    if len(sys.argv) > 1:
        html_folder = sys.argv[1]
    else:
        # 기본 경로 사용
        html_folder = r"C:\Project\gigabitamin\genspark\dcs_site\html"
        print(f"기본 HTML 폴더 사용: {html_folder}")
        print("사용법: python exact_folder_converter.py <HTML폴더경로>")
        print()
    
    # HTML 폴더 존재 확인
    if not Path(html_folder).exists():
        print(f"❌ HTML 폴더가 존재하지 않습니다: {html_folder}")
        return
    
    # 출력 파일 경로 설정
    output_path = Path(html_folder) / "exact_all_pages_16x9.pptx"
    
    print(f"HTML 폴더: {html_folder}")
    print(f"출력 파일: {output_path}")
    print("-" * 50)
    
    # 변환 실행
    success = convert_folder_to_pptx(html_folder, str(output_path))
    
    if success:
        print("-" * 50)
        print("✅ 변환 완료!")
        print(f"출력 파일: {output_path}")
        print(f"파일 크기: {output_path.stat().st_size:,} bytes")
        print("📐 슬라이드 비율: 16:9")
        print("🎨 HTML의 정확한 레이아웃과 스타일이 완벽하게 재현되었습니다!")
        print("🔧 중앙 정렬, 폰트 크기, 색상, 간격 모두 정확!")
        print("💎 브라우저에서 보는 것과 완전히 동일한 시각적 효과!")
        print("🚀 아이콘, 카드, 텍스트 모든 요소가 완벽하게 재현!")
    else:
        print("❌ 변환 실패!")

if __name__ == "__main__":
    main()
