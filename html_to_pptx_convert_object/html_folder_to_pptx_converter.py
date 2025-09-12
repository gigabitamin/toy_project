#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HTML 폴더 to Editable PPTX 변환기
폴더 내 모든 HTML 파일을 하나의 PPTX 파일로 변환
"""

import sys
import os
from pathlib import Path
from html_to_editable_pptx_converter_v6 import convert_folder_to_pptx

def main():
    print("HTML 폴더 to Editable PPTX 변환기")
    print("=" * 50)
    
    # 명령행 인수 확인
    if len(sys.argv) > 1:
        html_folder = sys.argv[1]
    else:
        # 기본 경로 사용
        html_folder = r"C:\Project\gigabitamin\genspark\dcs_site\html"
        print(f"기본 HTML 폴더 사용: {html_folder}")
        print("사용법: python html_folder_to_pptx_converter.py <HTML폴더경로>")
        print()
    
    # HTML 폴더 존재 확인
    if not Path(html_folder).exists():
        print(f"❌ HTML 폴더가 존재하지 않습니다: {html_folder}")
        return
    
    # 출력 파일 경로 설정
    output_path = Path(html_folder) / "all_pages_editable_v6.pptx"
    
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
    else:
        print("❌ 변환 실패!")

if __name__ == "__main__":
    main()

