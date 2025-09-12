#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HTML 여백 완전 제거 스크립트
오른쪽 여백 완전 제거
"""

import os
import re
from pathlib import Path

def fix_html_margins(html_file_path):
    """HTML 파일의 여백을 완전히 제거"""
    try:
        with open(html_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 전체 스타일을 완전히 재작성
        new_styles = '''        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        html {
            margin: 0;
            padding: 0;
            width: 100%;
            height: 100%;
            overflow: hidden;
        }
        body {
            margin: 0;
            padding: 0;
            font-family: 'Noto Sans KR', sans-serif;
            overflow: hidden;
            width: 100%;
            height: 100%;
        }
        .slide-container {
            width: 100%;
            height: 100%;
            background-color: white;
            position: relative;
            overflow: hidden;
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
        }
        .content-area {
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            height: 100%;
            width: 100%;
            padding: 20px;
            box-sizing: border-box;
            margin: 0;
        }
        .slide-header {
            padding: 20px 40px 10px;
            width: 100%;
            margin: 0;
            box-sizing: border-box;
        }
        .slide-title {
            color: #1e3a8a;
            font-weight: 700;
            font-size: 36px;
            margin: 0;
            width: 100%;
        }
        .title {
            color: #1e3a8a;
            font-weight: 700;
            margin-bottom: 20px;
            text-align: center;
            font-size: 48px;
            line-height: 1.2;
            width: 100%;
        }
        .subtitle {
            color: #3b82f6;
            font-weight: 500;
            margin-bottom: 60px;
            text-align: center;
            font-size: 28px;
            width: 100%;
        }
        .info-box {
            display: flex;
            flex-direction: column;
            align-items: center;
            margin-top: 20px;
            padding: 20px;
            border-top: 1px solid #e5e7eb;
            width: 100%;
            box-sizing: border-box;
        }
        .info-item {
            margin: 8px 0;
            color: #4b5563;
            font-size: 18px;
            width: 100%;
            text-align: center;
        }
        .highlight {
            color: #2563eb;
            font-weight: 500;
        }
        .footer {
            position: absolute;
            bottom: 30px;
            width: 100%;
            text-align: center;
            color: #6b7280;
            font-size: 16px;
            margin: 0;
            padding: 0;
        }
        .intro-text {
            color: #4b5563;
            font-size: 18px;
            margin-bottom: 30px;
            line-height: 1.6;
            width: 100%;
        }
        .section-title {
            color: #2563eb;
            font-weight: 600;
            font-size: 24px;
            margin-bottom: 15px;
            width: 100%;
        }
        .feature-list {
            list-style-type: none;
            padding: 0;
            width: 100%;
        }
        .feature-item {
            display: flex;
            align-items: flex-start;
            margin-bottom: 16px;
            color: #4b5563;
            font-size: 18px;
            width: 100%;
        }
        .feature-icon {
            color: #2563eb;
            margin-right: 12px;
            width: 22px;
            flex-shrink: 0;
        }'''
        
        # 기존 스타일을 새 스타일로 교체
        content = re.sub(r'<style>.*?</style>', f'<style>\n{new_styles}\n    </style>', content, flags=re.DOTALL)
        
        # 파일 저장
        with open(html_file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"여백 제거 완료: {html_file_path.name}")
        return True
        
    except Exception as e:
        print(f"여백 제거 오류 ({html_file_path.name}): {e}")
        return False

def main():
    html_dir = Path(r"C:\Project\gigabitamin\notion\doc_ppt\hearth_chat")
    
    print("HTML 여백 완전 제거 시작")
    print(f"대상 디렉토리: {html_dir}")
    print("-" * 50)
    
    # HTML 파일 목록 가져오기
    html_files = list(html_dir.glob("*.html"))
    
    if not html_files:
        print("HTML 파일을 찾을 수 없습니다.")
        return
    
    print(f"발견된 HTML 파일: {len(html_files)}개")
    
    # 각 HTML 파일 수정
    success_count = 0
    for html_file in html_files:
        if fix_html_margins(html_file):
            success_count += 1
    
    print("-" * 50)
    print(f"여백 제거 완료: {success_count}/{len(html_files)}개 파일")
    
    if success_count == len(html_files):
        print("모든 HTML 파일의 여백이 완전히 제거되었습니다!")
    else:
        print("일부 파일 수정에 실패했습니다.")

if __name__ == "__main__":
    main()
