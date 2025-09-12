#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HTML 스타일 수정 스크립트
파란색 헤더 라인 제거 및 여백 조정
"""

import os
import re
from pathlib import Path

def fix_html_file(html_file_path):
    """HTML 파일의 스타일을 수정"""
    try:
        with open(html_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 파란색 헤더 라인 제거
        content = re.sub(r'\.header-accent\s*\{[^}]*\}', '', content)
        content = re.sub(r'<div class="header-accent"></div>', '', content)
        
        # body 스타일 수정 - 여백 제거
        content = re.sub(
            r'body\s*\{[^}]*\}',
            '''body {
            margin: 0;
            padding: 0;
            font-family: 'Noto Sans KR', sans-serif;
            overflow: hidden;
            width: 100%;
            height: 100%;
        }''',
            content
        )
        
        # slide-container 스타일 수정 - 여백 제거
        content = re.sub(
            r'\.slide-container\s*\{[^}]*\}',
            '''.slide-container {
            width: 100%;
            height: 100%;
            background-color: white;
            position: relative;
            overflow: hidden;
            margin: 0;
            padding: 0;
        }''',
            content
        )
        
        # content-area 패딩 조정
        content = re.sub(
            r'\.content-area\s*\{[^}]*padding:[^;]*;',
            '.content-area {\n            padding: 20px;\n',
            content
        )
        
        # slide-header 패딩 조정
        content = re.sub(
            r'\.slide-header\s*\{[^}]*padding:[^;]*;',
            '.slide-header {\n            padding: 20px 40px 10px;\n',
            content
        )
        
        # 전체 페이지 여백 제거를 위한 추가 스타일
        if 'html {' not in content:
            # html 스타일 추가
            content = content.replace(
                'body {',
                '''html {
            margin: 0;
            padding: 0;
            width: 100%;
            height: 100%;
        }
        body {'''
            )
        
        # 파일 저장
        with open(html_file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"수정 완료: {html_file_path.name}")
        return True
        
    except Exception as e:
        print(f"수정 오류 ({html_file_path.name}): {e}")
        return False

def main():
    html_dir = Path(r"C:\Project\gigabitamin\notion\doc_ppt\hearth_chat")
    
    print("HTML 스타일 수정 시작")
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
        if fix_html_file(html_file):
            success_count += 1
    
    print("-" * 50)
    print(f"수정 완료: {success_count}/{len(html_files)}개 파일")
    
    if success_count == len(html_files):
        print("모든 HTML 파일이 성공적으로 수정되었습니다!")
    else:
        print("일부 파일 수정에 실패했습니다.")

if __name__ == "__main__":
    main()
