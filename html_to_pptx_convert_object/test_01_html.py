#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
01.html 전용 테스트
"""

from debug_html_to_pptx_converter import DebugHTMLConverter

def main():
    html_file = r"C:\Project\gigabitamin\genspark\dcs_site\html\01.html"
    output_path = r"C:\Project\gigabitamin\genspark\dcs_site\html\test_01.html.pptx"
    
    print("01.html 전용 테스트")
    print("=" * 50)
    
    converter = DebugHTMLConverter(html_file, output_path)
    success = converter.convert()
    
    if success:
        print("✅ 01.html 변환 완료!")
    else:
        print("❌ 01.html 변환 실패!")

if __name__ == "__main__":
    main()
