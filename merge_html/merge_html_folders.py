import os
from pathlib import Path

import re


TEMPLATE_HEAD = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>{{TITLE}}</title>
  <style>
    /* 페이지 공통 스타일 (컨테이너 및 iframe 레이아웃) */
    body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Noto Sans KR', Arial, 'Apple Color Emoji', 'Segoe UI Emoji', 'Segoe UI Symbol'; background: #f8fafc; color: #0f172a; }
    .page { max-width: 1280px; margin: 0 auto; padding: 24px; }
    h1 { font-size: 24px; margin: 16px 0 24px; }
    .section { background: #ffffff; border: 1px solid #e5e7eb; border-radius: 10px; padding: 16px; margin-bottom: 20px; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
    .section-title { font-size: 18px; margin: 0 0 12px; color: #334155; }
    .file-title { font-size: 14px; margin: 12px 0; color: #475569; }
    .frame-wrap { border: 1px solid #e2e8f0; border-radius: 8px; background: #f8fafc; overflow: hidden; }
    /* 기본 높이를 크게 지정 (파일 프로토콜 교차출처 제약으로 자동조절 실패 대비) */
    iframe { width: 100%; border: 0; display: block; background: #ffffff; height: 1200px; }
  </style>
  <script>
    // 화면에서 가로 스크롤이 강제로 막히는 경우를 대비한 보정
    (function() {
      function enableScrolling() {
        try {
          document.documentElement.style.setProperty('overflow-x', 'auto', 'important');
          document.documentElement.style.setProperty('overflow-y', 'auto', 'important');
          document.body.style.setProperty('overflow-x', 'auto', 'important');
          document.body.style.setProperty('overflow-y', 'auto', 'important');
          // 콘텐츠의 최대 가로 폭을 계산해 래퍼 폭을 설정 (가로 스크롤 유도)
          var containers = Array.prototype.slice.call(document.querySelectorAll('.slide-container'));
          if (containers.length) {
            var maxW = containers.reduce(function(m, el){ return Math.max(m, el.scrollWidth || el.clientWidth || 0); }, 0);
            var page = document.querySelector('.page');
            if (page && maxW > 0) {
              page.style.width = (maxW + 40) + 'px'; // 여유 40px
            }
          }
        } catch (e) {}
      }
      document.addEventListener('DOMContentLoaded', enableScrolling);
      window.addEventListener('load', enableScrolling);
      window.addEventListener('resize', enableScrolling);
    })();
  </script>
  <script>
    // 네트워크 환경에서 열릴 경우 자동 높이 조절을 시도하지만,
    // 파일 프로토콜(file://)에서는 교차출처 정책으로 실패할 수 있음.
    function resizeIframes() {
      const frames = document.querySelectorAll('iframe[data-auto-height]');
      frames.forEach((frame) => {
        try {
          const doc = frame.contentDocument || frame.contentWindow.document;
          if (!doc) return;
          const height = Math.max(
            doc.body ? doc.body.scrollHeight : 0,
            doc.documentElement ? doc.documentElement.scrollHeight : 0,
            0
          );
          if (height > 0) frame.style.height = (height + 10) + 'px';
        } catch (e) { /* ignore */ }
      });
    }
    window.addEventListener('load', () => {
      setTimeout(resizeIframes, 300);
      setTimeout(resizeIframes, 1200);
      setTimeout(resizeIframes, 2500);
    });
  </script>
</head>
<body>
  <div class="page">
    <h1>{{TITLE}}</h1>
"""

TEMPLATE_FOOT = """
  </div>
</body>
</html>
"""


def build_section_html(title: str, items: list[str]) -> str:
    html_parts: list[str] = []
    html_parts.append(f'<div class="section">')
    html_parts.append(f'  <div class="section-title">{title}</div>')
    for label, src in items:
        html_parts.append(f'  <div class="file-title">{label}</div>')
        html_parts.append('  <div class="frame-wrap">')
        # data-auto-height 속성으로 높이 자동 조절
        html_parts.append(f'    <iframe data-auto-height src="{src}"></iframe>')
        html_parts.append('  </div>')
    html_parts.append('</div>')
    return "\n".join(html_parts)


def find_html_files(folder: Path) -> list[Path]:
    return sorted([p for p in folder.glob('*.html') if p.is_file()])


def make_relative_src(base: Path, target: Path) -> str:
    try:
        return os.path.relpath(target, start=base)
    except Exception:
        # 실패 시 절대 경로 사용 (file:// scheme은 필요 없음, 브라우저가 로컬 경로를 열 수 있음)
        return str(target)


def merge_folder(folder_html: Path, out_name: str) -> Path:
    files = find_html_files(folder_html)
    # index/merged 파일 등은 제외 (자기 자신 포함 방지)
    files = [f for f in files if not f.name.lower().endswith(('_all.html', 'all.html', 'merged.html', 'index.html'))]

    title = f"{folder_html.parent.name} - 통합 페이지"
    parts: list[str] = [TEMPLATE_HEAD.replace("{{TITLE}}", title)]

    if not files:
        parts.append('<div class="section"><div class="section-title">HTML 파일이 없습니다.</div></div>')
    else:
        items = []
        for f in files:
            label = f.name
            # 출력 파일은 상위 폴더에 생성되므로, 상위 폴더 기준의 상대경로를 생성
            rel_src = make_relative_src(folder_html.parent, f)
            items.append((label, rel_src))
        parts.append(build_section_html('페이지 목록', items))

    parts.append(TEMPLATE_FOOT)

    out_file = folder_html.parent / out_name
    out_file.write_text("\n".join(parts), encoding='utf-8')
    return out_file


# -------------------- 인쇄용(iframe 미사용, 인라인 병합) --------------------
PRINT_HEAD = """<!DOCTYPE html>
<html lang=\"ko\">
<head>
  <meta charset=\"UTF-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />
  {{HEAD_ASSETS}}
  <style>
    /* 가로 인쇄만 강제 (최소 설정) */
    @page { size: A4 landscape; }
    html, body { margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Noto Sans KR', Arial; background: #fff; color: #0f172a; }
    .page { margin: 0; padding: 0; }
    h1 { font-size: 22px; margin: 0 0 16px; }
    .doc { margin: 0 0 16px; }
    .doc-body { }

    /* 인쇄 시 16:9(1920x1080) 기준 페이지를 용지 폭에 맞게 균일 축소 */
    @media print {
      :root {
        /* A4 가로 용지의 인쇄 가능 폭(여백 제외) */
        --print-width: 277mm;
        --design-width: 1920px; /* 원본 화면 기준 */
        --scale: calc(var(--print-width) / var(--design-width));
      }
      .doc-body { display: flex; justify-content: center; }
      .print-fit {
        width: var(--design-width);
        display: inline-block;
        zoom: var(--scale);
        -webkit-transform-origin: top center;
      }
      /* 분할 억제만 하고 강제 개행은 하지 않음 */
      .doc { break-inside: avoid; page-break-inside: avoid; }
    }

    /* 화면에서 스크롤이 막히는 문제 해결: 강제 오버라이드 */
    @media screen {
      /* 화면에서 가로/세로 스크롤 모두 강제 허용 */      
      body { overflow-x: scroll !important; overflow-y: auto !important; }
      /* 컨텐츠 최소 폭 보장 (슬라이드 기준 1920px) */
      .page { width: max-content; min-width: 1920px; }
      .doc, .doc-body { min-width: 1920px; }
      /* 원본 슬라이드가 내부에서 잘리지 않도록 */
      .slide-container { overflow: visible !important; min-width: 1920px; }
    }
  </style>
</head>
<body>
  <div class=\"page\">
"""

PRINT_FOOT = """
  </div>
</body>
</html>
"""


def extract_body_html(html_text: str) -> str:
    # body 태그 내부만 추출. 없으면 전체 반환
    m = re.search(r"<body[^>]*>([\s\S]*?)</body>", html_text, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    # body가 없으면 html 전체에서 html/head 태그 제거
    cleaned = re.sub(r"</?html[^>]*>", "", html_text, flags=re.IGNORECASE)
    cleaned = re.sub(r"</?head[^>]*>[\s\S]*?</head>", "", cleaned, flags=re.IGNORECASE)
    return cleaned


def remove_footers(html_text: str) -> str:
    """불필요한 푸터/참조 블록 제거"""
    text = html_text
    # <footer> 태그 전체 제거
    text = re.sub(r"<footer[^>]*>[\s\S]*?</footer>", "", text, flags=re.IGNORECASE)
    # class에 footer 포함된 블록 제거
    text = re.sub(r"<([a-zA-Z0-9]+)([^>]*class=\"[^\"]*footer[^\"]*\"[^>]*)>[\s\S]*?</\\\1>", "", text, flags=re.IGNORECASE)
    # 하단 고정 참조: bottom-*, right-* 유틸 클래스 조합 제거 (Tailwind)
    text = re.sub(r"<div[^>]*class=\"[^\"]*(bottom-\d+[^\"]*right-\d+|right-\d+[^\"]*bottom-\d+)[^\"]*\"[^>]*>[\s\S]*?</div>", "", text, flags=re.IGNORECASE)
    # 특정 문구 포함 요소 제거
    text = re.sub(r"<div[^>]*>[^<]*개발 프로젝트[^<]*</div>", "", text, flags=re.IGNORECASE)
    return text

def extract_head_assets(html_text: str, base_path: Path) -> tuple[list[str], list[str]]:
    """각 문서의 <link rel="stylesheet"> href들과 <style> 내용을 추출"""
    links: list[str] = []
    styles: list[str] = []

    head_match = re.search(r"<head[^>]*>([\s\S]*?)</head>", html_text, flags=re.IGNORECASE)
    if not head_match:
        return links, styles
    head = head_match.group(1)

    # link rel=stylesheet
    for m in re.finditer(r"<link[^>]+rel=\"?stylesheet\"?[^>]*>", head, flags=re.IGNORECASE):
        tag = m.group(0)
        href_m = re.search(r"href=\"([^\"]+)\"|href='([^']+)'", tag, flags=re.IGNORECASE)
        if not href_m:
            continue
        href = href_m.group(1) or href_m.group(2)
        # 절대/상대 경로 처리 (상대경로는 통합 파일 위치 기준으로 재계산 필요 없으므로 원본 상대 그대로 사용)
        links.append(href)

    # style 태그 내용 수집
    for m in re.finditer(r"<style[^>]*>([\s\S]*?)</style>", head, flags=re.IGNORECASE):
        styles.append(m.group(1))

    return links, styles


def merge_folder_print(folder_html: Path, out_name: str) -> Path:
    files = find_html_files(folder_html)
    files = [f for f in files if not f.name.lower().endswith(('_all.html', 'all.html', 'merged.html', 'index.html'))]

    title = f"{folder_html.parent.name} - 통합 페이지(인쇄용)"

    # 모든 문서의 CSS 링크/스타일 수집 (중복 제거)
    link_set: dict[str, None] = {}
    style_list: list[str] = []
    for f in files:
        text = f.read_text(encoding='utf-8', errors='ignore')
        links, styles = extract_head_assets(text, f.parent)
        for href in links:
            link_set[href] = None
        style_list.extend(styles)

    head_assets = []
    for href in link_set.keys():
        head_assets.append(f'<link rel="stylesheet" href="{href}">')
    if style_list:
        head_assets.append('<style>\n' + "\n".join(style_list) + '\n</style>')

    head_final = PRINT_HEAD.replace("{{TITLE}}", title).replace("{{HEAD_ASSETS}}", "\n".join(head_assets))
    parts: list[str] = [head_final]

    if not files:
        parts.append('<div class="doc"><div class="doc-title">HTML 파일이 없습니다.</div></div>')
    else:
        for idx, f in enumerate(files, start=1):
            text = f.read_text(encoding='utf-8', errors='ignore')
            body_inner = extract_body_html(text)
            body_inner = remove_footers(body_inner)
            # 개별 페이지 미세 조정: 파일명 기반 규칙 예시
            data_attrs = ''
            name = f.name.lower()
            if name.startswith('01'):
                data_attrs = ' data-scale="95"'
            elif name.startswith('02'):
                data_attrs = ' data-scale="92"'
            # 필요 시 추가 규칙을 아래에 확장
            parts.append(f'<div class="doc"{data_attrs}>')
            parts.append('  <div class="doc-body">')
            parts.append('    <div class="print-fit">')
            parts.append(body_inner)
            parts.append('    </div>')
            parts.append('  </div>')
            parts.append('</div>')

    parts.append(PRINT_FOOT)

    out_file = folder_html.parent / out_name
    out_file.write_text("\n".join(parts), encoding='utf-8')
    return out_file


def main():
    targets = [
        (Path(r"C:\\Project\\gigabitamin\\genspark\\dcs_site\\html"),        "dcs_site_all.html"),
        (Path(r"C:\\Project\\gigabitamin\\genspark\\doc_analystic\\html"),   "doc_analystic_all.html"),
        (Path(r"C:\\Project\\gigabitamin\\genspark\\farm_quest\\html"),     "farm_quest_all.html"),
        (Path(r"C:\\Project\\gigabitamin\\genspark\\hearth_chat\\html"),    "hearth_chat_all.html"),
        (Path(r"C:\\Project\\gigabitamin\\genspark\\inneats\\html"),        "inneats_all.html"),
        (Path(r"C:\\Project\\gigabitamin\\genspark\\smart_gate\\html"),     "smart_gate_all.html"),
    ]

    for folder_html, out_name in targets:
        if not folder_html.exists():
            print(f"경고: 폴더가 존재하지 않습니다: {folder_html}")
            continue
        try:
            out_path = merge_folder(folder_html, out_name)
            print(f"✅ 생성됨: {out_path}")
        except Exception as e:
            print(f"❌ 실패: {folder_html} -> {out_name} | {e}")

    # 인쇄용 파일도 함께 생성 (*.*_all_print.html)
    print("\n인쇄용 파일 생성 중...")
    print_targets = [
        (Path(r"C:\\Project\\gigabitamin\\genspark\\dcs_site\\html"),        "dcs_site_all_print.html"),
        (Path(r"C:\\Project\\gigabitamin\\genspark\\doc_analystic\\html"),   "doc_analystic_all_print.html"),
        (Path(r"C:\\Project\\gigabitamin\\genspark\\farm_quest\\html"),     "farm_quest_all_print.html"),
        (Path(r"C:\\Project\\gigabitamin\\genspark\\hearth_chat\\html"),    "hearth_chat_all_print.html"),
        (Path(r"C:\\Project\\gigabitamin\\genspark\\inneats\\html"),        "inneats_all_print.html"),
        (Path(r"C:\\Project\\gigabitamin\\genspark\\smart_gate\\html"),     "smart_gate_all_print.html"),
    ]

    for folder_html, out_name in print_targets:
        if not folder_html.exists():
            continue
        try:
            out_path = merge_folder_print(folder_html, out_name)
            print(f"🖨️ 인쇄용 생성됨: {out_path}")
        except Exception as e:
            print(f"❌ 인쇄용 실패: {folder_html} -> {out_name} | {e}")


if __name__ == "__main__":
    main()


