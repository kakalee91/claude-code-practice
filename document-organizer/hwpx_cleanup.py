"""
KCS 문서(hwpx) 일괄 정리 스크립트

동작:
  1) 표지는 남기되, 표지 안의 워터마크/로고 이미지, 색깔 채워진 셀(파란 바탕 등),
     "표준시방서 Korean Construction Specification" 부제, "OOOO년 O월 O일 개정" 날짜,
     "http://..." URL 텍스트를 제거해서 표지를 심플하게 만든다.
  2) 표지와 목차 사이에 있는 "경과조치/건설기준 연혁" 페이지를 삭제한다. (본문/목차는 안 건드림)
  3) 본문 끝의 "집필위원" 이후 페이지들(집필위원/자문위원 표, 마지막 작성기관 페이지)을 통째로 삭제한다.
  4) 마스터페이지(바탕쪽/배경쪽)에 박혀있는 워터마크 이미지, 색/이미지로 채워진
     배경(표 형태로 들어있는 경우 포함), "제정/개정/심의/소관부서" 등 이력 정보란
     표(순수 텍스트 표인 경우도 포함)도 제거한다.
  5) 문서에 설정된 꼬리말(쪽 하단 반복 영역)과 쪽번호(자동 채번 필드)를 삭제한다.

사용법:
    pip install lxml
    python hwpx_cleanup.py "폴더경로"                # 폴더 안의 모든 .hwpx 처리
    python hwpx_cleanup.py "폴더경로" --recursive     # 하위 폴더까지

결과물은 원본을 건드리지 않고, 같은 폴더에 "원본이름_정리됨.hwpx"로 저장한다.
처리가 애매한 파일(마커를 못 찾은 경우)은 건드리지 않고 경고만 출력한다.
"""

import sys
import os
import glob
import re
import shutil
import zipfile
import argparse
import tempfile

try:
    from lxml import etree
except ImportError:
    sys.exit("lxml이 필요합니다. 먼저 'pip install lxml'을 실행하세요.")

HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'
HC = '{http://www.hancom.co.kr/hwpml/2011/core}'

DATE_REVISION_RE = re.compile(r'\d{4}\s*년\s*\d{1,2}\s*월\s*\d{1,2}\s*일\s*(개정|제정)')
SUBTITLE_RE = re.compile(r'표준시방서\s*Korean\s*Construction\s*Specification', re.I)

# 표지에 섞여 남아있는 '이 문서와 무관한 다른 기준 계열' 잔재 텍스트 표시.
# (예: KCS 표준시방서 표지에 KDS 설계기준 옛 표지 텍스트가 안 지워지고 같이 남아있는 경우)
FOREIGN_STANDARD_RE = re.compile(
    r'KDS\s*\d|Design\s*Standard|설계기준|건축구조기준|국가건설기준(?!센터)',
    re.I
)

# "KCS 61 10 05 : 2017" 처럼 문서 고유 코드+연도가 표기된 문구.
CODE_YEAR_RE = re.compile(r'[A-Za-z]{2,4}\s*\d{2}\s*\d{2}\s*\d{2}\s*[:：]\s*\d{4}')
# "KCS 61 00 00" 처럼 대분류(끝 두 자리가 00 00)만 표기된 코드.
DIVISION_CODE_RE = re.compile(r'^[A-Za-z]{2,4}\s*\d{2}\s*00\s*00\s*$')

# 표지/바탕쪽(마스터페이지)에 "제정 : 2016년 6월 30일 / 심의 : .../ 소관부서 : .../
# 관련단체(작성기관) : ..." 식으로 박혀있는 제정·개정 이력 정보란 표를 찾기 위한 라벨들.
# 글자 사이에 공백이 섞여 있는 경우가 많아 각 글자 사이에 공백을 허용한다.
INFO_TABLE_LABEL_RES = [
    re.compile(r'제\s*정'),
    re.compile(r'개\s*정'),
    re.compile(r'심\s*의'),
    re.compile(r'자\s*문\s*검\s*토'),
    re.compile(r'소\s*관\s*부\s*서'),
    re.compile(r'관\s*련\s*단\s*체'),
]


def norm(s):
    return (s or '').replace(' ', '').replace(' ', '')


def para_text(p):
    return ''.join(p.itertext())


def find_plain_borderfill_id(header_path, used_ids):
    """cover 표에서 실제 쓰인 borderFillIDRef 중, 색/이미지 채우기가 전혀 없는 id를
    가장 많이 쓰인 순서로 골라 '민짜(투명)' 스타일로 사용한다."""
    data = open(header_path, encoding='utf-8').read()
    plain_ids = []
    for m in re.finditer(r'<hh:borderFill id="(\d+)"[^>]*>(.*?)</hh:borderFill>', data, re.S):
        bid, block = m.group(1), m.group(2)
        if bid in used_ids and '<hc:fillBrush>' not in block:
            plain_ids.append(bid)
    if not plain_ids:
        return None
    # used_ids는 Counter이므로 가장 많이 쓰인 것을 우선
    plain_ids.sort(key=lambda i: -used_ids[i])
    return plain_ids[0]


def find_colored_borderfill_ids(header_path):
    """색(faceColor) 또는 이미지(imgBrush)가 채워진 borderFill id 집합을 반환."""
    data = open(header_path, encoding='utf-8').read()
    colored = set()
    for m in re.finditer(r'<hh:borderFill id="(\d+)"[^>]*>(.*?)</hh:borderFill>', data, re.S):
        bid, block = m.group(1), m.group(2)
        if '<hc:fillBrush>' in block:
            colored.add(bid)
    return colored


# 각 정리 항목의 기본 켬/끔 상태. GUI 체크박스와 1:1로 대응한다.
DEFAULT_OPTIONS = {
    'cover_image': True,      # 표지: 워터마크/로고 이미지 제거
    'cover_color': True,      # 표지: 색깔 채워진 배경 제거
    'cover_infobox': True,    # 표지/바탕쪽: 제정/개정/심의 등 정보란 표 제거
    'cover_subtitle': True,   # 표지: "표준시방서 Korean Construction Specification" 부제 제거
    'cover_date': True,       # 표지: "OOOO년 O월 O일 개정/제정" 날짜 제거
    'cover_url': True,        # 표지: "http://..." URL 제거
    'cover_foreign': True,    # 표지: 다른 기준(KDS 등) 잔재 텍스트 제거
    'cover_duplicate': True,  # 표지: 같은 표지 블록이 통째로 중복된 경우 앞쪽(옛) 블록 제거
    'frontmatter_gap': True,  # 표지~목차 사이 경과조치/연혁 페이지 삭제
    'trailing_matter': True,  # 본문 끝 집필위원 이후 페이지 삭제
    'remove_footer': True,       # 꼬리말 삭제 (안의 쪽번호도 함께 삭제됨)
    'remove_page_number': True,  # 꼬리말과 별도로 남은 쪽번호 필드 삭제
}

OPTION_LABELS = {
    'cover_image': '표지/바탕쪽: 워터마크/로고 이미지 제거',
    'cover_color': '표지/바탕쪽: 색깔·이미지로 채워진 배경 제거 (파란 바, 회색 줄 등)',
    'cover_infobox': '표지/바탕쪽: 제정/개정/심의/소관부서 등 정보란 표 제거',
    'cover_subtitle': '표지: "표준시방서 Korean Construction Specification" 부제 제거',
    'cover_date': '표지: 개정/제정 날짜 텍스트 제거',
    'cover_url': '표지: URL(http://...) 텍스트 제거',
    'cover_foreign': '표지: 다른 기준(KDS 설계기준 등) 잔재 텍스트 제거',
    'cover_duplicate': '표지: 중복된 옛 표지 블록(코드/제목 반복) 제거',
    'frontmatter_gap': '표지~목차 사이 경과조치/연혁 페이지 삭제',
    'trailing_matter': '본문 끝 집필위원 이후 페이지(+ 앞 빈 페이지) 삭제',
    'remove_footer': '꼬리말 삭제 (안의 쪽번호도 함께 삭제)',
    'remove_page_number': '쪽번호 삭제 (꼬리말과 별도로 남은 경우)',
}


def remove_duplicate_cover_block(cover):
    """표지 전체가 (대분류코드/제목/코드:연도/제목) 4개 묶음으로 두 번 이상
    들어있는 경우, 실제 문서 고유 코드(코드:연도)가 있는 마지막 묶음만
    남기고 그 앞의 옛 묶음 텍스트를 지운다.

    (예: KCS 61 00 00/하수관로공사/KCS 61 00 00:2018/하수관로공사 <- 옛 묶음
         KCS 61 00 00/하수관로 공사/KCS 61 10 05:2017/공통사항      <- 진짜 묶음)
    """
    t_elems = list(cover.iter(HP + 't'))
    anchor_idxs = [i for i, t in enumerate(t_elems) if CODE_YEAR_RE.search(t.text or '')]
    if len(anchor_idxs) < 2:
        return False

    last_anchor = anchor_idxs[-1]
    division_start = None
    for i in range(last_anchor - 1, -1, -1):
        if DIVISION_CODE_RE.match((t_elems[i].text or '').strip()):
            division_start = i
            break

    if division_start is None or division_start == 0:
        return False

    for i in range(division_start):
        t_elems[i].text = ''
    return True


def remap_colored_fills(root, header_path):
    """root 안에서 borderFillIDRef 속성을 쓰는 모든 요소(표 셀/표/도형 등)를 대상으로,
    색이나 이미지로 채워진 fill을 '민짜(투명)' 스타일로 바꾼다.

    표지 본문뿐 아니라 마스터페이지(바탕쪽/배경쪽)에도 같은 방식으로 워터마크/색배경이
    표 형태로 박혀있는 경우가 있어서, 표지 전용이 아니라 공용 함수로 뺐다."""
    from collections import Counter
    used = Counter()
    tagged = []
    for el in root.iter():
        ref = el.get('borderFillIDRef')
        if ref:
            used[ref] += 1
            tagged.append(el)

    if not tagged:
        return False

    colored_ids = find_colored_borderfill_ids(header_path)
    plain_id = find_plain_borderfill_id(header_path, used)
    if not plain_id:
        return False

    changed = False
    for el in tagged:
        if el.get('borderFillIDRef') in colored_ids:
            el.set('borderFillIDRef', plain_id)
            changed = True
    return changed


def remove_pics(root):
    removed = 0
    for pic in list(root.iter(HP + 'pic')):
        parent = pic.getparent()
        if parent is not None:
            parent.remove(pic)
            removed += 1
    return removed


def remove_cover_info_table(root):
    """"제정 : 2016년 6월 30일 / 심의 : ... / 소관부서 : ... / 관련단체(작성기관) : ..."
    처럼 라벨이 모여있는 제정·개정 이력 정보란 표를 통째로 지운다. 그림도 색배경도 아닌
    순수 텍스트 표라서 remove_pics/remap_colored_fills로는 안 지워지고, 표지 본문뿐
    아니라 바탕쪽(마스터페이지)에도 박혀있는 경우가 있어 양쪽 다 검사한다."""
    removed = 0
    for tbl in list(root.iter(HP + 'tbl')):
        text = ''.join(tbl.itertext())
        matched = sum(1 for pat in INFO_TABLE_LABEL_RES if pat.search(text))
        if matched >= 2:
            parent = tbl.getparent()
            if parent is not None:
                parent.remove(tbl)
                removed += 1
    return removed


def remove_footer_ctrl(root):
    """머리말/꼬리말 통제 개체(<hp:ctrl><hp:header>...</hp:header></hp:ctrl> 또는
    <hp:footer> 버전)를 통째로 지운다. 안에 들어있는 쪽번호 필드도 함께 사라진다."""
    removed = 0
    for ctrl in list(root.iter(HP + 'ctrl')):
        if ctrl.find(HP + 'footer') is not None or ctrl.find(HP + 'header') is not None:
            parent = ctrl.getparent()
            if parent is not None:
                parent.remove(ctrl)
                removed += 1
    return removed


def remove_page_number_fields(root):
    """쪽번호를 지운다. HWP에는 쪽번호를 넣는 방식이 두 가지 있다:
      1) 꼬리말/머리말 안에 <hp:autoNum numType="PAGE"> 채번 필드를 넣는 방식
         (꼬리말/머리말 자체를 지우면 이 필드도 같이 사라지지만, 머리말/꼬리말 없이
         쪽번호만 단독으로 들어있는 경우를 위해 여기서도 한 번 더 확인한다)
      2) "쪽 번호 매기기" 기능으로 본문에 직접 <hp:pageNum pos="BOTTOM_CENTER" .../>
         지시자를 심는 방식 (예: "- 1 -" 처럼 페이지 하단 중앙에 표시). 이건 꼬리말이
         아니라 섹션 본문 문단에 바로 박히기 때문에 꼬리말 삭제로는 안 지워진다.
    두 경우 모두 해당 <hp:ctrl>을 통째로 지운다."""
    removed = 0
    for ctrl in list(root.iter(HP + 'ctrl')):
        auto = ctrl.find(HP + 'autoNum')
        is_auto_page = auto is not None and auto.get('numType') == 'PAGE'
        is_page_num = ctrl.find(HP + 'pageNum') is not None
        if is_auto_page or is_page_num:
            parent = ctrl.getparent()
            if parent is not None:
                parent.remove(ctrl)
                removed += 1
    return removed


def clean_cover(root, header_path, opts):
    """root[0]을 표지로 간주하고 opts에서 켜진 항목만 제거."""
    if len(root) == 0:
        return
    cover = root[0]

    if opts.get('cover_duplicate', True):
        remove_duplicate_cover_block(cover)

    if opts.get('cover_color', True):
        remap_colored_fills(cover, header_path)

    if opts.get('cover_image', True):
        remove_pics(cover)

    if opts.get('cover_infobox', True):
        remove_cover_info_table(cover)

    for t in cover.iter(HP + 't'):
        txt = t.text or ''
        remove = False
        if opts.get('cover_subtitle', True) and SUBTITLE_RE.search(txt):
            remove = True
        if opts.get('cover_date', True) and DATE_REVISION_RE.search(txt):
            remove = True
        if opts.get('cover_url', True) and txt.strip().startswith('http'):
            remove = True
        if opts.get('cover_foreign', True) and FOREIGN_STANDARD_RE.search(txt):
            remove = True
        if remove:
            t.text = ''


def remove_frontmatter_gap(root):
    """표지 ~ 목차 사이의 경과조치/연혁 페이지 삭제. 마커 못 찾으면 아무 것도 안 함.

    목차가 시작되는 문단(idx_toc)을 먼저 찾고, 표지(index 0) 다음으로 나오는
    첫 페이지 나눔(pageBreak=1) 지점을 삭제 시작점으로 삼는다. 표지 바로 뒤에
    나오는 빈 문단들도 "경과조치 페이지"의 일부(문단 앞의 여백)이기 때문이다.
    다만 오탐을 막기 위해 그 구간 안에 실제로 '경과조치' 관련 문구가 있을 때만 삭제한다.
    """
    children = list(root)

    idx_toc = None
    for i, p in enumerate(children):
        if i == 0:
            continue
        if '목차' in norm(para_text(p)):
            idx_toc = i
            break
    if idx_toc is None:
        return False

    idx_start = None
    has_revision_marker = False
    for i in range(1, idx_toc):
        p = children[i]
        if idx_start is None and p.get('pageBreak') == '1':
            idx_start = i
        txt = norm(para_text(p))
        if '경과' in txt and '조치' in txt:
            has_revision_marker = True

    if idx_start is not None and has_revision_marker:
        # idx_start 위치의 pageBreak=1 표시(새 페이지 시작)를 지금 지우는 것이므로,
        # 목차 문단이 그 표시를 이어받아야 표지와 같은 페이지에 붙지 않는다.
        children[idx_toc].set('pageBreak', '1')
        for i in sorted(range(idx_start, idx_toc), reverse=True):
            root.remove(children[i])
        return True
    return False


def is_visually_empty(p):
    """텍스트도 없고, 그림/표도 없는 순수 빈 문단인지."""
    if para_text(p).strip():
        return False
    if p.find('.//' + HP + 'pic') is not None:
        return False
    if p.find('.//' + HP + 'tbl') is not None:
        return False
    return True


def remove_trailing_blank_page(root):
    """문서 맨 끝이 '내용도 없이 페이지 나눔만 있는 빈 페이지'로 끝나면 그 페이지를 삭제.
    ('집필위원' 같은 문구가 없는 문서에서, 뒷장이 통째로 잘려 빈 페이지만 남은 경우 처리)"""
    children = list(root)
    n = len(children)
    i = n - 1
    while i >= 0 and is_visually_empty(children[i]):
        i -= 1
    # children[i+1 .. n-1] 는 전부 빈 문단. 그 중 pageBreak=1로 시작하는 지점을 찾는다.
    start = None
    for j in range(i + 1, n):
        if children[j].get('pageBreak') == '1':
            start = j
            break
    if start is not None:
        for k in sorted(range(start, n), reverse=True):
            root.remove(children[k])
        return True
    return False


DEBUG = os.environ.get('HWPX_DEBUG', '').strip() not in ('', '0')


def dbg(*args):
    if DEBUG:
        print('[DEBUG]', *args)


def safe_parse(path):
    """etree.parse를 감싸서, section*.xml이 비어있거나 손상돼 못 여는 경우
    lxml의 알아보기 힘든 원본 에러 대신 원인과 대처법을 알려주는 메시지를 낸다.
    (원본 hwpx 파일 자체가 손상된 경우, 압축은 풀려도 그 안의 xml이 0바이트이거나
    깨져있는 경우가 있다.)"""
    try:
        return etree.parse(path)
    except Exception as e:
        name = os.path.basename(path)
        if not os.path.exists(path):
            reason = "파일이 존재하지 않습니다"
        elif os.path.getsize(path) == 0:
            reason = "파일이 비어 있습니다(0바이트)"
        else:
            reason = f"XML 구조가 손상되어 읽을 수 없습니다 ({os.path.getsize(path)}바이트)"
        raise RuntimeError(
            f"'{name}' {reason}. 원본 hwpx 파일 자체가 손상됐을 가능성이 높습니다. "
            f"한글에서 이 hwpx 파일을 열어 '다른 이름으로 저장'을 한 번 한 뒤, "
            f"그렇게 새로 저장된 파일로 다시 시도해보세요. (원본 오류: {e})"
        ) from e


def remove_trailing_matter(section_paths):
    """'집필위원'이 나오는 section 파일을 찾아 그 문단부터 끝까지, 그리고 그 뒤에 오는
    section 파일 전체를 삭제 대상으로 표시한다. (뒤에 오는 section 파일은 흔치 않음)
    '집필위원' 문구를 못 찾으면, 마지막 section 파일 끝에 내용 없는 빈 페이지만
    남아있는지 확인해서 있으면 그것만이라도 제거한다."""
    dbg("remove_trailing_matter: section_paths =", section_paths)
    for idx, path in enumerate(section_paths):
        tree = safe_parse(path)
        root = tree.getroot()
        children = list(root)
        cut_at = None
        for i, p in enumerate(children):
            if '집필위원' in para_text(p):
                cut_at = i
                break
        dbg(f"  [{idx}] {os.path.basename(path)}: children={len(children)} cut_at={cut_at}")
        if cut_at is not None:
            for i in sorted(range(cut_at, len(children)), reverse=True):
                root.remove(children[i])
            # '집필위원' 바로 앞에 내용 없는 빈 페이지가 하나 더 끼어있는 경우가 있어서
            # (본문 마지막 페이지와 집필위원 페이지 사이의 구분용 빈 페이지) 그것도 마저 제거.
            trimmed = remove_trailing_blank_page(root)
            dbg(f"  -> '집필위원' 발견, {cut_at}부터 끝까지 삭제. 추가 빈 페이지 제거={trimmed}")
            tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)
            return idx, True  # 이 인덱스 뒤에 오는 section 파일들은 통째로 삭제 대상

    # '집필위원'을 못 찾은 경우: 마지막 section 파일에서 빈 트레일링 페이지만 제거 시도
    last_path = section_paths[-1]
    tree = safe_parse(last_path)
    root = tree.getroot()
    before = len(root)
    result = remove_trailing_blank_page(root)
    after = len(root)
    dbg(f"  fallback: last_path={last_path} before={before} after={after} result={result}")
    if result:
        tree.write(last_path, xml_declaration=True, encoding='UTF-8', standalone=True)
        dbg(f"  -> 빈 페이지 제거 후 {last_path} 저장 완료")
        return None, True
    return None, False


def process_one(src_path, dst_path, work_dir, opts=None):
    opts = opts if opts is not None else DEFAULT_OPTIONS
    extract_dir = os.path.join(work_dir, 'extract')
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)

    with zipfile.ZipFile(src_path) as z:
        order = [(info.filename, info.compress_type) for info in z.infolist()]
        z.extractall(extract_dir)

    header_path = os.path.join(extract_dir, 'Contents', 'header.xml')

    section_files = sorted(
        glob.glob(os.path.join(extract_dir, 'Contents', 'section*.xml')),
        key=lambda p: int(re.search(r'section(\d+)\.xml', p).group(1))
    )
    if not section_files:
        raise RuntimeError("Contents/section*.xml 을 찾지 못함")

    warnings = []

    # 1) 앞부분(표지+경과조치+목차)은 항상 첫 section 파일에 있다고 가정
    first_tree = safe_parse(section_files[0])
    first_root = first_tree.getroot()
    clean_cover(first_root, header_path, opts)
    if opts.get('frontmatter_gap', True):
        ok = remove_frontmatter_gap(first_root)
        if not ok:
            warnings.append("경과조치/연혁 구간을 못 찾아서 표지~목차 사이는 그대로 두었습니다.")
    first_tree.write(section_files[0], xml_declaration=True, encoding='UTF-8', standalone=True)

    # 2) masterpage(바탕쪽/배경쪽)의 워터마크 이미지, 색/이미지 배경, 제정·개정 이력
    #    정보란 표 제거. 워터마크가 <hp:pic> 그림 개체로 박혀있는 경우도 있고, 표(hp:tbl)
    #    셀에 색/이미지가 채워진 경우도 있고, 순수 텍스트로 된 정보란 표가 그대로 박혀있는
    #    경우도 있어서 표지와 동일한 방식으로 세 가지 다 처리한다.
    remove_footer = opts.get('remove_footer', True)
    remove_page_number = opts.get('remove_page_number', True)

    if opts.get('cover_image', True) or opts.get('cover_color', True) or \
            opts.get('cover_infobox', True) or remove_footer or remove_page_number:
        for mp_path in glob.glob(os.path.join(extract_dir, 'Contents', 'masterpage*.xml')):
            mtree = safe_parse(mp_path)
            mroot = mtree.getroot()
            changed = False
            if opts.get('cover_image', True):
                changed = remove_pics(mroot) > 0 or changed
            if opts.get('cover_color', True):
                changed = remap_colored_fills(mroot, header_path) or changed
            if opts.get('cover_infobox', True):
                changed = remove_cover_info_table(mroot) > 0 or changed
            if remove_footer:
                changed = remove_footer_ctrl(mroot) > 0 or changed
            if remove_page_number:
                changed = remove_page_number_fields(mroot) > 0 or changed
            if changed:
                mtree.write(mp_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    # 3) 꼬리말/쪽번호 삭제. 머리말/꼬리말은 마스터페이지가 아니라 실제로 쓰이는
    #    section 파일 쪽에 통제 개체로 박혀있는 경우가 많아서 모든 section 파일을 훑는다.
    if remove_footer or remove_page_number:
        for sp in section_files:
            stree = safe_parse(sp)
            sroot = stree.getroot()
            changed = False
            if remove_footer:
                changed = remove_footer_ctrl(sroot) > 0 or changed
            if remove_page_number:
                changed = remove_page_number_fields(sroot) > 0 or changed
            if changed:
                stree.write(sp, xml_declaration=True, encoding='UTF-8', standalone=True)

    # 4) 집필위원 이후 삭제 (없으면 마지막 빈 페이지만이라도 제거)
    extra_removed_files = []
    if opts.get('trailing_matter', True):
        cut_section_idx, handled = remove_trailing_matter(section_files)
        if cut_section_idx is None:
            if handled:
                warnings.append("'집필위원' 문구는 없었지만, 끝에 남아있던 빈 페이지를 제거했습니다.")
            else:
                warnings.append("'집필위원' 문구를 못 찾았고, 끝에 빈 페이지도 없어서 마지막 부분은 그대로 두었습니다.")
        else:
            # 집필위원이 발견된 section 파일 뒤에 오는 section 파일은 전부 삭제
            for path in section_files[cut_section_idx + 1:]:
                name = os.path.relpath(path, extract_dir).replace(os.sep, '/')
                os.remove(path)
                extra_removed_files.append(name)

    # content.hpf / manifest.xml에서 삭제된 section 파일 참조 제거
    if extra_removed_files:
        for meta_name in ('Contents/content.hpf', 'META-INF/manifest.xml'):
            meta_path = os.path.join(extract_dir, meta_name)
            if not os.path.exists(meta_path):
                continue
            data = open(meta_path, encoding='utf-8').read()
            for fn in extra_removed_files:
                data = re.sub(r'<[^<>]*' + re.escape(fn) + r'[^<>]*/>\s*', '', data)
            open(meta_path, 'w', encoding='utf-8').write(data)

    dbg(f"final section1 children just before repack: "
        f"{len(safe_parse(section_files[-1]).getroot())}")

    # 5) 재압축 (원본 순서/압축방식 유지, 삭제된 파일은 건너뜀)
    with zipfile.ZipFile(dst_path, 'w') as zout:
        for name, ctype in order:
            if name in extra_removed_files:
                continue
            fpath = os.path.join(extract_dir, name)
            if not os.path.exists(fpath):
                continue
            with open(fpath, 'rb') as f:
                data = f.read()
            zi = zipfile.ZipInfo(name)
            zi.compress_type = ctype
            zout.writestr(zi, data)

    dbg(f"wrote dst_path={dst_path!r} exists={os.path.exists(dst_path)} "
        f"size={os.path.getsize(dst_path) if os.path.exists(dst_path) else -1} "
        f"mtime={os.path.getmtime(dst_path) if os.path.exists(dst_path) else -1}")

    return warnings


def run(paths, recursive=False, opts=None, output_dir=None):
    """GUI 등 다른 스크립트에서 그대로 불러 쓰기 위한 진입점.
    paths: 폴더 경로 또는 .hwpx 파일 경로의 리스트 (섞여 있어도 됨).
    opts: DEFAULT_OPTIONS와 같은 키를 가진 dict. None이면 전체 기본값(전부 켬) 사용.
    output_dir: 결과("..._정리됨.hwpx") 파일을 저장할 폴더. None이면 기존처럼 원본과
    같은 폴더에 저장한다. 지정하면 재실행할 때마다 그 폴더로 바로 덮어써지므로,
    오류를 고치고 다시 돌릴 때 결과 파일을 수동으로 옮길 필요가 없다.
    (--recursive로 여러 하위 폴더를 함께 처리하는 경우, 이름이 겹치지 않도록
    output_dir 밑에 원본 폴더 기준 상대 경로를 그대로 재현한다.)"""
    opts = opts if opts is not None else DEFAULT_OPTIONS
    pattern = "**/*.hwpx" if recursive else "*.hwpx"
    file_entries = []  # (src_path, 상대경로 계산 기준이 되는 폴더)
    for path in paths:
        if os.path.isfile(path):
            if path.lower().endswith('.hwpx'):
                file_entries.append((path, os.path.dirname(path) or '.'))
        else:
            for f in glob.glob(os.path.join(path, pattern), recursive=recursive):
                file_entries.append((f, path))
    file_entries = [(f, base) for f, base in file_entries if not f.endswith('_정리됨.hwpx')]

    if not file_entries:
        print("hwpx 파일을 찾지 못했습니다.")
        return

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    print(f"총 {len(file_entries)}개 파일 처리 시작.\n")

    dbg("files found:", [f for f, _ in file_entries])

    # 시스템 임시 폴더(%TEMP%)가 일부 PC에서는 압축/보안 유틸리티가 관리하는
    # 폴더(예: ...\ESTsoft\CreatorTemp)로 리디렉션되어 있어서, 그 안에 막 풀어놓은
    # section*.xml이 처리 도중 사라지는 경우가 있었다(그 유틸리티의 자동 임시파일
    # 정리 기능으로 추정). 이를 피하기 위해 시스템 임시 폴더 대신, 실제 처리 대상
    # 파일과 같은 폴더(또는 지정된 저장 폴더) 밑에 작업용 임시 폴더를 만든다.
    tmp_base = output_dir or os.path.dirname(os.path.abspath(file_entries[0][0])) or '.'
    try:
        tmpdir_ctx = tempfile.TemporaryDirectory(dir=tmp_base)
    except OSError:
        tmpdir_ctx = tempfile.TemporaryDirectory()

    ok_count = 0
    fail_count = 0
    with tmpdir_ctx as work_dir:
        dbg("work_dir =", work_dir)
        for src, base_dir in file_entries:
            name = os.path.splitext(os.path.basename(src))[0] + '_정리됨.hwpx'
            if output_dir:
                rel_dir = os.path.relpath(os.path.dirname(src) or '.', base_dir)
                out_dir = output_dir if rel_dir in ('.', '') else os.path.join(output_dir, rel_dir)
                os.makedirs(out_dir, exist_ok=True)
                dst = os.path.join(out_dir, name)
            else:
                dst = os.path.join(os.path.dirname(src), name)
            dbg(f"processing src={src!r} dst={dst!r} src_mtime_before={os.path.getmtime(src)}")
            try:
                warnings = process_one(src, dst, work_dir, opts)
                ok_count += 1
                status = "완료" if not warnings else "완료(확인 필요)"
                print(f"[{status}] {os.path.basename(src)} -> {dst}")
                for w in warnings:
                    print(f"         ! {w}")
            except Exception as e:
                fail_count += 1
                print(f"[실패] {os.path.basename(src)} -> {e}")

    print(f"\n총 {len(file_entries)}개 중 성공 {ok_count}개, 실패 {fail_count}개")


def main():
    parser = argparse.ArgumentParser(description="KCS hwpx 문서 일괄 정리 (표지 정리 + 경과조치/집필위원 페이지 삭제)")
    parser.add_argument("folders", nargs="+", help="처리할 .hwpx 파일들이 있는 폴더 (여러 개 가능)")
    parser.add_argument("--recursive", action="store_true", help="하위 폴더까지 포함")
    parser.add_argument("--output", "-o", metavar="폴더",
                         help="결과(_정리됨.hwpx) 파일을 저장할 폴더. 생략하면 원본과 같은 폴더에 저장")
    args = parser.parse_args()
    run(args.folders, args.recursive, output_dir=args.output)


if __name__ == "__main__":
    main()
