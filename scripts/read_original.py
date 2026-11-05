"""逐段阅读随附 Word，包括篇章标题、普通短段及表格。W 段号从 1 起，空段也占一个号。"""

import argparse
from bisect import bisect_right
from hashlib import sha256
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


BOOK = Path(__file__).resolve().parents[1] / 'references' / '公文写作词句应用大全（500页25万字）.docx'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
# Verified one-based Word paragraph numbers: TOC chapter name, then start of chapter.
CHAPTERS = ((9,355),(32,1656),(37,1828),(95,2320),(117,2570),(143,4843),
            (151,5682),(191,6088),(225,6344),(231,6664),(301,9741),(312,9775),
            (339,10107),(342,10531),(349,10609))
BODY_SHA = 'e5971a7728c73bd91ecbe549c5e76324201c82d23f8e3f7f9e0b8219f50df448'
BREAKS = {'tab':'\t', 'br':'\n', 'cr':'\n', 'noBreakHyphen':'\u2011', 'softHyphen':'\u00ad'}


def paragraph_text(paragraph):
    # Read only run contents; paragraph-format tabs are not characters on the page.
    parts = []
    def walk(node):
        for child in node:
            if child.tag in (W+'p', W+'pPr', W+'del', W+'moveFrom'):
                continue
            if child.tag == W+'r':
                for e in child:
                    if e.tag == W+'t':
                        parts.append(e.text or '')
                    elif e.tag.startswith(W) and e.tag[len(W):] in BREAKS:
                        parts.append(BREAKS[e.tag[len(W):]])
            else:
                walk(child)
    walk(paragraph)
    return ''.join(parts)


def load():
    with ZipFile(BOOK) as archive:
        info = archive.getinfo('word/document.xml')
        if info.file_size > 64*1024*1024:
            raise ValueError('Word 正文部件超过 64 MiB。')
        xml = archive.read(info)
    body = ET.fromstring(xml).find(W+'body')
    if body is None:
        raise ValueError('文件没有 Word 正文。')
    paragraphs = [paragraph_text(p) for p in body.iter(W+'p')]
    if not paragraphs:
        raise ValueError('文件没有正文段落。')
    if sha256(xml).hexdigest() == BODY_SHA:
        spans = [(1,'书名与目录')] + [(start,paragraphs[toc-1].strip()) for toc,start in CHAPTERS]
    else:
        spans = [(1,'全文（此版本篇界未经校准）')]
    return paragraphs, spans


def number(value):
    digits = value.upper().removeprefix('W')
    if not digits.isascii() or not digits.isdigit() or int(digits)<1:
        raise argparse.ArgumentTypeError('段号必须为正整数或 W00001 形式。')
    return int(digits)


def parse_range(value):
    try:
        first,last = value.split(':')
        start,end = number(first),number(last)
    except (ValueError, argparse.ArgumentTypeError) as exc:
        raise argparse.ArgumentTypeError('范围示例：W00355:W00362。') from exc
    if start>end:
        raise argparse.ArgumentTypeError('起始段号不能大于结束段号。')
    return start,end


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('query',nargs='?',help='在单个原段内作逐字查询，保留字词、空格与标点')
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument('--list-chapters', action='store_true',help='列出原书各篇段号范围')
    mode.add_argument('--paragraph',type=number,help='读取指定 Word 段，即使是标题或空段')
    mode.add_argument('--range',type=parse_range,help='完整输出首末段号之间的所有原段')
    ap.add_argument('--chapter',help='按篇名或篇名片段浏览或限定查询；单独使用可阅读整篇')
    ap.add_argument('--context',type=int,default=0,help='查询或段号周围附带的非空段数')
    ap.add_argument('--limit',type=int,help='查询最多显示多少个命中段，默认 8')
    a = ap.parse_args()
    lookup = a.query is not None or a.paragraph is not None
    if not (lookup or a.range or a.list_chapters or a.chapter):
        ap.error('请填查询短语、段号、段号范围或篇名。')
    if (a.list_chapters or a.range or a.paragraph) and (a.query is not None or a.chapter is not None):
        ap.error('列篇、按段号及按范围阅读时，请勿同时指定短语或篇名。')
    if a.query == '' or (a.chapter is not None and not a.chapter.strip()):
        ap.error('短语或篇名不能为空。')
    if a.context < 0 or (a.context and not lookup):
        ap.error('--context 仅配合短语或段号查询，且必须为非负整数。')
    if a.limit is not None and (a.limit<1 or a.query is None):
        ap.error('--limit 只用于短语查询，且必须为正整数。')
    paragraphs, spans = load()
    n = len(paragraphs)
    starts = [s for s,_ in spans]
    if a.list_chapters:
        for i,(s,label) in enumerate(spans):
            end = starts[i+1]-1 if i+1<len(starts) else n
            print(f'W{s:05d}:W{end:05d}\t{label}')
        return 0
    bounds = a.range or ((a.paragraph,a.paragraph) if a.paragraph else (1,n))
    if bounds[1]>n:
        ap.error(f'本书段号范围为 W00001 至 W{n:05d}。')
    chapter_at = lambda i: spans[bisect_right(starts,i)-1][1]
    hits = [i for i in range(bounds[0],bounds[1]+1)
            if (not a.chapter or a.chapter.strip() in chapter_at(i))
            and (a.query is None or a.query in paragraphs[i-1])]
    if not hits:
        print('未找到匹配项，可缩短短语或用 --list-chapters 查看篇名。')
        return 1
    if a.query is not None:
        limit = a.limit or 8
        if len(hits)>limit:
            print(f'共命中 {len(hits)} 段，显示前 {limit} 段。可用 --limit 增加条数。')
        hits = hits[:limit]
    shown = set(hits)
    if a.context:
        for hit in hits:
            for step in (-1,1):
                count = 0
                i = hit+step
                while 1<=i<=n and count<a.context:
                    if a.chapter and a.chapter.strip() not in chapter_at(i):
                        break
                    if paragraphs[i-1].strip():
                        shown.add(i)
                        count += 1
                    i += step
    for i in sorted(shown):
        print(f'[W{i:05d}] {chapter_at(i)}')
        print(paragraphs[i-1])
    return 0


if __name__=='__main__':
    for stream in (sys.stdout,sys.stderr):
        if hasattr(stream,'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    try:
        sys.exit(main())
    except (OSError, ValueError, KeyError, BadZipFile, ET.ParseError) as exc:
        print(f'无法读取随附原书：{exc}', file=sys.stderr)
        sys.exit(2)
