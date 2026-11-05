"""把自己有权使用的 Word 资料转换为本机的公文词句索引。"""

import argparse
from collections import Counter
import json
import os
from pathlib import Path
import re
import sys
import tempfile
import unicodedata
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


DEFAULT_OUTPUT = Path(__file__).resolve().parents[1] / 'references' / 'local-phrasebook.jsonl'
BUNDLED_WORD = DEFAULT_OUTPUT.with_name('公文写作词句应用大全（500页25万字）.docx')
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
LABEL = re.compile(r'^(?:第[一二三四五六七八九十百〇零两\d]+[章节篇部分]|[一二三四五六七八九十百]+[、．.]|[（(][一二三四五六七八九十百]+[)）])')
PAGE = re.compile(r'[\s.．…·]*(\d+)\s*$')


def key(value):
    return ''.join(c for c in unicodedata.normalize('NFKC', value) if c.isalnum())


def val(node, path, default=''):
    child = node.find('/'.join(W + part for part in path.split('/')))
    return default if child is None else child.get(W + 'val', default)


def xml_part(archive, name, optional=False):
    try:
        meta = archive.getinfo('word/' + name)
    except KeyError:
        if optional:
            return ET.Element('empty')
        raise ValueError('找不到 Word 正文，请使用未加密的 .docx 文件。') from None
    if meta.file_size > 64 * 1024 * 1024:
        raise ValueError(f'Word 部件 {name} 超过 64 MiB；请分册导入。')
    return ET.fromstring(archive.read(meta))


def paragraph_text(paragraph):
    parts = []
    def visit(node):
        if node.tag in (W + 'del', W + 'moveFrom', W + 'drawing', W + 'pict', W + 'txbxContent'):
            return
        if node.tag == W + 't':
            parts.append(node.text or '')
        elif node.tag == W + 'tab':
            parts.append('\t')
        elif node.tag in (W + 'br', W + 'cr'):
            parts.append('\n')
        else:
            for child in node:
                visit(child)
    visit(paragraph)
    return ''.join(parts).strip()


def style_info(styles, style):
    """Follow basedOn, including localized or custom paragraph styles."""
    level = None
    toc = False
    seen = set()
    while style in styles and style not in seen:
        seen.add(style)
        node = styles[style]
        name = val(node, 'name', style)
        toc |= bool(re.fullmatch(r'(?:toc|目录)\s*\d+', name, re.I))
        outline = val(node, 'pPr/outlineLvl')
        heading = re.fullmatch(r'(?:heading|标题)\s*([1-9])', name, re.I)
        if level is None:
            if outline.isdigit():
                level = int(outline) + 1
            elif heading:
                level = int(heading[1])
        style = val(node, 'basedOn')
    return level, toc


def list_starts(numbering):
    abstract = {x.get(W + 'abstractNumId'): x for x in numbering.findall(W + 'abstractNum')}
    info = {}
    for num in numbering.findall(W + 'num'):
        num_id = num.get(W + 'numId')
        definition = abstract.get(val(num, 'abstractNumId'))
        if definition is None:
            continue
        for lvl in definition.findall(W + 'lvl'):
            level = lvl.get(W + 'ilvl', '0')
            if val(lvl, 'numFmt') not in ('bullet', 'none'):
                info[num_id, level] = int(val(lvl, 'start', '1'))
        for override in num.findall(W + 'lvlOverride'):
            start = val(override, 'startOverride')
            level = override.get(W + 'ilvl', '0')
            if start.isdigit() and (num_id, level) in info:
                info[num_id, level] = int(start)
    return info


def get_paragraphs(path):
    with ZipFile(path) as archive:
        body = xml_part(archive, 'document.xml').find(W + 'body')
        if body is None:
            raise ValueError('找不到 Word 正文。')
        styles = {s.get(W + 'styleId'): s for s in xml_part(archive, 'styles.xml', True).findall(W + 'style')}
        starts = list_starts(xml_part(archive, 'numbering.xml', True))
    numbering = {}
    paragraphs = []
    for ix,p in enumerate(body.iter(W + 'p'), 1):
        text = paragraph_text(p)
        if not text:
            continue
        level, toc = style_info(styles, val(p, 'pPr/pStyle'))
        explicit = val(p, 'pPr/outlineLvl')
        if explicit.isdigit():
            level = int(explicit) + 1
        num = (val(p, 'pPr/numPr/numId'), val(p, 'pPr/numPr/ilvl', '0'))
        example = None
        if num in starts:
            example = numbering.get(num, starts[num] - 1) + 1
            numbering[num] = example
            for other in list(numbering):
                if other[0] == num[0] and int(other[1]) > int(num[1]):
                    del numbering[other]
        paragraphs.append((ix,text,level if level and level <= 9 else None,toc,example))
    return paragraphs


def find_catalog(rows):
    marker = next((i for i,r in enumerate(rows) if key(r[1]) in ('目录','目次','Contents','TableofContents')), None)
    if marker is None:
        return [], rows
    start = marker + 1
    stop = start
    for _, text, _, is_toc, _ in rows[start:]:
        if is_toc or PAGE.search(text) or (len(text) <= 40 and text.endswith('篇')):
            stop += 1
        else:
            break
    # Some Word TOCs mark only their later lines with TOC paragraph styles.
    if stop == start or not any(r[3] or PAGE.search(r[1]) for r in rows[start:stop]):
        return [], rows
    return rows[start:stop], rows[stop:]


def index_document(path, id_prefix='L'):
    catalog, rows = find_catalog(get_paragraphs(path))
    section_pages = {}
    chapters = []
    for _, text, _, _, _ in catalog:
        pg = PAGE.search(text)
        title = text[:pg.start()].strip() if pg else text
        if title.endswith('篇') and len(title) <= 40 and not LABEL.match(title):
            chapters.append(title)
        if pg:
            section_pages[key(title)] = (title,int(pg[1]))
    series = [title for title,_ in section_pages.values() if title.endswith('系列')]
    split_series = re.compile('('+'|'.join(re.escape(x) for x in sorted(series,key=len,reverse=True))+')[：:]') if series else None
    chapter = path.stem
    section = chapter
    page = None
    next_chapter = 0
    output = []

    def append(text, ix, example=None):
        if not text.strip():
            return
        row = {'id':f'{id_prefix}{len(output)+1:05d}', 'chapter':chapter, 'section':section,
               'text':text.strip(), 'paragraph':ix}
        if page is not None:
            row['page'] = page
        if example is not None:
            row['example'] = example
        output.append(row)

    for ix,text,level,is_toc,example in rows:
        if is_toc:
            continue
        if next_chapter < len(chapters):
            expected = key(chapters[next_chapter])
            label = key(text)
            # Body names sometimes omit a modifier or the final 篇 printed in the TOC.
            chars = iter(expected)
            shortened = len(label)>=4 and text.endswith('篇') and all(c in chars for c in label)
            if label == expected or label+'篇' == expected or shortened:
                chapter = section = chapters[next_chapter]
                page = None
                next_chapter += 1
                continue
        markers = list(split_series.finditer(text)) if split_series else []
        if markers:
            prefix = text[:markers[0].start()]
            if prefix and not (level and LABEL.match(prefix)):
                append(prefix,ix,example)
            for i,match in enumerate(markers):
                section,page = section_pages[key(match[1])]
                end = markers[i+1].start() if i+1<len(markers) else len(text)
                append(text[match.end():end],ix)
            continue
        toc_title = section_pages.get(key(text.rstrip('：:')))
        if not chapters and (level == 1 or re.fullmatch(r'第[一二三四五六七八九十百〇零两\d]+[章篇].{0,45}',text)):
            chapter = section = text
            page = toc_title[1] if toc_title else None
            continue
        if toc_title or (level and len(text)<=90 and example is None):
            section,page = toc_title if toc_title else (text,None)
            continue
        if example is None and len(text)<=46 and LABEL.match(text) and not re.search(r'[。；！？]',text):
            section,page = text,None
            continue
        append(text,ix,example)
    if not output:
        raise ValueError('没有可索引的正文段落。扫描图片需要先做文字识别。')
    return output


def read_bundled_word():
    """Read the distributed document in memory without creating a local index."""
    try:
        return index_document(BUNDLED_WORD, id_prefix='B')
    except (BadZipFile, ET.ParseError, RuntimeError) as exc:
        raise ValueError(f'无法解析随附原书 {BUNDLED_WORD.name}：{exc}') from exc


def save(rows, target, replace):
    target.parent.mkdir(parents=True,exist_ok=True)
    tmp = None
    try:
        with tempfile.NamedTemporaryFile('w', encoding='utf-8', dir=target.parent, delete=False) as stream:
            tmp = Path(stream.name)
            for row in rows:
                stream.write(json.dumps(row,ensure_ascii=False)+'\n')
        if replace:
            os.replace(tmp,target)
        elif os.name == 'nt':
            os.rename(tmp,target)  # Windows rename fails if the target already exists.
        else:
            os.link(tmp,target)  # Atomic no-clobber publication on Unix.
    finally:
        if tmp is not None:
            tmp.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx',type=Path,help='自己的 .docx 材料')
    parser.add_argument('--output',type=Path,default=DEFAULT_OUTPUT,help='本地 JSONL 索引路径')
    parser.add_argument('--replace',action='store_true',help='允许替换已有索引')
    args = parser.parse_args()
    if args.docx.suffix.lower() != '.docx' or args.output.suffix.lower() != '.jsonl':
        parser.error('输入必须是 .docx 文件；输出必须是 .jsonl 文件。')
    try:
        rows = index_document(args.docx)
        save(rows,args.output,args.replace)
    except FileExistsError:
        print('目标索引已存在。需要替换时添加 --replace。',file=sys.stderr)
        return 1
    except (OSError,BadZipFile,ET.ParseError,ValueError,RuntimeError) as exc:
        print(f'导入失败：{exc}',file=sys.stderr)
        return 1
    print(f'已导入 {len(rows)} 条，涉及 {len(Counter(r["chapter"] for r in rows))} 个篇目。')
    print(args.output)
    return 0


if __name__ == '__main__':
    sys.exit(main())
