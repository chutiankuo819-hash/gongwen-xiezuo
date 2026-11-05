"""按词语、段落用途或篇目检索个人资料、随附原书和写作示例。"""

import argparse
import json
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path

sys.dont_write_bytecode = True
from import_docx import BUNDLED_WORD, read_bundled_word

REFERENCES = Path(__file__).resolve().parents[1] / 'references'
LIBRARIES = {
    'bundled': ('随附原书', BUNDLED_WORD),
    'local': ('本地导入', REFERENCES / 'local-phrasebook.jsonl'),
    'examples': ('内置示例', REFERENCES / 'phrasebook.jsonl'),
}


def read_book(source):
    records = []
    label, path = LIBRARIES[source]
    if source == 'bundled':
        for entry in read_bundled_word():
            entry['origin'] = label
            records.append(entry)
        return records
    with path.open(encoding='utf-8-sig') as stream:
        for number, line in enumerate(stream, 1):
            if not line.strip():
                continue
            entry = json.loads(line)
            if not isinstance(entry, dict) or not all(
                key in entry for key in ('id', 'chapter', 'section', 'text')
            ) or not all(isinstance(entry[key], str) for key in ('chapter', 'section', 'text')):
                raise ValueError(f'第 {number} 行缺少有效的编号、篇目、小节或原文。')
            entry['origin'] = label
            records.append(entry)
    if not records:
        raise ValueError('词句库为空。')
    return records


def rank(entry, terms):
    chapter, section, body = (entry[k].casefold() for k in ('chapter', 'section', 'text'))
    exact, near, weight = 0, 0, 0
    anchors = []
    for term in terms:
        position = body.find(term)
        in_section, in_chapter = term in section, term in chapter
        if position >= 0 or in_section or in_chapter:
            exact += 1
            weight += 5 * in_section + 2 * in_chapter + 7 * (position >= 0)
            if position >= 0:
                anchors.append(position)
        elif len(term) >= 4 and re.fullmatch(r'[\u3400-\u9fff]+', term):
            # A reordered phrase such as 落实责任 can match both parts of 责任落实.
            pairs = {term[i:i + 2] for i in range(len(term) - 1)}
            found = {pair for pair in pairs if pair in body or pair in section}
            if len(found) >= max(2, (len(pairs) + 1) // 2):
                near += 1
                anchors.extend(body.find(pair) for pair in found if pair in body)
    return (exact + near, exact, weight), anchors


def show(entry, body):
    place = f"[{entry['id']}] {entry['origin']} · {entry['chapter']} / {entry['section']}"
    if entry.get('example') is not None:
        place += f" / 第{entry['example']}例"
    if entry.get('page'):
        place += f"（目录节起页 {entry['page']}）"
    print(place)
    print(body)
    print()


def excerpt(value, anchors, width=200):
    if len(value) <= width:
        return value
    start = max(0, min(min(anchors, default=0) - 40, len(value) - width))
    end = start + width
    return ('…' if start else '') + value[start:end] + ('…' if end < len(value) else '')


def title_key(value):
    """Allow the book's printed spacing, e.g. '41 例' and '41例', to match."""
    return ''.join(unicodedata.normalize('NFKC', value).casefold().split())


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('keywords', nargs='*', help='主题词或段落用途，多个词用空格分隔')
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--list-chapters', action='store_true', help='列出篇目及条目数')
    mode.add_argument('--id', help='按编号逐字显示整条原文')
    parser.add_argument('--chapter', help='限定篇目，支持篇名中的一段连续文字')
    parser.add_argument('--section', help='限定小节，支持节名片段；无关键词时按原序展示整节原文')
    parser.add_argument('--source', choices=('auto', *LIBRARIES, 'all'), default='auto',
                        help='auto 按随附原书、个人资料、内置示例依次检索；all 同时查全部可用资料')
    parser.add_argument('--limit', type=int, help='最多返回几条；默认查词 8 条、单独浏览小节时返回整节')
    args = parser.parse_args(argv)
    terms = list(dict.fromkeys(' '.join(args.keywords).casefold().split()))
    if args.limit is not None and args.limit < 1:
        parser.error('--limit 必须是正整数。')
    for label, value in (('--chapter', args.chapter), ('--section', args.section)):
        if value is not None and not title_key(value):
            parser.error(f'{label} 需要填写非空的标题或标题片段。')
    if (args.id is not None or args.list_chapters) and (terms or args.chapter or args.section):
        parser.error('列篇目或按编号查看时，请勿同时指定关键词、篇名或节名。')
    if not (terms or args.id is not None or args.list_chapters or args.chapter or args.section):
        parser.error('请输入检索词；可先用 --list-chapters 查看篇目。')
    if args.source in ('auto', 'all'):
        sources = [k for k,(_,p) in LIBRARIES.items() if p.exists()]
    else:
        sources = [args.source]
    if args.source == 'local' and not LIBRARIES['local'][1].is_file():
        print('尚无本地资料，请先运行 scripts/import_docx.py 导入 Word 文档。', file=sys.stderr)
        return 2
    chapter = title_key(args.chapter or '')
    section = title_key(args.section or '')
    results = []
    id_source = {'L':'local', 'B':'bundled', 'P':'examples'}.get((args.id or '')[:1])
    try:
        if not sources:
            raise ValueError('未找到资料，请检查 references 目录或导入 Word 文档。')
        for source in sources:
            if args.id is not None and id_source and id_source != source:
                continue
            entries = read_book(source)
            if args.list_chapters:
                for name, count in Counter(item['chapter'] for item in entries).items():
                    print(f'{LIBRARIES[source][0]} · {name}\t{count} 条')
                continue
            if args.id is not None:
                for item in entries:
                    if str(item['id']) == args.id:
                        show(item, item['text'])
                        return 0
                continue
            for item in entries:
                if chapter and chapter not in title_key(item['chapter']):
                    continue
                if section and section not in title_key(item['section']):
                    continue
                score, anchors = rank(item, terms)
                if score[0] or not terms:
                    results.append((score, item, anchors))
            if args.source == 'auto' and results:
                break
    except (OSError, ValueError) as error:
        print(f'无法读取词句库：{error}', file=sys.stderr)
        return 2
    if args.list_chapters:
        return 0
    if args.id is not None:
        print('未找到该编号。请先用主题词检索，再复制结果里的完整编号。')
        return 1

    if not results:
        print('没有匹配条目。请缩短关键词或节名，去掉篇目限制，或用 --list-chapters 查看篇名。')
        return 1

    if terms:
        results.sort(key=lambda match: match[0], reverse=True)
    browsing_section = bool(section and not terms)
    limit = args.limit if args.limit is not None or browsing_section else 8
    for score, item, anchors in results[:limit]:
        if terms and score[1] < len(terms):
            print(f'（完整命中 {score[1]}/{len(terms)} 个检索词）')
        show(item, item['text'] if browsing_section else excerpt(item['text'], anchors))
    return 0


if __name__ == '__main__':
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    sys.exit(main())
