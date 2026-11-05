"""查写作词语、使用频次和例篇，也可按主题找到讲话原文。"""
import argparse
import json
import sys
from collections import Counter
from pathlib import Path

sys.dont_write_bytecode = True
from import_docx import BUNDLED_WORD
from read_original import load as read_full_book

REFERENCES = Path(__file__).resolve().parents[1] / 'references'
LIBRARY = REFERENCES / 'xi-phrases.jsonl'
LOCAL_INDEX = REFERENCES / 'local-phrasebook.jsonl'
SOURCES = {'primary': '本人原文样本（1,201份）', 'reports': '报道直接引语样本（1,579篇报道）'}
KEYWORDS = {
    '调查研究': '调研 立意 找问题 重点 方法',
    '规划部署': '工作计划 安排任务 分工 谋划 路径',
    '任务落实': '执行 整改 督办 通知 责任',
    '协作分工': '合作 团结 协调 配合 共同',
    '学习交流': '培训 读书 教育 学用 实践',
    '改革创新': '突破 转型 探索 动力',
    '变化与成效': '总结 汇报 进展 变化 同尾 小标题',
    '制度监督': '风险 防控 规矩 纪律 廉洁 治理',
    '接续奋斗': '长期 坚持 接力 攻坚 激励',
}


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('keywords', nargs='*', help='词形、用途或主题；多个词须同时匹配')
    parser.add_argument('--preferred', action='store_true', help='只查通用写作精选')
    parser.add_argument('--length', type=int, choices=(3, 4, 5, 6), help='限定短语的汉字数')
    parser.add_argument('--source', choices=SOURCES, default='primary', help='本人原文或报道直接引语')
    parser.add_argument('--sentences', action='store_true', help='查看按主题推荐的原文出处及古诗文用例')
    parser.add_argument('--usage', action='store_true', help='列出当前结果的用途')
    parser.add_argument('--book', action='store_true', help='查找也见于随附全书的词句；原书优先，个人资料补充')
    parser.add_argument('--book-index', type=Path, help='配合 --book 只查指定的 JSONL 索引')
    parser.add_argument('--limit', type=int, default=8, help='展示1至100条，默认8条')
    args = parser.parse_args()
    if not 1 <= args.limit <= 100:
        parser.error('--limit 须在1至100之间。')
    if args.sentences and (args.length or args.preferred or args.source == 'reports'):
        parser.error('--sentences 使用本人原文出处，不能与 --length、--preferred 或 --source reports 合用。')
    if not args.book and args.book_index is not None:
        parser.error('--book-index 应与 --book 合用。')
    return args


def read_records(path):
    with path.open(encoding='utf-8-sig') as file:
        for line in file:
            if line.strip():
                yield json.loads(line)


def filter_material(matches, custom_index):
    if custom_index is not None:
        materials = [('指定索引', custom_index)]
    else:
        materials = [(label, path) for label, path in (
            ('随附原书', BUNDLED_WORD), ('本地导入资料', LOCAL_INDEX)
        ) if path.is_file()]
    if not materials:
        raise ValueError('缺少个人资料及随附原书。可用 scripts/import_docx.py 导入资料，或使用 --book-index。')
    for label, path in materials:
        pending = [row for _, row in matches if row.get('text') and 'book_origin' not in row]
        if not pending and custom_index is None:
            break
        # Search the whole book, including headings and table paragraphs.
        if path == BUNDLED_WORD:
            paragraphs, _ = read_full_book()
            text = '\n'.join(paragraphs)
        else:
            text = '\n'.join(record['text'] for record in read_records(path))
        for row in pending:
            if row['text'] in text:
                row['book_origin'] = label
    return [(score, row) for score, row in matches if 'book_origin' in row]


def display(row, source):
    kind = row['kind']
    name = row.get('text') or row['topic']
    label = {'reference': '原文线索，标题为主题提示', 'classic': '讲话中引用的古诗文', 'phrase': '短语'}[kind]
    tags = [label, row['usage']]
    if row.get('featured'):
        tags.append('写作精选')
    if row.get('book_origin'):
        tags.append('见于' + row['book_origin'])
    print(f"{name}  [{' / '.join(tags)}]")
    ref = row[source]
    if kind == 'phrase':
        print(f"  {SOURCES[source]}：出现{ref['count']}次，覆盖{ref['documents']}篇。")
    print(f"  {ref['title']}（资料日期 {ref['date']}） {ref['url']}")
    if kind == 'classic':
        print(f"  典出 {row['origin']}  {row['origin_url']}")


def main():
    args = arguments()
    default_browse = not (args.keywords or args.length or args.book or args.usage or args.sentences) and args.source == 'primary'
    matches = []
    for row in read_records(LIBRARY):
        is_sentence_resource = row['kind'] in ('classic', 'reference')
        if args.source not in row or is_sentence_resource != args.sentences:
            continue
        if args.length and row['length'] != args.length:
            continue
        if (args.preferred or default_browse) and not row.get('featured'):
            continue
        term = row.get('text', '')
        subject = ' '.join((row['usage'], KEYWORDS.get(row['usage'], ''), row.get('topic', ''), row.get('origin', '')))
        if not all(k in term or k in subject for k in args.keywords):
            continue
        relevance = sum(3 if k == term else 2 if k in term else 1 for k in args.keywords)
        matches.append((relevance, row))
    if args.book:
        matches = filter_material(matches, args.book_index)
    if not matches:
        print('没有匹配项。可以缩短关键词，或用 --usage 查看用途分类。')
        return 1
    if args.usage:
        counts = Counter(row['usage'] for _, row in matches)
        for label, count in sorted(counts.items()):
            print(f'{label}  {count}项')
        return 0
    matches.sort(key=lambda item: (item[0], bool(item[1].get('featured')), item[1][args.source].get('documents', 0)), reverse=True)
    if default_browse or args.preferred and not args.keywords:
        by_group, rest = {}, []
        for item in matches:
            usage = item[1]['usage']
            if usage in by_group:
                rest.append(item)
            else:
                by_group[usage] = item
        matches = list(by_group.values()) + rest
    for _, row in matches[:args.limit]:
        display(row, args.source)
    return 0


if __name__ == '__main__':
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    try:
        sys.exit(main())
    except (OSError, ValueError, KeyError, TypeError) as error:
        print(f'读取资料失败：{error}', file=sys.stderr)
        sys.exit(2)
