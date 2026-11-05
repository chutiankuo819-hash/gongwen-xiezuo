"""从个人 Word 范文提取可核对的写作习惯，生成仅供本地使用的文风简表。"""

import argparse
from collections import Counter, defaultdict
from pathlib import Path
import os
import re
from statistics import median
import sys
import tempfile
from xml.etree import ElementTree
from zipfile import BadZipFile

from import_docx import find_catalog, get_paragraphs


OUTPUT = Path(__file__).resolve().parents[1] / 'references' / 'local-style.md'
GENRES = ('整改报告', '调研报告', '工作报告', '工作汇报', '工作总结', '工作方案',
          '实施方案', '请示', '报告', '通知', '纪要', '复函', '函', '讲话', '致辞', '发言')
NUMBER = r'[一二三四五六七八九十百〇零]+'
HEADING_MARKERS = (
    (1, re.compile(r'^' + NUMBER + r'[、．.]\s*')),
    (2, re.compile(r'^[（(]' + NUMBER + r'[）)]\s*')),
    (2, re.compile(r'^[1-9]\d?[.、．]\s*(?!\d)')),
    (2, re.compile(r'^' + NUMBER + r'是')),
)
TRANSITIONS = ('在此基础上', '从这个意义上讲', '为此', '与此同时', '在看到成绩的同时',
               '从实际情况看', '针对上述问题', '由此可见', '具体来说', '下一步', '下面',
               '总的来看', '一方面', '另一方面', '更重要的是')
PUNCT = '，。；：！？、'
SENTENCE = re.compile(r'[。！？!?]+')
CLAUSE = re.compile(r'[，,、；;：:。！？!?]+')
DATE = re.compile(r'^[（(]?[〇○零一二三四五六七八九十\d\s年月日./-]+[）)]?$')


def letters(text):
    return ''.join(c for c in text if c.isalnum())


def quote(text):
    return '“' + re.sub(r'\s+', ' ', text).strip()[:110] + '”'


def identify_genre(name, first):
    for text in (first, name):
        for genre in GENRES:
            if genre in text:
                return genre
    return '未标明'


def classify_heading(text, level):
    for inferred, pat in HEADING_MARKERS:
        marker = pat.match(text)
        if marker:
            title, sep, rest = text[marker.end():].partition('。')
            title = title.strip()
            if 0 < len(letters(title)) <= 48:
                return level or inferred, title, rest.strip() if sep else ''
            break
    if level and len(letters(text)) <= 60:
        return level, text.rstrip('。'), ''
    return None


def read_sample(path, supplied_genre):
    rows = get_paragraphs(path)
    _, body = find_catalog(rows)
    first = rows[0][1].splitlines()[0] if rows else ''
    genre = supplied_genre or identify_genre(path.stem, first)
    parents = {}
    groups = defaultdict(list)
    paras = []
    for ix, paragraph, level, toc, _ in body:
        if toc:
            continue
        for part_index, line in enumerate(paragraph.splitlines()):
            text = line.strip()
            if not text or DATE.fullmatch(text) or text in ('目录', '目次') or text.startswith(('附件：', '附件:')):
                continue
            if (ix == rows[0][0] and part_index == 0 and len(text) < 80 and not any(c in text for c in '。！？；')
                    and (text == path.stem or text.startswith('关于') or
                         (any(g in text for g in GENRES) and not any(p.match(text) for _, p in HEADING_MARKERS)))):
                continue
            heading = classify_heading(text, level if part_index == 0 else None)
            if heading:
                depth, title, rest = heading
                parents = {d: p for d, p in parents.items() if d < depth}
                parent = parents[max(parents)] if parents else 0
                groups[depth, parent].append((ix, title))
                parents[depth] = ix
                text = rest
            if len(letters(text)) >= 5 and not (len(text) < 26 and text.endswith(('：', ':'))):
                paras.append((ix, text))
    if not paras:
        raise ValueError(f'{path.name} 没有检测到可用于评估句段的正文。请检查是否为扫描件或仅有提纲。')
    return path.name, genre, groups, paras


def repeated_edges(values):
    findings = []
    for edge, label in [('start', '起句'), ('end', '收尾')]:
        counts = Counter()
        for value in values:
            for size in range(1, min(5, len(value)) + 1):
                counts[value[:size] if edge == 'start' else value[-size:]] += 1
        choices = [(s, n) for s, n in counts.items() if n >= 2]
        if choices:
            s, n = max(choices, key=lambda v: (v[1], len(v[0])))
            findings.append(f'{n}/{len(values)}项{label}共享“{s}”')
    return findings


def describe(sample):
    name, genre, groups, paras = sample
    out = [f'## 样本《{name}》', '', f'文种线索：{genre}。']
    lengths = [len(letters(t)) for _, t in paras]
    sentences = [len(letters(s)) for _, t in paras for s in SENTENCE.split(t) if letters(s)]
    out.append(f'正文{len(paras)}个有效段落，段长中位数{median(lengths):g}字；'
               f'{len(sentences)}个句段，句长中位数{median(sentences):g}字。')
    counts = Counter(c for _, text in paras for c in text if c in PUNCT)
    out.append('标点用量：' + ('、'.join(f'“{k}”{v}次' for k, v in counts.most_common()) or '未检出常见中文标点') + '。')
    if len(paras) < 4:
        out.append('可用正文较少，句段统计只说明这份短样本。')

    out.extend(['', '### 同级标题', ''])
    chains = [(k,v) for k,v in groups.items() if len(v) >= 2]
    if not chains:
        out.append('未识别出成组的同级标题。')
    for (depth, _), group in chains[:4]:
        shown = group[:6]
        out.append(f'- 第{depth}层，同组{len(group)}条。示例：' +
                   ' / '.join(f'段{ix} {quote(title)}' for ix,title in shown))
        first = [letters(CLAUSE.split(title)[0]) for _, title in shown]
        sizes = '、'.join(str(len(x)) for x in first)
        signals = repeated_edges(first)
        out.append(f'  示例标题首分句依次为{sizes}字。' + ('；'.join(signals)+'。' if signals else ''))

    out.extend(['', '### 段内节奏与过渡', ''])
    pairs = []
    openings = []
    recurring = defaultdict(set)
    for ix,t in paras:
        for s in SENTENCE.split(t):
            parts = [p.strip() for p in CLAUSE.split(s) if p.strip()]
            for part in parts:
                value = letters(part)
                for width in range(3, min(6, len(value)) + 1):
                    recurring['句首', value[:width]].add(ix)
                    recurring['句尾', value[-width:]].add(ix)
            for left,right in zip(parts, parts[1:]):
                n,m = len(letters(left)),len(letters(right))
                if 3<=n<=16 and 3<=m<=16 and abs(n-m)<=1 and len(pairs)<3:
                    pairs.append(f'- 段{ix}：{quote(left)} / {quote(right)}（{n}/{m}字）。')
        if len(openings)<4 and any(t.startswith(c) for c in TRANSITIONS):
            openings.append(f'- 段{ix}过渡起句：{quote(t[:40])}')
    out.extend(pairs or ['未检出长度相近的相邻分句。'])
    out.extend(openings or ['未检出预设过渡词开头的段落。'])
    selected = []
    for (edge, phrase), seen in sorted(recurring.items(), key=lambda x: (-len(x[1]), -len(x[0][1]))):
        if len(seen) < 2 or len(selected) >= 4:
            break
        if any(edge == prior_edge and phrase in prior for prior_edge, prior in selected):
            continue
        selected.append((edge, phrase))
        out.append(f'- {len(seen)}个段落的分句{edge}出现“{phrase}”（段号 {", ".join(map(str, sorted(seen)[:5]))}）。')
    return out


def render(paths, genre):
    samples, skipped, seen = [], [], set()
    for path in paths:
        if path.suffix.lower() != '.docx':
            raise ValueError(f'{path.name} 不是 .docx 文件。')
        sample=read_sample(path,genre)
        signature=(tuple(t for _,t in sample[3]),tuple(t for group in sample[2].values() for _,t in group))
        if signature in seen:
            skipped.append(path.name)
        else:
            seen.add(signature)
            samples.append(sample)
    output=['# 本地范文文风简表','',
            f'有效范文{len(samples)}篇。字数按汉字、字母和数字计，句长以句号、问号、叹号或段尾切分。'
            '文种取自显式参数或题名，仅作线索。段号按 Word 正文顺序定位。',
            '写新稿时，优先参考同文种样本的标题组织、节奏与过渡。例句是待分析的范文素材，'
            '原稿中的人名、数据、成绩和承诺只属于原稿。分句字数相近只提示形式，对仗是否合意要结合语义判断。']
    if len(samples)==1:
        output.append('目前只有一篇独立范文，可据此模仿本篇写法；作者长期偏好仍需其他范文印证。')
    if skipped:
        output.append('内容重复、已合并的输入：'+'、'.join(skipped)+'。')
    for sample in samples:
        output.append('')
        output.extend(describe(sample))
    return '\n'.join(output)+'\n'


def write_profile(text, replace):
    OUTPUT.parent.mkdir(parents=True,exist_ok=True)
    tmp=None
    try:
        with tempfile.NamedTemporaryFile('w',encoding='utf-8',dir=OUTPUT.parent,delete=False) as file:
            tmp=Path(file.name)
            file.write(text)
        if replace:
            os.replace(tmp,OUTPUT)
        elif os.name=='nt':
            os.rename(tmp,OUTPUT)
        else:
            os.link(tmp,OUTPUT)
    finally:
        if tmp is not None:
            tmp.unlink(missing_ok=True)


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('documents',nargs='+',type=Path,help='一篇或多篇本人定稿（.docx）')
    parser.add_argument('--genre',help='给这一批范文明示文种，例如工作汇报')
    modes=parser.add_mutually_exclusive_group()
    modes.add_argument('--stdout',action='store_true',help='将简表写到终端，不保存')
    modes.add_argument('--replace',action='store_true',help='替换此前生成的本地文风简表')
    args=parser.parse_args()
    try:
        text=render(args.documents,args.genre)
        if args.stdout:
            print(text,end='')
        else:
            write_profile(text,args.replace)
            print(f'已保存文风简表：{OUTPUT}')
    except FileExistsError:
        print('本地文风简表已经存在。确认要更新时加 --replace。',file=sys.stderr)
        return 1
    except (OSError, ValueError, BadZipFile, RuntimeError, ElementTree.ParseError) as error:
        print(f'文风分析失败：{error}',file=sys.stderr)
        return 1
    return 0


if __name__=='__main__':
    for stream in (sys.stdout,sys.stderr):
        if hasattr(stream,'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    sys.exit(main())
