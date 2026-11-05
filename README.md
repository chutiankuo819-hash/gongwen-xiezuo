# 公文写作 · gongwen-xiezuo

一套用于中文公文和机关、企事业单位常用材料的 Codex 技能。可以从零起草通知、请示、报告、函、纪要、总结、汇报和讲话，也可以只改一组小标题或一处过渡。克隆后即可用内置写作示例和带来源的习语索引，需要贴合单位文风时，可再导入自己的 Word 参考材料。

## 安装

在 Codex 中粘贴当前仓库的 GitHub 地址，并说“请用 `$skill-installer` 安装这个仓库根目录下的 `gongwen-xiezuo` 技能”。也可以在 GitHub 页点 **Code → Download ZIP**，解压，将含有 `SKILL.md` 的文件夹命名为 `gongwen-xiezuo`，放进 Codex 用户技能目录（`$CODEX_HOME/skills`）。如已有同名个人版，先保留该版本及其本地资料。

在新对话中输入：

> 用 $gongwen-xiezuo 起草一份图书馆通知。11月18日至20日盘点馆藏，阅览室照常开放。各阅览室负责人20日16时前将差异清单报办公室。正文写短一点。

也可以说“用 $gongwen-xiezuo 把这三个小标题改成判断句，每条表达一个不同意思，比喻用同一系列、同一长度”。

## 自带资料

- **45条原创示例**按标题、比喻、过渡和文种编排，可按篇目、小节或关键词检索。
- **306个习语短语**按9类写作用途组织，其中50个是便于拟稿的精选表达。索引列出例篇和出处。统计基于1,201份本人讲话、文章、书信等文本及单次讲话节选；另有1,579篇报道中的直接引语。同一讲话可能在两组都出现，两组数据分别使用。次数按连续汉字逐字命中计算，覆盖篇数指至少出现过一次的文本数。可按出处回查原文。[共产党员网讲话专栏](https://www.12371.cn/special/xxzd/jh/) · [《求是》署名文章专栏](https://www.qstheory.cn/zt2019/qskfzsjwz/index.htm) · [央广网《一习话》](https://news.cnr.cn/mryxh/)
- **用典与好句出处**提供12条当代讲话的主题线索，以及3条同时标明原典和讲话出处的古诗古语。引用前打开原文核对语境，普通词语自然用进正文即可。

## 导入自己的 Word 资料

在技能目录用 Python 3 运行，检索和导入均使用标准库：

```text
python scripts/import_docx.py "自己的写作参考.docx"
python scripts/search_phrases.py --source local --list-chapters
python scripts/search_phrases.py --section 过渡 --limit 5
python scripts/search_xiyu.py --preferred --length 5 学习
python scripts/search_xiyu.py --sentences 用典
```

导入器提取文档正文，按可识别的标题和小节建立本地索引，保存在 `references/local-phrasebook.jsonl`。替换已有索引时显式加 `--replace`。扫描图片需先做文字识别。导入一本有权使用的词句集后，可以优先查其中的原词和对偶句；没有 Word 文件时，程序会使用内置示例。原 Word 和本地索引已列入 Git 忽略规则。

## 许可与来源

本项目原创代码、写作说明和原创示例采用 [MIT 许可](LICENSE)。链接指向的文章、讲话等资料保留各自来源。使用者导入的文档由其自行管理。
