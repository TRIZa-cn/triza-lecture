# TRIZa AI4Science camera-ready V8 / 正式出版稿V8

2026-10-12 · Asia/Shanghai

## 中文

在Overleaf上传完整源码包，选择 **XeLaTeX** 编译器；参考文献由BibTeX处理。

- 主文档 `main.tex`：会议提供的ACM `sigconf` 双栏，编译结果8页，包含参考文献；参考文献后没有附录。
- 主文档 `supplement.tex`：独立在线补充附录，20页（含封面、A—H和参考文献），不读取 `main.aux`。
- `acmart.cls` 原样来自正式出版模板；原有参考文献文件及三张图的内容未改。
- `OUTLINE.md` 记录缩减前的结构与证据安排，不参与编译。

作者信息和通讯作者保留。DOI、ISBN为空，`\setcopyright{none}` 暂不生成未确认的版权声明。出版方提供rights设置后，应替换相应元数据并再次检查页数；未复制模板的2018年、Woodstock或虚构DOI/ISBN。

在线附录、双语说明及既有实验材料统一入口：
https://triza.cn/triza-paper-exp-data2026/

没有修改实验原始测量、重新训练或新增搜索。FI的一般保守上/下界公式与历史实现的点估计、成本代理严格区分；不把计划中的实现写成已经验证。局限、比较对象、资格/赢家标签和晋级边界均保留。

## English

Upload the complete archive to Overleaf and select **XeLaTeX**, with BibTeX for the bibliography.

- `main.tex`: native conference-supplied ACM `sigconf`; **8 pages including references**, with no appended supplement.
- `supplement.tex`: independently compiled online supplement; **20 pages** including the cover, Appendices A–H and references. It does not require `main.aux`.
- The supplied `acmart.cls`, existing bibliography and original three figure assets are preserved.
- `OUTLINE.md` documents content organization and evidence boundaries; it is not compiled.

DOI and ISBN remain empty. The temporary `\setcopyright{none}` avoids asserting an unconfirmed rights grant. Replace publication metadata when instructed by the publisher and recheck pagination. No sample rights identifiers were invented.

Expanded methods, bilingual reader guidance, retained measurements and decision replay are available at the shared public entry above. No new training/search was performed. General bound-based FI specifications are distinguished from the historical implementation and observed results.
