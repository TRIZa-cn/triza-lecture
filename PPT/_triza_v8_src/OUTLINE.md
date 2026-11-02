# V8正式出版稿与独立补充附录

日期：2026-10-12，Asia/Shanghai。采用用户提供的正式出版模板，主稿严格8页（含参考文献），参考文献之后不接附录。补充材料独立PDF/TeX，发布同一根入口。

## 故事线：主张与证据

| 主张 | 证据与范围 |
|---|---|
| 主贡献是TRIZ指导Agent向哪些方向探索新/改进方案 | 原typed proposal、Figure1、PowerDuck记录实例；不是全部40类已实验验证 |
| FI/Ideality用于评价，依赖固定参考总体和协议 | 原公式2/3/5、已保留代码；历史点估计与拟议上界区分，不改实验 |
| 指标改善不等于晋级 | formal45及原门槛；四控制政策各自规则、资格非winner或部署授权 |
| 追溯是schema和已留存实例 | 公共originals、108文件清单、2040逐条重放；不宣称所有候选完整原件 |
| 三任务primary metrics沿用既有协议 | PowerDuck/DDoS weighted、CIC binary macro后平均family/seed；不可事后换主指标或跨任务排名 |
| 显示舍入不可逆 | formal精确端点差0.01583770699，表值五位；保留差+0.01584并明确先算差后舍入 |
| V18对两种MCC参照的delta方向不同 | same-fit core正、旧V17缓存负；不是同一候选MCC绝对值因比较而变 |
| 五项贡献与证据边界都保留 | 缩写重复解释、移出细节/表到独立Appendix；不弱化整体策略或删掉限制 |

## 主稿结构

准确的小节标题：Framework Overview；Theory-Guided Transformations；Functional Information and FI/Ideality；Traceability and Evidence State；Tasks and Paired Evaluation；Legacy Paired Results；Formal V2 and Campaign08；Downstream Evidence；Evidence Scope and Reproducibility。

Introduction：背景/问题/整体策略，五贡献简洁保留。
Related Work：TRIZ、program evolution、FI三类，保留15引用。
Methodology：三层四阶段、typed transformations、FI/Ideality核心公式、hard gates和traceability。
Experiments：task-specific metrics、三个比较参考、formal45/label规则、四政策统计、下游负结果及局限。
Conclusion与AI disclosure：沿用事实，修support operation歧义；不机械改合理时态。
References结束，不含Appendix。

## 独立Supplementary Appendix

保留原A—F完整材料，合并移出正文的程序、统计条件、case record和详细比较表。读者能从主稿明确定位在线章节；不留未定义的跨文件ref。
修改同样落入附录：指标理由、FI选择指代、MCC比较、舍入说明、复算与原SHA/publicSHA区分。
沿用完整中英双语网站读者指南与所有旧数据；新增独立附录下载与说明，不公开内部包。

## 格式和诚信门

- 原版acmart.cls，正式sigconf双栏，不改字号/边距/行距/vspace。
- 用内容组织和合理float placement满足8页，不堆无意义填充、不裁剪图表。
- 作者/通讯星号、AI disclosure、公式语义、实验数字和原图不改；不跑新训练或搜索。
- 官网日期October23–25,2026；未收到rights DOI/ISBN时不填样例值。
- 通用paperlint要求重复introduction/related图与本任务8页约束冲突，记录例外；参考文献不新增，bibguard不可用。
- 文档成果不改Triza仓库源码、不递增build；独立网站静态材料更新不部署应用。
