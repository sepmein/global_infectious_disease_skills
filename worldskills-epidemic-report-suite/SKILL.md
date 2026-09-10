---
name: worldskills-epidemic-report-suite
description: 基于动态国家和地区名单，融合本地 WHO/HealthMap 爬取数据与实时网络证据，生成境外传染病每日监测日报或世界技能大赛多国疫情摸底报告及可追溯底表。用于跨国疫情监测和赛事公共卫生研判，不用于个体诊疗。
metadata:
  short-description: 融合本地和网络证据生成境外疫情日报与多国摸底报告
---

# 世赛境外疫情报告套件

以同一套证据库生成两种交付物：每日监测日报强调相对上一期的新增和变化；多国摸底报告强调指定期间的逐国覆盖、重点国家和重点疫情。报告中的每项事实都必须能追溯到来源记录和原始位置。

`daily` 模式生成 Word 日报前必须读取并执行 [日报规范.md](日报规范.md)；该文件是日报内容结构和版式的最高优先级规范。`baseline` 模式必须以 [世界技能大赛传染病疫情摸底报告_写作模板.docx](assets/世界技能大赛传染病疫情摸底报告_写作模板.docx) 为结构和版式的最高优先级依据，并读取 [baseline-report.md](references/baseline-report.md)。这些 Word 版式不适用于配套 Excel 证据底表。

## 输入与模式

先读取 [input-contract.md](references/input-contract.md)，确定日期、模式、名单、本地数据目录、历史运行目录和输出目录。

- 用户要求“日报、每日监测、当日变化”时使用 `daily` 模式，并读取 [日报规范.md](日报规范.md)、[daily-report.md](references/daily-report.md) 与 [shanghai-risk.md](references/shanghai-risk.md)。三者冲突时以 `日报规范.md` 为准。
- 用户要求“摸底、指定期间、重点国家、世赛疫情报告”时使用 `baseline` 模式，并读取 [baseline-report.md](references/baseline-report.md)、`references/participants.csv`、`references/priority-diseases.xlsx` 和模板 `assets/世界技能大赛传染病疫情摸底报告_写作模板.docx`。
- 用户同时要求两种报告时，先构建一次共享证据库，再分别生成两种交付物；不得重复计数或分别维护事实。

## 强制原则

开始检索和归一化前，读取 [source-strategy.md](references/source-strategy.md) 与 [evidence-schema.md](references/evidence-schema.md)。

- 动态读取名单，不把“113 个”或“115 个”写死。默认名单为 `references/participants.csv`。
- 国家级网络检索必须先读取 [country-monitoring-sources.xlsx](references/country-monitoring-sources.xlsx) 的“国家监测数据源”工作表。以“地区代码”优先匹配参赛名单；每个匹配国家的“国家监测数据源 URL”和“区域交叉来源 URL”都是本轮必须执行的检索入口，不得仅作为备注保存。
- 来源参考表中的“实际访问状态”“实际最终 URL”“页面标题”和“访问日期”是历史核验记录，不代表本轮已完成访问。疫情数据具有时效性，每次报告必须重新访问并记录本轮结果。
- `baseline` 模式采用封闭范围：国家和地区只能来自 `references/participants.csv`，疾病只能来自 `references/priority-diseases.xlsx`。本地扫描、在线实时检索、证据归一化、重点国家筛选和正文写作均不得纳入名单外国家或病种表外疾病。来源材料中同时出现的其他疾病只作忽略处理，不进入候选事件、底表或报告。
- `daily` 模式仍可把 `references/priority-diseases.xlsx` 作为定向检索表，并按日报规范处理其他达到报告阈值的重要事件；除非用户另行要求，不把 `baseline` 的疾病白名单规则扩展到日报。
- `baseline` 模式中，每个疾病的“疾病概述”必须先按 `references/disease-intro-index.csv` 读取 `references/疾病简介/` 中对应的本地疾病文档。存在对应文档时，以该文档为内容依据，不用互联网资料替代或任意扩写；仅当映射文件指定的文档确实缺失时，才检索权威互联网资料补齐疾病简介。
- `baseline` 模式中，全球流行概况、重点国家流行情况、趋势、医疗负担、监测能力或风险因素若不能由本地报告充分支撑，必须实时检索互联网并回到可核验的官方来源，不得以材料不足为由推测或留用模板示例。
- `baseline` 模式的疾病分类只能使用 `references/priority-diseases.xlsx` 的“所属分类”字段，不得凭疾病名称自行归类或调整分类。
- 本地 WHO 与本地 HealthMap 是两个不同来源渠道。本地文件是证据载体，不自动等于原始发布机构。
- WHO 本地快照保留规范原文 URL、内容和抓取时间时，可按 WHO 官方证据评级；HealthMap 默认作为发现线索，除非记录能够直接回溯并核实原始官方页面。
- 搜索引擎、搜索摘要、HealthMap 摘要和媒体转载只能发现候选事件。病例、死亡、日期和趋势必须优先回到参考表指定的国家卫生主管部门、国家疾控机构、国家公共卫生研究院、WHO 或正式区域监测系统核实。出现重复、高严重性或快速变化的线索时，国家官方来源核实不得省略。
- 将来源、疫情事件和每日观测分开保存。不得用标题相似直接删除记录，也不得平均冲突数字。
- 未检索到信号不等于无疫情；访问失败不等于零病例。逐国状态必须使用规定的覆盖状态。
- 疫情数据时效性高且涉及公共卫生，网络部分必须在本轮实时访问。不得使用搜索摘要或上次报告替代本轮访问。

## 执行流程

1. **配置校验**：运行 `scripts/validate_scope.py` 校验名单和重点病种配置。`baseline` 必须使用默认的 `references/participants.csv` 和 `references/priority-diseases.xlsx`，并确认每个病种均有非空“所属分类”。记录实际国家和地区数量、病种数量、分类、监测起止日期、时区和运行时间。
2. **本地证据清点**：提供本地目录时运行 `scripts/inventory_local_sources.py`，保留绝对路径、相对路径、SHA-256、原始 URL 和抓取时间。未提供时明确记录 `SKIPPED_NOT_PROVIDED`，不得猜测目录。
3. **强制来源台账加载**：读取 `references/country-monitoring-sources.xlsx`，按地区代码与本轮名单连接。逐国保存国家机构、国家入口、区域交叉来源和历史访问记录。匹配不到的国家必须先发现并确认国家官方来源，标记 `SOURCE_REGISTRY_MISSING`，不得直接跳过国家级检索。
4. **指定来源检索**：本轮实际访问每个国家的“国家监测数据源 URL”，继续进入疫情新闻、法定监测周报、病例数据、暴发通报、仪表板以及 PDF/XLSX/CSV 附件；只打开机构首页不算完成。随后访问该行“区域交叉来源 URL”进行补充发现和交叉核验。入口失效时依次尝试“实际最终 URL”、同一官方机构域内搜索和官网名称定向搜索，并记录失败原因及替代入口。
5. **高召回发现**：`daily` 先处理本地 WHO/HealthMap 增量，再处理 WHO、参考表指定的国家与区域机构及正式数据源，最后执行逐国多语种通用搜索。`baseline` 只处理同时命中参赛名单和重点病种表的本地记录，并只执行名单内国家与表内疾病相关的在线查询；不得使用名单外国家或表外疾病扩展检索范围。
6. **定向核实**：对已发现信号、重点病种、活跃事件和高风险国家回到官方原页核实。HealthMap 或媒体出现病例、死亡、聚集性疫情或本土传播线索时，必须检查该国参考表中的国家官方来源。保存页面标题、发布机构、发布日期、统计截止日、病例口径、原始 URL、实际访问 URL、定位信息和访问时间。
7. **归一化与去重**：按 [evidence-schema.md](references/evidence-schema.md) 建立 Source、Event、Observation 和 Coverage 记录；运行 `scripts/generate_overlap_candidates.py` 生成重叠候选，再进行语义复核。
8. **变化识别**：`daily` 模式使用 `scripts/detect_daily_changes.py` 比较当前与上一期观测，只把新事件或实质性变化写入“当日新增及变化疫情动态”；较上一版无变化但仍需关注的事件按 [日报规范.md](日报规范.md) 写入“其他未更新传染病疫情”。“当日新增及变化疫情动态”是面向读者的疫情事实模块，只能写国家或地区、疾病、病例/死亡/暴发等最新数据、统计截止日期、与上一期的变化、当前趋势或状态及必要的不确定性说明。不得写入文件扫描数量、检索国家数量、访问成功或失败、网页打不开、搜索关键词、来源筛选过程、去重/冲突处理、证据链构建、覆盖状态、模型判断过程、写作过程或任何其他检索与思考记录；这些内容只能保存在 Excel 证据底表、运行清单或内部校验记录中。
9. **覆盖校验**：运行 `scripts/validate_coverage.py`。每个名单条目必须恰有一个本轮状态；`NOT_CHECKED` 禁止正式成稿，其他缺口必须如实披露。覆盖记录必须同时说明国家来源和区域交叉来源的本轮检查结果。
10. **证据校验**：运行 `scripts/validate_evidence_package.py`，确保报告事实能够通过 Event_ID、Observation_ID 和 Source_ID 回到原始证据。
11. **生成与验收**：按模式生成 Word 和 Excel。`daily` Word 按 [日报规范.md](日报规范.md) 编排；`baseline` Word 必须从 `assets/世界技能大赛传染病疫情摸底报告_写作模板.docx` 派生，保留模板的页面、字体、段落、层级和表格规则，并用实际内容替换全部说明、占位和虚构示例。随后读取并执行 [acceptance-gates.md](references/acceptance-gates.md)。Word 使用 documents Skill 的模板复刻流程渲染并逐页检查；Excel 使用 spreadsheets Skill 检查公式、筛选、列宽和全部工作表。`daily` 交付前必须对第二部分执行正文净化检查：逐段确认每段都包含可核验的疫情信息或其必要的变化说明；发现“检索、访问、扫描、核验、覆盖、证据、来源筛选、搜索、判断、生成”等过程性表述时，删除或改写为对应的疫情事实，不能仅靠增加“信息来源”条目掩盖过程性内容。

## 交付约束

- `daily`：交付可编辑 Word 日报和支持核查的 Excel 证据底表；无实质变化时仍生成日报，但必须明确“本轮未发现符合报告阈值的新变化”，不能写“无疫情”。
- `baseline`：交付可编辑 Word 摸底报告和完整 Excel 底表；只包含名单内国家和重点病种表内疾病，重点国家最多 10 个，实际不足时如实列出。Word 必须严格沿用指定模板，不得改用通用公文模板重建。
- 每次运行使用独立输出目录和版本化文件名，不覆盖历史运行。保留标准化来源、事件、观测、覆盖记录和重叠复核结果。
- 正式交付前运行 `scripts/validate_deliverables.py --require-official-format`。结构、证据或公文版式校验失败时只能交付诊断结果或明确标注的送审稿，不得把部分文件标为正式报告。

## 停止条件

只有在名单或日期无效、本地目录由用户指定但不可读、关键证据无法追溯、仍有 `NOT_CHECKED` 国家/地区，或交付校验失败时停止正式发布。官网受限、信号待核实或单国资料缺失应记录为覆盖缺口并继续处理其他国家，不得伪造完成状态。
