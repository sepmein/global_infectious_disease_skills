# 可追溯证据模型

来源、事件和时间观测必须分开存储。一条来源可支持多个事件，一个事件可由多个来源支持，一个事件可包含多个日期的观测。

## Source

必需字段：

- `source_id`：稳定编号，如 `S000001`。
- `source_channel`：`LOCAL` 或 `WEB`。
- `publisher`、`source_type`、`title`、`language`、`evidence_grade`。
- `publication_date`、`statistic_cutoff`、`accessed_at`、`collected_at`。
- `original_url`、`retrieval_url`、`locator`。
- 本地来源的 `local_absolute_path`、`local_relative_path`、`sha256`。
- `access_status`、`archive_status`、`notes`。

`original_url` 是发布机构的规范地址；`retrieval_url` 是本轮实际访问地址。两者不能混为一列。

## Event

必需字段：

- `event_id`、`canonical_event_id`。
- `country_code`、`country_name_zh`、`country_name_en`。
- `disease_canonical`、`disease_as_reported`。
- `event_location`、`geographic_level`、`event_status`、`event_nature`。
- `first_observed_at`、`latest_observed_at`。
- `source_ids`、`confidence`、`inclusion_status`、`limitations`。
- `overlap_group_id`、`overlap_status`、`merge_decision`。

## Observation

必需字段：

- `observation_id`、`event_id`、`observed_at`。
- `statistic_start`、`statistic_end`、`epi_week`。
- `case_count`、`case_metric`、`case_definition`。
- `death_count`、`death_definition`。
- `affected_areas`、`event_status`、`pathogen_status`。
- `source_ids`、`summary`、`limitations`。

空值表示来源未报告或无法提取，不表示零。累计病例、期间新增、疑似、确诊、发病率和不同地理层级不得相加。

## Coverage

每个名单条目每轮恰有一条记录：

- `country_code`、`country_name_zh`、`checked_at`。
- `coverage_status`：`VERIFIED_EVENT`、`CHECKED_NO_SIGNAL`、`SIGNAL_UNVERIFIED`、`ACCESS_FAILED` 或 `NOT_CHECKED`。
- `channels_checked`：本地 WHO、本地 HealthMap、WHO/区域、国家官方、补充网络来源中的实际检查项。
- `reviewed_urls`、`latest_data_date`、`gap_reason`、`next_review_due`。

`CHECKED_NO_SIGNAL` 必须有实际检查渠道和页面记录。`ACCESS_FAILED` 与 `NOT_CHECKED` 不得产生无疫情或零病例结论。

## 追溯链

每个报告事实必须满足：

`报告文字或表格 -> Observation_ID/Event_ID -> Source_ID -> 原始 URL 或本地文件 -> locator -> 支持该事实的原文或数据行`

## 重叠状态

- `SAME_SNAPSHOT`：相同规范 URL 或内容哈希。
- `CORROBORATED`：独立来源对同一口径事件相互印证。
- `PARTIAL_OVERLAP`：同一疫情但时间、地区或病例定义不同。
- `CONFLICT`：同一截止日和口径出现实质不一致。
- `DISTINCT`：相近记录实际属于不同事件。
- `REVIEW_REQUIRED`：需要语义判断。

冲突数字不得平均。明确修订时保留旧值并标记 `SUPERSEDED`。
