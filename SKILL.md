---
name: deepdraw-price-audit
description: Audit Deepdraw product prices in bulk from a listing-plan workbook and style-code list, including SKU/platform prices, product-line-specific rules, and an Excel report with error rows highlighted. Use when the user asks to 排查/核查深绘价格 or validate Deepdraw prices against 上市计划表.
---

# 深绘价格排查

输入只需要：

1. 上市计划表 `.xlsx` 文件；
2. 一个或多个待排查款号。

运行环境还必须已有可用的深绘登录会话。若未登录，请用户在浏览器中完成登录后继续；不要索取账号、密码、Cookie 或令牌。

## 工作流

1. 从上市计划表中同时检查期货与全域数据。工作表名称可能是“期货/全域”，也可能是“2026-线上/2026-全域”等包含“线上”或“全域”的名称。
2. 按完整款号精确匹配，取得每款的吊牌价与产品线。同款多色必须逐行核对：
   - 吊牌价和产品线一致时，使用该值；
   - 存在冲突时，停止对该款作价格结论，在结果中明确标记“上市计划表数据冲突”，不要自行选择。
3. 在已登录的深绘中逐款搜索并进入商品编辑页。确认页面款号等于当前款号后，再读取价格；不要沿用上一个商品页面的数据。
4. 等待 SKU 表和平台字段完成动态加载。按字段语义读取，不要依赖某一个品类的固定字段 ID。详细方法见 [references/deepdraw-extraction.md](references/deepdraw-extraction.md)。
5. 按 [references/rules-and-output.md](references/rules-and-output.md) 的 33 项规则核查。
6. 将抓取结果保存为规范化 JSON，并运行：

   ```bash
   python scripts/build_report.py \
     --plan /path/to/上市计划表.xlsx \
     --codes /path/to/款号.txt \
     --actual /path/to/deepdraw_actual.json \
     --output /path/to/价格排查结果.xlsx
   ```

7. 打开生成的 Excel 做最终检查：款号不能显示为科学计数法，错误行必须整行标红，产品线应位于款号后。

## 深绘查找与异常处理

- 搜索时使用完整款号；有多个精确匹配商品时，不要猜测，列出候选记录并请用户确认。
- 页面加载未完成时至少重试一次，并重新确认页面款号。字段仍未展示时，填写值写“未展示”，判定“错误”，说明“页面未展示该字段，无法取得应填价格”。
- SKU 字段必须核对全部 SKU，不得只抽查第一个。覆盖范围写成“18个SKU，均为239.9”或“18个SKU，存在2种填写值”。
- 唯品会市场价字段可显示为“划线价”，但需根据所属平台区域识别，不能与京东或有赞的划线价混淆。
- 1688 核对价格区间第一档的产品单价，并在覆盖范围中注明“价格区间首档，购买数量≥X”。
- 不修改或提交深绘商品；本 Skill 仅做只读排查和报告。

## 交付

- 默认输出一个 Excel 文件，不生成无关工作表。
- 表头固定为：`序号、款号、产品线、平台/区域、检查字段、填写值、目标值/要求、覆盖范围、核查结果、说明`。
- 保留所有正确和错误项目；不要只输出异常。
- 向用户简要汇报排查款数、正确项数、错误项数，并提供结果文件。
