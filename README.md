# Deepdraw Price Audit Skill

一个可复用的深绘商品价格排查 Skill。用户提供上市计划表和待排查款号后，Skill 会从深绘读取各平台价格，按吊牌价、产品线及平台规则核查，并输出错误行标红的 Excel 结果表。

主要规则包括：

- 拼多多单买价 = 吊牌价 - 1 元；
- 拼多多团购价 = 吊牌价 - 2 元；
- 天猫专柜价 = 10000；
- 唯品会市场价为空；
- 鞋品产品线的京东市场价为空；
- 其余受检价格原则上等于上市计划表吊牌价。

Skill 入口为 [`SKILL.md`](SKILL.md)。运行环境需要可访问上市计划表、电子表格处理能力，以及已经登录深绘的浏览器会话。

自动报告脚本使用 Python 3 和 `openpyxl`：

```bash
python -m pip install -r requirements.txt
```
