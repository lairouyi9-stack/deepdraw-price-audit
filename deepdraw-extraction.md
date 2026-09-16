# 深绘页面取值指南

## 可靠性原则

深绘不同品类的字段编号和输入框名称前缀会变化。按页面标签、`data-field-name`、`data-sku-code` 和当前表头的 `data-option-id` 识别，不要把某个款的字段 ID 当成全局常量。

每次进入编辑页后先确认：

- 页面款号等于目标款号；
- SKU 表已经出现数据行；
- “专柜价”等平台字段完成动态加载；
- 加载过程中没有跳回登录页。

## SKU 字段

SKU 表通常为 `#skuContainer`。表头单元格含：

- `data-sku-code`：字段语义代码；
- `data-option-id`：当前品类对应的字段编号。

读取某字段时：

1. 找到目标 `data-sku-code` 的表头；
2. 取得其 `data-option-id`；
3. 在 `#skuContainer` 内读取名称以 `_<optionId>` 结尾的文本输入框；
4. 保存全部 SKU 值、SKU 数量、唯一值和空值数量。

字段代码：

```text
PRICE
RETAIL_PRICE
TM_DISCOUNT_PRICE
TM_SHOP_PRICE
JD_PRICE
PURCHASE_PRICE
LINEATION_PRICE
JD_VC_MARKET_PRICE
YOUZAN_PRICE
YOUZAN_STANDARD_PRICE
INDIVIDUAL_PRICE
GROUP_PRICE
PRICE_OF_LITTLE_RED_BOOK
ORIGINAL_PRICE
DOUYIN_PRICE
DOUYIN_SETTLEMENT_PRICE
KUAISHOU_NEW_PRICE
AIKUCUN_SUPPLY_PRICE
AIKUCUN_FLOOR_PRICE
AIKUCUN_RETAIL_PRICE
HAOYK_PRICE
HAOYK_ORIGINAL_PRICE
HAOYK_SUPPLY_PRICE
HAOYK_SETTLEMENT_PRICE
WX_PRICE
```

## 单值字段

优先按 `input[data-field-name="字段名称"]` 读取：

| JSON键 | 页面字段名称 |
|---|---|
| `douyinReferencePrice` | 抖音参考价 |
| `tmallOutletDiscountPrice` | 奥莱店折扣价 |
| `jdMarketPrice` | 京东市场价 |
| `vipMarketPrice` | 唯品会市场价 |
| `pddMarketPrice` | 商品市场价 |
| `tmallCounterPrice` | 专柜价 |

商品价格通常为商品基本信息区的“商品价格/零售价”输入框。保存为 `singles.productPrice`。

1688 产品单价位于“价格区间”区域：取第一组“购买数量”和“产品单价”，分别保存为 `alibabaMinQty` 和 `alibabaUnitPrice`。如果字段区域尚未出现，继续等待动态加载；不要直接判为空。

## 规范化 JSON

报告脚本接受如下结构：

```json
[
  {
    "code": "208426108218",
    "skuFields": [
      {"code": "PRICE", "count": 18, "values": ["239.9", "239.9"]}
    ],
    "singles": {
      "productPrice": "239.9",
      "douyinReferencePrice": "239.9",
      "alibabaMinQty": "1",
      "alibabaUnitPrice": "239.9",
      "tmallOutletDiscountPrice": "239.9",
      "jdMarketPrice": "239.9",
      "vipMarketPrice": "",
      "pddMarketPrice": "239.9",
      "tmallCounterPrice": "10000"
    }
  }
]
```

字段未展示时可使用 `null`；字段存在但未填写时使用空字符串。两者在说明中应区分。
