// Run in the already authenticated Deepdraw product edit page.
// Wait until the page style code and dynamic fields match the requested product.
(() => {
  const skuCodes = [
    "PRICE", "RETAIL_PRICE", "TM_DISCOUNT_PRICE", "TM_SHOP_PRICE",
    "JD_PRICE", "PURCHASE_PRICE", "LINEATION_PRICE", "JD_VC_MARKET_PRICE",
    "YOUZAN_PRICE", "YOUZAN_STANDARD_PRICE", "INDIVIDUAL_PRICE", "GROUP_PRICE",
    "PRICE_OF_LITTLE_RED_BOOK", "ORIGINAL_PRICE", "DOUYIN_PRICE",
    "DOUYIN_SETTLEMENT_PRICE", "KUAISHOU_NEW_PRICE", "AIKUCUN_SUPPLY_PRICE",
    "AIKUCUN_FLOOR_PRICE", "AIKUCUN_RETAIL_PRICE", "HAOYK_PRICE",
    "HAOYK_ORIGINAL_PRICE", "HAOYK_SUPPLY_PRICE", "HAOYK_SETTLEMENT_PRICE",
    "WX_PRICE"
  ];
  const wanted = new Set(skuCodes);
  const byName = name => document.querySelector(`input[data-field-name="${name}"]`)?.value ?? null;
  const skuFields = [...document.querySelectorAll("#skuContainer td[data-sku-code]")]
    .filter(header => wanted.has(header.dataset.skuCode))
    .map(header => {
      const optionId = header.dataset.optionId;
      const values = [...document.querySelectorAll(`#skuContainer input[type="text"][name$="_${optionId}"]`)].map(input => input.value);
      return { code: header.dataset.skuCode, count: values.length, values };
    });
  const priceRange = [...document.querySelectorAll("li.form_li")]
    .find(item => item.querySelector("em")?.innerText.trim().startsWith("价格区间"));
  const rangeValues = priceRange
    ? [...priceRange.querySelectorAll('input[type="text"]')].map(input => input.value)
    : [];
  return {
    code: document.querySelector("#code")?.value ?? null,
    skuFields,
    singles: {
      productPrice: document.querySelector("#retailPrice")?.value ?? null,
      douyinReferencePrice: byName("抖音参考价"),
      alibabaMinQty: rangeValues[0] ?? null,
      alibabaUnitPrice: rangeValues[1] ?? null,
      tmallOutletDiscountPrice: byName("奥莱店折扣价"),
      jdMarketPrice: byName("京东市场价"),
      vipMarketPrice: byName("唯品会市场价"),
      pddMarketPrice: byName("商品市场价"),
      tmallCounterPrice: byName("专柜价")
    }
  };
})();
