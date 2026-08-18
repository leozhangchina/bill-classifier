# 账单自动分类器（Python）

读取 Excel（`.xlsx`）或 CSV/TSV 账单，自动识别表头并把每笔交易归入 ezBookkeeping 中现有的一个二级分类。输出保留全部原始列，只在末尾新增：

- `最终分类`
- `是否大模型判断`

`categories.default.json` 已按本地软件当前分类页校准，共 65 个二级分类：42 个支出、13 个收入、10 个转账。最终分类只能是其中一项。

## 运行

```powershell
Set-Location 'D:\zqy\codex\bill-classifier'
python -m pip install -r requirements.txt
python .\classify_bill.py "D:\path\账单.xlsx"
```

输入 CSV、关闭大模型并跳过交互确认：

```powershell
python .\classify_bill.py "D:\path\账单.csv" --llm off --no-prompt
```

默认输出到输入文件旁的 `outputs` 目录，文件名为 `*_已分类_Python.xlsx`。可用 `--output` 指定路径。

## 分类顺序

1. 判断交易是支出、收入还是转账，只开放该分组下的二级分类。
2. 优先匹配 `merchant-overrides.json` 中经人工确认并学习的交易对方。
3. 匹配退款状态、固定商户规则和 `categories.default.json` 中的关键词。
4. 规则结果达到阈值（默认 82%）时直接采用，“是否大模型判断”为“否”。
5. 无匹配或存在歧义时，才调用大模型；大模型被约束为只能返回对应分组中的现有二级分类，“是否大模型判断”为“是”。
6. 大模型仍不确定时，在交互终端中允许人工确认。

这意味着大多数重复商户会在本地完成分类，不会发送到 API。交易字段只有在规则置信度不足时才会提交。

## 大模型配置

在当前 PowerShell 会话中设置：

```powershell
$env:OPENAI_API_KEY = '你的密钥'
$env:OPENAI_MODEL = '你的模型名'
python .\classify_bill.py "D:\path\账单.xlsx" --llm required
```

默认使用 Responses API 和严格 JSON Schema。兼容服务如果只支持 Chat Completions，可设置：

```powershell
$env:OPENAI_API_STYLE = 'chat'
$env:OPENAI_BASE_URL = 'https://服务地址/v1'
```

`--llm auto` 是默认值：环境变量完整时启用，否则保留低置信度候选并进入人工确认。`--llm required` 在 API 配置缺失或请求失败时直接报错；`--llm off` 完全禁用大模型。

## 动态更新关键词

- 通用关键词：直接编辑 `categories.default.json` 中相应分类的 `keywords`。
- 固定商户：编辑同一文件的 `merchantRules`。
- 自动学习：运行时加 `--learn`；人工确认后的“交易对方 → 二级分类”会写入 `merchant-overrides.json`，下次优先匹配。

```powershell
python .\classify_bill.py "D:\path\账单.xlsx" --learn
```

程序不会把未经人工确认的大模型结果自动写入学习规则，以免错误结果污染后续账单。

## 常用参数

```text
--output PATH          指定输出文件
--categories PATH      指定分类 JSON/TXT/CSV
--sheet NAME           指定 Excel 工作表
--threshold 0.82       规则/人工复核阈值
--llm auto|off|required
--no-prompt            不进行终端人工确认
--learn                学习本次人工确认的交易对方
```

## 轻量版（无大模型、无交互）

`classify_bill_light.py` 专门适配微信支付导出的 `.xlsx` 和支付宝导出的 `.csv`：

- 自动定位真实表头；微信样本会删除前 17 行说明，支付宝样本会删除前 23 行说明；
- 自动读取支付宝常见的 GB18030 编码并移除空白尾列；
- 支付宝“余额宝…收益发放”记录会把原始“收/支”从“不计收支”修正为“收入”；
- 转账不再使用单独分类组：原始“收/支”优先，方向不明确时再按“转入/收款”或“转出/还款”等词归入收入或支出；
- 不连接大模型，也不会弹出人工确认；
- 启动时读取 `merchant-overrides.json`，优先应用人工维护的“交易对方关键词 → 二级分类”；
- 每笔交易始终输出本地规则中置信度最高的分类；
- 只新增 `最终分类` 和 `是否需要人工确认` 两列；低于阈值（默认 82%）时标记“是”。

```powershell
python .\classify_bill_light.py "D:\path\微信支付账单.xlsx"
python .\classify_bill_light.py "D:\path\支付宝交易明细.csv"
```

如需调整人工确认敏感度，例如改为 75%：

```powershell
python .\classify_bill_light.py "D:\path\账单.xlsx" --threshold 0.75
```

商户覆盖规则既可以填写完整交易对方，也可以填写稳定片段。例如：

```json
{
  "KUMO KUMO": "食品",
  "某某咖啡店": "饮料"
}
```

匹配不区分英文字母大小写；如果多个关键词同时命中，优先使用最长的关键词。轻量版的覆盖分类必须属于该笔交易对应的支出或收入分类组，转账分类不会被采用。可用 `--overrides` 指定另一份规则文件。
