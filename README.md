# 账单轻量分类器

以 `classify_bill_light.py` 为主入口，在本地用商户规则和关键词给账单自动分类。无需 API 密钥或联网，不弹出分类确认；不确定的交易在 Excel 中标记为需要人工确认。

支持微信支付 `.xlsx`、支付宝 `.csv`、招商银行交易流水 `.pdf`、中国银行交易流水明细清单 `.pdf`。银行 PDF 需要有文字层，目前不支持扫描图片或其他银行的版式。

## 在另一台设备上安装

安装 Python 3.10 或更新版本，推荐 Python 3.12；Windows 安装时勾选将 Python 添加到 PATH。安装 Git 后，克隆默认分支：

```powershell
git clone https://github.com/leozhangchina/bill-classifier.git
cd bill-classifier
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe .\classify_bill_light.py --help
```

不需要激活虚拟环境，直接使用其中的 Python 即可。若 Windows 只有 `py` 命令，创建环境时改用 `py -3 -m venv .venv`。依赖已固定为本项目验证过的版本；首次安装依赖需要网络，分类过程离线运行。

macOS / Linux 使用：

```sh
git clone https://github.com/leozhangchina/bill-classifier.git
cd bill-classifier
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python classify_bill_light.py --help
```

复制文件迁移时，需要主脚本、两份 JSON 配置和 `requirements.txt`。虚拟环境应在新设备重新创建，不要复制旧设备的 `.venv`。

## 使用

Windows 示例（把路径改成自己的账单路径）：

```powershell
.\.venv\Scripts\python.exe .\classify_bill_light.py "D:\账单\微信支付账单.xlsx"
.\.venv\Scripts\python.exe .\classify_bill_light.py "D:\账单\支付宝交易明细.csv"
.\.venv\Scripts\python.exe .\classify_bill_light.py "D:\账单\招商银行交易流水.pdf"
.\.venv\Scripts\python.exe .\classify_bill_light.py "D:\账单\中国银行交易流水.pdf"
```

Windows 也可以用 `run.cmd "D:\账单\账单.xlsx"`，它优先使用项目内的 `.venv`，并把其余参数传给轻量脚本。PowerShell 包装脚本示例：

```powershell
.\run.ps1 -InputFile "D:\账单\账单.xlsx" -Threshold 0.75
```

默认结果放在**输入文件旁的 `outputs` 文件夹**，名称为 `原文件名_已分类_轻量版.xlsx`。中国银行结果会根据账单内最早和最晚的交易日期命名，例如 `中国银行交易流水(20260101-20260131)_已分类_轻量版.xlsx`，不沿用乱码文件名。同名结果会覆盖，保留多个版本时请用 `--output` 指定新名称。程序拒绝把输出写回原始账单。

```powershell
.\.venv\Scripts\python.exe .\classify_bill_light.py "D:\账单\账单.xlsx" --output "D:\结果\已分类.xlsx"
```

## 输出格式与特殊处理

四种账单统一输出七列：`交易时间`、`交易类型`、`金额`、`交易对方`、`交易说明`、`二级分类`、`是否需要人工确认`。

- `交易类型` 表示收入、支出或不计收支；交易金额按绝对值输出为可计算数值。
- 自动定位真实表头，删除微信、支付宝账单开头的说明及空白尾列，不固定依赖说明行数。
- 招行从“客户摘要”最后一个 `-` 后提取交易对方；中行去掉交易对方中的常见支付渠道前缀。
- 支付宝“余额宝…收益发放”修正为收入；退款归入收入下的“退款”。
- 普通转账按原始收支或转入/转出语义判断。微信“转入零钱通-来自零钱”属于账户内划转，记为不计收支，二级分类使用现有的“其他转账”。
- 不输出交易单号、商户单号、账户余额、对方账号、网点等栏位。
- 规则置信度低于阈值（默认 `0.82`）时，人工确认列为“是”。这是复核提示，程序仍给出候选分类。

## 分类依据和动态规则

先判断收支方向，限制可用的二级分类，再按以下顺序判断：人工维护的商户覆盖规则、退款规则、固定商户规则、关键词评分。关键词越具体、匹配越多，评分越高；候选相近、群收款或只有粗略原始分类时，置信度降低。没有匹配时使用对应收支组的兜底分类，并标记人工确认。置信度是规则评分，不是经过统计校准的正确率。

`categories.default.json` 保存二级分类、收支分组、关键词、固定商户规则及兜底分类；`merchant-overrides.json` 保存已确认的“交易对方关键词 → 二级分类”。迁移时请一起保留你修改过的这两份文件。

商户覆盖规则支持完整名称或稳定片段，英文字母不区分大小写；多个关键词命中时采用最长者，分类必须属于该笔交易对应的收支组。示例：

```json
{
  "某某咖啡店": "饮料",
  "某某球馆": "运动健身"
}
```

轻量脚本读取规则但不自动修改或学习规则。配置路径默认相对于脚本所在目录，所以从其他工作目录运行也能找到配置。

## 常用参数

```text
--output PATH       指定输出 .xlsx 文件
--categories PATH   指定分类配置 JSON
--overrides PATH    指定商户覆盖规则 JSON
--sheet NAME        指定 Excel 工作表
--threshold 0.75    调整人工复核阈值，范围 0 到 1
--pdf-password TEXT 指定加密 PDF 的打开密码
```

加密 PDF 默认在交互终端中隐藏输入密码。批量运行时也可以临时设置 `BILL_CLASSIFIER_PDF_PASSWORD` 环境变量；参数中的密码可能进入终端历史，交互输入更适合日常使用。密码不保存到代码、配置或 Excel。

```powershell
$env:BILL_CLASSIFIER_PDF_PASSWORD = '自己的密码'
.\.venv\Scripts\python.exe .\classify_bill_light.py "D:\账单\中国银行交易流水.pdf"
Remove-Item Env:BILL_CLASSIFIER_PDF_PASSWORD
```

## 项目结构与检查

```text
classify_bill_light.py      主程序
categories.default.json    分类和关键词配置
merchant-overrides.json    商户覆盖规则
requirements.txt           已验证的依赖版本
run.cmd / run.ps1          Windows 启动包装脚本
tests/test_light.py         不包含真实账单的回归检查
legacy/                    暂存旧版大模型脚本及配置示例
```

运行检查：

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m pip check
```

账单、输出 Excel、密码文件、虚拟环境和 JavaScript 本地文件均被 Git 忽略。仓库只保存程序及分类规则。`legacy/classify_bill.py` 保留原来的可选大模型功能，使用方式见 `legacy/README.md`；当前轻量版不导入或调用它。
