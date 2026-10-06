#!/usr/bin/env python3
"""微信、支付宝、招商银行和中国银行账单轻量分类器。"""

from __future__ import annotations

import argparse
import csv
import getpass
import json
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

if sys.platform == "win32":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError as exc:
    raise SystemExit("缺少 openpyxl，请运行：python -m pip install openpyxl") from exc


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_CATEGORIES = SCRIPT_DIR / "categories.default.json"
DEFAULT_OVERRIDES = SCRIPT_DIR / "merchant-overrides.json"
OUTPUT_HEADERS = [
    "交易时间", "交易类型", "金额", "交易对方", "交易说明",
    "二级分类", "是否需要人工确认",
]

# 不同账单采用不同列名，这里统一为程序内部字段。
HEADER_ALIASES = {
    "time": ["交易时间", "时间", "支付时间", "记账时间", "记账日期"],
    "type": ["交易类型", "交易分类", "交易摘要", "类型"],
    "counterparty": ["交易对方", "商户名称", "对方", "收款方"],
    "account": ["对方账号", "交易对方账号", "对手信息"],
    "product": ["商品", "商品说明", "商品名称", "客户摘要", "摘要", "交易说明"],
    "direction": ["收/支", "收支", "交易方向"],
    "amount": ["金额(元)", "金额（元）", "金额", "交易金额"],
    "balance": ["联机余额", "账户余额", "余额"],
    "currency": ["货币", "币种"],
    "payment": ["支付方式", "付款方式", "收/付款方式", "收付款方式"],
    "status": ["当前状态", "交易状态", "状态"],
    "transaction_id": ["交易单号", "交易订单号", "交易号", "订单号"],
    "merchant_id": ["商户单号", "商家订单号", "商家订单号"],
    "remark": ["备注", "附言", "说明"],
}

CMB_PDF_HEADERS = [
    "记账日期", "货币", "交易金额", "联机余额", "交易摘要",
    "对手信息", "客户摘要", "交易对方", "收/支",
]

BOC_PDF_HEADERS = [
    "交易时间", "币种", "交易金额", "账户余额", "交易类型", "支付方式",
    "网点名称", "备注", "交易对方", "对方账号", "对方开户行", "收/支",
]

PDF_PASSWORD_ENV = "BILL_CLASSIFIER_PDF_PASSWORD"

# 对两个样本中常见、但通用分类表尚未覆盖的表达做少量补充。
# 用户仍可直接编辑 categories.default.json 的 keywords 动态扩充规则。
EXTRA_KEYWORDS = {
    "食品": ["餐饮美食", "一餐", "美食", "小杨生煎", "萨莉亚", "肯德基", "麦当劳", "汉堡王", "麻辣烫", "螺蛳粉"],
    "饮料": ["coffee", "manner"],
    "水果零食": ["果优鲜", "油桃", "蟠桃"],
    "家居用品": ["日用百货"],
    "电子产品": ["充电宝"],
    "私家车费用": ["爱车养车", "充电订单", "停车场"],
    "运动健身": ["体育系", "体育场", "场地费"],
    "会员订阅": ["连续包月", "自动续费"],
    "利息收入": ["收益发放", "余额宝收益"],
}

# 支付宝“交易分类”可作为辅助证据，但不会覆盖更具体的商品关键词。
SOURCE_CATEGORY_HINTS = {
    "餐饮美食": "食品",
    "日用百货": "家居用品",
    "爱车养车": "私家车费用",
    "信用借还": "其他支出",
}


@dataclass
class Category:
    label: str
    flow: str
    keywords: list[str] = field(default_factory=list)


@dataclass
class Config:
    categories: list[Category]
    fallback_by_flow: dict[str, str]
    merchant_rules: list[dict[str, Any]]

    @property
    def labels(self) -> list[str]:
        return [item.label for item in self.categories]

    def categories_for_flow(self, flow: str) -> list[Category]:
        return [item for item in self.categories if item.flow == flow]

    def labels_for_flow(self, flow: str) -> set[str]:
        return {item.label for item in self.categories_for_flow(flow)}


@dataclass
class SourceData:
    headers: list[str]
    lookup: dict[str, int]
    rows: list[list[Any]]
    sheet_name: str
    provider: str
    removed_leading_rows: int


@dataclass
class Result:
    category: str
    confidence: float
    method: str


def normalize_text(value: Any) -> str:
    text = "" if value is None else str(value)
    return re.sub(r"\s+", " ", text.lstrip("`").replace("\u3000", " ")).strip().lower()


def normalize_header(value: Any) -> str:
    return re.sub(r"[\s:：]", "", normalize_text(value))


def header_lookup(headers: list[Any]) -> dict[str, int]:
    normalized = [normalize_header(item) for item in headers]
    result: dict[str, int] = {}
    for key, aliases in HEADER_ALIASES.items():
        choices = {normalize_header(alias) for alias in aliases}
        result[key] = next((index for index, item in enumerate(normalized) if item in choices), -1)
    return result


def header_score(row: list[Any] | tuple[Any, ...]) -> int:
    return sum(index >= 0 for index in header_lookup(list(row)).values())


def find_header_row(rows: list[list[Any]]) -> int:
    score, index = max(
        ((header_score(row), index) for index, row in enumerate(rows[:100])),
        default=(-1, -1),
    )
    if score < 6:
        raise ValueError(f"无法识别账单表头（最多匹配 {score} 个字段）")
    return index


def decode_csv(raw: bytes) -> str:
    for encoding in ("utf-8-sig", "gb18030", "utf-16"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("无法识别 CSV 编码")


def join_pdf_words(words: list[dict[str, Any]]) -> str:
    """按 PDF 中的纵向位置拼接同一单元格里被折行的文字。"""
    lines: list[list[dict[str, Any]]] = []
    for word in sorted(words, key=lambda item: (item["top"], item["x0"])):
        if not lines or abs(word["top"] - lines[-1][0]["top"]) > 1.5:
            lines.append([word])
        else:
            lines[-1].append(word)

    text = ""
    for line_words in lines:
        line = " ".join(str(word["text"]) for word in sorted(line_words, key=lambda item: item["x0"]))
        if not text:
            text = line
        elif text.endswith("-"):
            text += line
        elif re.search(r"[A-Za-z0-9]$", text) and re.match(r"[A-Za-z0-9]", line):
            text += " " + line
        elif re.search(r"[\u4e00-\u9fff）)]$", text) and re.match(r"\d", line):
            text += " " + line
        else:
            text += line
    text = text.replace("－", "-").replace("–", "-").replace("—", "-")
    return re.sub(r"\s*-\s*", "-", text).strip()


def cmb_counterparty(customer_summary: str, counter_info: str) -> str:
    """招商银行客户摘要最后一个连字符后的内容是原交易对方。"""
    if "-" in customer_summary:
        candidate = customer_summary.rsplit("-", 1)[-1].strip()
        if candidate:
            return candidate
    return re.sub(r"\s+\d[\d*]*$", "", counter_info).strip()


def pdf_password_error(exc: BaseException) -> bool:
    """判断 pdfplumber 包装的异常是否由 PDF 打开密码错误引起。"""
    current: BaseException | None = exc
    visited: set[int] = set()
    while current is not None and id(current) not in visited:
        visited.add(id(current))
        if current.__class__.__name__ == "PDFPasswordIncorrect":
            return True
        current = current.__context__ or current.__cause__
    return False


def open_pdf_document(path: Path, supplied_password: str | None) -> tuple[Any, str | None]:
    """打开 PDF；密码只从参数、环境变量或终端隐藏输入中取得。"""
    try:
        import pdfplumber
    except ImportError as exc:
        raise ValueError("读取银行 PDF 需要 pdfplumber，请运行：python -m pip install pdfplumber") from exc

    environment_password = os.environ.get(PDF_PASSWORD_ENV)
    password = supplied_password if supplied_password is not None else environment_password
    try:
        return pdfplumber.open(path, password=password or ""), password
    except Exception as exc:
        if not pdf_password_error(exc):
            raise ValueError(f"无法读取 PDF：{exc or exc.__class__.__name__}") from exc
        if password is not None:
            raise ValueError("PDF 打开密码不正确") from exc
        if not sys.stdin.isatty():
            raise ValueError(
                f"PDF 已加密；请使用 --pdf-password、环境变量 {PDF_PASSWORD_ENV}，"
                "或在终端运行后按提示输入密码"
            ) from exc

    entered = getpass.getpass(f"请输入 {path.name} 的 PDF 打开密码：")
    if not entered:
        raise ValueError("未输入 PDF 打开密码")
    try:
        return pdfplumber.open(path, password=entered), entered
    except Exception as exc:
        if pdf_password_error(exc):
            raise ValueError("PDF 打开密码不正确") from exc
        raise ValueError(f"无法读取 PDF：{exc or exc.__class__.__name__}") from exc


def pdf_cell_text(value: Any) -> str:
    text = "" if value is None else str(value)
    text = re.sub(r"\s*\n\s*", "", text).strip()
    return "" if re.fullmatch(r"-{5,}", text) else text


def boc_counterparty(counterparty: str, remark: str, transaction_name: str) -> str:
    """去掉常见支付渠道前缀，保留中国银行账单中的实际交易对方。"""
    candidate = counterparty or remark or transaction_name
    candidate = candidate.replace("（", "(").replace("）", ")").strip()
    for prefix in ("支付宝-", "抖音支付-", "微信支付-", "财付通-"):
        if candidate.startswith(prefix):
            return candidate[len(prefix):].strip()
    return candidate


def read_cmb_pdf_rows(path: Path, pdf_password: str | None = None) -> list[list[Any]]:
    try:
        import pdfplumber
    except ImportError as exc:
        raise ValueError("读取招商银行 PDF 需要 pdfplumber，请运行：python -m pip install pdfplumber") from exc

    rows: list[list[Any]] = []
    with pdfplumber.open(path, password=pdf_password or "") as pdf:
        first_text = (pdf.pages[0].extract_text() or "") if pdf.pages else ""
        if "招商银行交易流水" not in first_text:
            raise ValueError("目前仅支持带文字层的招商银行交易流水 PDF")

        for page in pdf.pages:
            words = page.extract_words(x_tolerance=1, y_tolerance=3, keep_blank_chars=False)
            anchors = sorted(
                (
                    word for word in words
                    if word["x0"] < 80 and re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(word["text"]))
                ),
                key=lambda item: item["top"],
            )
            if not anchors:
                continue

            boundaries = [anchors[0]["top"] - 18]
            boundaries.extend((left["top"] + right["top"]) / 2 for left, right in zip(anchors, anchors[1:]))
            boundaries.append(anchors[-1]["top"] + 18)

            for index, anchor in enumerate(anchors):
                row_words = [
                    word for word in words
                    if boundaries[index] <= word["top"] < boundaries[index + 1]
                ]
                columns = [
                    join_pdf_words([word for word in row_words if lower <= word["x0"] < upper])
                    for lower, upper in (
                        (0, 80), (80, 130), (130, 205), (205, 285),
                        (285, 370), (370, 475), (475, page.width + 1),
                    )
                ]
                date_text, currency, amount_text, balance_text, summary, counter_info, customer_summary = columns
                if date_text != str(anchor["text"]) or not currency:
                    raise ValueError(f"招商银行 PDF 第 {page.page_number} 页存在无法识别的交易行")

                signed_amount = parse_amount(amount_text)
                balance = parse_amount(balance_text)
                if not isinstance(signed_amount, (int, float)) or not isinstance(balance, (int, float)):
                    raise ValueError(f"招商银行 PDF 第 {page.page_number} 页金额或余额无法识别")
                direction = "收入" if signed_amount > 0 else "支出" if signed_amount < 0 else "不计收支"
                amount = abs(signed_amount)
                rows.append([
                    datetime.strptime(date_text, "%Y-%m-%d"), currency, amount, balance,
                    summary, counter_info, customer_summary,
                    cmb_counterparty(customer_summary, counter_info), direction,
                ])

    if not rows:
        raise ValueError("招商银行 PDF 中没有识别到交易记录")
    return rows


def read_boc_pdf_rows(path: Path, pdf_password: str | None = None) -> list[list[Any]]:
    try:
        import pdfplumber
    except ImportError as exc:
        raise ValueError("读取中国银行 PDF 需要 pdfplumber，请运行：python -m pip install pdfplumber") from exc

    rows: list[list[Any]] = []
    with pdfplumber.open(path, password=pdf_password or "") as pdf:
        first_text = (pdf.pages[0].extract_text() or "") if pdf.pages else ""
        if "中国银行交易流水明细清单" not in first_text:
            raise ValueError("目前仅支持带文字层的中国银行交易流水明细清单 PDF")

        for page in pdf.pages:
            page_text = page.extract_text() or ""
            expected_match = re.search(r"行数\s*[:：]\s*(\d+)", page_text)
            expected_rows = int(expected_match.group(1)) if expected_match else None
            page_rows: list[list[Any]] = []
            for table in page.extract_tables():
                if not table or not table[0] or pdf_cell_text(table[0][0]) != "记账日期":
                    continue
                for raw in table[1:]:
                    cells = [pdf_cell_text(value) for value in (list(raw) + [None] * 12)[:12]]
                    date_text, time_text = cells[0], cells[1]
                    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_text):
                        continue
                    if not re.fullmatch(r"\d{2}:\d{2}:\d{2}", time_text):
                        raise ValueError(f"中国银行 PDF 第 {page.page_number} 页存在无法识别的记账时间")

                    currency, amount_text, balance_text = cells[2], cells[3], cells[4]
                    transaction_name, channel, branch_name = cells[5], cells[6], cells[7]
                    remark, counterparty, counter_account, counter_bank = cells[8], cells[9], cells[10], cells[11]
                    signed_amount = parse_amount(amount_text)
                    balance = parse_amount(balance_text)
                    if not isinstance(signed_amount, (int, float)) or not isinstance(balance, (int, float)):
                        raise ValueError(f"中国银行 PDF 第 {page.page_number} 页金额或余额无法识别")

                    direction = "收入" if signed_amount > 0 else "支出" if signed_amount < 0 else "不计收支"
                    page_rows.append([
                        datetime.strptime(f"{date_text} {time_text}", "%Y-%m-%d %H:%M:%S"),
                        currency, abs(signed_amount), balance, transaction_name, channel,
                        branch_name, remark, boc_counterparty(counterparty, remark, transaction_name),
                        counter_account, counter_bank, direction,
                    ])

            if expected_rows is not None and len(page_rows) != expected_rows:
                raise ValueError(
                    f"中国银行 PDF 第 {page.page_number} 页声明 {expected_rows} 笔，实际识别 {len(page_rows)} 笔"
                )
            rows.extend(page_rows)

    if not rows:
        raise ValueError("中国银行 PDF 中没有识别到交易记录")
    return rows


def read_rows(
    path: Path,
    requested_sheet: str | None,
    pdf_password: str | None = None,
) -> tuple[list[list[Any]], str]:
    if path.suffix.lower() == ".pdf":
        pdf, resolved_password = open_pdf_document(path, pdf_password)
        with pdf:
            first_text = (pdf.pages[0].extract_text() or "") if pdf.pages else ""
        if "招商银行交易流水" in first_text:
            return [CMB_PDF_HEADERS, *read_cmb_pdf_rows(path, resolved_password)], "交易流水"
        if "中国银行交易流水明细清单" in first_text:
            return [BOC_PDF_HEADERS, *read_boc_pdf_rows(path, resolved_password)], "交易流水"
        raise ValueError("目前仅支持招商银行或中国银行的带文字层交易流水 PDF")

    if path.suffix.lower() == ".xlsx":
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            if requested_sheet:
                if requested_sheet not in workbook.sheetnames:
                    raise ValueError(f"不存在工作表：{requested_sheet}")
                sheet = workbook[requested_sheet]
            else:
                sheet = max(
                    workbook.worksheets,
                    key=lambda item: max(
                        (header_score(row) for row in item.iter_rows(min_row=1, max_row=min(100, item.max_row), values_only=True)),
                        default=-1,
                    ),
                )
            return [list(row) for row in sheet.iter_rows(values_only=True)], sheet.title
        finally:
            workbook.close()

    if path.suffix.lower() in {".csv", ".tsv"}:
        text = decode_csv(path.read_bytes())
        delimiter = "\t" if path.suffix.lower() == ".tsv" else ","
        if path.suffix.lower() == ".csv":
            try:
                delimiter = csv.Sniffer().sniff(text[:8192], delimiters=",\t;|").delimiter
            except csv.Error:
                pass
        return [list(row) for row in csv.reader(text.splitlines(), delimiter=delimiter)], "账单"

    raise ValueError("仅支持 .xlsx、.csv、.tsv 或招商银行/中国银行交易流水 .pdf")


def parse_datetime(value: Any) -> Any:
    if isinstance(value, datetime) or not isinstance(value, str):
        return value
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(value.strip(), fmt)
        except ValueError:
            continue
    return value.strip()


def parse_amount(value: Any) -> Any:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value
    cleaned = normalize_text(value).replace(",", "").replace("¥", "").replace("￥", "")
    cleaned = cleaned.replace("−", "-").replace("–", "-").replace("—", "-")
    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    return float(match.group(0)) if match else value


def clean_source(
    path: Path,
    requested_sheet: str | None,
    pdf_password: str | None = None,
) -> SourceData:
    raw_rows, sheet_name = read_rows(path, requested_sheet, pdf_password)
    header_index = find_header_row(raw_rows)
    header_values = list(raw_rows[header_index])
    while header_values and not normalize_text(header_values[-1]):
        header_values.pop()
    headers = [str(value).strip() if normalize_text(value) else f"未命名列{index + 1}" for index, value in enumerate(header_values)]
    lookup = header_lookup(headers)
    provider = (
        "中国银行" if path.suffix.lower() == ".pdf" and "对方开户行" in headers
        else "招商银行" if path.suffix.lower() == ".pdf"
        else "支付宝" if "支付宝" in path.name or "交易分类" in headers
        else "微信"
    )
    rows: list[list[Any]] = []
    for raw in raw_rows[header_index + 1:]:
        row = list(raw[:len(headers)]) + [None] * max(0, len(headers) - len(raw))
        if all(not normalize_text(value) for value in row):
            continue
        identity = [lookup[key] for key in ("time", "counterparty", "product", "amount") if lookup[key] >= 0]
        if identity and all(not normalize_text(row[index]) for index in identity):
            continue
        if lookup["time"] >= 0:
            row[lookup["time"]] = parse_datetime(row[lookup["time"]])
        if lookup["amount"] >= 0:
            row[lookup["amount"]] = parse_amount(row[lookup["amount"]])
        if lookup["balance"] >= 0:
            row[lookup["balance"]] = parse_amount(row[lookup["balance"]])
        # 支付宝原始账单把余额宝每日收益标成“不计收支”，但它实际增加资产，
        # 因此在输出清洗后的表格时统一修正为“收入”。
        product_text = normalize_text(row[lookup["product"]]) if lookup["product"] >= 0 else ""
        if (
            provider == "支付宝"
            and lookup["direction"] >= 0
            and "余额宝" in product_text
            and "收益发放" in product_text
        ):
            row[lookup["direction"]] = "收入"
        for key in ("transaction_id", "merchant_id"):
            index = lookup[key]
            if index >= 0 and row[index] is not None:
                row[index] = str(row[index]).strip().lstrip("`")
        rows.append(row)
    if not rows:
        raise ValueError("识别到表头，但没有交易记录")
    return SourceData(headers, lookup, rows, sheet_name, provider, header_index)


def load_config(path: Path) -> Config:
    raw = json.loads(path.read_text(encoding="utf-8-sig"))
    categories = [
        Category(
            label=str(item["label"]).strip(),
            flow=str(item.get("flow") or "expense").strip().lower(),
            keywords=[str(keyword) for keyword in item.get("keywords", [])],
        )
        for item in raw["categories"]
    ]
    labels = [item.label for item in categories]
    if len(labels) != len(set(labels)) or any(not label for label in labels):
        raise ValueError("分类名称不能为空或重复")
    fallbacks = raw.get("fallbackByFlow") or {
        "expense": "其他支出", "income": "其他收入", "transfer": "其他转账"
    }
    for flow in ("expense", "income", "transfer"):
        if fallbacks.get(flow) not in labels:
            raise ValueError(f"缺少 {flow} 对应的兜底分类")
    return Config(categories, fallbacks, raw.get("merchantRules", []))


def load_overrides(path: Path, config: Config) -> dict[str, str]:
    """读取人工维护的“交易对方关键词 → 二级分类”映射。"""
    if not path.is_file():
        return {}
    raw = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(raw, dict):
        raise ValueError("merchant-overrides.json 必须是 JSON 对象")
    overrides: dict[str, str] = {}
    for merchant, category in raw.items():
        keyword = normalize_text(merchant)
        label = str(category).strip()
        if not keyword:
            raise ValueError("merchant-overrides.json 中的交易对方关键词不能为空")
        if label not in config.labels:
            raise ValueError(f"商户覆盖规则引用了不存在的分类：{label}")
        overrides[keyword] = label
    return overrides


def transaction_from_row(row: list[Any], lookup: dict[str, int]) -> dict[str, Any]:
    def get(key: str) -> Any:
        return row[lookup[key]] if lookup[key] >= 0 else None

    return {key: get(key) for key in (
        "time", "type", "counterparty", "account", "product", "direction",
        "amount", "payment", "status", "remark",
    )}


def display_text(value: Any) -> str:
    text = "" if value is None else str(value)
    cleaned = re.sub(r"\s+", " ", text.lstrip("`").replace("\u3000", " ")).strip()
    return "" if re.fullmatch(r"[-/_.—－]+", cleaned) else cleaned


def transaction_description(transaction: dict[str, Any]) -> str:
    """保留对人工复核有用的商品/备注；原始交易类型仅在没有详情时兜底。"""
    details: list[str] = []
    for key in ("product", "remark"):
        value = display_text(transaction.get(key))
        if value and value not in details:
            details.append(value)
    if not details:
        value = display_text(transaction.get("type"))
        if value:
            details.append(value)
    return " / ".join(details)


def transaction_text(transaction: dict[str, Any]) -> str:
    return normalize_text(" ".join(str(transaction.get(key) or "") for key in (
        "type", "counterparty", "account", "product", "direction", "status", "remark"
    )))


def is_internal_change_transfer(transaction: dict[str, Any]) -> bool:
    """微信零钱转入零钱通只是账户内划转，不产生收入或支出。"""
    return (
        normalize_text(transaction.get("type")).replace(" ", "") == "转入零钱通-来自零钱"
        and normalize_text(transaction.get("direction")) in {"", "/", "不计收支"}
    )


def transaction_flow(transaction: dict[str, Any]) -> str:
    """普通交易区分收支；明确的账户内划转记为不计收支。"""
    if is_internal_change_transfer(transaction):
        return "transfer"
    text = transaction_text(transaction)
    if "退款" in text or "退货" in text:
        return "income"
    if any(marker in text for marker in ("收益发放", "利息收入", "结息", "存款利息")):
        return "income"

    # 账单原始方向最可靠；“不计收支”或空值再根据交易语义推断。
    direction = normalize_text(transaction.get("direction"))
    if "收入" in direction or direction in {"收", "+"}:
        return "income"
    if "支出" in direction or direction in {"支", "-"}:
        return "expense"

    income_markers = (
        "转入", "收款", "借入", "收回借款", "收债", "报销", "存入", "存款", "到账",
    )
    if any(marker in text for marker in income_markers):
        return "income"

    expense_markers = (
        "转出", "付款", "还款", "借出", "偿还借款", "垫付", "提现", "取款", "充值",
    )
    if any(marker in text for marker in expense_markers):
        return "expense"

    # 没有明确方向的普通转账按支出给出候选，并以低置信度要求人工复核。
    return "expense"


def classify(transaction: dict[str, Any], config: Config, overrides: dict[str, str]) -> Result:
    text = transaction_text(transaction)
    flow = transaction_flow(transaction)
    if flow == "transfer":
        return Result(config.fallback_by_flow["transfer"], 0.995, "账户内划转")
    eligible = config.categories_for_flow(flow)
    allowed = config.labels_for_flow(flow)

    # 覆盖规则由用户明确维护，优先级最高。既支持完整交易对方，也支持“🏐”、
    # “KUMO KUMO”这样的稳定片段；多个片段同时命中时取最长者。
    counterparty = normalize_text(transaction.get("counterparty"))
    matched_overrides = [
        (keyword, category)
        for keyword, category in overrides.items()
        if category in allowed and (keyword == counterparty or keyword in counterparty)
    ]
    if matched_overrides:
        _, category = max(matched_overrides, key=lambda item: len(item[0]))
        return Result(category, 0.995, "商户覆盖规则")

    if ("退款" in text or "退货" in text) and "退款" in allowed:
        return Result("退款", 0.995, "退款规则")

    for rule in config.merchant_rules:
        keyword = normalize_text(rule.get("contains"))
        if keyword and keyword in text and rule.get("category") in allowed:
            confidence = float(rule.get("confidence", 0.95))
            if "群收款" in text:
                confidence = min(confidence, 0.58)
            return Result(str(rule["category"]), confidence, "商户规则")

    source_category = normalize_text(transaction.get("type"))
    hint = SOURCE_CATEGORY_HINTS.get(source_category)
    if hint and hint not in allowed:
        hint = None
    scored: list[tuple[float, Category, list[str]]] = []
    for category in eligible:
        keywords = [*category.keywords, *EXTRA_KEYWORDS.get(category.label, [])]
        matches = [keyword for keyword in keywords if normalize_text(keyword) and normalize_text(keyword) in text]
        score = sum(min(4.0, 1.5 + len(normalize_text(keyword)) * 0.25) for keyword in matches)
        if hint == category.label:
            score += 2.4
        scored.append((score, category, matches))
    scored.sort(key=lambda item: item[0], reverse=True)
    top_score, top_category, matches = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0.0
    second_matches = scored[1][2] if len(scored) > 1 else []

    if top_score == 0:
        return Result(config.fallback_by_flow[flow], 0.25, "无关键词")

    confidence = min(0.96, 0.80 + top_score * 0.035)
    top_specificity = max((len(normalize_text(keyword)) for keyword in matches), default=0)
    second_specificity = max((len(normalize_text(keyword)) for keyword in second_matches), default=0)
    if second_score > 0 and top_score - second_score < 1.2 and top_specificity < second_specificity + 2:
        confidence -= 0.18
    if "群收款" in text or "二维码收款" in text:
        confidence = min(confidence, 0.58)
    if hint and not matches:
        confidence = min(confidence, 0.80)
    return Result(top_category.label, max(0.30, confidence), "关键词规则")


def excel_column(index: int) -> str:
    value = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        value = chr(65 + remainder) + value
    return value


def build_workbook(source: SourceData, results: list[Result], threshold: float) -> Workbook:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "分类结果"
    sheet.append(OUTPUT_HEADERS)
    for row_number, (source_row, result) in enumerate(zip(source.rows, results), start=2):
        transaction = transaction_from_row(source_row, source.lookup)
        amount = parse_amount(transaction.get("amount"))
        if not isinstance(amount, (int, float)):
            raise ValueError(f"第 {row_number} 笔交易金额无法识别")
        flow = transaction_flow(transaction)
        sheet.append([
            transaction.get("time"),
            {"income": "收入", "expense": "支出", "transfer": "不计收支"}[flow],
            abs(amount),
            display_text(transaction.get("counterparty")),
            transaction_description(transaction),
            result.category,
            "是" if result.confidence < threshold else "否",
        ])

    header_fill = PatternFill("solid", fgColor="0F766E")
    header_font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF")
    body_font = Font(name="Microsoft YaHei", size=10)
    bottom_border = Border(bottom=Side(style="thin", color="E2E8F0"))
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.row_dimensions[1].height = 30
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font
            cell.alignment = Alignment(vertical="center")
            cell.border = bottom_border
        sheet.row_dimensions[row[0].row].height = 38

    for row in range(2, sheet.max_row + 1):
        sheet.cell(row, 1).number_format = (
            "yyyy-mm-dd" if source.provider == "招商银行" else "yyyy-mm-dd hh:mm:ss"
        )
        sheet.cell(row, 2).alignment = Alignment(horizontal="center", vertical="center")
        sheet.cell(row, 3).number_format = "#,##0.00"
        sheet.cell(row, 3).alignment = Alignment(horizontal="right", vertical="center")
        for column in (4, 5):
            sheet.cell(row, column).alignment = Alignment(vertical="center", wrap_text=True)
        for column in (6, 7):
            sheet.cell(row, column).alignment = Alignment(horizontal="center", vertical="center")

    for column, width in {
        "A": 22, "B": 11, "C": 13, "D": 30, "E": 40, "F": 18, "G": 20,
    }.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    table = Table(displayName="LightClassificationResults", ref=f"A1:{excel_column(sheet.max_column)}{sheet.max_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    sheet.add_table(table)
    return workbook


def verify_output(path: Path, source: SourceData, labels: set[str]) -> None:
    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        if workbook.sheetnames != ["分类结果"]:
            raise RuntimeError("输出应且仅应包含“分类结果”工作表")
        sheet = workbook["分类结果"]
        if sheet.max_row != len(source.rows) + 1 or sheet.max_column != len(OUTPUT_HEADERS):
            raise RuntimeError("输出行列数与输入不一致")
        if [sheet.cell(1, column).value for column in range(1, sheet.max_column + 1)] != OUTPUT_HEADERS:
            raise RuntimeError("输出列标题或顺序不正确")
        for row in range(2, sheet.max_row + 1):
            if sheet.cell(row, 2).value not in {"收入", "支出", "不计收支"}:
                raise RuntimeError(f"第 {row} 行交易类型无效")
            amount = sheet.cell(row, 3).value
            if not isinstance(amount, (int, float)) or amount < 0:
                raise RuntimeError(f"第 {row} 行金额不是非负数值")
            if sheet.cell(row, 6).value not in labels:
                raise RuntimeError(f"第 {row} 行分类不在清单中")
            if sheet.cell(row, 7).value not in {"是", "否"}:
                raise RuntimeError(f"第 {row} 行人工确认标记错误")
    finally:
        workbook.close()


def default_output_path(input_path: Path, source: SourceData) -> Path:
    output_dir = input_path.parent / "outputs"
    if source.provider == "中国银行" and source.lookup["time"] >= 0:
        dates = [
            row[source.lookup["time"]]
            for row in source.rows
            if isinstance(row[source.lookup["time"]], datetime)
        ]
        if dates:
            start_date = min(dates).strftime("%Y%m%d")
            end_date = max(dates).strftime("%Y%m%d")
            return output_dir / f"中国银行交易流水({start_date}-{end_date})_已分类_轻量版.xlsx"
    return output_dir / f"{input_path.stem}_已分类_轻量版.xlsx"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="微信、支付宝、招商银行和中国银行账单轻量分类器（无大模型）")
    parser.add_argument(
        "input", type=Path,
        help="微信 .xlsx、支付宝 .csv 或招商银行/中国银行交易流水 .pdf 文件",
    )
    parser.add_argument("--output", type=Path, help="输出 .xlsx 路径")
    parser.add_argument("--categories", type=Path, default=DEFAULT_CATEGORIES, help="分类配置 JSON")
    parser.add_argument("--overrides", type=Path, default=DEFAULT_OVERRIDES, help="商户覆盖规则 JSON")
    parser.add_argument("--sheet", help="指定 Excel 工作表")
    parser.add_argument(
        "--pdf-password",
        help=f"PDF 打开密码；也可使用环境变量 {PDF_PASSWORD_ENV}，均不会写入输出文件",
    )
    parser.add_argument("--threshold", type=float, default=0.82, help="需要人工确认的置信度阈值，默认 0.82")
    args = parser.parse_args()
    if not 0 <= args.threshold <= 1:
        parser.error("--threshold 必须在 0 到 1 之间")
    return args


def main() -> int:
    args = parse_args()
    input_path = args.input.resolve()
    if not input_path.is_file():
        raise FileNotFoundError(f"输入文件不存在：{input_path}")
    config = load_config(args.categories.resolve())
    overrides = load_overrides(args.overrides.resolve(), config)
    source = clean_source(input_path, args.sheet, args.pdf_password)
    output_path = (args.output or default_output_path(input_path, source)).resolve()
    if output_path == input_path:
        raise ValueError("输出路径不能与原始账单相同，请选择另一个 .xlsx 文件")
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("输出文件必须使用 .xlsx 扩展名")
    results = [classify(transaction_from_row(row, source.lookup), config, overrides) for row in source.rows]
    workbook = build_workbook(source, results, args.threshold)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()
    verify_output(output_path, source, set(config.labels))
    pending = sum(result.confidence < args.threshold for result in results)
    source_note = (
        "解析 PDF 固定栏位"
        if source.provider in {"招商银行", "中国银行"}
        else f"删除表头前内容 {source.removed_leading_rows} 行"
    )
    print(f"完成：{output_path}\n来源：{source.provider}；{source_note}；交易 {len(results)} 笔；需要人工确认 {pending} 笔。")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        raise SystemExit(1) from exc
