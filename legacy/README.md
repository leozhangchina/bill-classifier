# 旧版可选大模型分类器

此目录保留 `classify_bill.py`，用于将来的大模型功能需求。当前项目主入口为根目录的 `classify_bill_light.py`，日常运行无需此目录或 API 密钥。

旧版支持 Excel / CSV / TSV，输出保留原始列并新增“最终分类”“是否大模型判断”。它与轻量版的七列输出格式不同，也不支持银行 PDF。默认仍读取项目根目录的分类和商户规则。

在项目根目录运行：

```powershell
.\.venv\Scripts\python.exe .\legacy\classify_bill.py "D:\账单\账单.xlsx" --llm off --no-prompt
```

可选参数包括 `--llm auto|off|required`、`--no-prompt`、`--learn`。启用大模型时设置 `OPENAI_API_KEY` 和 `OPENAI_MODEL`，需要时再设置 `OPENAI_BASE_URL`、`OPENAI_API_STYLE=responses|chat`。`.env.example` 只列出变量示例，脚本不会自动加载 `.env` 文件。

旧版 `--learn` 会把人工确认后的商户规则写回根目录 `merchant-overrides.json`。当前轻量脚本只读取这份规则，不调用 API，也不自动学习。
