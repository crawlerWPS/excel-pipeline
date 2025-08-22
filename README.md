# Flask Excel 处理流水线（A/B/C/D/E + Pivot + 模板下载）

- 首页提供 5 个输入模板下载（内存生成）。
- 上传 5 个 Excel 后，依次生成：A_对公聚合、B_利率互换聚合、C_汇总、D_对私映射、E_合并结果、Pivot_按支行汇总。

## 运行
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
python app.py
# http://127.0.0.1:5000
```
