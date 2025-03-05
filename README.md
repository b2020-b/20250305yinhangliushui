# 20250305yinhangliushui   Python 处理 Excel中银行流水 数据并合并






### **需求分析**
有一个 Excel 文件 `流水处理_示例.xlsx`，其中 `Sheet1` 包含以下数据：

| 银行参考号 | 客户参考号 | TRN 类型 | 生效日期 | 汇入金额 | 汇出金额 | 余额 | 时间 | 过账日期 |
|-----------|-----------|---------|---------|---------|---------|------|------|------|
| NON       | NON1      | S++     | 04 r 2025 | 4,028.29 |         | 2,649.73 | 18:52 | 04 r 2025 |
| 交易附言  | CREDIT, Pay |         |         |         |         |       |      |      |
| NO        | NREF      | P+      | 04 r 2025 | 1,854.28 |         | 24,951.44 | 14:47 | 04 r 2025 |
| 交易附言  | CREDIT, Pay |         |         |         |         |       |      |      |

可以看到，“交易附言”行并没有完整的交易信息，我们的目标是将“交易附言”合并到上一行的“交易附言”列，并删除原始的“交易附言”行。

### **实现步骤**
我们使用 Pandas 读取 `Sheet1`，并进行如下处理：

1. **读取 Excel 文件**
2. **去除自动生成的空白列（`Unnamed` 开头的列）**
3. **识别“交易附言”行**
4. **将“交易附言”内容合并到前一行**
5. **删除“交易附言”行**
6. **将处理后的数据保存到 `Sheet2`**

### **Python 实现代码**
```python
import pandas as pd

# 读取 Excel 文件
file_path = "流水处理_示例.xlsx"
df = pd.read_excel(file_path, sheet_name="sheet1", dtype=str, engine="openpyxl")

# 去除空白列
df = df.loc[:, ~df.columns.str.contains("^Unnamed")]  # 移除列名为 Unnamed 的列

# 确保列名正确去除空白
df.columns = df.columns.str.strip()

# 识别“交易附言”行（即第一列为“交易附言”）
mask = df.iloc[:, 0] == "交易附言"

# 前一行加入“交易附言”内容
df.loc[df.shift(-1).iloc[:, 0] == "交易附言", "交易附言"] = df.loc[mask, df.columns[1]].values

# 删除“交易附言”行
df = df[~mask]

# 读取原 Excel 文件，保留其他 Sheet
with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name="Sheet2", index=False)

print(f"处理完成，结果保存在 {file_path} 的 Sheet2")
```

### **代码解析**
1. **读取 Excel 数据**：`pd.read_excel()` 读取 `Sheet1` 的数据。
2. **去除无效列**：`df.loc[:, ~df.columns.str.contains("^Unnamed")]` 移除 Pandas 自动生成的空列。
3. **识别“交易附言”行**：使用 `mask = df.iloc[:, 0] == "交易附言"` 找出交易附言行。
4. **合并“交易附言”**：利用 `df.shift(-1)` 检测是否下一行为交易附言，并将其内容赋值给前一行。
5. **删除原始“交易附言”行**：`df = df[~mask]` 过滤掉这些无用行。
6. **保存到 `Sheet2`**：使用 `ExcelWriter` 以追加模式写入 `Sheet2`。

### **运行结果**
处理后，数据将被保存到 `Sheet2`，格式如下：

| 银行参考号 | 客户参考号 | TRN 类型 | 生效日期 | 汇入金额 | 汇出金额 | 余额 | 时间 | 过账日期 | 交易附言 |
|-----------|-----------|---------|---------|---------|---------|------|------|------|----------|
| NON       | NON1      | S++     | 04 r 2025 | 4,028.29 |         | 2,649.73 | 18:52 | 04 r 2025 | CREDIT, Pay |
| NO        | NREF      | P+      | 04 r 2025 | 1,854.28 |         | 24,951.44 | 14:47 | 04 r 2025 | CREDIT, Pay |

### **总结**
通过 Pandas，我们能够高效地处理 Excel 数据，完成合并、清理和格式化操作。这个方法适用于各种 Excel 数据清理场景，例如财务报表整理、数据预处理等。希望本文对你有所帮助！
