import pandas as pd

# 读取 Excel 文件
file_path = "流水处理_示例.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1", dtype=str, engine="openpyxl")

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

# 将结果写入Excel的新Sheet
# 读取原 Excel 文件，保留其他 Sheet
with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name="Sheet2", index=False)

print(f"处理完成，结果保存在 {file_path} 的 Sheet2")
# # 保存到新的 Excel
# output_path = "流水处理_结果.xlsx"
# df.to_excel(output_path, index=False, engine="openpyxl")

# print(f"处理完成，结果保存在 {output_path}")
