# 月末费用表复核

运营同事打开 benefits.xlsx，看到的费用合计是 600。请核对合计是否能由明细重算，找出所有直接引用自身的公式单元格，并计算三笔费用的正确合计。

在 output/answer.json 提交一个 JSON 对象，包含 circular_cells（形如 Sheet1!A1 的地址列表，无重复，顺序不限）和 corrected_total（数值，单位元）。没有自引用时提交空列表。
