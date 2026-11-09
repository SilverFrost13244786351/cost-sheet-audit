# -*- coding: utf-8 -*-
"""造一份带错的样例清单，用来测试核对工具。

故意埋的错（共 14 处）：
  合价算错   5 处
  重复列项   3 处
  单价异常   4 处
  单位缺失   2 处

答案会单独写在 data/答案.txt 里，不放在表格的备注列（真实清单不会自己告诉你哪里错了）。

用法：python data/make_sample.py
输出：data/清单样例.xlsx、data/答案.txt
"""
import os
import random
from openpyxl import Workbook

random.seed(20261102)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_XLSX = os.path.join(HERE, "清单样例.xlsx")
OUT_KEY = os.path.join(HERE, "答案.txt")

# 真实一点的清单项：(编码前缀, 名称, 单位, 单价区间)
ITEMS = [
    ("010101001", "土方开挖", "m3", (22, 34)),
    ("010201001", "边坡支护", "m2", (95, 135)),
    ("010501001", "喷射混凝土", "m3", (520, 620)),
    ("010515001", "钢筋制安", "t", (4600, 5300)),
    ("011702001", "模板工程", "m2", (42, 62)),
    ("010502001", "现浇混凝土", "m3", (470, 560)),
    ("010903001", "防水卷材", "m2", (28, 48)),
    ("040203001", "路面恢复", "m2", (85, 120)),
    ("040504001", "检查井砌筑", "座", (1200, 1800)),
    ("040501001", "管道敷设", "m", (160, 240)),
]


def main():
    rows = []          # [序号, 编码, 名称, 单位, 工程量, 单价, 合价, 备注]
    answer = []        # 答案行：类型, 行号, 说明
    seq = 1

    for _ in range(196):
        pre, name, unit, (lo, hi) = random.choice(ITEMS)
        code = f"{pre}{seq:04d}"   # 按行唯一，避免随机撞码造成的假重复
        qty = round(random.uniform(5, 500), 2)
        price = round(random.uniform(lo, hi), 2)
        rows.append([seq, code, name, unit, qty, price, round(qty * price, 2), ""])
        seq += 1

    idx_all = list(range(len(rows)))
    used = set()

    # ---- 埋错 1：合价算错 5 处 ----
    bad_amount = random.sample([i for i in idx_all if i not in used], 5)
    for i in bad_amount:
        used.add(i)
        rows[i][6] = round(rows[i][6] * random.choice([0.9, 1.1, 1.25]), 2)
        answer.append(f"合价错   第 {i + 2} 行（Excel 行号）  {rows[i][2]} 合价应为 {round(rows[i][4] * rows[i][5], 2)}")

    # ---- 埋错 2：重复列项 3 处（抄一行另起序号） ----
    dup_src = random.sample([i for i in idx_all if i not in used], 3)
    for i in dup_src:
        used.add(i)
        dup = list(rows[i])
        dup[0] = seq
        seq += 1
        rows.append(dup)
        answer.append(f"重复项   第 {len(rows) + 1} 行（Excel 行号）  与第 {i + 2} 行重复：{dup[2]} / {dup[1]}")

    # ---- 埋错 3：单价异常 4 处（同项价格翻 3 倍，合价同步重算） ----
    bad_price = random.sample([i for i in idx_all if i not in used], 4)
    for i in bad_price:
        used.add(i)
        rows[i][5] = round(rows[i][5] * 3, 2)
        rows[i][6] = round(rows[i][4] * rows[i][5], 2)
        answer.append(f"单价异常 第 {i + 2} 行（Excel 行号）  {rows[i][2]} 单价 {rows[i][5]}（同类约 {round(rows[i][5] / 3, 2)}）")

    # ---- 埋错 4：单位缺失 2 处 ----
    bad_unit = random.sample([i for i in idx_all if i not in used], 2)
    for i in bad_unit:
        used.add(i)
        rows[i][3] = ""
        answer.append(f"单位缺失 第 {i + 2} 行（Excel 行号）  {rows[i][2]} 计量单位为空")

    total = round(sum(r[6] for r in rows), 2)

    wb = Workbook()
    ws = wb.active
    ws.title = "分部分项清单"
    ws.append(["序号", "项目编码", "项目名称", "计量单位", "工程量", "综合单价", "合价", "备注"])
    for r in rows:
        ws.append(r)
    ws.append(["合计", "", "", "", "", "", total, ""])

    for col, width in zip("ABCDEFGH", [8, 16, 16, 10, 12, 12, 14, 18]):
        ws.column_dimensions[col].width = width

    wb.save(OUT_XLSX)

    header = [
        f"样例清单：{OUT_XLSX}",
        f"明细 {len(rows)} 行，表内合计 {total}",
        f"埋错 {len(answer)} 处：",
        "",
    ]
    open(OUT_KEY, "w", encoding="utf-8").write("\n".join(header + answer))
    print("\n".join(header))
    for a in answer:
        print("  " + a)


if __name__ == "__main__":
    main()

