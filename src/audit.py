# -*- coding: utf-8 -*-
"""分部分项清单核对工具 v0.1

已实现：
  1. 读表
  2. 字段校验（缺列）
  3. 合价校验：工程量 × 综合单价 是否等于 合价
  4. 合计校验：明细求和 是否等于 表内合计行
  5. 输出 Excel 报告 + 控制台摘要

待实现（你来做，见 README）：
  6. 重复项检测
  7. 单价异常检测

用法：python src/audit.py
"""
import os
import sys

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "data", "清单样例.xlsx")
OUT = os.path.join(ROOT, "out", "核对报告.xlsx")

REQUIRED_COLS = ["序号", "项目编码", "项目名称", "计量单位", "工程量", "综合单价", "合价"]
TOL = 0.01  # 金额容差，分


# ---------------------------------------------------------------- 1. 读表
def load_statement(path):
    """读取清单，拆成「明细」和「合计」两部分。

    真实清单最后一行通常是「合计」，它没有工程量，要单独拿出来，
    否则会污染后面的计算。
    """
    df = pd.read_excel(path, sheet_name=0, dtype={"项目编码": str})  # 编码必须当文本读，否则前导零丢失
    df["序号"] = df["序号"].astype(str).str.strip()

    is_total = df["序号"].str.startswith("合计")
    detail = df[~is_total].copy()
    total_row = df[is_total]

    for c in ["工程量", "综合单价", "合价"]:
        detail[c] = pd.to_numeric(detail[c], errors="coerce")
    return detail, total_row


# ---------------------------------------------------------------- 2. 字段校验
def check_columns(df):
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    problems = []
    if missing:
        problems.append({"问题": "缺少必需列", "详情": "、".join(missing)})
    return pd.DataFrame(problems, columns=["问题", "详情"])


# ---------------------------------------------------------------- 3. 合价校验
def check_amount(df):
    """合价 应等于 工程量 × 综合单价。"""
    d = df.copy()
    d["应为合价"] = (d["工程量"] * d["综合单价"]).round(2)
    d["差额"] = (d["合价"] - d["应为合价"]).round(2)
    bad = d[d["差额"].abs() > TOL]
    return bad[["序号", "项目编码", "项目名称", "计量单位", "工程量", "综合单价", "合价", "应为合价", "差额"]]


# ---------------------------------------------------------------- 4. 合计校验
def check_total(df, total_row):
    detail_sum = round(df["合价"].sum(), 2)
    if total_row.empty:
        return pd.DataFrame([{"项目": "表内合计", "数值": None, "问题": "清单里没有合计行"}]), detail_sum
    declared = pd.to_numeric(total_row["合价"], errors="coerce").iloc[0]
    diff = round(detail_sum - declared, 2) if pd.notna(declared) else None
    problem = "" if diff is not None and abs(diff) <= TOL else "明细求和与表内合计不一致"
    return pd.DataFrame([{
        "项目": "表内合计", "数值": declared,
    }, {
        "项目": "明细求和", "数值": detail_sum,
    }, {
        "项目": "差额", "数值": diff, "问题": problem,
    }]), detail_sum


# ---------------------------------------------------------------- 5. 缺值检查
def check_blank(df):
    bad = df[df["计量单位"].isna() | (df["计量单位"].astype(str).str.strip() == "")]
    return bad[["序号", "项目编码", "项目名称", "工程量", "综合单价", "合价"]].assign(问题="计量单位为空")


# ---------------------------------------------------------------- 6. 重复项
def check_duplicate(df):
    """同一「项目编码 + 项目名称 + 计量单位」出现多次，判为重复列项。

    keep=False：一组里所有成员都标出来（6 行 = 3 对，两行都报）。
    重复组列：把同一组的所有序号拼在一起，看出“跟谁撞了”。
    """
    key = ["项目编码", "项目名称", "计量单位"]
    dup = df[df.duplicated(subset=key, keep=False)]
    g = df.groupby(key)["序号"].transform(lambda s: "、".join(map(str, s)))
    return dup.assign(重复组=g)[["序号", "项目编码", "项目名称", "计量单位",
                                 "工程量", "合价", "重复组"]]


# ---------------------------------------------------------------- 7. 单价异常（TODO）
def check_price_outlier(df, ratio=0.3):
    """TODO（D6–D7）：同一「项目名称 + 计量单位」的单价，偏离中位数 ±30% 以上判为异常。

    提示：
      - 先按 [项目名称, 计量单位] 分组，算 price 的中位数
      - 再算每行偏离比例，筛出 |偏离| > ratio 的行
    现在还没实现，先返回空表。
    """
    return pd.DataFrame(columns=["序号", "项目名称", "计量单位", "综合单价", "同类中位数", "偏离比例"])


# ---------------------------------------------------------------- 主流程
def main():
    if not os.path.exists(SRC):
        sys.exit(f"找不到样例文件：{SRC}\n先跑：python data/make_sample.py")

    detail, total_row = load_statement(SRC)

    col_problems = check_columns(detail)
    bad_amount = check_amount(detail)
    total_check, detail_sum = check_total(detail, total_row)
    blank_unit = check_blank(detail)
    dup = check_duplicate(detail)
    outliers = check_price_outlier(detail)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as w:
        summary = pd.DataFrame([
            {"检查项": "明细行数", "结果": len(detail)},
            {"检查项": "明细合价求和", "结果": detail_sum},
            {"检查项": "字段缺失", "结果": len(col_problems)},
            {"检查项": "合价与量价不符", "结果": len(bad_amount)},
            {"检查项": "计量单位为空", "结果": len(blank_unit)},
            {"检查项": "重复列项", "结果": len(dup)},
            {"检查项": "单价异常", "结果": len(outliers)},
        ])
        summary.to_excel(w, sheet_name="汇总", index=False)
        bad_amount.to_excel(w, sheet_name="合价异常", index=False)
        blank_unit.to_excel(w, sheet_name="单位缺失", index=False)
        total_check.to_excel(w, sheet_name="合计校验", index=False)
        if len(dup):
            dup.to_excel(w, sheet_name="重复项", index=False)
        if len(outliers):
            outliers.to_excel(w, sheet_name="单价异常", index=False)

    print("=" * 46)
    print(f"明细 {len(detail)} 行，合价求和 {detail_sum}")
    print(f"合价与量价不符：{len(bad_amount)} 行")
    print(f"计量单位为空  ：{len(blank_unit)} 行")
    print(f"重复列项      ：{len(dup)} 行")
    print(f"单价异常      ：{len(outliers)} 行")
    print(f"报告已生成    ：{OUT}")
    print("=" * 46)


if __name__ == "__main__":
    main()
