# -*- coding: utf-8 -*-
"""清单核对工具 · 网页界面（Streamlit 版）

用法：
    python -m streamlit run src/app.py
或双击项目根目录的 启动界面.bat

界面只做三件事：
    1. 收一个 Excel
    2. 调 audit.py 里现成的六个检查函数
    3. 把结果摊在页面上，并给一个"下载报告"的按钮
"""
import io
import os
import sys

import pandas as pd
import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)          # 保证能 import 到同目录的 audit.py

from audit import (              # noqa: E402
    TOL,
    check_amount,
    check_blank,
    check_columns,
    check_duplicate,
    check_price_outlier,
    check_total,
    load_statement,
)

st.set_page_config(page_title="清单核对工具", page_icon="📋", layout="wide")

st.title("📋 分部分项清单核对工具")
st.caption("上传一份清单 Excel，自动查出：合价不符、单位缺失、重复列项、单价异常。")

# ---------------------------------------------------------------- 上传
uploaded = st.file_uploader("把清单拖进来（.xlsx）", type=["xlsx"])

if not uploaded:
    st.info("还没收到文件。先在下面看看它能查出什么，再上传你自己那份。")
    st.markdown(
        """
        **它会查四类问题**

        1. **合价不符** — 合价 ≠ 工程量 × 综合单价（容差 1 分钱）
        2. **单位缺失** — 计量单位是空的
        3. **重复列项** — 同项目编码 + 名称 + 单位出现多次
        4. **单价异常** — 同名称同单位的单价，偏离中位数 50% 以上
        """
    )
    st.stop()

# ---------------------------------------------------------------- 读表
try:
    detail, total_row = load_statement(io.BytesIO(uploaded.getvalue()))
except Exception as e:
    st.error(f"读表失败：{e}")
    st.stop()

if detail.empty:
    st.warning("这份表里没有明细行，检查一下是不是选错了文件。")
    st.stop()

# ---------------------------------------------------------------- 跑检查
col_problems = check_columns(detail)
bad_amount = check_amount(detail)
total_check, detail_sum = check_total(detail, total_row)
blank_unit = check_blank(detail)
dup = check_duplicate(detail)
outliers = check_price_outlier(detail)

# ---------------------------------------------------------------- 概览
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("明细行数", len(detail))
c2.metric("合价不符", len(bad_amount))
c3.metric("单位缺失", len(blank_unit))
c4.metric("重复列项", len(dup))
c5.metric("单价异常", len(outliers))

st.caption(f"明细合价求和：{detail_sum:,.2f}")

if len(bad_amount) + len(blank_unit) + len(dup) + len(outliers) == 0:
    st.success("这份表没查出问题。")
else:
    st.warning("查出了一些问题，逐项看下面。")

# ---------------------------------------------------------------- 明细
tabs = st.tabs(["合价不符", "单位缺失", "重复项", "单价异常", "合计校验"])

with tabs[0]:
    if len(bad_amount):
        st.dataframe(bad_amount, use_container_width=True, hide_index=True)
    else:
        st.success("没有合价不符的行。")

with tabs[1]:
    if len(blank_unit):
        st.dataframe(blank_unit, use_container_width=True, hide_index=True)
    else:
        st.success("没有单位缺失的行。")

with tabs[2]:
    if len(dup):
        st.dataframe(dup, use_container_width=True, hide_index=True)
    else:
        st.success("没有重复列项。")

with tabs[3]:
    if len(outliers):
        st.dataframe(outliers, use_container_width=True, hide_index=True)
    else:
        st.success("没有单价异常。")

with tabs[4]:
    st.dataframe(total_check, use_container_width=True, hide_index=True)
    st.caption(f"判定容差：{TOL} 元")

# ---------------------------------------------------------------- 下载
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as w:
    pd.DataFrame([
        {"检查项": "明细行数", "结果": len(detail)},
        {"检查项": "明细合价求和", "结果": detail_sum},
        {"检查项": "字段缺失", "结果": len(col_problems)},
        {"检查项": "合价与量价不符", "结果": len(bad_amount)},
        {"检查项": "计量单位为空", "结果": len(blank_unit)},
        {"检查项": "重复列项", "结果": len(dup)},
        {"检查项": "单价异常", "结果": len(outliers)},
    ]).to_excel(w, sheet_name="汇总", index=False)
    bad_amount.to_excel(w, sheet_name="合价异常", index=False)
    blank_unit.to_excel(w, sheet_name="单位缺失", index=False)
    total_check.to_excel(w, sheet_name="合计校验", index=False)
    if len(dup):
        dup.to_excel(w, sheet_name="重复项", index=False)
    if len(outliers):
        outliers.to_excel(w, sheet_name="单价异常", index=False)

st.download_button(
    "⬇️ 下载 Excel 报告",
    data=buf.getvalue(),
    file_name="核对报告.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
