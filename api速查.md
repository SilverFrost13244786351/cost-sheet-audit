# pandas / os 速查（项目 01 专用）

> 只收这个项目真用得上的。忘了就翻这里，别再翻教程。

---

## 0. 先学会「查」，别背

```python
dir(对象)              # 它身上有哪些名字
help(对象.方法)         # 这个方法怎么用、参数是什么
df.dtypes              # 每列是什么类型
df.head()              # 前 5 行
df.shape               # (行数, 列数)
```

VS Code 里敲 `.` 会弹补全——**那就是全部可用的名字**，比背靠谱。

---

## 1. 读表 / 写表

| 写法 | 作用 |
|---|---|
| `pd.read_excel(路径, sheet_name=0)` | 读第 1 个工作表（从 0 数） |
| `pd.read_excel(路径, dtype={"项目编码": str})` | 指定某列按文本读（保住前导零） |
| `pd.to_numeric(列, errors="coerce")` | 转数字，转不了的变 NaN |
| `pd.DataFrame(记录列表, columns=["问题","详情"])` | 用列表造表 |
| `with pd.ExcelWriter(路径, engine="openpyxl") as w:` | 一次写多页 |
| `df.to_excel(w, sheet_name="汇总", index=False)` | 写一页；`index=False` 不写行号 |

**易错**：编码、电话、身份证这类列必须指定 `dtype=str`，否则 `010501` 会变成 `10501`。

---

## 2. 取数据

```python
df["合价"]                 # 取一列（Series）
df[["序号","合价"]]        # 取多列（DataFrame，注意双层方括号）
df[df["合价"] > 1000]      # 按条件筛行
df[~df["合计"]]            # 加 ~ 取反
df.iloc[0]                 # 第 0 行（按位置）
df.loc[5]                  # 索引名是 5 的那行（按标签）
列 = df["合价"].values     # 转成普通数组
```

---

## 3. 改数据（都会返回新表，不改原表）

| 写法 | 作用 |
|---|---|
| `df.copy()` | 复制一份，后面要改就先复制 |
| `df.assign(新列=值)` | 加一列 |
| `df.rename(columns={"旧":"新"})` | 改列名 |
| `df.drop(columns=["x"])` | 删列 |
| `df.sort_values("合价", ascending=False)` | 排序 |
| `df.reset_index(drop=True)` | 重排索引 |

---

## 4. 类型与空值

| 写法 | 作用 |
|---|---|
| `s.astype(str)` | 整列转文本 |
| `s.isna()` / `s.notna()` | 是不是空（文本列也管用） |
| `s.fillna(0)` | 空值填 0 |
| `s.dropna()` | 丢掉空值所在行 |

**记住**：`NaN` 参与比较永远 `False`，所以用 `> 0.01` 筛错行时，空值会被**静默漏掉**。空值要单独查一遍。

---

## 5. 统计

```python
s.sum()        # 求和（空值自动跳过）
s.mean()       # 平均
s.median()     # 中位数
s.min() / s.max()
s.round(2)     # 保留两位
s.abs()        # 绝对值
s.value_counts()   # 每个值出现几次
s.describe()   # 一把梭的统计摘要
```

---

## 6. 字符串（前面加 `.str`）

```python
s.str.strip()               # 去前后空格
s.str.startswith("合计")     # 以什么开头
s.str.contains("混凝土")     # 包含
s.str.replace("万", "")      # 替换
s.str.len()                 # 长度
s.str.upper()               # 转大写
```

---

## 7. 去重 / 分组

```python
df.duplicated(subset=["项目编码","项目名称"], keep=False)   # 所有重复行都标出来
df.drop_duplicates()                                        # 真正删掉重复

df.groupby(["项目名称","计量单位"])["综合单价"].median()      # 每组一个值
df.groupby(键)["序号"].transform(lambda s: "、".join(s))      # 每组结果铺回每行
df.groupby(键).size()                                       # 每组几行
```

**duplicated 的 keep**：`False` = 一组里全部标出；`"first"` = 只标后面重复的；`"last"` = 只标前面重复的。

---

## 8. 路径（os）

```python
os.path.abspath(__file__)      # 本文件绝对路径
os.path.dirname(路径)           # 上一层目录
os.path.join(a, "data", "x.xlsx")   # 拼路径（自动用对分隔符）
os.path.exists(路径)            # 存不存在
os.makedirs(目录, exist_ok=True)    # 建目录，已存在不报错
```

---

## 9. 报错对照表

| 报错 | 真实含义 | 先查什么 |
|---|---|---|
| `ModuleNotFoundError: No module named 'pandas'` | 这个 Python 里没装 | 装库，或换解释器 |
| `KeyError: '合价'` | 表里没有这一列 | 打印 `df.columns` 看真实列名 |
| `AttributeError: 'DataFrame' object has no attribute 'x'` | 名字写错，或对象的类型不对 | `dir()` 查，或 `type()` 看类型 |
| `IndexError: single positional indexer is out-of-bounds` | 想取第 0 个，但它是空的 | 先判 `.empty` |
| `FileNotFoundError` | 路径不对 | 打印路径，用 `os.path.exists` 验 |
| `PermissionError` | 文件被 Excel 占着 | 关掉 Excel 再跑 |
| `SettingWithCopyWarning` | 你在改一个切片 | 先 `.copy()` |

---

## 10. 本项目用到的地方

| 名字 | 位置 |
|---|---|
| `read_excel` / `dtype` | `audit.py:38` |
| `astype` / `str.strip` | `audit.py:39` |
| `str.startswith` | `audit.py:41` |
| `.copy()` | `42`、`63` |
| `to_numeric` | `47`、`75` |
| `sum` / `round` / `abs` | `72`、`64`、`66` |
| `.empty` / `.iloc[0]` | `73`、`75` |
| `isna` | `89` |
| `duplicated` / `groupby.transform` / `assign` | `103–106` |
| `ExcelWriter` / `to_excel` | `136–153` |

---

## 11. Git（版本控制）

Git 的所有命令都认**当前目录**，敲之前先看提示符在哪。不在仓库里会报 atal: not a git repository。

| 命令 | 干什么 |
|---|---|
| `git init` | 把当前目录变成仓库（只需一次） |
| `git config user.name "名字"` / `git config user.email "邮箱"` | 第一次提交必须设，只设一次 |
| `git status` | 看哪些文件改了、哪些待提交 |
| `git add .` | 把改动放进暂存区（准备提交） |
| `git commit -m "说明"` | 存一个版本，带一句说明 |
| `git log --oneline` | 看提交历史，一行一个版本 |
| `git diff` | 看具体改了哪几行 |

**常见坑**

- `fatal: not a git repository` → 目录不对，先 `cd` 到项目根
- `warning: LF will be replaced by CRLF` → 换行符提示，忽略
- `Please tell me who you are` → 那两条 `git config` 没敲
- 提交前一定跑 `git status`，**别把 out/、*.xlsx 提交进去**（`.gitignore` 负责挡）

**提交信息怎么写**：一句话说清“这次做了什么”，比如 `修复：合计行重复计算`，别写 `update`。
