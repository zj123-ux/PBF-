# PBF 滚动预测自动编制工具

通过 Excel COM 驱动 SAP Analysis for Office（AO），按业务顺序把滚动预测数据填入 PBF 五张编制表、执行计算与运行包，最后读取损益核对指标。整个流程不需要人工点击。

---

## 1. 运行环境（缺一不可）

| 项目 | 要求 |
|---|---|
| Windows | 10 / 11 |
| Excel | 已安装 **SAP Analysis for Office** 加载项，且顶部能看到 `Analysis` 选项卡 |
| Python | 3.10 及以上，`python` 在 PATH 中（**用 `pbf-uploader.exe` 时不需要**） |
| WPS | **可开着使用**。工具已自动绕开 WPS（用真实 Excel 的版本专用 CLSID 连接），不会关闭/操作你开着的 WPS 表格 |
| 屏幕 | 保持解锁；本工具会模拟鼠标点击（折叠损益表年度列、点导航表"打开"） |
| 网络 | 能访问 BPC PRD |

安装依赖：

```bat
pip install -r requirements.txt
```

---

## 2. 准备工作（每次编制前）

### 2.1 配置账号

账号密码放在本机 `config/credentials.yaml`（**不要提交到代码库 / 不要发给别人**）：

```yaml
username: your.name
password: your_password
client: "333"
lang: zh
```

也可以用环境变量覆盖：`PBF_USERNAME`、`PBF_PASSWORD`、`PBF_CLIENT`、`PBF_LANG`。
优先级：环境变量 > `credentials.yaml` > `flow.yaml` 的 `login` 段。

### 2.2 准备数据

从《滚动预测模板表.xlsx》导出 **1–6 个期间**的上传数据（用 openpyxl 直接读文件，**无需先开 Excel**）：

```bat
python export_templates.py --periods 2027.01-2027.04          :: 区间（1–6 个月）
python export_templates.py --periods 2027.01,2027.03          :: 列表
python export_templates.py 2027.03                            :: 单期间（向后兼容旧用法）
python cli.py export --template 滚动预测模板表.xlsx --periods 2027.01-2027.04
```

期间语法支持单值 / 逗号列表 / `a-b` 区间 / 混合。输出目录：单期间 `data_2027.03/`，多期间 `data_2027.01_2027.04/`。

会在输出目录下生成 5 份 CSV：

| 文件 | 对应编制表 |
|---|---|
| `价格销量CT.csv` | 价格&销量&加权CT编制表 |
| `客户费用.csv` | 客户费用编制表 |
| `HQ毛上调整.csv` | HQ渠道毛上调整表 |
| `其他成本.csv` | 其他成本编制表 |
| `费用及损益其他项.csv` | 费用及损益其他项编制查询表 |

> 也可以自己准备 CSV，列名需与表单配置（`config/forms/*.yaml`）一致；期间列写成 `2027.03` 这样的文本。

### 2.3 确认导航表路径

`config/flow.yaml` 中：

```yaml
navigator:
  launch_file: "C:\\Users\\<你>\\Desktop\\AR项目\\PBF Pro - 全面预算系统流程导航.xlsb"
```

路径不对时脚本无法自动启动 Excel。

---

## 3. 执行

### 方式一：双击 run.bat（推荐）

双击 `run.bat`，或带参数：

```bat
run.bat                                  :: 全流程
run.bat --from-step 4                    :: 从第 4 步（其他成本）开始
run.bat --data-dir data_2027.03          :: 指定数据目录（默认取最新的 data_YYYY.MM）
```

脚本会：有 `pbf-uploader.exe` 就直接用它，否则用 `python cli.py` → 运行流程 → 结束后暂停。
Excel 没开时会自动打开导航表，AO 登录框会自动填写。

### 方式二：直接双击 pbf-uploader.exe

**不需要安装 Python**。双击等价于跑全流程（结束后按回车退出）；带参数用法：

```bat
pbf-uploader.exe run-flow --data-dir data_2027.03
pbf-uploader.exe run-flow --only-verify
pbf-uploader.exe run-flow --from-step 4
```

> exe 必须和 `config/`、`data_2027.xx/`、`logs/` 放在同一目录。

### 方式三：命令行（开发/调试用）

```bat
python cli.py run-flow --data-dir data_2027.03
python cli.py run-flow --only-verify     :: 只确认版本 + 读损益指标
python cli.py run-flow --skip-fill       :: 跳过填表，只跑计算/运行包/核对
python cli.py run-flow --from-step 4     :: 断点续跑
```

### 导出与批量清除（数据准备，无需打开 Excel / AO）

这两个命令只操作本机 CSV，不连 Excel、不跑 AO：

```bat
:: 导出 1–6 个月的上传数据
python cli.py export --template 滚动预测模板表.xlsx --periods 2027.01-2027.04
python cli.py export --template 滚动预测模板表.xlsx --periods 2027.01,2027.03-2027.05 --zero   :: 全 0 占位

:: 批量清除（把已填数字清成 0），期间可自由选择
python cli.py clear --periods 2027.01-2027.04                     :: 只生成全 0 CSV 到 data_..._clear/
python cli.py clear --periods 2027.01-2027.04 --run               :: 生成后立即执行清除填表
python cli.py clear --periods 2027.02 --data-dirs data_2027.02    :: 只清指定数据目录
```

清除逻辑：从已有 `data_*` 目录取**已填项的并集**（各表单的维度键），对所选期间生成全 0 CSV；`--run` 时复用 `run-flow` 编排把 0 写回 AO 并跑计算/运行包。双击 `clear.bat` / `export.bat` 也可用（参数同命令行）。

> **数据目录命名规则**：上传数据目录必须是 `data_YYYY.MM`（单期间）或 `data_YYYY.MM_YYYY.MM`（多期间），月份两位补零——`data_2027.1` 这类不补零的名字不被识别。默认数据目录按目录名里的**目标期间**取最新（多期间取最大期间），不是按目录名字符串排序，跨年也不会排错；`data_*_clear` 是批量清除的输出目录，永远不作为数据源。不符合命名规则的 `data_` 前缀目录（旧版目录名、备份、临时目录）不参与选择，`export` / `clear` / `run-flow` 运行时会列出并提示。

### 重新打包 exe

改了 Python 代码后重新生成 `pbf-uploader.exe`：

```bat
pip install pyinstaller
python -m PyInstaller --noconfirm --onefile --name pbf-uploader --distpath . ^
   --workpath build --specpath build ^
   --hidden-import win32com.client --hidden-import win32com.server.exception ^
   --hidden-import pythoncom --hidden-import comtypes ^
   --collect-all uiautomation ^
   --collect-submodules engine cli.py
```

> `--collect-submodules engine` 不能省：`engine/__init__.py` 是 PEP 562 按需导入
> （`from engine import ExcelEngine` 走 `__getattr__` 懒加载），PyInstaller 的静态
> 分析发现不了 `engine.excel` 等子模块，漏掉这条打包出的 exe 启动即
> `ModuleNotFoundError: engine.excel`。

产物约 82 MB（含 pandas / uiautomation）。

---

## Web 界面（图形化，可选）

不想记命令行时，可以用浏览器界面：填账号 → 选版本/品牌×品类/期间（1–6 月，可单值/列表/区间）→ 上传模板表 → 点运行 / 批量清除 → 看损益。

```bat
run_web.bat
```

或手动：

```bat
pip install flask
python -m webapp.app
```

浏览器自动打开 `http://127.0.0.1:8765`（只绑本机 127.0.0.1，不对外；端口可用环境变量 `PBF_WEB_PORT` 覆盖）。

界面四个区块：

1. **账号** — 填用户名/密码/客户端/语言，点「保存账号」写入 `config/credentials.yaml`（与命令行同源）。
2. **编制参数** — 选品牌×品类（当前只有「乐华·电视」可用，其余灰显未配置）、版本、期间（1–6 月，支持单值/列表/区间，如 `2027.01-2027.04`）。
3. **上传模板表** — 选《滚动预测模板表.xlsx》，自动导出 5 份 CSV 到 `data_<首>_<末>/`（多期间）或 `data_<期间>/`（单期间），用 openpyxl 直接读，**无需先开 Excel**。
4. **运行** — 「环境自检」看 Excel/WPS/导航表状态；「运行全流程」会（可选）先切版本，再跑 `run-flow`，实时刷日志，结束显示损益核对表；「批量清除填 0」从既有 `data_*` 取已填项并集生成全 0 CSV 并运行（等价 `cli.py clear --periods ... --run`）。

> 网页底层仍是调用 `cli.py`（子进程），与命令行完全等价，运行结果同样写入 `logs/YYYYMMDD.jsonl`。

### 品牌 × 品类（业务单元）

业务单元（品牌×品类）集中配在 `config/business_units.yaml`。选中单元后，`steps[].nav_cell`
会覆盖 `config/flow.yaml` 的导航行，网页下拉与命令行效果一致：

```bat
python cli.py units                                   :: 查看/校验品牌×品类配置
python cli.py run-flow --unit 乐华·电视
python cli.py clear --periods 2027.01-2027.04 --unit 乐华·电视 --run
```

- 单元 `steps` 必须与 `flow.yaml` 的 `steps` **顺序一一对应**（`表单` 名用于校验，错位会直接报错）；
- 可选 `after_close_nav_cell`（运行包导航行）、`verify_nav_cell` / `verify_form_row`（损益核对位置）；
- 未指定 `--unit` 时沿用 `flow.yaml` 的默认导航行（当前即乐华），保持向后兼容；
- 导出/运行前会检查 CSV 里 `品牌`/`品类` 列是否整列为空，并在界面或控制台给出 `[数据告警]`。

要接新单元（TCL/雷鸟、空调/冰箱/洗衣机/TCL商用）：

1. 在 AO 里现场确认该单元在导航表上的业务单元行（nav_cell），以及各编制表的工作簿/Sheet/期间列；
2. 在 `config/business_units.yaml` 对应单元补 `steps` 的 `nav_cell` 并把 `启用` 改为 `true`；
3. 若该单元用到的编制表结构与现有 `config/forms/*.yaml` 不同，补一份表单配置（可用 `python cli.py scan ... --write` 辅助生成）。

启动 `run-flow` 时会先跑一遍单元配置校验（数量/表单对应/nav_cell 重复等），
网页层的下拉自动读 `business_units.yaml`，启用后即可选择。

---

## 4. 流程顺序（顺序敏感，勿改）

1. 价格&销量&加权CT → 保存 → `执行收入计算`
2. 客户费用 → 保存 → `执行客户费用分摊`
3. HQ渠道毛上调整 → 保存 → 导航表 AC62 → 运行包：收入计算、算到销售组织的成本
4. 其他成本 → 保存（本表打开即编辑态，**不要切模式**）
5. 费用及损益其他项 → 保存 → 导航表 AC70 → 运行包：计算滚动预测损益
6. 打开 YC_CBG损益查询表N → 刷新 → 折叠 2026 年度列 → 读 2027 指标

### 每步执行前的自动校验

| 校验 | 不通过时 |
|---|---|
| 版本确认（`--version`） | 中止 |
| **期间窗口**：目标数据期间必须落在版本起始期间之后（窗口默认 6 个月） | 早于版本 → 中止；超窗口 → 只告警 |
| 编制模式 / 版本单元格 | 中止 |
| 行匹配（科目/组合键全部命中） | 中止，并列出未匹配项 |
| **期间覆盖**：数据里的期间必须都有对应列 | 中止（避免“少写一部分还显示成功”） |
| **单元格唯一**：两行数据不能落到同一格 | 中止（避免互相覆盖） |
| 保存结果 | 未保存成功 → 中止，不再跑计算和运行包 |
| **假保存识别**：AO 提示“没有要保存的数据”，但本次确有格子新值≠旧值 | 中止（常见于表格未进入编制态、写入未生效） |
| **落库校验**：保存后原地 `Refresh_Sheet`（从 BPC 重新拉数）→ 读回本次写入的每一格 | 任一格对不上 → 中止 |
| **损益期间存在**：目标期间列必须出现在损益表头 | 缺失 → 中止（不再静默只读部分月份） |
| **损益指标合理**：目标期间指标不能全为空/为 0 | 中止（运行包可能未生效） |

落库校验可通过 `config/flow.yaml` 调整：

```yaml
verify_writeback:
  enabled: true
  mode: refresh     # refresh=保存后原地刷新读回（默认，不关表重开）
                    # reopen =关闭(不保存)→经导航重开→读回（最严格，但每表多 30-60 秒）
  tolerance: 0.01   # 数值绝对误差
  sample: 0         # 0=全量；>0=每表均匀抽样 N 格（提速但可能漏检）
```

单张表可在 `config/forms/*.yaml` 里用 `verify_mode: reopen` 单独覆盖。

---

## 5. 输出

- 控制台：填表/计算/运行包状态 + 损益核对表
- 审计日志：`logs/YYYYMMDD.jsonl`（每次运行一条完整记录）

正常结果示例：

```
[填表 / 计算]
  价格销量: 写入 201 格  ✓
  ...
[运行包]
  ✓ 已完成 时长=6 秒  PS_模型_收入_预测_滚动预测收入计算_价量一起编制
[损益核对]
  指标          2027.01    2027.02    2027.03
  净收入      3,500.0027 3,500.0027 3,500.0027
```

---

## 6. 常见问题

| 现象 | 原因 / 处理 |
|---|---|
| `无法连接真实 Excel：相关 CLSID 均被 WPS 抢占` | 极少见：WPS 把 Excel 所有版本专用 CLSID 的注册也改写了。临时关一次 WPS 重试，或修复 Office 的 COM 注册 |
| `Could not attach to running Excel` | Excel 没开；`run.bat` 会自动打开，命令行方式请手动打开导航表 |
| `选中 ACxx 并点击 AT52 后未打开新工作簿` | 导航表状态过期。已内置重试 3 次；仍失败请手动点一次导航表"刷新状态" |
| 打开表单后一直卡住 | Excel 里可能有提示弹窗；工具会自动关闭 `没有要保存的数据` 类弹窗，其它弹窗需人工确认 |
| 保存时报"没有要保存的数据" | 本次写入与已保存值一致 → 属正常，自动跳过保存；**但若本次确有新值≠旧值 → 判为失败并中止**（表格多半没进入编制态），需手动确认后 `--from-step N` 续跑 |
| 折叠 2026 年度列失败 | 需要 Excel 在前台且屏幕未锁。点 `/config/flow.yaml` 的 `verify.collapse_year` 可改目标年度 |
| 版本不符 | `--version` 与实际版本不一致；不带该参数则只记录不校验 |
| 中途失败 | 用 `--from-step N` 从失败那一步续跑，前面的步骤不会重做 |
| 损益 2027 列全是 0 / 提示"损益表未出现目标期间列" | 确认三个运行包都"已完成"、版本与期间匹配；工具现在会直接中止而不是输出 0 值。若确认无误仍报错，检查 `verify.header_row` 与年度列折叠是否成功（需 Excel 前台、屏幕未锁） |
| `配置缺少期间 [...]` | 数据里的期间在表单配置里没有对应列（版本滚动后列位置变了），需要更新 `config/forms/*.yaml` 的期间列 |
| `重复写入 [...]` | 两行数据映射到同一单元格（产品/科目重复），先清洗数据 |
| `保存失败：...` | 该表未保存成功，流程已中断；修复后按提示的步骤号用 `--from-step` 续跑 |
| `落库校验失败：...` | 保存没报错，但关闭重开后读回的数值与写入值不一致。先确认表格是否真的进入编制态；若确认是 AO 保存时自动重算导致，可临时把 `config/flow.yaml` 的 `verify_writeback.enabled` 设为 `false` 跳过校验（并反馈具体格子） |
| `目标期间 ... 早于版本起始期间` | 版本没切，先在 AO 里切换版本再跑 |
| Excel 中途闪退 / `已与其客户端断开连接` / 表单反复打不开 | 已内置自动恢复：结束本流程的 Excel 进程 → 重开导航表 → 重新登录 → **从失败的那一步重做**（最多 1 次，见 `flow.yaml` 的 `recovery` 段）。单步超过 `recovery.stall_timeout_sec`（默认 600 秒）无进展会自动判定卡死并重启。恢复也失败时按提示用 `--from-step N` 续跑 |

---

## 7. 配置说明

| 文件 | 作用 |
|---|---|
| `config/flow.yaml` | 主流程：登录、导航表、运行包、5 个填表步骤、核对期间与指标行 |
| `config/credentials.yaml` | 本机账号密码（**勿外传**） |
| `config/forms/*.yaml` | 每张编制表的布局：工作簿特征、Sheet、表头行、科目列、期间列、模式区 |
| `config/dimensions.yaml` | 版本成员、导航表版本单元格、运行包清单单元格 |

### 换期间 / 换版本时要改哪里

- 生成数据：`python export_templates.py --periods 2027.01-2027.04`（1–6 月，自动输出到 `data_2027.01_2027.04/`；单期间 `2027.04` → `data_2027.04/`）
- 批量清除：`python cli.py clear --periods 2027.01-2027.04 --run`（期间自由选择）
- 直接执行：`run.bat --data-dir data_2027.04`（多期间目录也可：`--data-dir data_2027.01_2027.04`）
- **目标期间从数据目录名自动推导**（`data_2027.01_2027.04` → 取最大期间 2027.04），要折叠的年度也自动推导（上一年度），不需要再改配置
- 确实需要强制指定期间时，在 `flow.yaml` 的 `verify` 下加一行：`target_period: "2027.04"`
- `verify.periods` 仅在表头解析失败时兜底，平时不用维护
- 版本切换建议在 AO 里手动确认；`python cli.py switch-version C03202701V1` 会自动校验显示文本，失败会还原

---

## 8. 目录结构

```
pbf-uploader/
├─ run.bat                 一键执行（优先用 exe）
├─ export.bat              从模板表导出多期间数据（入口）
├─ clear.bat               批量清除填 0（入口）
├─ pbf-uploader.exe        打包产物（无需 Python，约 82MB）
├─ cli.py                  命令行入口（run-flow / export / clear）
├─ paths.py                运行目录解析（兼容源码/exe）
├─ export_templates.py     从模板表导出 CSV（薄封装）
├─ requirements.txt
├─ config/
│  ├─ flow.yaml            主流程
│  ├─ credentials.yaml     本机凭据
│  ├─ dimensions.yaml      版本/导航/运行包单元格
│  └─ forms/*.yaml         5 张编制表配置
├─ engine/
│  ├─ excel.py             COM 连接、宏执行、弹窗看门狗、鼠标点击
│  ├─ bootstrap.py         Excel 启动与断线恢复
│  ├─ orchestrator.py      流程编排 + 损益核对 + 自动恢复
│  ├─ writer.py            填表（预检/写值/保存）
│  ├─ template_export.py   多期间模板导出
│  ├─ clear.py             批量清除（生成全 0 CSV）
│  ├─ periods.py           期间解析（单值/列表/区间）
│  ├─ ao_login.py          AO 登录框自动填写
│  ├─ navigator.py         导航表操作
│  ├─ scanner.py           表单结构扫描（生成配置用）
│  └─ reader.py / runner.py  读值、单独跑运行包
├─ data_2027.xx/           当期上传数据（多期间为 data_首_末/）
└─ logs/                   审计日志
```

---

## 9. 开发与测试

纯逻辑模块（如 `engine/periods.py`、配置校验、webapp 运行状态管理等）不依赖 COM/Office。
`engine/__init__.py` 采用按需导入（PEP 562），`import engine` 或
`from engine import periods` 不会加载 pywin32，因此可以在没有 Excel 的环境里测试：

```bat
python -m unittest discover -s tests -v
```

覆盖期间解析、数据目录选择（不会把 `data_*_clear` 当数据源）、模板导出/批量清除的
空表与陈旧文件处理。同一套测试可直接接入 CI（Windows + Python 3.10 / 3.12 / 3.14），
不需要安装 Office 或 pywin32。改完代码先跑一遍，再接真机验证。

审计日志 `logs/YYYYMMDD.jsonl` 是财务操作留痕，**默认永不清理**。如需控制
磁盘占用，显式设置环境变量 `PBF_AUDIT_RETENTION_DAYS`（正整数，单位天），
例如 `set PBF_AUDIT_RETENTION_DAYS=365`；此后启动任意命令时会自动清理过期
文件并打印清理数量。非数字 / 0 / 负数均视为未配置（不清理）。
