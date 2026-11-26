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

`flow.yaml` 的 `login` 段还支持几个时长参数（秒）：

- `timeout`：自动登录整体超时（默认 120）。
- `no_box_ok_sec`：新开导航表后最多等多久登录框（默认 30）。已用缓存凭据
  登录时 AO 不弹框，超时后按"已登录"继续；真正的会话失效会由后续操作
  暴露，届时按提示手动登录即可。
- `reuse_wait_sec`：复用已打开 Excel 时守候 AO 登录框的时长（默认 15）。
  登录框约 10s 才弹，已登录时不会弹，所以默认值偏短；若日志出现
  "复用 Excel：Ns 内未检测到 AO 登录框"且随后的操作卡死在登录框，
  把它调大（如 60）。

登录框出现后**最多自动填写一次**：凭据被拒时反复提交有账号锁定风险，
第二次仍见到登录框会立即停止并提示在 Excel 中手动登录。

### 2.2 准备数据

从《滚动预测模板表.xlsx》导出 **1–6 个期间**的上传数据（用 openpyxl 直接读文件，**无需先开 Excel**）：

```bat
python export_templates.py --periods 2027.01-2027.04          :: 区间（1–6 个月）
python export_templates.py --periods 2027.01,2027.03          :: 列表
python export_templates.py 2027.03                            :: 单期间（向后兼容旧用法）
python cli.py export --template 滚动预测模板表.xlsx --periods 2027.01-2027.04
```

期间语法支持单值 / 逗号列表 / `a-b` 区间 / 混合。输出目录：单期间 `data_2027.03/`，多期间 `data_2027.01_2027.04/`；加 `--unit 乐华·电视` 时目录名再带单元后缀（`data_2027.01_2027.04_乐华·电视/`），不同业务单元同一期间互不覆盖。

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

### Excel 实例与取消（`logs/excel_owner.json`）

Excel 是单实例应用：即使再 `Start excel.exe 导航表.xlsb`，文件也会被塞进
已在运行的实例。若那个实例是无窗口、0 工作簿的"僵尸"，AO 登录框就永远
弹不出来——现场表现为 **"AO 又没了"**。为此：

- 每次启动/复用 Excel 后，把 PID + 进程创建时间写入 `logs/excel_owner.json`，
  只有记录为 `pbf` 的实例才允许自动 Quit/结束（用户的 Excel 一律不动）。
- 清理前先 `Quit`，**确认 PID 已从进程表消失**；`Quit` 被外部 COM 引用忽略时
  强制结束该 PID（日志会写"Quit 未生效…强制结束该进程"）。
- 绑定到导航表实例后强制把 Excel 设为可见；无法恢复可见时直接报错，不在僵尸实例上继续。
  但**探测性连接**（找导航表、判断是否空实例）一律 `make_visible=False`：
  以前探测时 `Dispatch` 连上残留的隐藏空实例再设可见，用户屏幕上就会多出一个
  空白 Excel，和导航表并排出现。
- 归属记录里的 `source: reused` 表示复用用户已开的 Excel，退出时不会关闭它。
- 登录框关闭后加载项会自动刷新所有数据源（实测占住 Excel UI 线程 30s 级）。
  刷新期间调 `SAPSetVariable`/`PauseVariableSubmit` 会把加载项打崩，弹
  「SAP BI 加载项已断开所有数据源连接 / 未将对象引用设置到对象的实例」。
  工具现在登录成功后先等 `login.settle_start_sec`（默认 5s），再用普通 Excel
  调用当空闲屏障（`login.idle_timeout_sec`，默认 120s）；切版本前也有同样的
  空闲屏障（`ao_idle_timeout_sec`，默认 30s，空闲时立即返回）。
- Excel 同时存在多个进程时（例如上一轮被杀后残留的空实例），`Dispatch` 可能
  连到没有导航表的那个。打开/复用导航表时改为**按文件路径从运行对象表（ROT）
  绑定**到真正开着导航表的实例，因此不会再「等待导航表打开超时（180s）」；
  也不会因为文件已经开着而又启动一份。

Web 界面的"取消"不再直接 `taskkill`：子进程在安全点（写单元格前、保存前、
轮询运行包时）检查取消标志并优雅退出，未完成的表单工作簿会被关闭且不保存。
启动阶段（打开导航表、等 AO 登录）也会检查取消标志——以前这里没有检查点，
取消后 60s 宽限一过就会被 `taskkill /T` 连刚打开的 Excel 一起杀掉，表现为
"AO 又没了"；现在取消会优雅退出，**Excel 与 AO 保持打开**，可继续手动操作。
硬取消/超时后，槽位只等**本次运行启动的** Excel 退出；运行包已在 BPC 后台
触发时（退出码 3）无法本地探测进度，槽位保守占用，由用户确认后点"强制释放"。
`PBF_CANCEL_GRACE_SEC` 控制优雅取消的宽限秒数（默认 60）。

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

> **数据目录命名规则**：上传数据目录必须是 `data_YYYY.MM`（单期间）或 `data_YYYY.MM_YYYY.MM`（多期间），月份两位补零——`data_2027.1` 这类不补零的名字不被识别。按业务单元运行时目录名再加 `_<单元>`（如 `data_2026.11_2026.12_TCL·TCL商用`），乐华与 TCL商用 同一期间各用各的目录、上传互不覆盖。默认数据目录按目录名里的**目标期间**取最新（多期间取最大期间），不是按目录名字符串排序，跨年也不会排错；`run-flow` / `clear` 带 `--unit` 时优先该单元的专用目录，只有旧版无后缀目录唯一时才回退（否则报错要求 `--data-dir` 指定）；`data_*_clear` 是批量清除的输出目录，永远不作为数据源。不符合命名规则的 `data_` 前缀目录（旧版目录名、备份、临时目录、未配置的单元名）不参与选择，`export` / `clear` / `run-flow` 运行时会列出并提示。

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

> **本机信任模型**：只绑 127.0.0.1 挡住的是局域网和其它网站的浏览器请求
> （Host/Origin 双重校验防 DNS rebinding），挡不住**本机进程**——本机任何
> 程序都可以直接调用这些接口（包括触发上传的 `/api/run`）。这是有意为之
> （便于本地脚本调用），因此请只在可信的个人电脑上运行 web 服务；
> 公用或不可信的机器上不要启动。

界面四个区块：

1. **账号** — 填用户名/密码/客户端/语言，点「保存账号」写入 `config/credentials.yaml`（与命令行同源）。
2. **编制参数** — 选品牌×品类（`config/business_units.yaml` 中已启用的单元，未配置的灰显）、版本、期间（1–6 月，支持单值/列表/区间，如 `2027.01-2027.04`）。
3. **上传模板表** — 选《滚动预测模板表.xlsx》，自动导出 5 份 CSV 到 `data_<首>_<末>_<单元>/`（单期间为 `data_<期间>_<单元>/`），用 openpyxl 直接读，**无需先开 Excel**；不同业务单元同期间上传各自独立目录，互不覆盖。
4. **运行** — 「环境自检」看 Excel/WPS/导航表状态；「运行全流程」会（可选）先切版本，再跑 `run-flow`，实时刷日志，结束显示损益核对表；「批量清除填 0」从既有 `data_*` 取已填项并集生成全 0 CSV 并运行（等价 `cli.py clear --periods ... --run`）。

> 网页底层仍是调用 `cli.py`（子进程），与命令行完全等价，运行结果同样写入 `logs/YYYYMMDD.jsonl`。

### 品牌 × 品类（业务单元）

业务单元（品牌×品类）集中配在 `config/business_units.yaml`。选中单元后，`steps[].nav_cell`
会覆盖 `config/flow.yaml` 的导航行，网页下拉与命令行效果一致：

```bat
python cli.py units                                   :: 查看/校验品牌×品类配置
python cli.py nav-scan --unit 乐华·电视                :: 只读核对导航行 AA/AB/AC（换账号后用）
python cli.py run-flow --unit 乐华·电视
python cli.py clear --periods 2027.01-2027.04 --unit 乐华·电视 --run
```

- 单元 `steps` 必须与 `flow.yaml` 应用 `跳过步骤` 之后的 `steps` **顺序一一对应**（`表单` 名用于校验，错位会直接报错）；
- **启用的单元必须配置 `nav_expect`**：运行前会读取该行的 **AA(业务模块) + AB(流程步骤) + AC(业务单元)** 三列，标签必须包含全部关键字，不匹配立即中止——**导航表按账号权限生成，换账号后行号可能整体移动，没有这层校验会写进别的业务单元**。配置层次：
  - `unit.nav_expect`：所有导航行都要包含的（品牌，如 `[乐华]`）；
  - `steps[].nav_expect`：该步追加，可含 AB 关键字（如 `[乐华, 智屏, 编制毛利之下费用]`）；
  - `steps[].after_close_nav_expect` / `unit.verify_nav_expect`：运行包行/核对行专用；
  - **AC 相同的行必须靠 AB 区分**：AC65 与 AC67 的 AC 都是「CBG-智屏-乐华品牌」，AB 分别是「编制毛利之下费用 / 查询毛利之下费用」；
- 可选 `after_close_nav_cell`（运行包导航行）、`verify_nav_cell` / `verify_form_row`（损益核对位置）；
- **单元级 `跳过步骤`**：该单元流程里没有的步骤整步移除（连同该步的 `after_save` 计算宏和 `after_close` 运行包），如 TCL商用没有 HQ渠道毛上调整表：
  ```yaml
  跳过步骤: [HQ毛上调整]
  ```
  关键字必须唯一命中一个流程步骤，匹配 0 个或多个都会直接报错；
- **单元级 `表单覆盖`**：该单元这张表不编制的度量块，writer 的表头识别、写入和 CSV 列校验都不再要求它（如 TCL商用 价格销量表不编制 CT 成本）：
  ```yaml
  表单覆盖:
    E61_CBG_价格销量CT编制:
      跳过度量: [加权CT]
  ```
- 未指定 `--unit` 时沿用 `flow.yaml` 的默认导航行（当前即乐华），保持向后兼容；
- 导出/运行前会检查 CSV 里 `品牌`/`品类` 列是否整列为空，并在界面或控制台给出 `[数据告警]`。
- **不同单元的数据文件不同**：`--data-dir` 指向哪个目录，就把哪个目录的 CSV 写进该单元的表单。跑 TCL商用 时必须用商用的导出目录，不要拿乐华的 `data_*` 目录跑——维度对不上时 writer 会按 `allow_new_rows` 追加出错误行。

**按内容定位导航行（默认开启）**：运行时会扫描导航表 `AA(业务模块)+AB(流程步骤)+AC(业务单元)`
三列（`config/flow.yaml` 的 `navigator.locate_by_label` / `locate_scan_range`）：

| 标签命中情况 | 行为 |
|---|---|
| 唯一命中 | 选中命中行——换账号后行号整体移动也不需要改配置 |
| 多行命中，且配置行在其中 | 沿用配置行并告警（建议补 AB 关键字让命中唯一） |
| 零命中 / 多行命中但配置行不在其中 | **中止**，不选中、不打开任何表单（绝不猜行） |

固定的 `nav_cell` 变成兜底与配置记录，选中前仍会做一次三列标签校验。

**换账号 / 换权限（必读）**：导航表是按登录账号的权限实时生成的，换人后行号可能整体移动。建议流程：

1. 用新账号**手动**登录 AO，确认导航表已打开。不要在工具里反复重试登录——凭据被拒有账号锁定风险（工具自身也只在登录框出现时最多自动填一次）；
2. 只读扫描（不选中任何行、不打开表单）：
   ```bat
   python cli.py nav-scan --unit 乐华·电视
   ```
   会打印导航表 AA/AB/AC 全部非空行，以及每个「填表 / 运行包 / 损益核对」行的
   「配置 → 实际」是否一致。显示「实际 ACxx」时按提示更新 `business_units.yaml`
   的 `nav_cell`；显示「定位失败」时先补/改 `nav_expect`；
3. 零写入干跑确认 5 张表都能打开、命中数正常：
   ```bat
   python cli.py run-flow --preflight-only --unit 乐华·电视
   ```
4. 以上都通过后再正式运行。面板指纹缓存按「账号\|版本\|导航格」隔离，换账号无需手工清缓存。

要接新单元（TCL/雷鸟、空调/冰箱/洗衣机）：

1. 在 AO 里现场确认该单元在导航表上的业务单元行（nav_cell + AA/AB/AC 标签原文），以及各编制表的工作簿/Sheet/期间列；
2. 在 `config/business_units.yaml` 对应单元补 `steps` 的 `nav_cell`、`nav_expect`/`after_close_nav_expect` 与 `verify_nav_expect`（按三列原文取关键字，同 AC 的行必须带上 AB 关键字）并把 `启用` 改为 `true`；
3. 该单元比乐华少的步骤写进 `跳过步骤`，不编制的度量块写进 `表单覆盖`（见上一节，TCL商用 是现成样例）；
4. 若该单元用到的编制表结构与现有 `config/forms/*.yaml` 不同，补一份表单配置（可用 `python cli.py scan ... --write` 辅助生成）；
5. 用该账号跑 `python cli.py nav-scan --unit 新单元名` 确认每一步都能按内容唯一定位，再用该单元的数据目录 `--preflight-only` 干跑。

**已启用单元**：`乐华·电视`（5 步）、`TCL·TCL商用`（4 步，跳过 HQ毛上调整，价格销量表不编制加权CT）。

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
| **落库校验**：保存后关闭（不保存）→ 重开 → 读回本次写入的每一格 | 任一格对不上 → 中止 |
| **损益期间存在**：目标期间列必须出现在损益表头 | 缺失 → 中止（不再静默只读部分月份） |
| **损益指标合理**：目标期间指标不能全为空/为 0 | 中止（运行包可能未生效） |

落库校验可通过 `config/flow.yaml` 调整：

```yaml
verify_writeback:
  enabled: true
  tolerance: 0.01   # 数值绝对误差
  sample: 0         # 0=全量；>0=每表均匀抽样 N 格（提速但可能漏检）
```

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
| 长流程最后一步报"面板未刷新 / AO 会话可能已过期" | **AO/BPC 会话约 15 分钟过期**（2026-11-23 现场：17:50 登录 → 18:05 失效）。失效后导航表宏统一报 `0x800A9C68`、右侧步骤面板不再随选中刷新、版本被重置回月初版，但 Excel 本身没崩。工具会识别为 `SessionExpiredError` 并**强制重启 Excel + 自动重新登录**后重做当前步骤（`recovery.max_retries`，默认 1 次）。缓解：长流程尽量在会话新鲜时启动；`panel_wait_sec` 默认 15s |
| 上一次运行/手工操作把步骤面板停在目标行，下次运行误报会话失效 | 已修复：验证过的面板指纹按「账号\|版本\|导航格」持久化到 `logs/nav_panel_cache.json`，新进程能识别"面板已是该行正确内容"，不再误触发重启；换账号后缓存自动隔离 |
| 报「导航行 ACxx 的标签（AA+AB+AC）是「…」，未包含 […]」 | **换了账号/权限，导航表行号移动了**。工具在选中前核对 AA/AB/AC 三列标签，不匹配即中止（不会打开表单、不会写数据）。用该账号的导航表核对 `config/business_units.yaml` 的 `nav_cell` / `nav_expect`（AC 相同的行要带上 AB 关键字，如「编制毛利之下费用」）；也可先跑 `python cli.py nav-scan --unit …` 与 `--preflight-only` 干跑确认 |
| 报「ACxx 按标签定位失败：…」 | 该账号上找不到包含 `nav_expect` 全部关键字的行，或命中多行且配置行不在其中。**工具不会猜行**，已零副作用中止。先跑 `python cli.py nav-scan --unit …` 看该账号导航表真实的 AA/AB/AC，再改配置；AC 相同的行必须补 AB 关键字 |
| 日志「导航行按标签定位：配置 ACxx → 实际 ACyy」 | 换账号后行号整体移动，工具已按内容自动用实际行打开。建议按提示更新 `business_units.yaml` 的 `nav_cell`；若关键字本身指错了行，也会按关键字定位，请一并核对 `nav_expect` |
| 打开表单后一直卡住 | Excel 里可能有提示弹窗；工具会自动关闭 `没有要保存的数据` 类弹窗，其它弹窗需人工确认 |
| 保存时报"没有要保存的数据" | 本次写入与已保存值一致 → 属正常，自动跳过保存；**但若本次确有新值≠旧值 → 判为失败并中止**（表格多半没进入编制态），需手动确认后 `--from-step N` 续跑 |
| 折叠 2026 年度列失败 | 需要 Excel 在前台且屏幕未锁。点 `/config/flow.yaml` 的 `verify.collapse_year` 可改目标年度 |
| 版本不符 | `--version` 与实际版本不一致；不带该参数则只记录不校验 |
| 切版本报 `AR45 仍为 ''（期望…）` | 工具现已按导航表 VBA 的「刷新状态」完整复刻：写 C20/F20 → 给 `DS_PROCESSLIST` 设版本提示变量 `ZCCSO_ZDCATEGRY_010` 和期间变量 `ZCCSO_ZDTIME_048`（取 C19）→ `RefreshData` 刷新，并在 `version_switch_timeout_sec`（默认 60s，`dimensions.yaml`）内轮询 `AR45`。日志会打印 AO 变量回读值和刷新错误原文。注意 AO 的 `SAPGetVariable` 参数顺序是（数据源, 变量, 属性），且变量提交暂停期间回读恒为空——回读放在恢复提交之后。只写 C20/F20 + `RefreshAll` 不会真正切版本（AO 回显会把 C20 抄回旧值）。仍失败说明 AO 没连上 BPC 或版本成员键有误：确认 Excel 里 Analysis 已连上后重试，或手动选择版本 |
| AO 弹「SAP BI 加载项已断开所有数据源连接」（ID-111006）或「未将对象引用设置到对象的实例」 | 加载项在自动刷新数据源时被 `SAP*` 调用打断，常见于「刚登录就切版本」；上一次 Excel 被强杀也会留下坏会话。先点弹窗里的「重新启动」，必要时关闭 Excel 重开。新版工具在登录成功后和切版本前都会等 Excel UI 空闲（见「Excel 实例与取消」） |
| 打开导航表时多出一个空白 Excel 窗口 | 上一轮被强杀留下的隐藏空实例被探测性 `Dispatch` 连上并设成了可见。新版探测一律不改可见性，不会再弹出；残留进程可在确认没有未保存工作簿后手动关闭 |
| 报 `等待导航表打开超时（180s）`，但 Excel 里导航表明明开着 | 多半是同时存在多个 Excel 进程（残留的无窗口实例让 `Dispatch` 连错了实例）。工具现已按文件路径从 ROT 绑定到开着导航表的实例；若仍出现，请检查是否有其它 Excel 进程占用同名文件，或重启 Excel 后再试 |
| 中途失败 | 用 `--from-step N` 从失败那一步续跑，前面的步骤不会重做 |
| 损益 2027 列全是 0 / 提示"损益表未出现目标期间列" | 确认三个运行包都"已完成"、版本与期间匹配；工具现在会直接中止而不是输出 0 值。若确认无误仍报错，检查 `verify.header_row` 与年度列折叠是否成功（需 Excel 前台、屏幕未锁） |
| `配置缺少期间 [...]` | 数据里的期间在表单配置里没有对应列（版本滚动后列位置变了），需要更新 `config/forms/*.yaml` 的期间列 |
| `重复写入 [...]` | 两行数据映射到同一单元格（产品/科目重复），先清洗数据 |
| `保存失败：...` | 该表未保存成功，流程已中断；修复后按提示的步骤号用 `--from-step` 续跑 |
| `落库校验失败：...` | 保存没报错，但关闭重开后读回的数值与写入值不一致。先确认表格是否真的进入编制态；若确认是 AO 保存时自动重算导致，可临时把 `config/flow.yaml` 的 `verify_writeback.enabled` 设为 `false` 跳过校验（并反馈具体格子） |
| `目标期间 ... 早于版本起始期间` | 版本没切，先在 AO 里切换版本再跑 |
| Excel 中途闪退 / `已与其客户端断开连接` | 已内置自动恢复：重开导航表 → 重新登录 → 重做当前步骤（最多 1 次，见 `flow.yaml` 的 `recovery` 段）。恢复也失败时按提示用 `--from-step N` 续跑 |

---

## 7. 配置说明

| 文件 | 作用 |
|---|---|
| `config/flow.yaml` | 主流程：登录、导航表、运行包、5 个填表步骤、核对期间与指标行 |
| `config/credentials.yaml` | 本机账号密码（**勿外传**） |
| `config/forms/*.yaml` | 每张编制表的布局：工作簿特征、Sheet、表头行、科目列、期间列、模式区 |
| `config/dimensions.yaml` | 版本成员、导航表版本单元格、运行包清单单元格 |

### 换期间 / 换版本时要改哪里

#### 版本清单（月初版 / 财总版）

每个月都有两个滚动预测版本，Web 界面下拉框和 `cli.py switch-version` 都从 `config/dimensions.yaml` 读取：

| 月份 | 月初版（约每月 5 号） | 财总版（每月 20 号以后） |
|---|---|---|
| 2026年9月 | `C03202609V1` | `C03202609V2` |
| 2026年10月 | `C03202610V1` | `C03202610V2` |
| 2026年11月 | `C03202611V1` | `C03202611V2`（已现场核实 AR45/C20） |
| 2026年12月 | `C03202612V1` | `C03202612V2` |
| 2027年1月 | `C03202701V1` | `C03202701V2` |

- 财总版在 AO 里显示为「滚动预测YYYY年M月财总版」（口语中也叫“财总会版”），YAML 里的显示文本必须与导航表 `AR45` 一致，切版本时会据此校验，不一致会自动还原。
- 新增月份时：在导航表里手动切到该版本，读 `C20`（成员键）和 `AR45`（显示文本），按上面的格式追加到 `versions:` 即可。

- 生成数据：`python export_templates.py --periods 2027.01-2027.04`（1–6 月，自动输出到 `data_2027.01_2027.04/`；单期间 `2027.04` → `data_2027.04/`）
- 批量清除：`python cli.py clear --periods 2027.01-2027.04 --run`（期间自由选择）
- 直接执行：`run.bat --data-dir data_2027.04`（多期间目录也可：`--data-dir data_2027.01_2027.04`）
- **目标期间从数据目录名自动推导**（`data_2027.01_2027.04` → 取最大期间 2027.04），要折叠的年度也自动推导（上一年度），不需要再改配置
- 确实需要强制指定期间时，在 `flow.yaml` 的 `verify` 下加一行：`target_period: "2027.04"`
- `verify.periods` 仅在表头解析失败时兜底，平时不用维护
- 版本切换建议在 AO 里手动确认；`python cli.py switch-version C03202701V1` 会自动校验显示文本，失败会还原

#### 运行时的版本 / 新增行 / 锁定检查

- **逐表单核对版本**：每张表打开后读版本区（`F4` 显示文本、`E4` 目标成员键、`G4` SAP 当前变量值、`H4` 一致性 Y/N）。不是目标版本时自动写 `E4` → 执行 `Refresh_Sheet` → 轮询直到显示文本和当前值都等于目标键；90 秒内没切过来则报错中止（不会带着错误版本写数）。
- **新增行只在 AO 数据区内进行**：空白表单可先写维度列、等 2 秒让 AO 完成维度校验，再写期间数值；同一维度组合复用已有行。新增行候选必须与上方数据行的格式一致——AO 查询结果之后常残留文本格式（`@`）的模板行，写进去 `Plan_Save` 不报错但重开即清空（2026-11-24 价格表实测）。表单可用 `allow_new_rows: false` 禁止自动新增行；`append_rows: true` 则改为只在最后一条数据行下方表尾追加，不占查询区中间的空行（价格&销量表即如此：新维度组合追加到表尾，缺失的产品型号等属性按 `new_row_copy` 从同产品既有行复制）。数据键除 `key_columns` 外可用 `key_match_columns` 声明多组匹配列（价格表的「产品」可能是产品成员 AD，也可能是客户型号 AF）；若同一键在不同列组合下指向不同行，preflight 直接失败，不会串行写错。
- **写前按「当前模式」切编制态**：读页面顶部 `当前模式` 显示（价格表 `AM51`，其他表自动扫描表头；配置项 `mode_display_cell`）。显示“查询”且目标格里有蓝色只读格时，自动运行 `Switch_mode` 并轮询 `AM51` 变“编制”后才写值——否则写入不被 AO 识别，`Plan_Save` 会弹「没有要保存的数据」。显示“查询”但目标格全白（客户费用表）则直接写；显示“编制”绝不重复切换（会反向切回只读）。
- **表单锁定**：命中「被XXX用户锁定」提示（状态栏/表单文本框/整片蓝色只读区）时，跳过该表并结束流程，命令行退出码 2，Web 界面显示「变动费用表单被XXX用户锁定」。

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
