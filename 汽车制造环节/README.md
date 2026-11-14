# MFA · 汽车制造环节

汽车制造环节相关的资料与工作区。

## 目录结构

- `gcam_common.py` — 共享工具与全局常量（只读读取、年份归一、半入取整、路径）
- `convert_transport.py` — 新能源汽车服务量 → 车辆保有量
- `convert_power.py` — 发电量(EJ) → 装机容量(MW)，按 tech / Subtech 两级分类
- `build_and_verify.py` — 更新最终双面板工作簿 + 全量核验
- `README.md` — 本说明

## GCAM 保有量与装机容量转换

### 运行命令

使用本机自带 Python（`C:\Users\22965\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe`）：
```powershell
python convert_transport.py    # 服务量→保有量（中间产物 Output/_transport_vehicle_stock.csv）
python convert_power.py        # 发电量→装机明细+Subtech，输出 _power_detail.csv / _power_subtechs.csv
python build_and_verify.py     # 更新最终双面板工作簿并全量核验
```

`build_and_verify.py` 直接更新最终工作簿 `C:\Users\22965\OneDrive\桌面\MFA\Output\能源转型设备所有量.xlsx`：
保留 `车辆保有量_面板` 并扩展到 1990–2050（新增 1990、2005 两年，均为 0），替换 `装机容量_九类_面板`（计算 1990–2050）。运行前请先备份该文件；脚本不会重建车辆面板。

### 输入与年份规则

- 主工作簿：`C:\Users\22965\OneDrive\桌面\关键矿产\数据\数据.xlsx`（只读）
  - `transport_service`、`annual travel per vehicle`、`All factor`（交通）
  - `power generation`、`CF`、`CIPG`（电力；`power generation` 与 `CF` 均沿用 `subcetor` 原拼写）
- CCS 因子补充工作簿：`C:\Users\22965\OneDrive\桌面\关键矿产\数据\GCAM\电力\GCAM_发电量_容量因子匹配_2010-2050.xlsx`（`匹配明细` 工作表）
- 历史容量因子长表：`C:\Users\22965\OneDrive\桌面\关键矿产\数据\GCAM\电力\电力部门_容量因子.xlsx`（`全部` 工作表；仅用于补齐 1990/2005）
- 年份范围：**1990–2050（含首尾）**；统一年份字段后先按范围过滤，再匹配/计算/导出。
- 保留各表实际存在的 GCAM 模型年份：交通表保留 `2020`、电力表保留 `2021`；1990、2005 为源表已有年份，不插值、不补造逐年，也不把区间外年份补进来。

### 交通：服务量 → 车辆保有量

- 匹配：`Region + Sector + Subsector + Technology + Year` 精确匹配单车年行驶里程(VKT)与平均载运系数。
- 载运系数口径：`All factor` 表未标单位，采用「平均载运系数」解释——客运按 人/车、货运按 吨/车。
- 公式（客运/货运同式）：

  ```
  车辆保有量(辆) = 年服务量原值 × 1,000,000 ÷ (单车年行驶里程 × 平均载运系数)
  ```
  - 服务量原值：million pass-km（客运）/ million ton-km（货运）
  - 单车年行驶里程：车公里/辆/年（km/vehicle/year）
  - 平均载运系数：人/车（客运）/ 吨/车（货运）
- 状态：数值 0 服务量 → 输出 0；正服务量遇缺失/非正/冲突因子 → 留空+登记；空白服务量 → 单独标记。

### 电力：发电量 → 装机容量（原始技术逐条）

- 匹配：按 `region + subcetor + technology + year` 匹配主表 `CF`（保留发电表 `Scenario`）。
- **1990/2005 历史因子**：主表 `CF` 只有 2010 起，没有 1990/2005 列；这两年从
  `电力部门_容量因子.xlsx` 的 `全部` 长表补齐。先用主表 2010–2050 的数值反推每个
  `region+subcetor+technology` 对应的来源文件（要求该文件在主表全部年份都与主表一致），
  再用同一文件取 1990/2005 值；地区级缺失时退回全局（`region=None`）同技术。
  不插值、不填零；同键多来源且多值时不猜。
- **水电直接使用主表 `CF` 中匹配到的 `hydro` 因子。**
- **六种 CCS 补缺**：仅对主表缺失因子的以下六种技术，从补充工作簿 `匹配明细` 按
  `Scenario + region + subcetor + technology + year` 精确补齐（补充值唯一匹配）：
  `biomass (IGCC CCS)`、`biomass (conv CCS)`、`coal (IGCC CCS)`、`coal (conv pul CCS)`、
  `gas (CC CCS)`、`refined liquids (CC CCS)`。
- 公式（默认 8760 小时，CF 用 `0 < CF ≤ 1` 小数）：

  ```
  装机容量(MW) = 年发电量(EJ) × 10^9 ÷ (3.6 × 8760 × CF)
  ```
  每个原始技术记录计算未取整 MW。
- CIPG：按 `technology + Year` 关联，保留 `t/GW` 系数与匹配状态；缺年留空、装机仍按有效 CF 计算。
- 状态：数值 0 发电量 → 装机记 0 并记录因子匹配状态；正发电量遇缺失/无效/冲突因子 → 装机留空+登记；空白发电量 → 单独标记。

### 电力分类：tech / Subtech

原始 `technology` 名称保持不变，继续用于匹配 CF / CIPG。分类字段只用于结果汇总与面板输出：

| 原始 technology | tech | Subtech |
| --- | --- | --- |
| `biomass (IGCC CCS)` | `biomass_CCS` | `biomass_IGCC_CCS` |
| `biomass (conv CCS)` | `biomass_CCS` | `biomass_conv_CCS` |
| `coal (IGCC CCS)` | `coal_CCS` | `coal_IGCC_CCS` |
| `coal (conv pul CCS)` | `coal_CCS` | `coal_conv_pul_CCS` |
| `gas (CC CCS)` | `gas_CCS` | `gas_CC_CCS` |
| `refined liquids (CC CCS)` | `refined liquids_CCS` | `refined liquids_CC_CCS` |
| `hydro` | `hydro` | `hydro` |
| `Gen_III`、`Gen_II_LWR` | `nuclear` | `nuclear` |
| `rooftop_pv`、`PV`、`PV_storage` | `PV` | `PV` |
| `wind`、`wind_storage` | `wind` | `wind` |
| `wind_offshore` | `wind` | `wind_offshore` |

- 共 **8 类 tech、11 类 Subtech**；每个 Subtech 只对应一个 tech。
- 执行顺序固定：**匹配原始技术 CF → 分别计算未取整装机 → 按 `Scenario+region+Year+tech+Subtech` 加总 → 对 Subtech 合计四舍五入**。
- `biomass`、`coal` 的不同 Subtech 分别保存，不再先合计成一个 `biomass_CCS` / `coal_CCS` 再拆分。
- `wind` 与 `wind_offshore` 同属 `tech=wind`，但作为两个 Subtech 分别保存。
- 主面板不输出 tech 合计行；`tech` 只作为两级分类标签。
- 某 Scenario-region 下若某 Subtech 全部年份均无源记录，则该 Subtech 不出行；任一年有记录即保留整行，缺年留空。该口径与旧面板“源表无记录则不出行”一致。

### 取整规则

中间值保留全精度，仅最终结果按**传统四舍五入**取整（`.5` 向上，`floor(x+0.5)`）：
`0.5→1、1.5→2、2.5→3、2.4→2、2.6→3`。先汇总未取整值，再对每个 Subtech 取整。

`biomass_CCS` / `coal_CCS` 由“旧类别合计后取整”改为“各 Subtech 分别取整”，
整数合计可能出现 ±1 差异；这是原取整规则的正常结果，未人为调整。

### 输出与核验

最终成果工作簿：`C:\Users\22965\OneDrive\桌面\MFA\Output\能源转型设备所有量.xlsx`

- 两张工作表：
  - `车辆保有量_面板` — 年份扩展到 1990–2050：`Scenario`、`Region`、`Subsector`、`Technology`、`单位`、1990、2005、2010…2050（含 2020）；1990/2005 为 0，2010–2050 既有内容、顺序、数值不变。
  - `装机容量_九类_面板` — **历史表名保留**，实际包含 8 类 tech、11 类 Subtech；列为
    `Scenario`、`region`、`tech`、`Subtech`、`单位`、1990、2005、2010、2015、2021、2025、2030、2035、2040、2045、2050；
    单位保持 `MW`，电力年份保留 2021（不改为车辆的 2020）。
- 本次实际运行核验（修改前备份：`能源转型设备所有量_before_1990_extend_20261114.xlsx`）：
  - 车辆保有量面板 1260 行；2010–2050 与修改前备份逐单元格差异 **0**；新增 1990/2005 全部为 **0**。
  - 装机换算明细 13200 行（1200 源记录 × 11 年）；其中 2010–2050 共 10800 行与旧代码逐键比对，发电量、CF、未取整装机差异均为 **0**；1990/2005 新增 2400 行。
  - Subtech 聚合 11616 行；最终电力面板 852 行（62 个有 CCS 的 Scenario-region × 11 + 34 个无 CCS 的 Scenario-region × 5）。
  - 面板主键 `Scenario+region+tech+Subtech` 重复 **0**；tech 8 类、Subtech 11 类；Subtech→tech 唯一映射通过；2010–2050 既有 7668 个面板值逐格差异 **0**。
  - 1990/2005：483 条正发电记录可计算，1917 条为零发电；历史 CF 从长表匹配 **1656** 条；异常 **0**。
  - 2010–2050 六种 CCS 补充因子记录 **3348** 条；全部原始技术均匹配到有效 CF。
  - 未取整汇总：2010–2050 各 Subtech 与上一版逐组一致（7668 组，差异 **0**）；biomass/coal 拆分与 `tech=wind` 合计口径延续上一版。
  - 取整结果：1:1 映射类别整数差异 **0**；`biomass_CCS` 分别取整差异 59 组、`coal_CCS` 78 组，最大绝对差 **1**，属分别取整的正常差异。
- 年份扩展口径：1990、2005 是源表已有年份，不做 1995/2000 插值；车辆源表这两年服务量为 0，电力历史 CF 按 2010–2050 同源文件从长表取值，不填零。
- 下游兼容性提示：`制造铜需求\copper_demand.py` 仍按旧列名 `归并类别` 读取本面板；本面板现为 `tech` / `Subtech`，该程序需同步调整列名与分类口径后才能继续运行。本任务未修改该下游程序。

### 提交记录（按时间顺序）

| 提交短哈希 | 文件 | 变更说明 |
|---|---|---|
| `11ff247` | `gcam_common.py` | 新增共享读取/年份归一/半入取整工具 |
| `b6b7d54` | `.gitignore` | 忽略 Python 缓存与本地产物 |
| `87353dd` | `convert_transport.py` | 服务量转保有量（2010–2050） |
| `9f1c860` | `convert_power.py` | 发电量转装机容量（EJ→MW，8760h，仅保留匹配 CF 记录） |
| `5487bc8` | `build_and_verify.py` | 合成三表工作簿并全量核验/交叉核对 |
| `a73925b` | `README.md` | 记录转换流程、公式口径与全部提交哈希 |
| `283198a` | `gcam_common.py` | 新增九类装机输出名与 CCS 补充工作簿路径 |
| `97e5a21` | `convert_power.py` | 重写为 CF 表匹配 + 六 CCS 补充 + 九类归并 |
| `11ac151` | `build_and_verify.py` | 改为四表输出并核验九类归并与取整一致性 |
| `408840a` | `README.md` | 更新九类装机流程、映射、核验与提交哈希 |
| `e91dc5b` | `README.md` | 追加 BEV/PHEV 技术占比处理参考文献（前四项） |
| `845c71f` | `split_nev_stock_bev_phev.py` | 新增 BEV/PHEV 拆分脚本（保留小数精度不取整） |
| `927bb6a` | `gcam_common.py` | 新增最终双面板工作簿路径与电力面板年份常量 |
| `34059dc` | `convert_power.py` | 按 tech/Subtech 两级分类重算装机（wind 合并、biomass/coal 细分） |
| `bb0bb59` | `build_and_verify.py` | 补齐最终双面板导出，保留车辆面板并核验 tech/Subtech |
| `9aacd31` | `README.md` | 更新分类映射、最终面板结构、核验结果与下游兼容性说明 |
| `0ffb9ea` | `gcam_common.py`、`convert_transport.py` | 年份范围扩展到 1990–2050，车辆/电力面板加入 1990、2005 |
| `b1ff507` | `convert_power.py` | 1990/2005 从容量因子长表按 2010+ 同源文件补齐 CF 并重算 |
| `8995a04` | `build_and_verify.py` | 最终双面板扩展到 1990–2050，保留既有年份并核验 |
| （本次） | `README.md` | 补充 1990–2050 年份扩展、历史 CF 口径与核验结果 |

## BEV/PHEV 技术占比处理参考文献

本节内容引自《汽车技术占比处理方案》（`预备工作/output/新能源汽车BEV和PHEV占比情况/汽车技术占比处理方案.md`），
记录 BEV/PHEV 技术占比处理及其用于 GCAM 数据库修正时所依据的参考文献。

### 一、技术占比处理的主要参考文献（前三项）

BEV 与 PHEV 技术占比处理采用 Logit 线性化的 Logistic 技术替代模型，其方法依据为：

1. Fisher, J. C., & Pry, R. H. (1971). A simple substitution model of technological change. *Technological Forecasting and Social Change*, 3, 75–88. https://doi.org/10.1016/S0040-1625(71)80005-7
2. Marchetti, C., & Nakicenovic, N. (1979). *The Dynamics of Energy Systems and the Logistic Substitution Model*. IIASA Research Report RR-79-13. International Institute for Applied Systems Analysis.
3. Domarchi, C., & Cherchi, E. (2023). Electric vehicle forecasts: A review of models and methods including diffusion and substitution effects. *Transport Reviews*, 43(6), 1118–1143. https://doi.org/10.1080/01441647.2023.2195687

其中，Fisher 与 Pry（1971）提出技术替代模型，Marchetti 与 Nakicenovic（1979）研究能源系统技术替代与 Logistic 模型，
二者构成 Logit 技术替代模型的理论基础；Domarchi 与 Cherchi（2023）对新能源汽车预测模型进行系统综述，
为技术扩散、技术替代及不同预测假设的处理提供方法参考。

### 二、GCAM 数据库修正所依据的参考文献（第四项）

4. International Energy Agency (IEA). (2026). *Global EV Outlook 2026*. Paris: IEA. https://www.iea.org/reports/global-ev-outlook-2026

GCAM 原有交通技术分类包含 BEV 但缺少 PHEV 这一独立技术类别。本方案依据 IEA《Global EV Outlook 2026》所附
BEV/PHEV 历史存量与未来情景预测数据，构建 BEV 与 PHEV 的存量结构比例，用于向 GCAM 数据库补充 PHEV 技术类别。
该项仅作为 BEV/PHEV 存量数据来源，并非前三项技术占比处理方法论的共同来源。
