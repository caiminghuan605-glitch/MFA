# MFA · 汽车制造环节

汽车制造环节相关的资料与工作区。

## 目录结构

- `gcam_common.py` — 共享工具与全局常量（只读读取、年份归一、半入取整、路径）
- `convert_transport.py` — 新能源汽车服务量 → 车辆保有量
- `convert_power.py` — 发电量(EJ) → 装机容量(MW)，并按九类归并
- `build_and_verify.py` — 合成四表工作簿 + 全量核验 + 手算 + 一致性核对
- `README.md` — 本说明

## GCAM 保有量与九类装机容量转换

### 运行命令

使用本机自带 Python（`C:\Users\22965\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe`）：
```powershell
python convert_transport.py    # 服务量→保有量，输出 Output/_transport_vehicle_stock.csv
python convert_power.py        # 发电量→装机(明细+九类)，输出 _power_detail.csv / _power_categories.csv
python build_and_verify.py     # 合成最终四表工作簿并全量核验
```

### 输入与年份规则

- 主工作簿：`C:\Users\22965\OneDrive\桌面\关键矿产\数据\数据.xlsx`（只读）
  - `transport_service`、`annual travel per vehicle`、`All factor`（交通）
  - `power generation`、`CF`、`CIPG`（电力；`power generation` 与 `CF` 均沿用 `subcetor` 原拼写）
- CCS 因子补充工作簿：`C:\Users\22965\OneDrive\桌面\关键矿产\数据\GCAM\电力\GCAM_发电量_容量因子匹配_2010-2050.xlsx`（`匹配明细` 工作表）
- 年份范围：**2010–2050（含首尾）**；统一年份字段后先按范围过滤，再匹配/计算/导出。
- 保留各表实际存在的 GCAM 模型年份：交通表保留 `2020`、电力表保留 `2021`，不互改年、不补造逐年、不用区间外因子填补缺年。

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

### 九类归并

执行顺序固定：**匹配原始技术 CF → 分别计算未取整装机 → 按九类加总 → 对类别合计四舍五入**。

| 归并类别 | 原始技术 |
| --- | --- |
| `biomass_CCS` | `biomass (IGCC CCS)`、`biomass (conv CCS)` |
| `coal_CCS` | `coal (IGCC CCS)`、`coal (conv pul CCS)` |
| `gas_CCS` | `gas (CC CCS)` |
| `hydro` | `hydro` |
| `nuclear` | `Gen_III`、`Gen_II_LWR` |
| `refined liquids_CCS` | `refined liquids (CC CCS)` |
| `PV` | `rooftop_pv`、`PV`、`PV_storage` |
| `wind` | `wind`、`wind_storage` |
| `wind_offshore` | `wind_offshore` |

- 归并键：`Scenario + region + year + 归并类别`。
- 类别内有源记录装机无法计算 → 类别合计留空并注明缺项；某情景/地区/年份下完全无某类源记录 → 保留该类别并标记「源表无记录」；映射表之外的技术登记为异常。

### 取整规则

中间值保留全精度，仅最终结果按**传统四舍五入**取整（`.5` 向上，`floor(x+0.5)`）：
`0.5→1、1.5→2、2.5→3、2.4→2、2.6→3`。九类合计先汇总未取整值再取整。

### 输出与核验

成果工作簿：`C:\Users\22965\OneDrive\桌面\MFA\Output\GCAM_保有量与九类装机容量_2010-2050.xlsx`（Git 仓库外）

- 四个工作表：`车辆保有量`、`装机换算明细`、`装机容量_九类`、`异常与核验`
- SHA-256：`9ed7b96eca1fd871523896f1c4b66312888e48cf9e1231e200ed59bcdc5131d5`
- 核验摘要：
  - 车辆保有量 5670 行；装机换算明细 10800 行（1200 源记录 × 9 年）；装机容量_九类 7776 行（864 情景-地区-年份 × 9 类）
  - 交通年份含 2010/2020/2050、电力年份含 2010/2021/2050，三张结果表均无范围外记录，匹配键无重复，九类映射无未覆盖技术
  - 补充因子使用记录数 **3348**（六种 CCS）；异常 **0**（所有原始技术均匹配到有效 CF）
  - 手算（不同 CF 合并项）：`rooftop_pv(CF 0.2400)+PV(CF 0.2824)+PV_storage(CF 0.2824) = 20.9001 → 21 MW`
  - 取整前一致核对：原始技术装机合计与九类装机合计一致 6552 / 不一致 0

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
| （本次） | `README.md` | 更新九类装机流程、映射、核验与提交哈希 |

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
