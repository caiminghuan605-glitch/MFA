# MFA · 汽车制造环节

汽车制造环节相关的资料与工作区。

## 目录结构

- `gcam_common.py` — 共享工具与全局常量（只读读取、年份归一、半入取整、路径）
- `convert_transport.py` — 新能源汽车服务量 → 车辆保有量
- `convert_power.py` — 发电量(EJ) → 估算装机容量(MW)
- `build_and_verify.py` — 合成三表工作簿 + 全量核验 + 交叉核对 + 手算
- `README.md` — 本说明

## GCAM 保有量与装机容量转换

### 运行命令

使用本机自带 Python（`C:\Users\22965\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe`）：
```powershell
python convert_transport.py    # 服务量→保有量，输出 Output/_transport_vehicle_stock.csv
python convert_power.py        # 发电量→装机容量，输出 Output/_power_installed_capacity.csv
python build_and_verify.py     # 合成最终工作簿并全量核验
```

### 输入与年份规则

- 主输入：`C:\Users\22965\OneDrive\桌面\关键矿产\数据\数据.xlsx`（只读）
  - `transport_service`、`annual travel per vehicle`、`All factor`（交通）
  - `power generation`、`CIPG`（电力）
- 容量因子：`C:\Users\22965\OneDrive\桌面\关键矿产\数据\GCAM\电力\电力部门_容量因子.xlsx`（`electricity` 页）
- 年份范围：**2010–2050（含首尾）**；统一年份字段后先按范围过滤，再匹配/计算/导出。
- 基准年：GCAM 基准年为 2021（`BASE_YEAR=2021`）。各源表保留区间内实际存在的模型年份：
  - 交通表：2010、2015、**2020**、2025、2030、2035、2040、2045、2050（源表无 2021 列，保留 2020）
  - 电力表：2010、2015、**2021**、2025、2030、2035、2040、2045、2050（源表无 2020 列，保留 2021）
  - 不互改年、不补造逐年、不使用区间外因子填补缺年。

### 交通：服务量 → 车辆保有量

- 匹配：`Region + Sector + Subsector + Technology + Year` 精确匹配单车年行驶里程(VKT)与平均载运系数。
- 载运系数口径：`All factor` 表未标单位，采用「平均载运系数」解释——客运按 人/车、货运按 吨/车（依据：GCAM 服务量单位为 million pass-km / million ton-km，除以每车行驶里程后即得每车平均载客/载货量）。
- 公式（客运/货运同式）：

  ```
  保有量(辆) = 年服务量原值 × 1,000,000 ÷ (单车年行驶里程 × 平均载运系数)
  ```
  - 服务量原值：million pass-km（客运）/ million ton-km（货运）
  - 单车年行驶里程：车公里/辆/年（km/vehicle/year）
  - 平均载运系数：人/车（客运）/ 吨/车（货运）
- 状态：服务量数值 0 → 输出 0（`zero_service`）；正服务量遇缺失/非正/冲突因子 → 留空+登记；空白服务量 → `blank_service`。

### 电力：发电量 → 估算装机容量

- 保留字段：`Scenario、region、subcetor(原拼写)、output、technology、units`；区分 `electricity` 与 `elect_td_bld`（含屋顶光伏）。
- 容量因子匹配（从 `electricity` 页）：`region+output+subcetor+technology+Year` ↔ `region+sector+subsector+technology+year`。
  - 地区专用值优先；空白地区记录仅当确认为该技术/年份适用的 GCAM 通用值时采用，并注明回退来源。
  - 三级回退：地区专用 → 全局(技术级) → 全局(子部门级)；冲突值逐项核查，不取首条/平均。
  - **只保留能匹配到有效 CF（0<CF≤1）的记录；匹配不上则从结果表排除并登记**（hydro 无任何 CF，全部排除）。
- 公式（默认 8760 小时，依据 EIA 容量因子定义 + 与旧参考 CSV 反算一致）：

  ```
  估算装机容量(MW) = 年发电量(EJ) × 10^9 ÷ (3.6 × 8760 × CF)
  ```
- CIPG：单位 `t/GW`，按 `technology + Year` 精确关联，系数原值与单位保留供后续 MFA；缺失年份留空并标记（不阻断容量计算）。

### 取整规则

中间值保留全精度，仅最终结果按**传统四舍五入**取整（`.5` 向上，`floor(x+0.5)`）：
`0.5→1、1.5→2、2.5→3、2.4→2、2.6→3`。计算得到的零写为数值 `0`，缺失结果留空。

### 已知异常

- 交通：无因子缺失/冲突异常（210 个 BEV 组合全部被 VKT 与 All factor 覆盖）。
- 电力：`hydro` 无适用容量因子，96 条记录 × 9 年 = **864 条被排除**，列于「异常与核验」。
- CIPG 缺失年份：2010、2015、2025（CIPG 仅有 2021、2030–2060），对应装机行 CIPG 留空并标记。

### 输出与核验

成果工作簿：`C:\Users\22965\OneDrive\桌面\MFA\Output\GCAM_保有量与装机容量.xlsx`（位于 Git 仓库外）

- 三个工作表：`车辆保有量`、`发电装机容量`、`异常与核验`
- SHA-256：`dfd0e229fe4f013234bc60d5aa0216bfa20b469359cd089474ba58c60698697f`
- 核验摘要：
  - 车辆保有量 5670 行（630 源记录 × 9 年）、发电装机容量 18765 行（匹配 CF 成功）、异常 864 条（均为 hydro 排除）
  - 交通年份含 2010/2020/2050、电力年份含 2010/2021/2050，三表均无范围外记录，匹配键无重复
  - 手算：客运 `104.452×1e6÷(10391×2.3333)=4308` 辆、货运 `135.288×1e6÷(67800×5.38)=371` 辆、发电 `0.00063587×1e9÷(3.6×8760×0.8)=25` MW，均与表内一致
  - 交叉核对（旧参考 CSV）：容量因子差异 0；同发电量记录 MW 全吻合（7503/7503）；发电量/服务量存在版本差异（数据.xlsx 为更新版本），绝对装机按当前数据计算

### 提交记录（每次 Python 修改单独提交）

| 提交短哈希 | 文件 | 变更说明 |
|---|---|---|
| `11ff247` | `gcam_common.py` | 新增共享读取/年份归一/半入取整工具 |
| `87353dd` | `convert_transport.py` | 服务量转保有量（2010–2050） |
| `9f1c860` | `convert_power.py` | 发电量转装机容量（EJ→MW，8760h，仅保留匹配 CF 记录） |
| `5487bc8` | `build_and_verify.py` | 合成三表工作簿并全量核验/交叉核对 |
| （本次） | `README.md` | 记录流程、公式、口径与全部提交哈希 |
