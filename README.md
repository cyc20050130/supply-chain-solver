# supply-chain-solver

供应链 `.xls` 启发式求解器，用于读取采购、生产、销售、综合运营案例表格，快速生成可填报的 HTML 方案。

仓库包含 12 个课程/演示样例 `.xls`。这些样例用于验证求解器能从 clone 后直接运行；正式使用时可以换成同结构的新场次文件。

## 功能

- 支持中文路径和带特殊符号的 `.xls` 文件名。
- 自动识别采购、生产、销售、综合题型。
- 固定按 30 天常规周期处理，硫磺等特殊案例按内置场景配置。
- 销量预测使用移动加权平均和偏差修正，不使用神经网络。
- 生成同目录 `*_求解方案.html`，包含采购量、生产计划、路线、承运商、运输数量、起运日期、趟数、频率和结果状态。
- 使用人工确认的单段运输时间补中转起运日；没有图片证据的中转段不会编日期。

## 环境要求

- Python 3.10 或更高版本
- LibreOffice，用于把 `.xls` 转换为 `.xlsx`

Windows 默认会自动查找 `C:\Program Files\LibreOffice\program\soffice.com`。如果你的安装位置不同，可以设置环境变量：

```powershell
$env:SUPPLY_CHAIN_SOFFICE="C:\Program Files\LibreOffice\program\soffice.com"
```

macOS/Linux 只要 `soffice` 或 `libreoffice` 在 `PATH` 中即可。

## 安装

```powershell
git clone https://github.com/cyc20050130/supply-chain-solver.git
cd supply-chain-solver
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -X utf8 solve.py --check-env
```

macOS/Linux：

```bash
git clone https://github.com/cyc20050130/supply-chain-solver.git
cd supply-chain-solver
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -X utf8 solve.py --check-env
```

## 使用

运行单个案例：

```powershell
python -X utf8 solve.py "采购\白砂糖采购★★-标准版个人练习（17_38）场.xls"
```

运行仓库内全部样例：

```powershell
python -X utf8 solve.py --all
```

使用易木平台前端数据生成高分版或极限版：

```powershell
python -X utf8 solve.py --mode high --har "物流计划.html" "销售计划.html" "销售\热水器销售★★★-标准版个人练习（10_57）场.xls"
python -X utf8 solve.py --mode extreme --har "物流计划.html" "销售计划.html" "销售\热水器销售★★★-标准版个人练习（10_57）场.xls"
```

`--har` 可以同时接收一个或多个易木平台导出的 HAR、HTML 或 JS 前端文件。若同时有运输计划和销售计划前端代码，建议两份都传入；求解器会合并承运商、单段距离、单段运输天数和页面计划数据，并在每个案例求解前清空路线缓存，避免不同场次串数据。高分版优先快速给出满足率 100% 的可填方案；极限版在同样满足率约束下继续压低成本。

运行评分公式自检：

```powershell
python -X utf8 solve.py --self-test
python -m unittest test_score_platform.py
```

热水器销售极限版默认约 1-3 分钟生成；更慢的全局/子集反例审查通过 `SUPPLY_CHAIN_HEATWATER_*_AUDIT` 环境变量手动开启，不进入默认快路径。

输出 HTML 会写在输入 `.xls` 同目录，例如：

```text
采购\白砂糖采购★★-标准版个人练习（17_38）场_求解方案.html
```

## 目录

```text
solve.py                 主入口
carrier_infer.py         承运商推断辅助
score_replay.py          平台评分复现辅助
test_score_platform.py   评分与仿真回归测试
采购/ 生产/ 销售/ 综合/   12 个样例 .xls
```

## 注意

- 最终平台成绩以平台回执为准；求解器内的分数用于方案比较和已知规则复现。
- 本仓库不发布历史截图、OCR 缓存、平台回执 HTML、生成方案 HTML 或求解中间产物。
- 样例数据仅用于课程/演示场景，请按你的实际授权使用。

## 已知限制

运输计划中的中转起运日只根据已人工确认的单段运输日期生成。对于样例历史截图中暂未找到清晰单段日期证据的线路，求解器会保留路线和运输数量，但不会自动补写无法确认的后续起运日，避免把整线提前期或经验估计误写成图片证据。

当前仍待补充图片证据的单段包括：

- 节能灯：中山火车站 -> 福州火车站；福州火车站 -> 三明门店、泉州门店。
- 热水器：惠州火车站 -> 石家庄火车站；东莞工厂 -> 洋山码头、盐田码头；盐田码头 -> 横滨码头；苏州火车站 -> 武汉火车站。
- 羽绒服：东莞拉链供应商 -> 广州火车站。
- 蓄电池：广州火车站 -> 石家庄火车站；广州蓄电池厂 -> 洋山港、盐田港；盐田港 -> 横滨码头；东莞塑料厂 -> 广州火车站；广州火车站 -> 苏州火车站；常州塑料厂 -> 苏州火车站；贵阳化工厂 -> 贵阳火车站；贵阳火车站 -> 广州火车站、苏州火车站。
