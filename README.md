# supply-chain-solver

供应链 `.xls` 启发式求解器，用于读取采购、生产、销售、综合运营案例表格，快速生成可填报的 HTML 方案。

仓库包含 12 个课程/演示样例 `.xls`。这些样例用于验证求解器能从 clone 后直接运行；正式使用时可以换成同结构的新场次文件。

## 功能

- 支持中文路径和带特殊符号的 `.xls` 文件名。
- 自动识别采购、生产、销售、综合题型。
- 固定按 30 天常规周期处理，硫磺等特殊案例按内置场景配置。
- 销量预测使用移动加权平均和偏差修正，不使用神经网络。
- 生成同目录 `*_求解方案.html`，包含采购量、生产量、销售量、路线、承运商、运输数量、起运日期、趟数、频率和结果状态。
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

运行评分公式自检：

```powershell
python -X utf8 solve.py --self-test
python -m unittest test_score_platform.py
```

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

