# 退货入库分配系统

这是一个基于 Excel VBA 的退货入库分配系统。系统读取退单、质检库存和配置数据，完成标准化、校验、库存分配、回溯、整单回滚与结果输出。

## 当前进度

更新日期：2026-07-19

- M01～M15 核心业务模块已实现。
- M16 单元与端到端测试入口 `modTestRunner.bas` 已实现。
- 批量 Excel 回归入口 `modBatchTestRunner.bas` 已实现，2026-07-19 起正式编号 **M17**。
- 可直接使用的生产工作簿 `退货入库分配系统.xlsm` 已生成，含 9 张工作表、15 个生产模块和 3 个操作按钮。
- 测试索引共追踪 58 条 TC 行：53 条测试资产齐备、4 条嵌入覆盖、1 条不适合自动测试，**模块缺口已清零**。
- 调试日志代码、表头与预期数据行已统一为 19 列（38 行预期数据 2026-07-19 按实际运行重建）。
- 运行历史记录表口径已拍板为 20 列（需求 17 字段 + 3 个配置快照列），需求、代码、测试与生成器已同步。
- 2026-07-19 实际验收（无头执行）：`RunAllTests` 491/491、TC17 81/81、TC18 6/6；批量回归 9 计划 13 子批次 659/659 断言通过。
- 2026-07-19 对抗性端到端（生产工作簿）：S1~S6 全通过——正常业务流、重跑覆盖、E01 拦截、空输入、详细调试日志、500 物流单号性能基线（51 秒）。

> 回溯次数口径：实现按“单行撤销-重试”每回退一行计 1 次（如 SF0016=4、SF0027=59、SF0028=11）；
> 早期文档按“直达 target 的回溯事件数”叙述，两种口径终态一致，数据以实际运行计步为准。
> 调试日志第 12/13 列（分配前/后QC剩余）当前实现为占位符 `-`，见 `调试日志19列规格说明.md` §8。

## 新手阅读顺序

1. 先读本文件，了解入口和当前进度。
2. 阅读 [`规则-TC-模块-数据集映射表.md`](规则-TC-模块-数据集映射表.md)，查看测试覆盖和缺口。
3. 阅读 [`退货入库分配系统_需求与技术方案.md`](退货入库分配系统_需求与技术方案.md) 的第 1～4 章，理解业务规则和分配流程。
4. 阅读 [模块划分方案](.cursor/plans/模块划分方案_2bd2f2bf.plan.md)，理解 M01～M17 的职责。
5. 按需打开 `TC-*.md`，查看具体输入、推导和预期结果。
6. 调试日志相关修改以 [`调试日志19列规格说明.md`](调试日志19列规格说明.md) 为准。

## 主要文件

| 文件 | 作用 |
|---|---|
| `退货入库分配系统_需求与技术方案.md` | 业务规则与技术方案的权威来源 |
| `规则-TC-模块-数据集映射表.md` | TC、规则、模块、DataSet 和进度索引 |
| `测试场景清单` | 测试目的和边界场景说明 |
| `TC-*.md` | 每个测试用例的详细输入、推导和预期结果 |
| `modTypes.bas`～`modRunner.bas` | M01～M15 生产代码 |
| `modTestRunner.bas` | M16 内存单元测试与端到端测试 |
| `modBatchTestRunner.bas` | M17 基于 Excel 工作簿的批量回归测试 |
| `生成测试数据Excel.ps1` | 生成部分独立测试 DataSet |
| `生成生产工作簿.ps1` | 从 M01～M15 源码生成可直接使用的生产 `.xlsm` |
| `同步测试工作簿VBA.ps1` | 备份并同步 17 个测试 VBA 模块，处理 UTF-8/ANSI 导入兼容 |
| `合并预期结果.ps1` | 幂等合并标准输入、配置和预期结果（带占用检测与自动备份） |
| `运行无头验收.ps1` | 无头执行全部测试入口（RunAllTests/TC17/TC18/批量回归）并判定通过与否 |
| `无头业务验证.ps1` | 对生产工作簿做对抗性端到端验证（S1~S6 场景） |

## 业务用户如何使用

1. 打开 `退货入库分配系统.xlsm`，点击“启用内容/启用宏”。
2. 将数据粘贴到 `输入_退单表` 和 `输入_质检库存表`；行号必须是五位纯数字文本，例如 `00001`。
3. 在“操作面板”先点击“仅运行校验”，修正异常后再点击“开始分配”。
4. 需要重跑时点击“清空结果”；配置和运行历史不会被删除。

开发者如需从源码重建工作簿，请运行 `生成生产工作簿.ps1`。不要直接导入 UTF-8 `.bas` 文件；Excel VBA 会按 Windows ANSI 解码，可能破坏中文字符串。

按钮对应的无参数入口为 `StartValidationOnly`、`StartFullAllocation`、`ClearAllocationResults`。

## 如何运行测试

**推荐（无头，无需打开 Excel）**：在项目目录执行

```powershell
powershell -ExecutionPolicy Bypass -File .\运行无头验收.ps1
powershell -ExecutionPolicy Bypass -File .\无头业务验证.ps1
```

- `运行无头验收.ps1`：依次执行 `RunAllTests`、`RunSingleTest 17`、`RunSingleTest 18`、`RunBatchTestPlan`，输出计数并判定通过/失败（带防呆：无启用批次或结果非本次运行会判失败）。
- `无头业务验证.ps1`：在临时副本上对生产工作簿跑 S1~S6 对抗性场景（正常流、重跑覆盖、E01 拦截、空输入、详细调试日志、500 单性能基线）。

**人工方式**：在 VBA 编辑器中按 `Ctrl + G` 打开“立即窗口”，按需执行：

```vba
RunAllTests
RunSingleTest 17
RunSingleTest 18
RunBatchTestPlan
```

- `RunAllTests`：运行内存单元测试和端到端测试，不读取外部 DataSet。
- `RunSingleTest 17`：运行文件集成测试，需要相关 `.xlsm/.xlsx` 文件。
- `RunSingleTest 18`：运行批量测试器的内存冒烟测试。
- `RunBatchTestPlan`：按“批量测试计划”工作表执行 Excel 批量回归（标准回归批次已启用）。

不要把 `RunAllTests` 通过理解为“所有 Excel 回归均通过”；文件集成测试和批量回归需要另行执行。
`BATCH-ERR-STRUCTURE`、`BATCH-ERR-CONFIG` 依赖独立异常工作簿，请使用 `RunSingleTest 17`，不要在标准汇总工作簿中启用。

## PowerShell 工具

在项目目录打开 PowerShell 后运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\生成测试数据Excel.ps1 -Target OnlyM05
powershell -ExecutionPolicy Bypass -File .\生成测试数据Excel.ps1 -Target OnlyPendingTC
powershell -ExecutionPolicy Bypass -File .\合并预期结果.ps1
powershell -ExecutionPolicy Bypass -File .\生成生产工作簿.ps1
powershell -ExecutionPolicy Bypass -File .\同步测试工作簿VBA.ps1
powershell -ExecutionPolicy Bypass -File .\升级预期调试日志19列.ps1
powershell -ExecutionPolicy Bypass -File .\运行无头验收.ps1
powershell -ExecutionPolicy Bypass -File .\无头业务验证.ps1
```

运行会修改 Excel 的脚本前，请先关闭目标工作簿（相关脚本已内置占用检测，会在被占用时中止并提示）。日常合并使用已验证幂等的 `合并预期结果.ps1`（执行时自动在上级目录留时间戳备份）；`直接合并.ps1` 与 `修复合并.ps1` 是职责重叠的历史修复工具，目前仅保留供排障参考，未确认前不要运行、移动或删除。

## 当前待办

无。2026-07-19 的六项历史待办已全部闭环：

1. ~~补齐 TC-43、TC-45 的自动化回归资产~~（TC-43 → `TestRunner_RerunOverwrite_TC43`；TC-45 → BATCH-TC45A/B1/B2 + M17 扩展结构化断言）
2. ~~为 TC-35 设计专用 DataSet 和规范 TC~~（SF0035 + BATCH-TC35 + `TC-35_大小批量混合场景测试用例.md`）
3. ~~补全五套 19 列调试日志预期数据行~~（38 行按实际运行重建，修正批次列错位）
4. ~~决定 `modBatchTestRunner` 模块编号~~（拍板为独立 M17）
5. ~~运行历史记录表字段口径~~（拍板为 20 列：需求 17 字段 + 3 配置快照列，需求/代码/测试/生成器已同步）
6. ~~`合并预期结果.ps1` 占用检测与自动备份~~（已内置）

如新增需求或规则变更，请按“文档维护”流程同步需求、映射表、场景清单与相关 TC。

## 项目维护提醒

- 当前目录不是 Git 仓库，修改代码或 Excel 前应手工备份；更稳妥的做法是初始化 Git，并排除临时锁文件。
- 修改业务规则时，要同步检查需求文档、映射表、场景清单和相关 `TC-*.md`。
- 不要把 `.cursor/plans` 中的待办状态当作真实开发进度；真实状态以代码、测试执行结果和映射表为准。
