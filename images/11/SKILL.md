---
name: lark-inventory-replenish
version: 1.0.0
description: "仓库备货计算 —— 把快麦导出的库存 Excel 全量覆盖到飞书多维表格「商品采购总表」的「报表2」，触发下游重算后导出「商品采购清单表」Excel。当消息里出现「仓库备货计算」、收到「快麦导出_库存」文件、或编排层以 skill key `inventory_replenish` 派发时使用。"
metadata:
  requires:
    bins: ["lark-cli"]
  cliHelp: "lark-cli base --help"
agent_created: true
---

# 仓库备货计算（能力层）

> **一句话**：给一份快麦导出的库存 Excel，收回一份能直接发采购的「商品采购清单表」。
>
> **干什么**：把上传表的数据按行序整体覆盖到飞书多维表格「商品采购总表」的「报表2」（不删记录、不动字段结构，只改单元格内容），让下游 lookup / 公式自动重算，再把重算结果「商品采购清单表」导出成 xlsx。
>
> **什么时候用**：飞书里出现「仓库备货计算」指令、或收到 `快麦导出_库存*.xlsx` 文件、或编排层 `lark-bot-excel-pipeline` 以 skill key `inventory_replenish` 派发过来。
>
> **前置条件**：本机已安装并授权 lark-cli（用户身份需多维表格读写权限）；Python 环境 `~\.workbuddy\binaries\python\envs\default` 已装 pandas / openpyxl。
>
> **产出**：`<workdir>/商品采购清单表.xlsx`（主交付物）+ `export_summary.json`（给编排层写回传文案）。
>
> **分工**：监听消息、下载附件、回卡片、回传文件由编排层 `lark-bot-excel-pipeline` 负责；本 skill 只做「输入 Excel → 产出 Excel」，不管通讯。

> **分层定位**：本 skill 是**能力层**，只做「输入 Excel → 产出 Excel」。
> **监听、下载、回传、回卡片由编排层 `lark-bot-excel-pipeline` 负责**，`chat_id` / `message_id` 由编排层传入。
> 本 skill 唯一的对外交付物是工作区里的 xlsx 和 `*_summary.json`（供编排层生成回传文案）。

## 输入 / 输出契约

| 项 | 值 |
|---|---|
| 输入 | `<workdir>/input.xlsx`（或任意扩展名的同结构文件）—— 快麦导出，sheet「报表1」，**第1行标题、第2行表头、第3行起数据**，只用前 9 列 |
| 输出 | `<workdir>/商品采购清单表.xlsx`（主交付物）、`export_summary.json`（给编排层写文案用） |
| 常量 | base `CTrWbbPeMa8jG1swtELcdcDKnMK`；报表2 `tblRzg7Ya6YuJB5i`；采购清单表 `tblfXFA0BYiX9Ybo` |

Excel 列位（脚本内约定，列索引从 0）：

| 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 |
|---|---|---|---|---|---|---|---|---|
| 序号(不写) | **规格商家编码(要写)** | 实际可用数 | 采购在途数 | 昨日销量 | 3天销量 | 3天平均销量 | 7天销量 | 7天平均销量 |

## 写入语义（重要，别按旧印象操作）

1. **配对方式 = positional（按位置）**：上传表第 N 个数据行 → 报表2 序号为 N 的那一行。
   **不做任何编码匹配、不依赖两侧排序一致**（原因见避坑第 19 条：两侧行序本来就不同）。
2. **全量覆盖**：不比较新旧值，配上的行一律整体覆写。
3. **写入字段 = 规格商家编码 + 7 个度量，共 8 个**。编码必须一起写，否则会出现「数字是新的、编码还是旧的」这种半新半旧的行。
4. **不写「序号」、不删记录、不新建记录**（除非显式 `--append`）。`record_id` 原地覆写单元格，下游 lookup 引用不断。
5. Excel 行数 > 报表2 行数时多出来的行 → `overflow_rows`，默认**不进表**。要尾部追加得显式传 `--append`（会写「序号」= 当前最大序号 + k，否则下游 lookup 匹配不到）。

## 执行流程（5 步）

```powershell
$py  = "C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
$ps  = "C:\Users\Administrator\.workbuddy\skills\lark-inventory-replenish\scripts\pipeline.py"
$br  = "C:\Users\Administrator\.workbuddy\skills\lark-inventory-replenish\scripts\batch-run.ps1"
$node= "C:\Users\Administrator\.workbuddy\binaries\node\versions\22.22.2-3\node.exe"
$cli = "C:\Users\Administrator\.workbuddy\binaries\node\versions\22.22.2-3\node_modules\@larksuite\cli\scripts\run.js"
$flds = @(); foreach ($f in @("fld47Q4mNq","fld42dEGf0","fld1aXBYoL","fld2v20eBy","fld7AvAqMl","fld6rY4blT","fldqdXyDdF","fld4ewrgIi","fld6cmfHeM")) { $flds += @("--field-id",$f) }
```

> ⚠️ `.ps1` 必须存成 **UTF-8 with BOM**，否则 Windows PowerShell 5.1 按本地编码读取会直接报语法错。改完 `batch-run.ps1` 后务必转 BOM + 用 `[System.Management.Automation.Language.Parser]::ParseFile` 校验。

**1. 并发拉报表2 基线**（落 `p_*.ndjson`，只取 `record_id` + 序号）

```powershell
powershell -ExecutionPolicy Bypass -File $br -WorkDir <workdir> -Mode pull -Concurrency 8 -PageSize 2000
# -> pull done rows=20000 files=10 rounds=3 ms=12441
```

自动探测行数、轮空就停，**不要写死 20000**。旧的手工分页拉法要 39 秒，这个 12 秒。

**2. prep**：解析 Excel → 按位置生成批次

```powershell
& $py $ps prep <input.xlsx|input.csv> <workdir>            # 默认 positional
& $py $ps prep <src> <workdir> positional --append         # Excel 比报表2 长时尾部追加
```

`normalize_input()` 按**文件魔数**嗅探真实格式：快麦导出的「`.csv`」实测内容就是 xlsx（头 4 字节 `PK\x03\x04`），openpyxl 按扩展名校验会直接抛 `InvalidFileException`。所以**上游不用管文件叫什么**，直接把下载下来的原文件丢给 `prep`；真 CSV 会探测编码（`utf-8-sig`→`gbk`→`gb18030`→`utf-8`）后转成 `input.xlsx`。

看 `summary.json`：

| 字段 | 含义 | 正常值 |
|---|---|---|
| `write_rows` | 本轮写入行数 | 应 = `base_rows`（全覆盖） |
| `skipped` / `not_covered_rows` | 没被覆盖的行 | **positional 模式下必须是 0** |
| `overflow_rows` | Excel 比报表2 多出的行 | 容量不够，记进文案 |
| `batch_files` | 生成的 `upd_*.json` 数量 | = ceil(write_rows/200) |
| `base_coverage` | 覆盖率 | ≈1.0 |

- 报 `error: no_baseline` → 基线没拉到本目录（或被 sweep 当残留挪走），回去补第 1 步
- 报 `error: base_seq_not_continuous` → 报表2 序号不是严格 1..N，**立刻停手**。多半是拉基线漏了页（并发过高会静默丢数据），重跑第 1 步
- **硬失败**：`base_coverage < 0.5` 或 `write_rows == 0` → 别写，直接回报

> `changed_rows` 与 `write_rows` 同值，前者是历史键名（编排层文案模板可能引用），保留勿删。

**3. 并发写入**

```powershell
powershell -ExecutionPolicy Bypass -File $br -WorkDir <workdir> -Mode update -Concurrency 8
# -> update: batches=100 ok=100 fail=0 ms=...
```

> ⚠️ **不要用 `foreach` 逐个串行调 CLI**（每批 ~1.4s，100 批要 140s）。用 `batch-run.ps1`，内部走 runspace 线程池。串行时 Python/PS 起子进程会被沙箱拦，只有这一种并行方式可用（详见避坑第 21 条）。

**Excel 行数 > 报表2 行数时**（`overflow_rows > 0` 且传了 `--append`）：

```powershell
powershell -ExecutionPolicy Bypass -File $br -WorkDir <workdir> -Mode create -Concurrency 8
```

**4. 抽样验证**（只回读 ≤200 条 changed 行，不回读全表）

**必须先用 `sample` 子命令生成 `ids.json`**，不要自己写脚本抽样：

```powershell
& $py $ps sample <workdir> 200             # -> ids.json (固定种子, 可复现)
& $node $cli base +record-get --base-token CTrWbbPeMa8jG1swtELcdcDKnMK --table-id tblRzg7Ya6YuJB5i `
  --json "@ids.json" $flds --format ndjson --output v_sample.ndjson --overwrite --as user
& $py $ps verify <workdir>                 # -> verify.json, passed=true 即通过
```

> ⚠️ **踩过的坑**：早期 `verify` 内部自己再抽一次样（种子 20240917），而 agent 侧用别的种子生成 `ids.json`——两边抽到的不是同一批 ID，回读 200 条却报 193 条 `not_in_pulled_set`，`passed=false`，纯属假警报。
> 现已改为：**抽样只由 `sample` 负责**，`verify` 不再抽样，改为「回读多少就验证多少」，零误差即通过。若仍报 `pulled_not_in_expected > 0`，说明回读到了非本次变更的记录，才是真异常。

**5. 导出采购清单**

**不要手动列 `--field-id`**——实测该表只有 1018 行，不指定字段一次拉全即可，且 ndjson 直接给中文列名（`序号`/`系列`/`款号-颜色`/6 个尺码件数），正是 `export` 需要的 key：

```powershell
& $node $cli base +record-list --base-token CTrWbbPeMa8jG1swtELcdcDKnMK --table-id tblfXFA0BYiX9Ybo `
  --offset 0 --limit 2000 --format ndjson --output po.ndjson --overwrite --as user
& $py $ps export <workdir> po.ndjson
```

实际列（20 个）：`1码-S（件数）` `2XL色键` `2码-M（件数）` `3XL色键` `3码-L（件数）` `4码-XL（件数）` `5码-2XL（件数）` `6码-3XL（件数）` `L色键` `M色键` `S色键` `XL色键` `record_id` `备注` `序号` `来货时间` `款号-颜色` `父记录` `父记录 2` `系列`。

**6. 输出**：把 `export_summary.json` 交给编排层，由它回传 xlsx + 文案。

## 为什么可以省掉二次验证

| 原流程 | 现在 | 理由 |
|---|---|---|
| 先写 200 条试点再回读 | **删掉** | 字段映射已固化在脚本里，写入失败/类型不符飞书会直接报错，不会静默写错 |
| 全量回读 2 万行比对 | **抽样 200 条** | 写入是同质批量操作；抽样零误差即可证明格式映射正确，剩余风险极低 |
| 抽查下游清洗表 | **删掉** | 已验证过：清洗表对报表2 是 lookup 按 `序号` 值匹配，不依赖 record_id，原地更新必然生效 |
| 按编码配对、只写变化的行 | **positional 全量覆盖** | 用户要求「上传什么表，报表2 就是什么内容」，不比对、不增量、零残留。代价是每轮都写满（2 万行），但有并发写入兜底，实测写入只要 ~27 秒 |

**必须保留的唯一校验：prep 阶段的对齐检查**（纯本地、零成本）。它能拦住 90% 的坑——列错位、编码对不上、Excel 格式变了。

## 避坑经验（都是踩过的）

1. **Python 起 node 子进程会被沙箱拦**。写入批次必须由 agent 用 PowerShell 逐个调 CLI，不要写进 Python 脚本里跑。
2. **PowerShell 不回显 stdout**。所有 CLI 输出重定向到文件再用 Read 读；中文输出前先 `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8`。
3. **Bash 工具不可用**（PATH 被清空，`ls`/`head` 都 not found）。一律走 PowerShell。
4. **lark-cli 要绕过 `.cmd` 包装器**，直接 `node.exe run.js`；URL 里的 `&` 会被 cmd 当成命令分隔符截断。
5. **`--json`/`--filter-json` 直接传中文字符串会被 PowerShell 引号破坏**。走文件：`--json "@upd_000.json"`。
6. **件数字段是字符串且带尾随小数点**（`"5."` `"21."`）。必须 `.rstrip('.')` 再转数值，否则 `int("5.")` 直接抛错。
7. **`3天平均销量` / `7天平均销量` 是文本型**，要按字符串写（保留 1 位小数）；其余 5 个是数字型。类型搞反飞书会报错。
8. **只读模式读 xlsx 会错报 max_row=1**，`load_workbook(read_only=False)`。
9. **批量 filter 不支持多值**：`["规格商家编码","intersects",[多个]]` 报 800010522，该字段只接受单值。批量匹配只能整表拉下来本地 join。
10. **绝不改写「序号」**——下游 lookup 靠序号值对齐，改了会整条链路静默错位。仅 `--append` 新增行时才写序号（= 当前最大序号 + k），因为新行没有序号下游就匹配不到。
11. **报表2 卡在 20000 行**（该版本单表上限）。Excel 有 8 万行时只能更新已有的 2 万行，新编码进不来，这是容量问题不是 bug。
12. **同一份导出重复跑也会全量重写**（`write_rows` ≈ 2 万），**这不是出错**。全量覆盖模式下不再有「changed=0」这种信号，别再拿 `changed_rows == 0` 判断"数据已是最新"——那个判断已随增量逻辑一同废弃。
13. **采购清单表 1018 行里只有 128 行有效**（序号 129+ 是空占位行）。导出主表必须过滤掉，否则发出去一半是空行。
14. **回传用 `--file` 时路径必须是 cwd 相对路径**，绝对路径会被拒。
15. **prep / verify 会自动清残留**（脚本内置 `sweep_stale`）：超过 60 分钟的 `p_*.ndjson`、`v_*.ndjson`、`upd_*.json`、`expected.json`、`ids.json`、`po.ndjson` 会被挪到 `_stale/`（**只挪不删，可恢复**）。原因：这些文件全靠 glob 读取，上一轮残留会被当成当前数据——旧基线让增量算错，旧 `v_*.ndjson` 让 verify 拿旧回读结果误判通过，**两者都不报错**。
16. **别在工作区根目录直接跑**。曾经手工测试把基线落在根目录，隔几天再跑就读到脏基线。按目录规范用独立 workdir；确需重跑时先确认 `p_*.ndjson` 的时间戳是本次刚拉的。prep 报 `error: no_baseline` = 基线缺失或已被 sweep 挪走。
17. **批量拉基线（10 页 × 2000）不要放后台任务**：后台 PowerShell 会被沙箱拦（读取 `C:\Users\Administrator\.ssh` 被拒）。改前台分页跑，每批 3–4 页，前台一次约 40 秒可跑完 3 页。
18. **回传 `--file` 成功时退出码可能是 1**：CLI 把 `uploading file: xxx` 写进 stderr，宿主记成非 0。判断成败看 stdout 里的 `"ok": true`，不要看退出码。**更严重的后果**：宿主可能据此把整段命令重跑一遍，用户会收到两份文件 + 两条文案（2026-09-18 实测）。所以回传一律走编排层的 `scripts/send-result.js`（发送前查最近消息去重），不要裸调 `im +messages-send`。
19. **不要按扩展名判断表格式**。快麦原始导出叫 `.csv` 但内容是 xlsx；用户手工另存后才叫 `.xlsx`。一律交给 `prep` 嗅探，别在编排层做 `if name.endswith('.csv')` 之类的分支。
20. **抽样与验证必须用同一个 ID 集合**：回读前跑 `sample` 生成 `ids.json`，`verify` 只做比对不抽样。自己另写抽样脚本会导致集合错位、误报大量缺失（实测误报 193/200）。
21. **⚠️ 两表的「结构一样」≠「行序一样」**。实测（2026-09-21）：两者列结构确实一致（序号/规格商家编码/实际可用数/…），但**行顺序不同** —— 前 6 行（序号 1–6）编码相同，从第 7 行起就错位：Excel 序号 7 是 `CT159-白色-S-3673`，报表2 序号 7 是 `CT171-粉拼白-S`。快麦每次导出的排序会变。
    → 这条同时说明两件事：(a) 早期「按编码配对」会漏掉 Base 有、Excel 没有的编码，那些行残留旧值；(b) 因此用户明确改为 **positional 按位置覆盖**（第 N 行对第 N 行），这才是「上传什么表，报表2 就是什么内容」。代价是不再保留报表2 原有编码顺序——但反正两侧顺序本来就不同，那个顺序没有业务含义。
22. **报表2 实际就是 20000 行，不是"接口上限所以看不全"**。实测 `offset=20000` 返回 `records_count=0 / has_more=false`，`offset=22000` 同样为 0。所以「Excel 83081 行想完整进报表2」在当前容量下**物理上做不到**，必须先扩容；且**数据清洗表、备货计算表也各自卡在 20000 行**，只扩报表2 的话下游 lookup 覆盖不到新增序号。
23. **⚠️ 唯一可用的并行方式是 runspace 线程池**（`batch-run.ps1` 内部实现）。实测对比（2026-09-21，写入报表2）：
    | 方式 | 结果 |
    |---|---|
    | 工具层同时发多个调用 | ❌ 宿主串行执行，相邻调用实测间隔 **13.6s** |
    | `Start-Process` 后台进程 | ❌ 沙箱拦截，259ms 空返回、无 stdout |
    | Python 起 node 子进程 | ❌ 沙箱拦截 |
    | Python 直连 OpenAPI | ❌ 可行但 `appSecret` 在 `~/.lark-cli/config.json` 里是加密结构，不该绕过 |
    | **runspace 线程池（同进程多线程）** | ✅ 4 路 1.43s / 4 批；8 路 1.66s / 8 批，全部 ok |
    单次 CLI 调用约 1.3s 且绝大部分是 node 进程启动开销，所以并行是唯一有效的提速手段。
24. **并发上限必须是 8，不要开到 16**。实测 `-Mode pull -Concurrency 16` 会把 20000 行拉成 **14000 行**，而且过程不报错、只是静默丢页。丢页能被 prep 的「序号必须 1..N 连续」校验拦住（报 `base_seq_not_continuous`），但别指望它——`batch-run.ps1` 已用 `[ValidateRange(1,10)]` 锁死。
25. **`Remove-Item -LiteralPath` 不支持通配符**。清 `p_*.ndjson` 必须走 `Get-ChildItem -Filter`。曾因此漏删上一轮基线，导致行数多报 4000（旧文件行数被重复计入）。
26. **⚠️⚠️ 合并 ndjson 必须做「纯字节级追加」，绝不能用 `Get-Content -Raw` + `AppendAllText`**。
    PowerShell 5.1 的 `Get-Content` **不带 `-Encoding`** 时按本机 ANSI(GBK) 解码 UTF-8 字节，再以 UTF-8 写回 —— 字节被彻底改写。
    2026-09-21 事故：采购清单 1018 行**全部** `JSONDecodeError: Expecting ':' delimiter`，行内出现 `\ue187`（PUA 区）与 `?`(U+FFFD)，导出失败 → xlsx 没生成 → 回传报 `cannot read file`。
    正确写法（`worker.ps1` 2.5 段已实现）：
    ```powershell
    $bytes = [System.IO.File]::ReadAllBytes($page)
    $start = if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { 3 } else { 0 }
    $fs.Write($bytes, $start, $bytes.Length - $start)      # 剥 BOM：中间页的 BOM 会污染 JSON
    if ($bytes[$bytes.Length - 1] -ne 10) { $fs.WriteByte(10) }   # ndjson 末行常无换行，不补会粘行
    ```
    判据：合并完先用 Python `json.loads` 逐行验一遍再交给 `export`。
27. **⚠️ `-NoExit` 会造成「僵尸进程」**。常驻角色是用 `powershell.exe -NoExit -File xxx.ps1` 拉起的（保证窗口不关、方便看日志）。
    此时脚本里的裸 `exit 0` **只结束脚本、PowerShell 进程仍停在提示符活着** —— `ensure-running.ps1` 按命令行匹配判定它「在线」，于是**永不补拉，任务永久停摆**。
    2026-09-21 实际踩到：worker 记了「收到重启哨兵，优雅退出」但进程仍在，后面的任务再没人处理。
    修复（两处，缺一不可）：
    - 退出用 `[System.Environment]::Exit(0)`（真终止进程，不受 `-NoExit` 影响）；
    - worker 每轮循环刷新心跳 `state\worker-<Id>.alive`；`ensure-running.ps1` 判僵尸要求**三者同时满足**：心跳 >90s 未刷新 **且** `queue\running` 为空 **且** `base.lock` 不存在。这样正在跑长任务（2–4 分钟不转循环、心跳会停）的 worker 绝不会被误杀。
28. **导出文件名带「日期时间」前缀**：`2026-9-21-1535-商品采购清单表.xlsx`（`Get-Date -Format "yyyy-M-d-HHmm"`）。
    两个原因：同一天多次任务的结果在飞书会话里能区分；去重键虽用 `event_id`，但文件名固定会让用户侧看起来像同一个文件。
    **Windows 文件名禁止 `:`** —— 用户口语说的 `15:10` 只能写成 `1510`，别照抄冒号（会直接抛 `InvalidArgument`）。
29. **⚠️ worker 崩溃会在队列里留下「卡死组合」**，必须靠启动自愈清掉，否则任务永久卡住：
    - `queue\running\w<Id>_*.json` —— 任务卡在 running。**新 worker 只捡 incoming**，所以这个任务再没有任何进程会碰它，用户侧表现为「发了表格一直没反应」；
    - `state\base.lock` —— 僵尸锁，后续任务全部卡在「等锁」。
    2026-09-21 实际踩到：worker 7696 在 prep 阶段被外部结束，任务就这样挂着。
    `worker.ps1` 启动时已内置自愈：本 Id 的 `running` 任务一律回收（同一 Id 同时只应有一个实例，由 `ensure-running.ps1` 的命令行匹配保证）；锁只在「超龄 >15 分钟」时清（阈值远大于单次任务最长耗时 ~5 分钟，避免误清别的 worker 正在用的锁）。
30. **⚠️ 绝不要在工具会话里手工跑 `ensure-running.ps1` 来拉起常驻进程**。命令结束时它启动的子进程树会被一并回收 ——
    2026-09-21 连续踩两次：手工拉起 worker（pid 9568）13 秒后心跳就停了；pid 7696 撑到命令结束也没了。
    **常驻进程只能由 `guard-loop`（进程链挂在开机 vbs 上，与工具会话无关）拉起**。需要重启时：停掉旧进程，然后**等 guard 的下一轮巡检**（最长 5 分钟），不要手工拉。
31. **诊断 worker 死活看心跳，不看进程列表**：`state\worker-<Id>.alive` 的 mtime 就是最后活跃时刻。
    注意**跑任务期间心跳本来就不更新**（不转循环），单看心跳会误判；要结合 `queue\running` 是否非空、`state\base.lock` 是否存在一起看。

> 以下两节是**增量模式时代的实录**（`changed` 只统计新旧值不同的行），仅供容量与耗时参考。
> 现为全量覆盖：`write_rows` = `matched`，`changed=0` 分支已不存在。

## 首次完整写入实录（2026-09-18 17:29→17:36，含真实写入）

| 步 | 结果 |
|---|---|
| 输入 | `.csv`（嗅探为 xlsx）→ 83081 行 |
| prep | matched 19981 / **changed 4931** / base_coverage 0.999 / 25 个批次 |
| 写入 | 25 批 × 200 条全部 `ok: true` |
| 抽样验证 | 回读 200 条，**checked 200、mismatches 0、passed=true** |
| 导出 | 有效款色 115、采购总件数 **12264**（改前为 128 / 8419，证明下游已重算） |

耗时约 7 分钟，其中拉基线 10 页约占 2 分钟、25 批写入约 3 分钟。

## 首次闭环实录（2026-09-18 16:53→17:00，约 7 分钟）

| 步 | 结果 |
|---|---|
| 下载 | `input.xlsx` 6.36 MB |
| 拉基线 | 10 页 × 2000 = 20000 行，约 40 秒 |
| prep | excel 81736 行 / base 20000 行 / matched 19997 / **changed=0** / base_coverage 0.9999 |
| 写入 + 验证 | **跳过**（changed=0） |
| 导出 | 1018 行 → 有效 128 行 / 采购总件数 8419 |
| 回传 | xlsx 48 KB + 完成文案 |

结论：`changed=0` 分支已实测跑通，整链路约 7 分钟，其中大半是拉基线。

## 典型耗时（2026-09-21 实测，`-Concurrency 8`）

| 环节 | 报表2 = 2 万行 | 报表2 = 10 万行（扩容后推算） |
|---|---|---|
| 拉基线 `Mode pull` | **12.4s**（串行 39s） | ~60s（50 页 / 8 路） |
| prep | ~15s（Excel 8.3 万行） | ~40s |
| 写入 `Mode update` | **~27s**（100 批；串行 140s） | **~2.2 分钟**（500 批；串行 11.7 分钟） |
| 抽样验证 + 导出 | ~10s | ~15s |
| **合计** | **约 1 分钟** | **约 4 分钟** |

实测依据：12 批并发写入 4.3s（串行 16.8s）；100 批 = 20000 行由 prep 生成，实际全量写入曾耗时约 3 分钟（串行），并发后理论 ~27s。

## 扩到 10 万行时要做什么

用户的 Excel 可能涨到 10 万行，届时按下面 checklist 走：

1. **三张表一起扩容**：报表2 + 数据清洗表 + 备货计算表，都要 ≥ Excel 行数。只扩报表2 没用——下游 lookup 按序号匹配，够不着的行等于白写。
2. **流程本身不用改**：`batch-run.ps1 -Mode pull` 自动探测行数（轮空即停），**没有写死 20000**；`-Mode update` 按 `upd_*.json` 数量并发，行数翻 5 倍也只是批次数变多。
   - 已同步修掉的隐患：编排层常驻 `worker.ps1` 里原本有一份写死的 `for ($off = 0; $off -lt 20000; $off += 2000)` 手工分页，扩到 10 万行时只会拉前 2 万行、写入整体错位。2026-09-21 已改成调用 `batch-run.ps1 -Mode pull`；采购清单导出同理改成「分页拉到空 + 合并 po.ndjson」。**改动后必须重启 worker 进程才生效**（常驻进程加载的是旧代码，`ensure-running.ps1` 会自动补起）。
3. **Excel 比报表2 长时**：prep 默认把多出的行记为 `overflow_rows` 丢弃。要让它们进来，传 `--append` 生成 `cre_*.json`，再 `-Mode create` 写入（新行会自动补序号）。
4. **耗时预期约 4 分钟**。若嫌慢，先确认瓶颈：`pull` 页数 = ceil(行数/2000)，`update` 批数 = ceil(行数/200)。这两个都随行数线性增长，没有捷径（API 硬上限 200 条/批、页大小上限 2000）。
