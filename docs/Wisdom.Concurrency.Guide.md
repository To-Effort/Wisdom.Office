# 高并发部署与性能指南（Wisdom.Office 全家桶）

适用版本：1.0.x。测量环境：28 核 Windows，.NET 8，`benchmarks/ConcurrencyMicro`
（六模块并行写 / 并行读 / CPU 密集转换 / 门面加载，dop = 1/4/8/28，每场景含产物回读校验）。

## 一句话结论

库本身线程安全（全部并发场景 0 错误、产物回读全部一致），高并发下的瓶颈不在代码锁竞争，
而在 .NET 默认运行时配置：**服务端部署务必开启 Server GC 并预热线程池**，
否则吞吐损失 2~8 倍、p50 延迟劣化可达 8 倍。

## 实测数据（dop=28，ops/s）

| 场景 | 默认配置 | Server GC + SetMinThreads(64,64) | 提升 |
|---|---:|---:|---:|
| excel 读 | 64.5 | 565.1 | 8.8× |
| facade OfficeFile.Load | 61.6 | 491.3 | 8.0× |
| markdown ToHtml | 553.3 | 3350.5 | 6.1× |
| markdown 往返 | 656.2 | 3326.8 | 5.1× |
| ppt 读 | 566.3 | 2154.7 | 3.8× |
| word 写 | 1479.1 | 4866.1 | 3.3× |
| word 读 | 2208.1 | 4051.0 | 1.8× |
| pdf 转换（Word→PDF） | 493.5 | 1106.3 | 2.2× |
| excel 写 | 197.5 | 308.4 | 1.6× |
| ppt 写 | 585.0 | 772.9 | 1.3× |

p50 延迟（默认 → 调优后）：excel 读 390ms → 48ms；markdown 41ms → 4.6ms；
Gen0 回收次数从单场景 132 次降到 5 次。

## 推荐配置

### 1. Server GC（必做）

csproj：

```xml
<PropertyGroup>
  <ServerGarbageCollection>true</ServerGarbageCollection>
  <ConcurrentGarbageCollection>true</ConcurrentGarbageCollection>
</PropertyGroup>
```

或运行时环境变量 `DOTNET_gcServer=1`。
OOXML 读写是分配密集型（每次数 MB 到数十 MB 的 XML/ZIP 缓冲），workstation GC 在
多核并发下会串行化回收、把每个操作拖成秒级。桌面单线程应用可保持默认。

### 2. 预热线程池（并发 ≥ 8 时建议）

```csharp
System.Threading.ThreadPool.SetMinThreads(64, 64);
```

.NET 线程池默认每秒只注入约 1~2 个工作线程，并发任务的首波会被排队（表现为
p95 远大于 p50）。服务启动时调用一次即可。

### 3. 内存预算

并发写大文档按"单文档峰值驻留 × 并发数"预留内存：400 页全效果 PPT 写出峰值约 173MB，
28 路并发写同一量级文档时建议容器内存 ≥ 8GB，或限制并发度（`SemaphoreSlim`）。

## 线程安全语义（可以放心并发的边界）

- **静态读写入口全部无共享状态**：`DocxReader/DocxWriter`、`XlsxReader/XlsxWriter`、
  `PptxReader/PptxWriter`、`PdfDocument.Load`、`Markdown.*`、`OfficeFile.Load`、
  `OfficeConverter.*` 可任意并行调用（实测 28 线程共享同一模型实例并发写也零错误）。
- **文档模型实例不是线程安全的**：同一 `OfficeDocument/OfficeWorkbook/OfficePresentation`
  在一个线程构建、交给写线程后即不可再改；多线程不要并发修改同一实例。
- **PDF 字体解析器有全局锁**：`FontResolver.Default` 内部单锁保证正确性；首次调用会扫描
  字体目录并缓存，之后锁持有时间极短。服务启动时先做一次 Word→PDF 转换即可预热。
- **安全上限是进程级静态字段**：`OfficePackageLimits.MaxEntryBytes/MaxTotalBytes`、
  `PdfReadLimits.*` 修改后影响所有线程，应在启动时一次性设置，不要在请求路径里改。
- **不要并发写同一个目标文件路径**：Windows 文件锁会抛 `IOException`（PowerPoint 打开
  中的 .pptx 同理会阻止重写）。

## 复现方法

```
dotnet run -c Release --project benchmarks/ConcurrencyMicro            # 默认配置
DOTNET_gcServer=1 dotnet run -c Release --project benchmarks/ConcurrencyMicro mt   # 调优后
```
