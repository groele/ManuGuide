# ManuGuide v2.3 工业级重构方案：长文档性能、内部标记化、安全回写与可扩展规则内核

> 适用对象：Word VSTO 版 ManuGuide  
> 目标：在长文档、公式/引用密集文档、多规则扩展、多批量操作场景下，实现 **流畅、稳健、可审计、可扩展、可回滚** 的学术文档诊断系统。

---

## 0. v2.3 相比 v2.2 的关键修正

v2.2 已经建立了“快照化 + 内部标记化 + Regex/JS/C# 三层扩展”的总体框架。v2.3 进一步修正几个容易在真实工程落地中出问题的点：

1. **不再把 CoordinateMaskedText 作为主扫描路径**  
   等长掩码保留为坐标兼容层，但主扫描路径改为 `ScannableSegment`，每个 scanner 只扫描非保护区片段，从根源杜绝跨公式、跨引用、跨 DOI、跨 References 的误匹配。

2. **全文批量修复不再只依赖“批次内倒序”**  
   全文修复必须先生成全局 `MutationPlan`，按全文坐标全局降序执行；若必须按非降序执行，则启用 `MutationDeltaMap` 做动态坐标映射。

3. **JS 扩展从 MVP 移到 Phase 2**  
   Phase 1 只落地 Regex Rule Pack、虚拟高亮、MarkerLedger、ProtectedRangeIndex、BatchMutationApplier，先把性能和安全边界做稳。

4. **JS 运行时必须使用 Engine Pool 或单线程队列**  
   不允许多线程共享同一个 Jint Engine，也不允许每个 chunk 都重新 new Engine 造成 GC 雪崩。

5. **扩展加载改为两阶段生命周期**  
   静态加载规则包元数据；快照生成之后再进行 `BindSnapshot()` 动态绑定，避免扩展加载与快照依赖之间的循环矛盾。

6. **用户自定义 Regex 规则默认不能 SafeAutoFix**  
   只有官方签名规则包或通过测试认证的团队规则包才能声明 `SafeAutoFix`。用户本地规则默认 `NeedsUserConfirmation`。

7. **MarkerLedger 改为运行时内存主账本 + 延迟持久化**  
   避免每次 marker 修改都写入 CustomXMLPart 导致保存和自动保存变慢。

8. **首屏响应目标替代全文秒级完成目标**  
   工业级体验的核心不是全文瞬间完成，而是 `FirstIssueVisibleMs ≤ 1s`，全文扫描后台继续，用户可取消、可查看渐进式结果。

---

## 1. 核心北极星原则

ManuGuide 的底层哲学应该从“Word 查找替换工具”升级为：

> **基于不可变文档快照的非破坏性学术审计系统。**

所有实现必须遵守以下原则：

1. **Word COM 最小化原则**  
   尽量不碰 Word；必须碰 Word 时，在主线程、小批量、可撤销、可校验地碰。

2. **扫描与修改完全分离原则**  
   扫描阶段只产生 `IssueCandidate`；不高亮、不修复、不写评论、不创建 Range。

3. **快照世界与 Word 世界隔离原则**  
   Scanner、Regex、JS、C# Native 插件都运行在快照世界；真实写回只能通过 `WordMutationApplier`。

4. **保护区优先原则**  
   公式、引用、域、DOI、URL、LaTeX、References、用户配置的保护区永远优先于规则命中。

5. **虚拟标记优先原则**  
   默认只在侧边栏显示虚拟 issue；用户显式要求后才写入 Word 真实标记。

6. **MarkerLedger 唯一痕迹原则**  
   任何真实 Word 痕迹都必须进入 MarkerLedger；清除只能清除 ManuGuide 自己的 marker。

7. **扩展低权限原则**  
   Regex/JS/C# 扩展只能发现问题；不能直接改 Word、不能绕过保护区、不能绕过回写防线。

---

## 2. v2.3 总体架构

```text
Ribbon / TaskPane
   │
   ├─ JobController
   │    ├─ ScanJob
   │    ├─ PreviewJob
   │    ├─ MarkJob
   │    ├─ FixSafeItemsJob
   │    ├─ ClearMarkersJob
   │    └─ ExtensionReloadJob
   │
Word Boundary Layer
   │
   ├─ IWordDocumentAccessor
   │    └─ VstoDocumentAccessor
   ├─ DocumentSnapshotBuilder       // 必须在 Word 主线程 / STA 执行
   ├─ WordRangeResolver             // 坐标 → Range 的唯一入口
   ├─ WordMutationApplier           // 修改 Word 的唯一入口
   └─ UiThreadContext               // 所有 COM 调用的线程门禁
   │
Snapshot Core
   │
   ├─ DocumentSnapshot
   ├─ ParagraphSnapshot
   ├─ ProtectedRangeIndex
   ├─ ScannableSegmentIndex
   ├─ TokenLedger
   ├─ SnapshotVersionGuard
   └─ MarkerLedger
   │
Scanner Runtime
   │
   ├─ ChunkPlanner
   ├─ ScannableSegmentPlanner       // 主扫描路径
   ├─ ScannerOrchestrator
   ├─ RegexRuleScanner
   ├─ BuiltInNativeScanner
   ├─ JavaScriptRuleRuntime         // Phase 2
   └─ IssueReducer
   │
Mutation Runtime
   │
   ├─ MutationPlanner
   ├─ MutationPlan
   ├─ MutationDeltaMap
   ├─ DocumentMutationGuard
   ├─ AtomicUndoGroup
   └─ MarkerLedgerPersistenceService
   │
Extension Runtime
   │
   ├─ ExtensionCatalog
   ├─ RulePackLoader
   ├─ RegexRuleCompiler
   ├─ ExtensionPermissionGuard
   ├─ JavaScriptEnginePool          // Phase 2
   └─ RulePackTestRunner
   │
UI Runtime
   │
   ├─ VirtualIssueList
   ├─ ExtensionManager
   ├─ DiagnosticsPanel
   ├─ ProgressReporter
   └─ SnapshotStalenessNotifier
```

---

## 3. Word COM 硬约束

这是不可妥协的工程红线：

```text
1. DocumentSnapshotBuilder 必须在 Word 主线程 / STA 线程执行。
2. Scanners/ 下的实现禁止 using Microsoft.Office.Interop.Word。
3. Regex/JS/Native scanner 禁止创建 Word.Range。
4. 所有写操作必须进入 WordMutationApplier。
5. 所有写操作必须先经过 DocumentMutationGuard。
6. WPF UI 只能通过 Dispatcher 更新。
7. 后台线程只能处理不可变快照和纯内存对象。
```

建议定义：

```csharp
[AttributeUsage(AttributeTargets.Method | AttributeTargets.Class)]
public sealed class RequiresWordStaThreadAttribute : Attribute { }
```

并在调试模式加入断言：

```csharp
public static class WordThreadGuard
{
    public static void AssertOnWordThread()
    {
        if (!Thread.CurrentThread.GetApartmentState().Equals(ApartmentState.STA))
            throw new InvalidOperationException("Word COM access must run on STA thread.");
    }
}
```

---

## 4. DocumentSnapshot：不可变快照

```csharp
public sealed class DocumentSnapshot
{
    public string SnapshotId { get; init; } = Guid.NewGuid().ToString("N");
    public DateTimeOffset CreatedAt { get; init; } = DateTimeOffset.UtcNow;

    public int ContentStart { get; init; }
    public int ContentEnd { get; init; }

    public string FullText { get; init; } = string.Empty;

    // 兼容层：可保留，但不作为主扫描路径。
    public string CoordinateMaskedText { get; init; } = string.Empty;

    public IReadOnlyList<ParagraphSnapshot> Paragraphs { get; init; } = Array.Empty<ParagraphSnapshot>();
    public ProtectedRangeIndex ProtectedRanges { get; init; } = ProtectedRangeIndex.Empty;
    public ScannableSegmentIndex ScannableSegments { get; init; } = ScannableSegmentIndex.Empty;
    public TokenLedger Tokens { get; init; } = TokenLedger.Empty;

    public DocumentFingerprint Fingerprint { get; init; } = DocumentFingerprint.Empty;
}
```

---

## 5. 从 CoordinateMaskedText 主扫描改为 ScannableSegment 主扫描

### 5.1 为什么不能把空格掩码作为主扫描路径

如果把公式、引用替换为空格，容易触发：

```text
\s{2,}
\bword\b
Fig\.\s+\d+
[A-Za-z]+\s+[A-Za-z]+
```

这些规则可能跨越保护区误匹配，产生假阳性。

### 5.2 v2.3 主策略

```text
ProtectedRangeIndex
   ↓
ScannableSegmentIndex
   ↓
TextChunk
   ↓
ScannableSegment within chunk
   ↓
Scanner only scans segment.Text
```

### 5.3 数据结构

```csharp
public sealed class ScannableSegment
{
    public int SegmentId { get; init; }
    public int TextStart { get; init; }
    public int TextEnd { get; init; }
    public string Text { get; init; } = string.Empty;

    public int ParagraphIndex { get; init; }
    public SegmentScope Scope { get; init; }
}

public enum SegmentScope
{
    MainText,
    Caption,
    Abstract,
    Title,
    Appendix,
    Footnote,
    Endnote
}
```

### 5.4 Segment 规划逻辑

```csharp
public sealed class ScannableSegmentPlanner
{
    public IReadOnlyList<ScannableSegment> Build(
        string fullText,
        IReadOnlyList<ProtectedTextRange> protectedRanges)
    {
        var merged = ProtectedRangeMerger.Merge(protectedRanges);
        var segments = new List<ScannableSegment>();

        int cursor = 0;

        foreach (var range in merged)
        {
            if (cursor < range.Start)
            {
                segments.Add(CreateSegment(fullText, cursor, range.Start));
            }

            cursor = Math.Max(cursor, range.End);
        }

        if (cursor < fullText.Length)
        {
            segments.Add(CreateSegment(fullText, cursor, fullText.Length));
        }

        return segments;
    }
}
```

### 5.5 等长掩码仍然保留，但只作为兼容层

如果某些 legacy scanner 仍需要全文字符串，可以提供：

```text
CoordinateMaskedText
```

但填充字符应使用罕见控制字符，而不是空格：

```csharp
public const char ProtectedMaskChar = '\u0001';
```

并且输出 candidate 后仍要二次过滤：

```csharp
if (snapshot.ProtectedRanges.Intersects(candidate.TextStart, candidate.TextEnd))
    continue;
```

---

## 6. ProtectedRangeIndex：保护区索引

```csharp
public sealed class ProtectedRangeIndex
{
    public static readonly ProtectedRangeIndex Empty = new(Array.Empty<ProtectedTextRange>());

    private readonly ProtectedTextRange[] _ranges;

    public bool Intersects(int start, int end)
    {
        if (_ranges.Length == 0) return false;

        int low = 0;
        int high = _ranges.Length - 1;

        while (low <= high)
        {
            int mid = low + (high - low) / 2;
            var r = _ranges[mid];

            if (r.End > start && r.Start < end)
                return true;

            if (r.Start >= end)
                high = mid - 1;
            else
                low = mid + 1;
        }

        return false;
    }
}
```

保护区来源：

```text
Zotero Field
Mendeley Field
EndNote Field
Word OMath
MathType Object
LaTeX inline
LaTeX block
DOI
URL
Hyperlink
References / Bibliography section
User ignored region
Tracked marker region
```

---

## 7. 扩展加载生命周期：两阶段模型

v2.2 中扩展加载可能与快照依赖形成循环。v2.3 改为：

```text
Phase A：静态加载
   - 读取 manifest
   - 校验权限
   - 校验 schema
   - 编译 Regex
   - 不依赖当前文档

Phase B：快照绑定
   - 传入 DocumentSnapshot
   - 初始化 scannable segment cache
   - 初始化术语表
   - 初始化 JS shared cache
   - 生成本次 ScanSession 专属 scanner instance
```

### 7.1 接口

```csharp
public interface IScannerFactory
{
    string ExtensionId { get; }
    string Engine { get; }

    ScannerBindingResult BindSnapshot(
        DocumentSnapshot snapshot,
        ExtensionSettings settings);
}

public interface ISnapshotScanner
{
    string ScannerId { get; }
    string ModuleType { get; }

    bool CanSkip(ScannableSegment segment);

    ValueTask<IReadOnlyList<IssueCandidate>> ScanAsync(
        ScannableSegment segment,
        ScannerExecutionContext context,
        CancellationToken cancellationToken);
}
```

---

## 8. Regex Rule Pack：Phase 1 主扩展方式

### 8.1 用户规则默认安全等级

```text
官方签名规则包：允许 SafeAutoFix
团队签名规则包：通过测试后允许 SafeAutoFix
用户本地规则包：默认 NeedsUserConfirmation
未签名规则包：禁止 SafeAutoFix
```

### 8.2 单条规则示例

```json
{
  "ruleId": "unit-space-nm",
  "moduleType": "DataValue",
  "title": "Missing space between number and nm",
  "pattern": "(?<value>\\d+(?:\\.\\d+)?)\\s*(?<unit>nm)\\b",
  "replacement": "${value} ${unit}",
  "fixSafety": "NeedsUserConfirmation",
  "scope": {
    "include": ["MainText", "Caption"],
    "excludeProtected": true,
    "excludeReferences": true,
    "excludeUrls": true
  },
  "guards": {
    "maxMatchLength": 32,
    "requireExpectedTextBeforeFix": true,
    "doNotCrossProtectedBoundary": true
  },
  "examples": [
    {
      "input": "The thickness is 10nm.",
      "expectedIssues": 1,
      "expectedFix": "The thickness is 10 nm."
    }
  ]
}
```

### 8.3 Regex 编译限制

| 项目 | 建议值 |
|---|---:|
| Pattern 最大长度 | 2000 chars |
| Timeout | 100–300 ms |
| 单规则单 segment 最大命中数 | 100 |
| 单规则单文档最大命中数 | 500 |
| 用户规则包最大规则数 | 200 |
| 未签名规则默认 FixSafety | NeedsUserConfirmation |

---

## 9. JS Script Pack：Phase 2 才启用

### 9.1 为什么 JS 不进 MVP

JS 扩展带来以下额外复杂度：

1. Jint Engine 非线程安全；
2. 沙箱边界容易被误配置破坏；
3. JS 需要 timeout、memory、issue 数量限制；
4. JS 跨 chunk 的上下文缓存需要额外生命周期管理；
5. 用户脚本质量不可控。

因此 MVP 阶段应明确：

```text
Phase 1：Regex Rule Pack + BuiltIn C# Scanner
Phase 2：JS Script Pack
Phase 3：Native C# Plugin SDK
```

### 9.2 JS 运行池

```csharp
public sealed class JavaScriptEnginePool
{
    private readonly ConcurrentBag<Engine> _pool = new();
    private readonly string _preloadedScript;
    private readonly int _maxPoolSize;
    private int _created;

    public Engine Rent()
    {
        if (_pool.TryTake(out var engine))
            return engine;

        if (Interlocked.Increment(ref _created) <= _maxPoolSize)
            return CreateEngine();

        Interlocked.Decrement(ref _created);

        // 也可以改为等待队列，避免无限创建。
        SpinWait.SpinUntil(() => _pool.TryTake(out engine), TimeSpan.FromMilliseconds(50));

        if (engine == null)
            throw new TimeoutException("No JavaScript engine available.");

        return engine;
    }

    public void Return(Engine engine)
    {
        // 必须清理本次 chunk 的临时状态。
        ResetTransientBindings(engine);
        _pool.Add(engine);
    }

    private Engine CreateEngine()
    {
        var engine = new Engine(options =>
        {
            options.TimeoutInterval(TimeSpan.FromMilliseconds(200));
            options.LimitMemory(4_000_000);
            options.MaxStatements(50_000);
        });

        // 只注入显式允许的 Host API。
        engine.Execute(_preloadedScript);
        return engine;
    }
}
```

### 9.3 JS Host API 白名单

允许：

```text
context.intersectsProtectedRange(start, end)
context.getParagraphInfo(start)
context.getSetting(name)
context.normalizeText(text)
context.emitDiagnostic(message)
```

禁止：

```text
Word.Document
Word.Range
System.IO
System.Net
Process
Reflection
NativeInterop
fetch
eval 外部动态代码
```

### 9.4 JS 全局初始化

```javascript
export function init(snapshot, sharedCache) {
  sharedCache.abbreviations = [];
  // 可基于 snapshot metadata 做一次性分析，但不能保存正文原文到日志。
}

export function canSkip(segment, context, sharedCache) {
  return !segment.text.includes("Fig.");
}

export function scanSegment(segment, context, sharedCache) {
  return [];
}
```

---

## 10. ScannerOrchestrator：基于 segment 的主扫描流程

```csharp
public sealed class ScannerOrchestrator
{
    public async Task<ScanResult> RunAsync(
        Word.Document doc,
        ScanOptions options,
        IProgress<ScanProgress> progress,
        CancellationToken ct)
    {
        // 1. 主线程构建快照
        DocumentSnapshot snapshot = await _uiThread.InvokeAsync(
            () => _snapshotBuilder.Build(doc));

        // 2. 静态扩展绑定到快照
        var scannerFactories = _extensionCatalog.LoadEnabledFactories();
        var scanners = scannerFactories
            .Select(f => f.BindSnapshot(snapshot, options.ExtensionSettings))
            .SelectMany(r => r.Scanners)
            .ToArray();

        // 3. 规划可扫描片段
        var segments = snapshot.ScannableSegments
            .Prioritize(options.VisibleRange)
            .ToArray();

        var bag = new ConcurrentBag<IssueCandidate>();

        // 4. 后台并行扫描 segment，不碰 Word COM
        await Parallel.ForEachAsync(segments, ct, async (segment, token) =>
        {
            foreach (var scanner in scanners)
            {
                if (scanner.CanSkip(segment))
                    continue;

                var issues = await scanner.ScanAsync(
                    segment,
                    new ScannerExecutionContext(snapshot, options),
                    token);

                foreach (var issue in issues)
                    bag.Add(issue);
            }

            progress.Report(ScanProgress.FromSegment(segment));
        });

        // 5. 合并、过滤、去重
        var reduced = _issueReducer.Reduce(bag, snapshot);

        return new ScanResult(snapshot.SnapshotId, reduced);
    }
}
```

---

## 11. IssueReducer：不要暴力去重

### 11.1 冲突类型

```csharp
public enum IssueCollisionScope
{
    SameRuleSameSpan,
    SameModuleOverlap,
    DifferentModuleOverlap,
    ContainsAnother,
    Independent
}
```

### 11.2 策略

| 冲突类型 | 策略 |
|---|---|
| SameRuleSameSpan | 去重 |
| SameModuleOverlap | 选更高置信度或更精准 span |
| DifferentModuleOverlap | 通常保留 |
| Formatting + Wording 重叠 | 保留 |
| SafeAutoFix 与 SuggestOnly 同类重叠 | SafeAutoFix 优先，但保留 SuggestOnly 为附加说明 |
| 用户规则与官方规则冲突 | 官方规则优先，用户规则降级显示 |

### 11.3 关键原则

不要因为 `10nm` 被 DataValueScanner 命中，就丢弃 `thickness is 10nm` 的 Wording 建议。格式问题和措辞问题是不同维度。

---

## 12. 批量修复：MutationPlan，而不是简单循环

### 12.1 错误做法

```text
按当前 UI 列表顺序修复
每批内部倒序
批次之间不考虑前文修改导致的坐标整体漂移
```

### 12.2 正确做法 A：全文全局降序计划

对于文本替换类修复，默认：

```text
1. 收集所有 SafeAutoFix candidates
2. 过滤保护区
3. 按 TextStart 全局降序排序
4. 按降序切 batch
5. 每个 batch 内继续降序
6. 从文档末尾向文档开头写
```

这种方式下，后文修改不会影响前文坐标，最稳定。

```csharp
public sealed class MutationPlanner
{
    public MutationPlan BuildSafeFixPlan(IEnumerable<IssueCandidate> candidates)
    {
        var ordered = candidates
            .Where(c => c.FixSafety == FixSafetyLevel.SafeAutoFix)
            .OrderByDescending(c => c.TextStart)
            .ToList();

        return new MutationPlan(ordered, MutationOrder.GlobalDescending);
    }
}
```

### 12.3 正确做法 B：MutationDeltaMap

如果必须按自然顺序执行，例如用户要求“从当前页开始修”，则使用 DeltaMap。

```csharp
public sealed class MutationDeltaMap
{
    private readonly List<MutationDelta> _deltas = new();

    public int MapStart(int originalStart)
    {
        int delta = 0;

        foreach (var d in _deltas)
        {
            if (d.OriginalStart < originalStart)
                delta += d.LengthDelta;
        }

        return originalStart + delta;
    }

    public void Add(int originalStart, int originalEnd, int lengthDelta)
    {
        _deltas.Add(new MutationDelta(originalStart, originalEnd, lengthDelta));
    }
}
```

长文档或大量变动时可改为 Fenwick Tree / Interval Delta Index。

### 12.4 DocumentMutationGuard

```csharp
public sealed class DocumentMutationGuard
{
    public ValidationResult Validate(
        Word.Document doc,
        IssueCandidate candidate,
        DocumentSnapshot snapshot,
        MutationCoordinateMapper mapper)
    {
        var mapped = mapper.Map(candidate.TextStart, candidate.TextEnd);

        if (mapped.Start < snapshot.ContentStart || mapped.End > doc.Content.End)
            return ValidationResult.RangeShifted;

        // 注意：保护区也应使用 mapper 映射后的当前文档区间判断。
        if (snapshot.ProtectedRanges.IntersectsOriginal(candidate.TextStart, candidate.TextEnd))
            return ValidationResult.InProtectedZone;

        var range = doc.Range(mapped.Start, mapped.End);

        if (Normalize(range.Text) != Normalize(candidate.OriginalText))
            return ValidationResult.RangeShifted;

        return ValidationResult.Valid;
    }
}
```

---

## 13. UndoRecord：原子批次策略

### 13.1 原则

```text
小批量
可撤销
失败不污染
失败后自动跳过问题项并继续
```

### 13.2 推荐策略

| 批次类型 | 批次大小 |
|---|---:|
| SafeAutoFix 文本替换 | 5–10 |
| Marker-only 高亮 | 20–50 |
| 清除 marker | 20–50 |
| 评论插入 | 10–20 |

### 13.3 失败处理

```text
1. 尝试执行 batch；
2. 若 batch 中某一项失败，不中断整个任务；
3. 该项记录为 Skipped / RangeShifted / InProtectedZone；
4. 成功项保留；
5. UndoRecord 按 batch 合并；
6. UI 显示成功、失败、跳过数量。
```

比“任何一项失败回滚整批”更适合用户体验。只有在用户启用严格事务模式时，才使用整批失败回滚。

---

## 14. MarkerLedger：运行时主账本 + 延迟持久化

### 14.1 为什么不能频繁写 CustomXMLPart

当 marker 数量很多时，每次修改都序列化 CustomXMLPart 会拖慢保存和自动保存。

### 14.2 v2.3 策略

```text
运行时：内存 MarkerLedger 为主
写入 Word 时：创建 mg_ bookmark / comment / content-control tag
保存前：DocumentBeforeSave 统一持久化 ledger
打开文档：从 CustomXMLPart / bookmarks 重建 ledger
清除时：ledger 优先，bookmark fallback
```

### 14.3 Marker 结构

```csharp
public sealed class MarkerRecord
{
    public string MarkerId { get; init; } = string.Empty;
    public string OperationId { get; init; } = string.Empty;
    public string SnapshotId { get; init; } = string.Empty;
    public string IssueId { get; init; } = string.Empty;

    public int OriginalStart { get; init; }
    public int OriginalEnd { get; init; }

    public MarkerType Type { get; init; }
    public MarkerState State { get; set; }

    public DateTimeOffset CreatedAt { get; init; }
}
```

### 14.4 Missing marker 防腐处理

```csharp
try
{
    var bookmark = doc.Bookmarks[marker.MarkerId];
    ClearRange(bookmark.Range);
}
catch
{
    marker.State = MarkerState.Missing;
    _ledger.MarkDirty();
    continue;
}
```

清除任务绝不能因为单个 marker 丢失而中断。

---

## 15. 快照失效与用户编辑同步

### 15.1 快照状态

```csharp
public enum SnapshotState
{
    Current,
    PossiblyStale,
    Stale,
    Invalid
}
```

### 15.2 触发条件

```text
SelectionChange：不一定失效
DocumentChange：可能失效
BeforeSave：检查 ledger
Batch mutation：快照变 stale
User manual edit：快照 possibly stale
ExpectedText mismatch：相关 issue stale
```

### 15.3 UI 提示

```text
当前结果基于 14:32:18 的文档快照。
文档似乎已被修改，批量修复前建议重新扫描。
[重新扫描] [仍查看旧结果]
```

### 15.4 批量修复策略

如果文档已变更：

```text
SafeAutoFix 批量修复：默认要求重新扫描
单条修复：允许 Guard 校验通过后执行
真实全文高亮：建议重新扫描
清除 marker：不依赖快照，可继续
```

---

## 16. 性能目标重设

### 16.1 不再承诺全文扫描 5 秒内必完成

更合理的目标：

| 指标 | 目标 |
|---|---:|
| FirstIssueVisibleMs | ≤ 1000 ms |
| VisibleRangeScanMs | ≤ 800 ms |
| TaskPane 不白屏 | 必须 |
| UI 批量追加间隔 | 100–200 ms |
| 全文扫描 | 后台继续，可取消 |
| JS 扩展 | 默认关闭或高级模式 |
| SafeFix 写回 | 小批量执行，可取消下一批 |

### 16.2 诊断指标

```text
SnapshotBuildMs
ProtectedIndexBuildMs
ScannableSegmentBuildMs
ExtensionStaticLoadMs
ExtensionBindMs
RegexCompileMs
VisibleScanMs
FirstIssueVisibleMs
FullScanMs
IssueReduceMs
BatchMutationMs
ClearMarkersMs
RegexTimeoutCount
JsTimeoutCount
SkippedProtectedCount
SkippedStaleCount
RangeMismatchCount
```

---

## 17. Extension 安全模型

### 17.1 权限

```csharp
public enum ExtensionPermission
{
    ReadSnapshotMetadata,
    ReadSegmentText,
    ReadSettings,
    EmitIssues,
    EmitDiagnostics,

    // Phase 2 以后仍默认禁止
    FileRead,
    FileWrite,
    NetworkAccess,
    WordComAccess,
    NativeInterop,
    ProcessStart
}
```

### 17.2 默认权限

Regex Rule Pack：

```text
ReadSegmentText
EmitIssues
EmitDiagnostics
```

JS Script Pack：

```text
ReadSnapshotMetadata
ReadSegmentText
ReadSettings
EmitIssues
EmitDiagnostics
```

Native C# Plugin：

```text
仅开发者模式 / 官方签名
```

### 17.3 禁止规则

```text
禁止 JS 访问 Word.Document
禁止 JS 访问文件系统
禁止 JS 访问网络
禁止 JS 执行进程
禁止 JS 保存正文日志
禁止扩展直接写 MarkerLedger
禁止扩展直接写 Word
```

---

## 18. MVP 范围重新定义

### Phase 1：性能与安全核心

必须做：

```text
1. DocumentSnapshotBuilder 主线程化
2. ProtectedRangeIndex
3. ScannableSegmentIndex
4. Regex Rule Pack
5. Regex timeout
6. VirtualIssueList
7. IssueReducer
8. MarkerLedger 基础版
9. BatchMutationApplier
10. SnapshotStalenessNotifier
```

明确不做：

```text
JS Script Pack
Native C# Plugin SDK
Office.js 迁移
复杂 marketplace
增量段落缓存
Aho-Corasick 高级词库
```

### Phase 2：高级扩展

```text
1. JS Script Pack
2. JavaScriptEnginePool
3. JS sharedCache
4. Extension Manager UI
5. RulePackTestRunner
6. Developer Diagnostics
7. Ledger delayed persistence
```

### Phase 3：生态化

```text
1. Native C# Plugin SDK
2. 官方规则包签名
3. 团队规则包共享
4. Aho-Corasick 多关键词引擎
5. Office.js companion architecture
6. Performance benchmark suite
```

---

## 19. 故障模式与降级策略

| 故障 | 处理 |
|---|---|
| 快照构建失败 | 降级为当前选择区域扫描 |
| ProtectedRange 构建失败 | 禁用自动修复，只允许建议 |
| Regex timeout | 禁用该 rule 当前 segment，记录诊断 |
| Regex 规则包加载失败 | 禁用规则包，不影响主插件 |
| JS timeout | 禁用该 JS pack 本次运行 |
| JS engine pool 耗尽 | 跳过 JS 扩展，核心扫描继续 |
| RangeMismatch | 跳过该 issue，标记 stale |
| Bookmark missing | ledger 标记 Missing，继续清除 |
| UndoRecord 失败 | 降级为无合并 Undo 的单项操作，并警告 |
| UI 渲染过载 | 限制显示前 N 条，分页加载 |
| 文档编辑导致快照失效 | UI 提示重新扫描 |

---

## 20. 多 Agent 开发分工

### Agent A：Snapshot + Segment

```text
DocumentSnapshot
DocumentSnapshotBuilder
ProtectedRangeIndex
ScannableSegmentIndex
SnapshotStalenessNotifier
```

红线：

```text
必须主线程读取 Word
不得修改 UI
不得实现 JS
```

### Agent B：Regex Rule Engine

```text
RulePackLoader
RegexRuleCompiler
RegexRuleScanner
RulePack schema
RulePack tests
```

红线：

```text
用户规则默认不能 SafeAutoFix
所有 Regex 必须 timeout
扫描 segment，不扫描 Word
```

### Agent C：IssueReducer + Diagnostics

```text
IssueReducer
CollisionScope
Telemetry
DeveloperDiagnostics
```

红线：

```text
不同 ModuleType 的重叠 issue 不得暴力去重
不得记录正文内容
```

### Agent D：Mutation + MarkerLedger

```text
MutationPlanner
MutationDeltaMap
WordMutationApplier
DocumentMutationGuard
MarkerLedger
ClearMarkers
```

红线：

```text
所有 Word 写操作必须主线程
所有写操作必须 Guard
清除不得逐字符遍历
```

### Agent E：UI

```text
VirtualIssueList
ProgressReporter
TaskState
ExtensionManager Phase 1 shell
SnapshotStaleness UI
```

红线：

```text
UI 线程不得执行长扫描
Issue List 必须虚拟化
```

### Agent F：Phase 2 JS

```text
JavaScriptEnginePool
ScriptPackLoader
ExtensionPermissionGuard
JS sharedCache
JS tests
```

红线：

```text
不得向 JS 暴露 Word/File/Network/Process
不得共享单例 Engine 多线程扫描
```

---

## 21. 关键代码骨架

### 21.1 ISnapshotScanner

```csharp
public interface ISnapshotScanner
{
    string ScannerId { get; }
    string ModuleType { get; }

    bool CanSkip(ScannableSegment segment);

    ValueTask<IReadOnlyList<IssueCandidate>> ScanAsync(
        ScannableSegment segment,
        ScannerExecutionContext context,
        CancellationToken cancellationToken);
}
```

### 21.2 IssueCandidate

```csharp
public sealed class IssueCandidate
{
    public string IssueId { get; init; } = Guid.NewGuid().ToString("N");
    public string RuleId { get; init; } = string.Empty;
    public string ModuleType { get; init; } = string.Empty;

    public int TextStart { get; init; }
    public int TextEnd { get; init; }

    public string OriginalText { get; init; } = string.Empty;
    public string? Suggestion { get; init; }

    public IssueSeverity Severity { get; init; }
    public FixSafetyLevel FixSafety { get; init; }

    public double Confidence { get; init; } = 1.0;
    public string SnapshotId { get; init; } = string.Empty;
    public ulong ParagraphHash { get; init; }

    public string SourceExtensionId { get; init; } = string.Empty;
}
```

### 21.3 MutationPlan

```csharp
public sealed class MutationPlan
{
    public string OperationId { get; init; } = Guid.NewGuid().ToString("N");
    public string SnapshotId { get; init; } = string.Empty;
    public IReadOnlyList<IssueCandidate> OrderedCandidates { get; init; } = Array.Empty<IssueCandidate>();
    public MutationOrder Order { get; init; }
    public int BatchSize { get; init; } = 10;
}
```

---

## 22. CI 与测试

### 22.1 必须加入的测试类型

```text
LongDocPerformanceTests
ProtectedRangeTests
ScannableSegmentTests
RegexRulePackTests
IssueReducerCollisionTests
MutationPlanDescendingTests
MutationDeltaMapTests
MarkerLedgerRecoveryTests
SnapshotStalenessTests
```

### 22.2 关键测试案例

```text
1. Fig.[OMath]1 不应被 Fig\.\s+\d+ 命中
2. URL 中的 10nm 不应被 unit-space-nm 命中
3. Zotero citation 中的括号不应被标点规则命中
4. References 后文不应被正文规则命中
5. 全文 1000 个 SafeFix 全局降序后不漂移
6. 用户删除 bookmark 后 ClearMarkers 不崩溃
7. 用户编辑文档后批量修复提示重新扫描
8. Regex timeout 不影响其他规则
9. 未签名规则包 SafeAutoFix 被降级
10. IssueReducer 保留 Formatting 与 Wording 重叠问题
```

---

## 23. 最终结论

v2.3 的核心定位是：

> **先把性能和安全做稳，再做扩展生态。**

最终推荐路线：

```text
Phase 1：
  ScannableSegment 主扫描路径
  Regex Rule Pack
  VirtualIssueList
  MarkerLedger
  BatchMutationApplier
  SnapshotStalenessNotifier

Phase 2：
  JS Script Pack
  Engine Pool
  Strict sandbox
  sharedCache
  Extension Manager

Phase 3：
  Native Plugin SDK
  Team rule distribution
  Office.js companion
  Advanced benchmark suite
```

最终原则：

> **所有扩展只能在快照世界发现问题；所有真实修改必须回到 WordMutationApplier；所有痕迹必须进入 MarkerLedger；所有批量修复必须由 MutationPlan 管理；所有长文档体验必须优先保证首屏响应。**

这样 ManuGuide 才能真正达到商业级 Word VSTO 插件的性能、安全与扩展性标准。
