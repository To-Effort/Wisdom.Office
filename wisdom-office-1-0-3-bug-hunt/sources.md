# Sources

本篇为第一手工程记录，事实来源以本仓库的代码、测试输出与提交历史为主；外部检索没有用于事实取证，标准条款仅按通用知识引用主题。

## 一手来源（可复核）

| # | 来源 | 类型 | 用途 | 备注 |
| --- | --- | --- | --- | --- |
| 1 | 仓库 `To-Effort/Wisdom.Office`（https://github.com/To-Effort/Wisdom.Office.git） | 代码仓库 | 全部代码级结论 | 用户在本次会话中提供的地址 |
| 2 | 镜像 `YFKael/WisdomOffice`（https://gitee.com/YFKael/WisdomOffice.git） | 代码仓库 | 发布同步核对 | 同上 |
| 3 | `docs/superpowers/specs/2026-11-02-wisdom-pdf-ppt-word-error-remediation-design.md` | 内部设计文档 | 条目编号（PD/PT/WE 前缀）、每条的目标行为与取舍 | 随修复持续更新，含红证原文与未修项清单 |
| 4 | `tests/Wisdom.Office.Tests/Remediation*.cs` | 测试代码 | 断言形态与失败原文 | 文中所有 `Test failed:` 均出自这些文件 |
| 5 | 本地提交记录（`f9fc99a`、`04ff01b`、`5cf54f5`、`91675ca`、`0505e02`、`0390975`、`e400f19`、`4ab840b`、`d2a5917`、`a892c2e`、`6dce2b6`、`71cdda7`、`256e1aa` 等） | 版本历史 | 修复顺序与拆分方式 | 提交信息里记录了红测/修复/文档的配对 |
| 6 | `nupkg/*.1.0.3.nupkg` 的 nuspec（本地解包核对） | 构建产物 | 版本号、六份目标框架、依赖区间 | 逐包读取 `<id>` / `<version>` / `<dependency>` |

## 标准与规范（按通用知识引用，不逐字引条文）

| # | 规范 | 涉及主题 | 用途 |
| --- | --- | --- | --- |
| 7 | ISO/IEC 29500（ECMA-376）OOXML | `CT_Properties`（`docProps/app.xml` 元素序列）、`CT_TLTimeNode*`（`p:timing` 结构）、`CT_TLCommonTimeNodeData`（`nodeType` / `presetID`）、`ST_SlideLayoutType`、`a:rPr/@baseline`、`a:tc@rowSpan` / `vMerge` / `hMerge` | 判断"这个 XML 结构合不合法" |
| 8 | ISO 32000-1（PDF 1.7） | 字体描述符 `/Flags` 位义、`/BBox` 与表单坐标系、`/Rotate` 只接受 90 的倍数 | 用于本系列早前条目；本文只在"同类问题的其他例子"里提及 |
| 9 | .NET 运行时行为 | `File.Replace` / `File.Move` 的同卷原子性、`DateTime.FromOADate` 的取值范围、`double → int` 的越界饱和 | 解释修复为什么这么写 |

## 未使用

- 没有引用任何第三方博客、论坛帖或厂商宣传材料作为事实依据。
- 没有引用其他 Office 库的缺陷率或性能对比数据（本会话未采集，故不写）。
- 没有给出 nuget.org 的包页面链接：1.0.3 尚未推送成功，链接在发布后才有意义；发布完成后可按包 ID 检索。

## 复核方式

- 失败原文：把注释掉的一行修复改回去，重跑测试即可复现（脚本与命令见 `post.md` 的"怎么证明修好了"一节）。
- 性能数字：本机单次测量（8000 行、两列），只用于说明数量级，不作为跨机器基准。
