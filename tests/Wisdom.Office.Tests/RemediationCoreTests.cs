using System;
using System.Globalization;
using Wisdom.Office.Core;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// `Wisdom.Office.Core` 的 Templates / Utilities 审计条目。
    /// Audit entries for `Wisdom.Office.Core`'s Templates and Utilities.
    /// </summary>
    internal static partial class Program
    {
        /// <summary>Core 审计入口。Entry point for the Core audit.</summary>
        private static void RunRemediationCoreTest()
        {
            RunBadTemplateFormatDegradesTest();

            Console.WriteLine("[remediation] core template formatting degradation ok");
        }

        /// <summary>
        /// 模板格式串来自模板、取值来自数据，一个坏格式不该让整份文档渲染失败：`{{when:'unclosed}}` 里那个
        /// 没闭合的引号会被 `DateTime.ToString` 判为非法格式并抛 `FormatException`，过去它直接穿出
        /// `TemplateEngine.Replace`，一张报表就因为一个占位符全废。现在退回不带格式的默认转换，值仍然出现。
        /// A template format string comes from the template and the value from the data, so one bad format
        /// must not fail the whole render: the unclosed quote in `{{when:'unclosed}}` makes
        /// `DateTime.ToString` throw `FormatException`, which used to escape `TemplateEngine.Replace` and lose
        /// an entire report over one placeholder. It now falls back to the unformatted conversion so the value
        /// still appears.
        /// </summary>
        private static void RunBadTemplateFormatDegradesTest()
        {
            TemplateContext context = new TemplateContext();
            context["when"] = new DateTime(2026, 3, 4, 5, 6, 7);
            context["amount"] = 12.5;
            TemplateOptions options = new TemplateOptions { Culture = CultureInfo.InvariantCulture };

            string rendered = TemplateEngine.Replace("date=[{{when:'unclosed}}] amount=[{{amount:0.00}}]", context, options);
            Require(rendered.IndexOf("date=[", StringComparison.Ordinal) == 0,
                "the text around the token was lost: " + rendered);
            Require(rendered.IndexOf("2026", StringComparison.Ordinal) >= 0,
                "the value disappeared instead of falling back to its default format: " + rendered);
            Require(rendered.IndexOf("amount=[12.50]", StringComparison.Ordinal) >= 0,
                "a valid format in the same template stopped working: " + rendered);

            // 控制：格式串为空时仍按默认转换，坏格式的分支不能改变这条老路。
            // Control: an empty format still goes through the default conversion, so the new branch did not
            // change the old path.
            string plain = TemplateEngine.Replace("v={{amount}}", context, options);
            Require(plain == "v=12.5", "the unformatted substitution changed: " + plain);
        }
    }
}
