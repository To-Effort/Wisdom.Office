using System;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 挂起类回归的子进程入口：主运行器里一次假死会带走整个套件，所以这类输入只在显式请求时单独跑。
    /// Child-process entry point for hang-class regressions: one freeze inside the main runner takes the
    /// whole suite down, so these inputs run only when explicitly requested.
    /// </summary>
    internal static partial class Program
    {
        /// <summary>按编号分发挂起类探针。Dispatches the hang-class probes by entry number.</summary>
        /// <param name="id">探针编号。The probe id.</param>
        /// <returns>探针通过返回 0，否则返回非 0。Zero when the probe passes, non-zero otherwise.</returns>
        private static int RunCrashProbe(string id)
        {
            switch (id)
            {
                case "WE1":
                    return ProbeOdtSpanProduct() ? 0 : 1;
                default:
                    Console.WriteLine("[crash-probe] unknown entry: " + id);
                    return 2;
            }
        }
    }
}
