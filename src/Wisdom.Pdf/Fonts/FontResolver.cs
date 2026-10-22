using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 把字体家族解析为本机已安装的 TrueType/TTC 字面，支持注册自定义字体文件，全程不调用系统字体 API。
    /// Resolves font families to TrueType/TTC faces installed on the machine, with support for registering
    /// custom font files. No system font APIs are used.
    /// </summary>
    /// <remarks>
    /// 候选按显式注册、家族名匹配、粗斜体子族等打分择优；文件字节与字面列表均有缓存，线程安全由单一锁保证。
    /// Candidates are scored by explicit registration, family-name match and bold/italic subfamily hints;
    /// file bytes and face lists are cached, and thread safety relies on a single lock.
    /// </remarks>
    public sealed class FontResolver
    {
        /// <summary>字体来源登记项：文件路径或内存字节，加家族与粗斜体标记。A font source registration: file path or in-memory bytes plus family and bold/italic flags.</summary>
        private sealed class FontSource
        {
            /// <summary>缓存键（含路径/字节标识与面号）。The cache key, including the path or bytes identity and face index.</summary>
            public string Key;
            /// <summary>字体文件路径，字节来源时为 null。The font file path; null for byte sources.</summary>
            public string FilePath;
            /// <summary>内存字体字节，文件来源时为 null。The in-memory font bytes; null for file sources.</summary>
            public byte[] Data;
            /// <summary>集合字体中的面号，-1 表示全部加载。The face index within a collection; -1 loads every face.</summary>
            public int FaceIndex;
            /// <summary>是否为用户显式注册（优先于目录扫描）。Whether the user registered it explicitly, taking priority over directory scanning.</summary>
            public bool Explicit;
            /// <summary>注册时指定的家族名。The family name given at registration.</summary>
            public string Family;
            /// <summary>注册时指定的粗体标记。The bold flag given at registration.</summary>
            public bool Bold;
            /// <summary>注册时指定的斜体标记。The italic flag given at registration.</summary>
            public bool Italic;
        }

        /// <summary>已知字体家族登记：名称、别名与候选文件名。A known font family registration: name, aliases and candidate file names.</summary>
        private sealed class FamilyRegistration
        {
            /// <summary>规范家族名。The canonical family name.</summary>
            public string Name;
            /// <summary>可匹配的全部别名（含规范名自身）。All matchable aliases, including the canonical name.</summary>
            public List<string> Aliases = new List<string>();
            /// <summary>按序尝试的字体文件名。The font file names tried in order.</summary>
            public List<string> FileNames = new List<string>();
        }

        /// <summary>解析过程中的一个候选来源。One candidate source during resolution.</summary>
        private sealed class Candidate
        {
            /// <summary>候选对应的字体来源。The font source of this candidate.</summary>
            public FontSource Source;
            /// <summary>候选所属家族名。The family name of the candidate.</summary>
            public string Family;
            /// <summary>是否来自显式注册。Whether it comes from explicit registration.</summary>
            public bool Explicit;
        }

        /// <summary>打分后的最优候选字面。The best-scoring candidate face.</summary>
        private sealed class CandidateScore
        {
            /// <summary>胜出的字面对象。The winning face.</summary>
            public TrueTypeFace Face;
            /// <summary>该字面的得分。The face's score.</summary>
            public int Score;
        }

        /// <summary>保护共享集合与缓存的锁。The lock guarding shared collections and caches.</summary>
        private readonly object _sync = new object();
        /// <summary>显式注册的字体来源列表。The list of explicitly registered font sources.</summary>
        private readonly List<FontSource> _explicitSources = new List<FontSource>();
        /// <summary>已知家族登记表。The known-family registrations.</summary>
        private readonly List<FamilyRegistration> _families = new List<FamilyRegistration>();
        /// <summary>扫描用字体目录列表。The font directories to scan.</summary>
        private readonly List<string> _fontDirectories = new List<string>();
        /// <summary>路径到文件字节的缓存。Cache from path to file bytes.</summary>
        private readonly Dictionary<string, byte[]> _fileBytes = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        /// <summary>来源键到字面列表的缓存。Cache from source key to loaded faces.</summary>
        private readonly Dictionary<string, List<TrueTypeFace>> _faceCache = new Dictionary<string, List<TrueTypeFace>>(StringComparer.OrdinalIgnoreCase);
        /// <summary>小写文件名到完整路径的目录索引，惰性构建。The lowercase file-name to full-path index, built lazily.</summary>
        private Dictionary<string, string> _fileIndex;

        /// <summary>进程级默认解析器的构建锁。The lock building the process-wide default resolver.</summary>
        private static readonly object DefaultSync = new object();
        /// <summary>惰性创建的进程级默认解析器。The lazily created process-wide default resolver.</summary>
        private static FontResolver _default;

        /// <summary>进程级共享的默认字体解析器。The process-wide shared default font resolver.</summary>
        public static FontResolver Default
        {
            get
            {
                lock (DefaultSync)
                {
                    if (_default == null) _default = new FontResolver();
                    return _default;
                }
            }
        }

        /// <summary>
        /// 创建解析器：扫描各平台字体目录并登记已知家族。
        /// Creates a resolver: scans the platform font directories and registers the known families.
        /// </summary>
        public FontResolver()
        {
            InitializeFontDirectories();
            RegisterKnownFamilies();
        }

        /// <summary>
        /// 登记 Windows/Linux/macOS 的常见字体目录，目录不存在则跳过。
        /// Registers the common font directories for Windows/Linux/macOS, skipping missing ones.
        /// </summary>
        private void InitializeFontDirectories()
        {
            string windows = Environment.GetEnvironmentVariable("WINDIR");
            if (!string.IsNullOrEmpty(windows))
            {
                AddDirectory(Path.Combine(windows, "Fonts"));
            }
            string localAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA");
            if (!string.IsNullOrEmpty(localAppData))
            {
                AddDirectory(Path.Combine(localAppData, "Microsoft", "Windows", "Fonts"));
            }

            string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (!string.IsNullOrEmpty(home))
            {
                AddDirectory(Path.Combine(home, ".fonts"));
                AddDirectory(Path.Combine(home, ".local", "share", "fonts"));
                AddDirectory(Path.Combine(home, "Library", "Fonts"));
            }

            AddDirectory("/usr/share/fonts");
            AddDirectory("/usr/local/share/fonts");
            AddDirectory("/Library/Fonts");
            AddDirectory("/System/Library/Fonts");
            AddDirectory("/System/Library/Fonts/Supplemental");
        }

        /// <summary>
        /// 加入一个字体目录，已存在、不存在或无权访问时静默跳过。
        /// Adds a font directory, silently skipping duplicates, missing paths or access failures.
        /// </summary>
        /// <param name="path">目录路径。The directory path.</param>
        private void AddDirectory(string path)
        {
            try
            {
                if (Directory.Exists(path) && !_fontDirectories.Contains(path))
                {
                    _fontDirectories.Add(path);
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// 登记内置的常见字体家族及其文件名（西文、中日韩与 Linux 发行版字体）。
        /// Registers the built-in common font families and their file names (Latin, CJK and Linux fonts).
        /// </summary>
        private void RegisterKnownFamilies()
        {
            AddFamily("Arial", new[] { "Helvetica", "Liberation Sans" }, "arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf");
            AddFamily("Arial Black", new string[0], "ariblk.ttf");
            AddFamily("Times New Roman", new[] { "Times", "Liberation Serif", "serif" }, "times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf");
            AddFamily("Courier New", new[] { "Courier", "Liberation Mono", "monospace" }, "cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf");
            AddFamily("Calibri", new string[0], "calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf");
            AddFamily("Cambria", new string[0], "cambria.ttc", "cambriab.ttf", "cambriai.ttf", "cambriaz.ttf");
            AddFamily("Candara", new string[0], "candara.ttf", "candarab.ttf", "candarai.ttf", "candaraz.ttf");
            AddFamily("Constantia", new string[0], "constantia.ttf", "constantiab.ttf", "constantiab.ttf", "constantiaz.ttf");
            AddFamily("Corbel", new string[0], "corbel.ttf", "corbelb.ttf", "corbeli.ttf", "corbelz.ttf");
            AddFamily("Consolas", new string[0], "consola.ttf", "consolab.ttf", "consolai.ttf", "consolaz.ttf");
            AddFamily("Georgia", new string[0], "georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf");
            AddFamily("Verdana", new string[0], "verdana.ttf", "verdanab.ttf", "verdanai.ttf", "verdanaz.ttf");
            AddFamily("Tahoma", new string[0], "tahoma.ttf", "tahomabd.ttf");
            AddFamily("Trebuchet MS", new string[0], "trebuc.ttf", "trebucbd.ttf", "trebucit.ttf", "trebucbi.ttf");
            AddFamily("Segoe UI", new string[0], "segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf");
            AddFamily("Segoe UI Symbol", new string[0], "seguisym.ttf");
            AddFamily("Segoe UI Emoji", new string[0], "seguiemj.ttf");
            AddFamily("Microsoft Sans Serif", new string[0], "micross.ttf");
            AddFamily("Arial Unicode MS", new string[0], "arialuni.ttf");

            AddFamily("Microsoft YaHei", new[] { "微软雅黑", "Microsoft YaHei UI" }, "msyh.ttc", "msyhbd.ttc", "msyhl.ttc");
            AddFamily("SimSun", new[] { "宋体", "NSimSun", "新宋体" }, "simsun.ttc", "simsunb.ttf");
            AddFamily("SimHei", new[] { "黑体" }, "simhei.ttf");
            AddFamily("KaiTi", new[] { "楷体" }, "simkai.ttf");
            AddFamily("FangSong", new[] { "仿宋" }, "simfang.ttf");
            AddFamily("DengXian", new[] { "等线" }, "Deng.ttf", "Dengb.ttf", "Dengl.ttf");
            AddFamily("Microsoft JhengHei", new[] { "微软正黑体" }, "msjh.ttc", "msjhbd.ttc");
            AddFamily("MingLiU", new[] { "細明體", "细明体" }, "mingliu.ttc", "mingliub.ttc");
            AddFamily("Malgun Gothic", new[] { "맑은 고딕" }, "malgun.ttf", "malgunbd.ttf");
            AddFamily("Yu Gothic", new[] { "游ゴシック" }, "YuGothR.ttc", "YuGothB.ttc", "YuGothM.ttc");
            AddFamily("MS Gothic", new[] { "ＭＳ ゴシック" }, "msgothic.ttc");
            AddFamily("Meiryo", new string[0], "meiryo.ttc", "meiryob.ttc");
            AddFamily("PingFang SC", new[] { "苹方" }, "PingFang.ttc", "PingFangSC-Regular.ttf");
            AddFamily("Hiragino Sans GB", new string[0], "Hiragino Sans GB.ttc");

            AddFamily("Noto Sans CJK SC", new[] { "Noto Sans SC", "NotoSansCJKsc" }, "NotoSansCJK-Regular.ttc", "NotoSansCJK-Bold.ttc", "NotoSansSC-Regular.ttf");
            AddFamily("Noto Sans CJK JP", new string[0], "NotoSansCJK-Regular.ttc", "NotoSansCJK-Bold.ttc");
            AddFamily("Noto Sans", new string[0], "NotoSans-Regular.ttf", "NotoSans-Bold.ttf", "NotoSans-Italic.ttf", "NotoSans-BoldItalic.ttf");
            AddFamily("WenQuanYi Zen Hei", new[] { "文泉驿正黑" }, "wqy-zenhei.ttc");
            AddFamily("WenQuanYi Micro Hei", new[] { "文泉驿微米黑" }, "wqy-microhei.ttc");
            AddFamily("Droid Sans Fallback", new string[0], "DroidSansFallbackFull.ttf", "DroidSansFallback.ttf");
            AddFamily("DejaVu Sans", new string[0], "DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf");
            AddFamily("DejaVu Serif", new string[0], "DejaVuSerif.ttf", "DejaVuSerif-Bold.ttf", "DejaVuSerif-Italic.ttf");
            AddFamily("DejaVu Sans Mono", new string[0], "DejaVuSansMono.ttf", "DejaVuSansMono-Bold.ttf");
            AddFamily("Liberation Sans", new string[0], "LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf");
            AddFamily("Liberation Serif", new string[0], "LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf", "LiberationSerif-Italic.ttf");
            AddFamily("Liberation Mono", new string[0], "LiberationMono-Regular.ttf", "LiberationMono-Bold.ttf");
        }

        /// <summary>
        /// 登记一个家族：别名与规范名一并可匹配，文件名去重后按序保存。
        /// Registers a family: aliases and the canonical name all match, file names are deduplicated and kept in order.
        /// </summary>
        /// <param name="name">规范家族名。The canonical family name.</param>
        /// <param name="aliases">别名列表，可为 null。The alias list; may be null.</param>
        /// <param name="fileNames">候选字体文件名。The candidate font file names.</param>
        private void AddFamily(string name, string[] aliases, params string[] fileNames)
        {
            FamilyRegistration registration = new FamilyRegistration();
            registration.Name = name;
            if (aliases != null) registration.Aliases.AddRange(aliases);
            registration.Aliases.Add(name);
            for (int i = 0; i < fileNames.Length; i++)
            {
                if (!string.IsNullOrEmpty(fileNames[i]) && !registration.FileNames.Contains(fileNames[i]))
                {
                    registration.FileNames.Add(fileNames[i]);
                }
            }
            _families.Add(registration);
        }

        /// <summary>
        /// 注册一个字体文件，自动识别其中全部字面。
        /// Registers a font file, automatically discovering all faces inside it.
        /// </summary>
        /// <param name="family">家族名。The family name.</param>
        /// <param name="bold">是否作为粗体注册。Whether it registers the bold face.</param>
        /// <param name="italic">是否作为斜体注册。Whether it registers the italic face.</param>
        /// <param name="filePath">字体文件路径。The font file path.</param>
        public void RegisterFont(string family, bool bold, bool italic, string filePath)
        {
            RegisterFont(family, bold, italic, filePath, -1);
        }

        /// <summary>
        /// 注册字体文件中指定序号的字面。
        /// Registers one face of a font file by index.
        /// </summary>
        /// <param name="family">家族名。The family name.</param>
        /// <param name="bold">是否作为粗体注册。Whether it registers the bold face.</param>
        /// <param name="italic">是否作为斜体注册。Whether it registers the italic face.</param>
        /// <param name="filePath">字体文件路径。The font file path.</param>
        /// <param name="faceIndex">TTC 集合中的面号，-1 表示全部。The face index within a TTC collection; -1 takes all faces.</param>
        /// <exception cref="ArgumentException">family 或 filePath 为空时抛出。Thrown when family or filePath is empty.</exception>
        public void RegisterFont(string family, bool bold, bool italic, string filePath, int faceIndex)
        {
            if (string.IsNullOrEmpty(family)) throw new ArgumentException("family");
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath");
            FontSource source = new FontSource();
            source.Key = "file:" + filePath + "#" + faceIndex;
            source.FilePath = filePath;
            source.FaceIndex = faceIndex;
            source.Explicit = true;
            source.Family = family;
            source.Bold = bold;
            source.Italic = italic;
            lock (_sync)
            {
                _explicitSources.Add(source);
            }
        }

        /// <summary>
        /// 注册内存中的字体数据，自动识别其中全部字面。
        /// Registers in-memory font data, automatically discovering all faces inside it.
        /// </summary>
        /// <param name="family">家族名。The family name.</param>
        /// <param name="bold">是否作为粗体注册。Whether it registers the bold face.</param>
        /// <param name="italic">是否作为斜体注册。Whether it registers the italic face.</param>
        /// <param name="data">字体文件字节。The font file bytes.</param>
        public void RegisterFont(string family, bool bold, bool italic, byte[] data)
        {
            RegisterFont(family, bold, italic, data, -1);
        }

        /// <summary>
        /// 注册内存字体数据中指定序号的字面。
        /// Registers one face of in-memory font data by index.
        /// </summary>
        /// <param name="family">家族名。The family name.</param>
        /// <param name="bold">是否作为粗体注册。Whether it registers the bold face.</param>
        /// <param name="italic">是否作为斜体注册。Whether it registers the italic face.</param>
        /// <param name="data">字体文件字节。The font file bytes.</param>
        /// <param name="faceIndex">TTC 集合中的面号，-1 表示全部。The face index within a TTC collection; -1 takes all faces.</param>
        /// <exception cref="ArgumentException">family 为空时抛出。Thrown when family is empty.</exception>
        /// <exception cref="ArgumentNullException">data 为 null 时抛出。Thrown when data is null.</exception>
        public void RegisterFont(string family, bool bold, bool italic, byte[] data, int faceIndex)
        {
            if (string.IsNullOrEmpty(family)) throw new ArgumentException("family");
            if (data == null) throw new ArgumentNullException("data");
            FontSource source = new FontSource();
            source.Key = "bytes:" + family + "|" + bold + "|" + italic + "|" + data.Length + "|" + data.GetHashCode() + "#" + faceIndex;
            source.Data = data;
            source.FaceIndex = faceIndex;
            source.Explicit = true;
            source.Family = family;
            source.Bold = bold;
            source.Italic = italic;
            lock (_sync)
            {
                _explicitSources.Add(source);
            }
        }

        /// <summary>
        /// 按规范化名查找家族登记项（匹配任一别名）。
        /// Finds a family registration by normalized name, matching any alias.
        /// </summary>
        /// <param name="family">家族名。The family name.</param>
        /// <returns>登记项或 null。The registration, or null.</returns>
        private FamilyRegistration FindFamilyRegistration(string family)
        {
            string normalized = NormalizeFamily(family);
            for (int i = 0; i < _families.Count; i++)
            {
                for (int a = 0; a < _families[i].Aliases.Count; a++)
                {
                    if (NormalizeFamily(_families[i].Aliases[a]) == normalized)
                    {
                        return _families[i];
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 解析家族名为最匹配的字面。
        /// Resolves a family name to the best-matching face.
        /// </summary>
        /// <param name="family">家族名，空时按 Arial 处理。The family name; empty falls back to Arial.</param>
        /// <param name="bold">是否要求粗体。Whether bold is requested.</param>
        /// <param name="italic">是否要求斜体。Whether italic is requested.</param>
        /// <returns>最优字面；完全找不到时返回 null。The best face, or null when nothing matches.</returns>
        /// <remarks>候选依次来自显式注册、已知家族文件名，最后按文件名猜测；每个候选的每个字面经 <see cref="ScoreFace"/> 打分取最高。
        /// Candidates come from explicit registrations, then known-family file names, then name guesses; every face of
        /// each candidate is scored via <see cref="ScoreFace"/> and the best wins.</remarks>
        public TrueTypeFace Resolve(string family, bool bold, bool italic)
        {
            if (string.IsNullOrEmpty(family)) family = "Arial";
            string normalized = NormalizeFamily(family);

            List<Candidate> candidates = new List<Candidate>();
            lock (_sync)
            {
                for (int i = 0; i < _explicitSources.Count; i++)
                {
                    FontSource source = _explicitSources[i];
                    if (NormalizeFamily(source.Family) == normalized)
                    {
                        Candidate candidate = new Candidate();
                        candidate.Source = source;
                        candidate.Family = family;
                        candidate.Explicit = true;
                        candidates.Add(candidate);
                    }
                }

                for (int i = 0; i < _families.Count; i++)
                {
                    FamilyRegistration registration = _families[i];
                    bool matches = false;
                    for (int a = 0; a < registration.Aliases.Count; a++)
                    {
                        if (NormalizeFamily(registration.Aliases[a]) == normalized)
                        {
                            matches = true;
                            break;
                        }
                    }
                    if (!matches) continue;

                    for (int f = 0; f < registration.FileNames.Count; f++)
                    {
                        string path = FindFontFile(registration.FileNames[f]);
                        if (path == null) continue;
                        FontSource source = new FontSource();
                        source.Key = "file:" + path + "#-1";
                        source.FilePath = path;
                        source.FaceIndex = -1;
                        Candidate candidate = new Candidate();
                        candidate.Source = source;
                        candidate.Family = registration.Name;
                        candidates.Add(candidate);
                    }
                }
            }

            if (candidates.Count == 0)
            {
                FindFileCandidates(normalized, candidates, family);
            }

            CandidateScore best = new CandidateScore();
            best.Score = int.MinValue;
            for (int i = 0; i < candidates.Count; i++)
            {
                List<TrueTypeFace> faces = LoadFaces(candidates[i].Source);
                for (int f = 0; f < faces.Count; f++)
                {
                    TrueTypeFace face = faces[f];
                    if (face == null) continue;
                    int score = ScoreFace(face, family, bold, italic, f, candidates[i].Explicit, candidates[i].Source);
                    if (score > best.Score)
                    {
                        best.Score = score;
                        best.Face = face;
                    }
                }
            }

            return best.Face;
        }

        /// <summary>
        /// 家族名未命中登记表时，按去空格/连字符的常见文件名形式在目录索引中猜测候选。
        /// When the family is not in the registration table, guesses file-name candidates by stripping spaces
        /// and hyphens against the directory index.
        /// </summary>
        /// <param name="normalizedFamily">规范化家族名。The normalized family name.</param>
        /// <param name="candidates">输出候选列表。Outputs candidates into this list.</param>
        /// <param name="originalFamily">原始家族名。The original family name.</param>
        private void FindFileCandidates(string normalizedFamily, List<Candidate> candidates, string originalFamily)
        {
            Dictionary<string, string> index = GetFileIndex();
            if (index == null) return;

            string compact = normalizedFamily.Replace(" ", string.Empty).Replace("-", string.Empty);
            List<string> guesses = new List<string>();
            guesses.Add(normalizedFamily.Replace(" ", string.Empty) + ".ttf");
            guesses.Add(normalizedFamily.Replace(" ", string.Empty) + ".ttc");
            guesses.Add(compact + ".ttf");
            guesses.Add(compact + ".ttc");

            for (int i = 0; i < guesses.Count; i++)
            {
                string path;
                if (index.TryGetValue(guesses[i].ToLowerInvariant(), out path))
                {
                    Candidate candidate = new Candidate();
                    candidate.Source = new FontSource
                    {
                        Key = "file:" + path + "#-1",
                        FilePath = path,
                        FaceIndex = -1
                    };
                    candidate.Family = originalFamily;
                    candidates.Add(candidate);
                }
            }
        }

        /// <summary>
        /// 给候选字面打分：家族名精确/包含匹配、显式注册来源、粗斜体一致性与子族提示词逐级加减分。
        /// Scores a candidate face using exact/substring family matches, explicit-source bonuses, bold/italic
        /// agreement and subfamily hint words.
        /// </summary>
        /// <param name="face">待打分字面。The face to score.</param>
        /// <param name="family">请求的家族名。The requested family name.</param>
        /// <param name="bold">是否请求粗体。Whether bold is requested.</param>
        /// <param name="italic">是否请求斜体。Whether italic is requested.</param>
        /// <param name="faceIndex">字面在文件中的序号。The face index within the file.</param>
        /// <param name="explicitSource">候选是否来自显式注册。Whether the candidate is from an explicit registration.</param>
        /// <param name="source">候选来源登记项。The candidate source entry.</param>
        /// <returns>相对得分，越大越优。The relative score; higher is better.</returns>
        private int ScoreFace(TrueTypeFace face, string family, bool bold, bool italic, int faceIndex, bool explicitSource, FontSource source)
        {
            int score = 0;
            string requested = NormalizeFamily(family);
            string faceFamily = NormalizeFamily(face.FamilyName);
            if (faceFamily == requested) score += 500;
            else if (faceFamily.Contains(requested) || requested.Contains(faceFamily)) score += 200;
            else if (!explicitSource) score -= 300;

            if (source.Explicit)
            {
                if (source.Bold == bold) score += 80; else score -= 40;
                if (source.Italic == italic) score += 80; else score -= 40;
                score += 200;
            }

            if (face.IsBold == bold) score += 120; else score -= 70;
            if (face.IsItalic == italic) score += 120; else score -= 60;

            string subfamily = (face.SubfamilyName ?? string.Empty).ToLowerInvariant();
            if (!bold && !italic && subfamily.Contains("regular")) score += 30;
            if (!bold && (subfamily.Contains("light") || subfamily.Contains("thin") || subfamily.Contains("black"))) score -= 50;
            if (subfamily.Contains("condensed") || subfamily.Contains("narrow")) score -= 20;
            score += Math.Max(0, 10 - faceIndex);
            return score;
        }

        /// <summary>
        /// 按码点类别（CJK/emoji/其他符号）沿预设家族顺序寻找含该字形的回退字面。
        /// Finds a fallback face containing the code point by walking the per-category family priority list
        /// (CJK, emoji, other symbols).
        /// </summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <param name="bold">是否要求粗体。Whether bold is requested.</param>
        /// <param name="italic">是否要求斜体。Whether italic is requested.</param>
        /// <returns>首个含该字形的字面，找不到返回 null。The first face with the glyph, or null.</returns>
        public TrueTypeFace ResolveFallback(int codePoint, bool bold, bool italic)
        {
            string[] names;
            if (IsCjkCodePoint(codePoint))
            {
                names = new[]
                {
                    "Microsoft YaHei", "SimSun", "SimHei", "Noto Sans CJK SC", "Noto Sans CJK JP",
                    "WenQuanYi Zen Hei", "Droid Sans Fallback", "PingFang SC", "Hiragino Sans GB",
                    "Yu Gothic", "MS Gothic", "Malgun Gothic", "Arial Unicode MS"
                };
            }
            else if (codePoint >= 0x1F000 && codePoint <= 0x1FAFF)
            {
                names = new[] { "Segoe UI Emoji", "Segoe UI Symbol", "Noto Sans", "DejaVu Sans", "Arial Unicode MS", "Microsoft YaHei" };
            }
            else
            {
                names = new[] { "Segoe UI Symbol", "Arial Unicode MS", "Arial", "DejaVu Sans", "Noto Sans", "Liberation Sans" };
            }

            for (int i = 0; i < names.Length; i++)
            {
                TrueTypeFace face = Resolve(names[i], bold, italic);
                if (face != null && face.HasGlyph(codePoint))
                {
                    return face;
                }
            }
            return null;
        }

        /// <summary>判断码点是否落在 CJK 区间（含假名、韩文、兼容表意文字、全角与扩展 B 及以后）。Whether the code point falls in a CJK range, including kana, Hangul, compatibility ideographs, fullwidth forms and extension B onward.</summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <returns>属于 CJK 返回 true。True when the code point is CJK.</returns>
        public static bool IsCjkCodePoint(int codePoint)
        {
            return (codePoint >= 0x2E80 && codePoint <= 0x9FFF)
                || (codePoint >= 0xAC00 && codePoint <= 0xD7AF)
                || (codePoint >= 0xF900 && codePoint <= 0xFAFF)
                || (codePoint >= 0xFE30 && codePoint <= 0xFE4F)
                || (codePoint >= 0xFF00 && codePoint <= 0xFFEF)
                || (codePoint >= 0x20000 && codePoint <= 0x3FFFF);
        }

        /// <summary>
        /// 加载来源的全部字面并按键缓存；解析失败的位置以 null 占位，文件字节另有独立缓存。
        /// Loads and caches every face of a source keyed by its cache key; unparseable slots stay null, and
        /// file bytes have their own cache.
        /// </summary>
        /// <param name="source">字体来源。The font source.</param>
        /// <returns>字面列表，可能含 null 元素。The face list, which may contain nulls.</returns>
        private List<TrueTypeFace> LoadFaces(FontSource source)
        {
            string key = source.Key;
            lock (_sync)
            {
                List<TrueTypeFace> cached;
                if (_faceCache.TryGetValue(key, out cached))
                {
                    return cached;
                }
            }

            byte[] data = source.Data;
            if (data == null)
            {
                try
                {
                    if (!_fileBytes.TryGetValue(source.FilePath, out data))
                    {
                        data = File.ReadAllBytes(source.FilePath);
                        lock (_sync) _fileBytes[source.FilePath] = data;
                    }
                }
                catch (Exception)
                {
                    data = null;
                }
            }

            List<TrueTypeFace> faces = new List<TrueTypeFace>();
            if (data != null)
            {
                int count = source.FaceIndex >= 0 ? source.FaceIndex + 1 : TrueTypeFace.GetFaceCount(data);
                if (source.FaceIndex >= 0) count = Math.Max(count, 1);
                for (int i = 0; i < count; i++)
                {
                    if (source.FaceIndex >= 0 && i != source.FaceIndex) 
                    {
                        faces.Add(null);
                        continue;
                    }
                    try
                    {
                        int offset = TrueTypeFace.GetFaceOffset(data, i);
                        TrueTypeFace face = new TrueTypeFace(data, offset);
                        faces.Add(face);
                    }
                    catch (Exception)
                    {
                        faces.Add(null);
                    }
                }
            }

            lock (_sync)
            {
                _faceCache[key] = faces;
            }
            return faces;
        }

        /// <summary>
        /// 读取并缓存字体文件字节。
        /// Reads and caches the bytes of a font file.
        /// </summary>
        /// <param name="path">文件路径。The file path.</param>
        /// <returns>文件字节。The file bytes.</returns>
        private byte[] GetFileBytes(string path)
        {
            lock (_sync)
            {
                byte[] data;
                if (_fileBytes.TryGetValue(path, out data)) return data;
            }
            byte[] bytes = File.ReadAllBytes(path);
            lock (_sync) _fileBytes[path] = bytes;
            return bytes;
        }

        /// <summary>
        /// 按文件名定位字体：绝对路径直接验证，否则查目录索引。
        /// Locates a font by file name: rooted paths are validated directly, otherwise the directory index
        /// is consulted.
        /// </summary>
        /// <param name="fileName">文件名或绝对路径。The file name or absolute path.</param>
        /// <returns>实际路径，未找到返回 null。The real path, or null when missing.</returns>
        private string FindFontFile(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return null;
            try
            {
                if (Path.IsPathRooted(fileName) && File.Exists(fileName)) return fileName;
            }
            catch (Exception)
            {
            }

            Dictionary<string, string> index = GetFileIndex();
            if (index == null) return null;
            string path;
            if (index.TryGetValue(Path.GetFileName(fileName).ToLowerInvariant(), out path))
            {
                return path;
            }
            return null;
        }

        /// <summary>
        /// 惰性构建「小写文件名 → 路径」索引：递归扫描各字体目录的 ttf/ttc，失败时退化为只扫顶层。
        /// Builds the "lowercase file name to path" index lazily by scanning each font directory for
        /// ttf/ttc recursively, falling back to the top level when that fails.
        /// </summary>
        /// <returns>文件名索引。The file-name index.</returns>
        private Dictionary<string, string> GetFileIndex()
        {
            lock (_sync)
            {
                if (_fileIndex != null) return _fileIndex;
                Dictionary<string, string> index = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int d = 0; d < _fontDirectories.Count; d++)
                {
                    try
                    {
                        string[] files;
                        try
                        {
                            files = Directory.GetFiles(_fontDirectories[d], "*.ttf", SearchOption.AllDirectories);
                        }
                        catch (Exception)
                        {
                            files = Directory.GetFiles(_fontDirectories[d], "*.ttf", SearchOption.TopDirectoryOnly);
                        }
                        for (int i = 0; i < files.Length; i++)
                        {
                            string name = Path.GetFileName(files[i]).ToLowerInvariant();
                            if (!index.ContainsKey(name)) index[name] = files[i];
                        }

                        string[] collections;
                        try
                        {
                            collections = Directory.GetFiles(_fontDirectories[d], "*.ttc", SearchOption.AllDirectories);
                        }
                        catch (Exception)
                        {
                            collections = Directory.GetFiles(_fontDirectories[d], "*.ttc", SearchOption.TopDirectoryOnly);
                        }
                        for (int i = 0; i < collections.Length; i++)
                        {
                            string name = Path.GetFileName(collections[i]).ToLowerInvariant();
                            if (!index.ContainsKey(name)) index[name] = collections[i];
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
                _fileIndex = index;
                return index;
            }
        }

        /// <summary>
        /// 规范化家族名：转小写并只保留字母与数字，令 "Times New Roman" 与 "timesnewroman" 等同。
        /// Normalizes a family name to lowercase letters and digits only, so "Times New Roman" compares
        /// equal to "timesnewroman".
        /// </summary>
        /// <param name="value">原始名称。The raw name.</param>
        /// <returns>规范化结果；null 或空返回空字符串。The normalized form; empty string for null or empty input.</returns>
        public static string NormalizeFamily(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            StringBuilder builder = new StringBuilder(value.Length);
            for (int i = 0; i < value.Length; i++)
            {
                char c = char.ToLowerInvariant(value[i]);
                if (char.IsLetterOrDigit(c))
                {
                    builder.Append(c);
                }
            }
            return builder.ToString();
        }
    }
}
