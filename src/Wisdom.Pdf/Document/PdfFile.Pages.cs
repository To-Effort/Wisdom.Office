using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 对象级文档的页面操作：页面树扁平化、查询、重排、导入、合并与拆分。
    /// Page operations for the object-level document: flattening, inspection, reordering, importing,
    /// merging and splitting.
    /// </summary>
    public sealed partial class PdfFile
    {
        /// <summary>页面数量。The number of pages.</summary>
        public int PageCount
        {
            get { return GetPageNodes().Count; }
        }

        /// <summary>
        /// 取得指定页的页面字典，下标从 0 开始。
        /// Gets the page dictionary at the given zero-based index.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <returns>页面字典。The page dictionary.</returns>
        public PdfDictionary GetPageDictionary(int index)
        {
            return GetPageNode(index).Dictionary;
        }

        /// <summary>
        /// 取得指定页在页面树中的引用；直接对象页会先被提升为间接对象。
        /// Gets the page's reference in the page tree; a direct page is first promoted to an indirect object.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <returns>页面引用。The page reference.</returns>
        public PdfReference GetPageReference(int index)
        {
            PageNode node = GetPageNode(index);
            if (node.Reference == null)
            {
                int number = AddObject(node.Dictionary);
                node.Reference = new PdfReference(number, 0);
            }
            return node.Reference;
        }

        /// <summary>
        /// 取得页面尺寸（MediaBox 宽高）。
        /// Gets the page size from the MediaBox width and height.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <param name="width">页面宽度。The page width.</param>
        /// <param name="height">页面高度。The page height.</param>
        public void GetPageSize(int index, out double width, out double height)
        {
            PageNode node = GetPageNode(index);
            width = 612;
            height = 792;
            PdfArray box = Resolve(node.MediaBox) as PdfArray;
            if (box == null || box.Count < 4) return;
            PdfNumber x1 = Resolve(box[0]) as PdfNumber;
            PdfNumber y1 = Resolve(box[1]) as PdfNumber;
            PdfNumber x2 = Resolve(box[2]) as PdfNumber;
            PdfNumber y2 = Resolve(box[3]) as PdfNumber;
            if (x1 == null || y1 == null || x2 == null || y2 == null) return;
            width = Math.Abs(x2.Value - x1.Value);
            height = Math.Abs(y2.Value - y1.Value);
            if (width <= 0) width = 612;
            if (height <= 0) height = 792;
        }

        /// <summary>取得页面旋转角度（0/90/180/270）。Gets the page rotation (0/90/180/270).</summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <returns>旋转角度。The rotation angle.</returns>
        public int GetPageRotation(int index)
        {
            return GetPageNode(index).Rotation;
        }

        /// <summary>取得页面沿继承链合并后的资源字典。Gets the page resource dictionary after inheritance.</summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <returns>资源字典。The resource dictionary.</returns>
        public PdfDictionary GetPageResources(int index)
        {
            return GetPageNode(index).Resources;
        }

        /// <summary>
        /// 取得页面内容流解码后的字节；多个内容流以换行连接。
        /// Gets the decoded page content; multiple content streams are joined with newlines.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <returns>解码后的内容流。The decoded content.</returns>
        public byte[] GetPageContentBytes(int index)
        {
            PdfDictionary page = GetPageDictionary(index);
            PdfObject contents = Resolve(page.Get("Contents"));
            PdfStream stream = contents as PdfStream;
            if (stream != null) return DecodeStream(stream);
            PdfArray array = contents as PdfArray;
            if (array == null) return new byte[0];
            using (MemoryStream memory = new MemoryStream())
            {
                for (int i = 0; i < array.Count; i++)
                {
                    PdfStream item = Resolve(array[i]) as PdfStream;
                    if (item == null) continue;
                    byte[] data = DecodeStream(item);
                    memory.Write(data, 0, data.Length);
                    memory.WriteByte(10);
                }
                return memory.ToArray();
            }
        }

        /// <summary>
        /// 向页面追加内容流，保留原有内容。
        /// Appends a content stream to a page while keeping the existing content.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <param name="content">内容流字节。The content bytes.</param>
        /// <param name="compress">是否压缩写出。Whether to compress on write.</param>
        public void AppendPageContent(int index, byte[] content, bool compress = true)
        {
            if (content == null) throw new ArgumentNullException("content");
            PdfDictionary page = GetPageDictionary(index);
            PdfReference appended = AddObjectReference(new PdfStream(content, compress));
            PdfObject existing = page.Get("Contents");
            if (existing == null)
            {
                page.Set("Contents", appended);
            }
            else
            {
                PdfArray array = new PdfArray();
                PdfArray existingArray = Resolve(existing) as PdfArray;
                if (existingArray != null)
                {
                    for (int i = 0; i < existingArray.Count; i++) array.Add(existingArray[i]);
                }
                else
                {
                    array.Add(existing);
                }
                array.Add(appended);
                page.Set("Contents", array);
            }
            MarkModified(GetPageReference(index).Number);
        }

        /// <summary>
        /// 在页面原内容之前插入内容流，保留原有内容。
        /// Inserts a content stream before the existing page content while keeping that content.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <param name="content">内容流字节。The content bytes.</param>
        /// <param name="compress">是否压缩写出。Whether to compress on write.</param>
        /// <remarks>页面内容流数组内的各段是首尾相接的同一绘图状态，因此新内容画在正文之下，适合背景水印或底纹。
        /// The streams of a page's /Contents array share one graphics state and run in order, so the new content
        /// paints beneath the existing text, which is what a background watermark or tint needs.</remarks>
        public void PrependPageContent(int index, byte[] content, bool compress = true)
        {
            if (content == null) throw new ArgumentNullException("content");
            PdfDictionary page = GetPageDictionary(index);
            PdfReference prepended = AddObjectReference(new PdfStream(content, compress));
            PdfObject existing = page.Get("Contents");
            if (existing == null)
            {
                page.Set("Contents", prepended);
            }
            else
            {
                PdfArray array = new PdfArray();
                array.Add(prepended);
                PdfArray existingArray = Resolve(existing) as PdfArray;
                if (existingArray != null)
                {
                    for (int i = 0; i < existingArray.Count; i++) array.Add(existingArray[i]);
                }
                else
                {
                    array.Add(existing);
                }
                page.Set("Contents", array);
            }
            MarkModified(GetPageReference(index).Number);
        }

        /// <summary>
        /// 替换页面内容流。
        /// Replaces the page content with a single stream.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <param name="content">新的内容流字节。The new content bytes.</param>
        /// <param name="compress">是否压缩写出。Whether to compress on write.</param>
        public void SetPageContent(int index, byte[] content, bool compress = true)
        {
            if (content == null) throw new ArgumentNullException("content");
            PdfDictionary page = GetPageDictionary(index);
            page.Set("Contents", AddObjectReference(new PdfStream(content, compress)));
            MarkModified(GetPageReference(index).Number);
        }

        /// <summary>
        /// 向页面资源字典登记一个资源。Registers a resource in the page resource dictionary.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <param name="category">资源类别，如 Font、XObject、ExtGState。The resource category, such as Font, XObject or ExtGState.</param>
        /// <param name="name">资源名（不含斜杠）。The resource name without the slash.</param>
        /// <param name="value">资源对象。The resource object.</param>
        public void SetPageResource(int index, string category, string name, PdfObject value)
        {
            PdfDictionary page = GetPageDictionary(index);
            PdfDictionary resources = Resolve(page.Get("Resources")) as PdfDictionary;
            if (resources == null)
            {
                resources = GetPageResources(index);
                if (resources == null) resources = new PdfDictionary();
                page.Set("Resources", resources);
            }
            PdfDictionary group = Resolve(resources.Get(category)) as PdfDictionary;
            if (group == null)
            {
                group = new PdfDictionary();
                resources.Set(category, group);
            }
            group.Set(name, value);
            MarkModified(GetPageReference(index).Number);
        }

        /// <summary>
        /// 设置页面旋转，角度会规整到 0/90/180/270。
        /// Sets the page rotation, normalized to 0/90/180/270.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <param name="rotation">旋转角度。The rotation angle.</param>
        public void SetPageRotation(int index, int rotation)
        {
            int value = rotation % 360;
            if (value < 0) value += 360;
            PdfDictionary page = GetPageDictionary(index);
            page.SetNumber("Rotate", value);
            MarkModified(GetPageReference(index).Number);
            _pageNodes = null;
        }

        /// <summary>
        /// 设置页面裁剪框（左下原点坐标）。
        /// Sets the page crop box in bottom-left origin coordinates.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <param name="x">左边界。The left edge.</param>
        /// <param name="y">下边界。The bottom edge.</param>
        /// <param name="width">宽度。The width.</param>
        /// <param name="height">高度。The height.</param>
        public void SetPageCropBox(int index, double x, double y, double width, double height)
        {
            if (width <= 0 || height <= 0) throw new ArgumentOutOfRangeException("width");
            PdfDictionary page = GetPageDictionary(index);
            page.Set("CropBox", new PdfArray().AddNumber(x).AddNumber(y).AddNumber(x + width).AddNumber(y + height));
            MarkModified(GetPageReference(index).Number);
        }

        /// <summary>
        /// 删除页面并更新页面树。
        /// Removes a page and updates the page tree.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        public void RemovePage(int index)
        {
            EnsureFlatPageTree();
            PageNode node = GetPageNode(index);
            List<PdfReference> kids = GetPageKids();
            if (node.Reference != null) kids.RemoveAll(delegate(PdfReference reference) { return reference.Number == node.Reference.Number; });
            ApplyPageKids(kids);
            _pageNodes = null;
        }

        /// <summary>
        /// 把页面移动到新的下标。
        /// Moves a page to a new zero-based index.
        /// </summary>
        /// <param name="fromIndex">原下标。The source index.</param>
        /// <param name="toIndex">目标下标。The destination index.</param>
        public void MovePage(int fromIndex, int toIndex)
        {
            EnsureFlatPageTree();
            List<PageNode> nodes = GetPageNodes();
            if (fromIndex < 0 || fromIndex >= nodes.Count) throw new ArgumentOutOfRangeException("fromIndex");
            if (toIndex < 0 || toIndex >= nodes.Count) throw new ArgumentOutOfRangeException("toIndex");
            if (fromIndex == toIndex) return;
            List<PdfReference> kids = GetPageKids();
            PdfReference moved = kids[fromIndex];
            kids.RemoveAt(fromIndex);
            kids.Insert(toIndex, moved);
            ApplyPageKids(kids);
            _pageNodes = null;
        }

        /// <summary>
        /// 按给定顺序重排全部页面。
        /// Reorders all pages according to the supplied order.
        /// </summary>
        /// <param name="order">新的页面下标顺序。The new page index order.</param>
        public void ReorderPages(IEnumerable<int> order)
        {
            if (order == null) throw new ArgumentNullException("order");
            EnsureFlatPageTree();
            List<PdfReference> kids = GetPageKids();
            List<PdfReference> sorted = new List<PdfReference>();
            bool[] used = new bool[kids.Count];
            foreach (int index in order)
            {
                if (index < 0 || index >= kids.Count) throw new ArgumentOutOfRangeException("order", "Page index is outside the document.");
                if (used[index]) throw new ArgumentException("A page index is repeated in the order.", "order");
                used[index] = true;
                sorted.Add(kids[index]);
            }
            for (int i = 0; i < kids.Count; i++)
            {
                if (!used[i]) sorted.Add(kids[i]);
            }
            ApplyPageKids(sorted);
            _pageNodes = null;
        }

        /// <summary>
        /// 从另一个文档导入指定页并插入到目标下标处。
        /// Imports pages from another document and inserts them at the destination index.
        /// </summary>
        /// <param name="source">来源文档。The source document.</param>
        /// <param name="destinationIndex">目标插入位置。The destination insertion index.</param>
        /// <param name="pageIndexes">来源页下标；为 null 时导入全部页。The source page indexes; null imports every page.</param>
        public void InsertPagesFrom(PdfFile source, int destinationIndex, IEnumerable<int> pageIndexes = null)
        {
            if (source == null) throw new ArgumentNullException("source");
            EnsureFlatPageTree();
            List<PdfReference> kids = GetPageKids();
            if (destinationIndex < 0 || destinationIndex > kids.Count) throw new ArgumentOutOfRangeException("destinationIndex");

            List<int> indexes = pageIndexes == null ? null : new List<int>(pageIndexes);
            int count = indexes == null ? source.PageCount : indexes.Count;
            List<PdfReference> imported = new List<PdfReference>(count);
            for (int i = 0; i < count; i++)
            {
                int sourceIndex = indexes == null ? i : indexes[i];
                imported.Add(ImportPage(source, sourceIndex));
            }
            kids.InsertRange(destinationIndex, imported);
            ApplyPageKids(kids);
            _pageNodes = null;
        }

        /// <summary>
        /// 追加另一个文档的全部页面。
        /// Appends every page of another document.
        /// </summary>
        /// <param name="source">来源文档。The source document.</param>
        public void AppendPagesFrom(PdfFile source)
        {
            InsertPagesFrom(source, GetPageNodes().Count);
        }

        /// <summary>
        /// 追加另一个文档的指定页。
        /// Appends selected pages from another document.
        /// </summary>
        /// <param name="source">来源文档。The source document.</param>
        /// <param name="pageIndexes">来源页下标。The source page indexes.</param>
        public void AppendPagesFrom(PdfFile source, params int[] pageIndexes)
        {
            InsertPagesFrom(source, GetPageNodes().Count, pageIndexes);
        }

        /// <summary>
        /// 合并多个文档为一个新文档。
        /// Merges several documents into a new document.
        /// </summary>
        /// <param name="sources">来源文档，按顺序拼接。The source documents, concatenated in order.</param>
        /// <returns>包含全部页面的新文档。A new document containing all pages.</returns>
        public static PdfFile Merge(params PdfFile[] sources)
        {
            if (sources == null || sources.Length == 0) throw new ArgumentException("At least one source document is required.", "sources");
            PdfFile result = Create();
            for (int i = 0; i < sources.Length; i++)
            {
                if (sources[i] != null) result.AppendPagesFrom(sources[i]);
            }
            if (sources[0] != null && sources[0].Info != null)
            {
                result.Info = sources[0].Info;
                result.UpdateInfo();
            }
            return result;
        }

        /// <summary>
        /// 按页数拆分文档，每个结果至少包含一页。
        /// Splits the document by page count; every result contains at least one page.
        /// </summary>
        /// <param name="pagesPerDocument">每个结果文档的页数。The page count per result document.</param>
        /// <returns>拆分后的文档列表。The resulting documents.</returns>
        public IReadOnlyList<PdfFile> Split(int pagesPerDocument)
        {
            if (pagesPerDocument <= 0) throw new ArgumentOutOfRangeException("pagesPerDocument");
            List<PdfFile> result = new List<PdfFile>();
            for (int start = 0; start < PageCount; start += pagesPerDocument)
            {
                PdfFile part = Create();
                int end = Math.Min(PageCount, start + pagesPerDocument);
                for (int i = start; i < end; i++) part.AppendPagesFrom(this, i);
                part.Info = Info;
                part.UpdateInfo();
                result.Add(part);
            }
            return result;
        }

        /// <summary>
        /// 提取指定页面组成新文档。
        /// Extracts the selected pages into a new document.
        /// </summary>
        /// <param name="pageIndexes">要提取的页面下标。The page indexes to extract.</param>
        /// <returns>新文档。The new document.</returns>
        public PdfFile ExtractPages(IEnumerable<int> pageIndexes)
        {
            if (pageIndexes == null) throw new ArgumentNullException("pageIndexes");
            PdfFile result = Create();
            foreach (int index in pageIndexes) result.AppendPagesFrom(this, index);
            result.Info = Info;
            result.UpdateInfo();
            return result;
        }

        /// <summary>
        /// 把页面树扁平化为单层 /Pages，便于重排和删页。
        /// Flattens the page tree into a single /Pages node for easy reordering and deletion.
        /// </summary>
        public void EnsureFlatPageTree()
        {
            List<PageNode> nodes = GetPageNodes();
            PdfDictionary pages = new PdfDictionary();
            pages.SetName("Type", "Pages");
            PdfArray kids = new PdfArray();
            pages.Set("Kids", kids);
            pages.SetNumber("Count", nodes.Count);
            int pagesNumber = AddObject(pages);
            PdfReference pagesReference = new PdfReference(pagesNumber, 0);

            for (int i = 0; i < nodes.Count; i++)
            {
                PageNode node = nodes[i];
                if (node.Reference == null)
                {
                    node.Reference = new PdfReference(AddObject(node.Dictionary), 0);
                }
                PdfDictionary page = node.Dictionary;
                page.Set("Parent", pagesReference);
                if (page.Get("Resources") == null && node.Resources != null) page.Set("Resources", node.Resources);
                if (page.Get("MediaBox") == null && node.MediaBox != null) page.Set("MediaBox", node.MediaBox);
                if (page.Get("Rotate") == null && node.Rotation != 0) page.SetNumber("Rotate", node.Rotation);
                MarkModified(node.Reference.Number);
                kids.Add(node.Reference);
            }

            Catalog.Set("Pages", pagesReference);
            MarkModified(_catalogNumber);
            _pageNodes = null;
        }

        /// <summary>取得页面树子节点引用数组。Gets the page reference array from the tree.</summary>
        /// <returns>页面引用列表。The page references.</returns>
        private List<PdfReference> GetPageKids()
        {
            PdfDictionary pages = Resolve(Catalog.Get("Pages")) as PdfDictionary;
            if (pages == null) throw new InvalidDataException("The PDF page tree is missing.");
            PdfArray kids = Resolve(pages.Get("Kids")) as PdfArray;
            List<PdfReference> result = new List<PdfReference>();
            if (kids == null) return result;
            for (int i = 0; i < kids.Count; i++)
            {
                PdfReference reference = kids[i] as PdfReference;
                if (reference != null) result.Add(reference);
            }
            return result;
        }

        /// <summary>把引用数组写回页面树并更新计数。Writes the reference array back into the page tree and updates the count.</summary>
        /// <param name="kids">页面引用列表。The page references.</param>
        private void ApplyPageKids(List<PdfReference> kids)
        {
            PdfDictionary pages = Resolve(Catalog.Get("Pages")) as PdfDictionary;
            if (pages == null) throw new InvalidDataException("The PDF page tree is missing.");
            PdfArray array = new PdfArray();
            for (int i = 0; i < kids.Count; i++) array.Add(kids[i]);
            pages.Set("Kids", array);
            pages.SetNumber("Count", kids.Count);
            PdfReference pagesReference = Catalog.Get("Pages") as PdfReference;
            if (pagesReference != null) MarkModified(pagesReference.Number);
        }

        /// <summary>取得页面节点缓存，首次访问时遍历页面树。Gets the page node cache, walking the tree on first access.</summary>
        /// <returns>页面节点列表。The page nodes.</returns>
        private List<PageNode> GetPageNodes()
        {
            if (_pageNodes == null) BuildPageNodes();
            return _pageNodes;
        }

        /// <summary>取得指定页节点。Gets a page node by index.</summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <returns>页面节点。The page node.</returns>
        private PageNode GetPageNode(int index)
        {
            List<PageNode> nodes = GetPageNodes();
            if (index < 0 || index >= nodes.Count) throw new ArgumentOutOfRangeException("index");
            return nodes[index];
        }

        /// <summary>从目录开始遍历页面树。Walks the page tree from the catalog.</summary>
        private void BuildPageNodes()
        {
            _pageNodes = new List<PageNode>();
            PdfDictionary pages = Resolve(Catalog.Get("Pages")) as PdfDictionary;
            if (pages == null) return;
            CollectPageNodes(pages, null, null, null, 0, new HashSet<PdfDictionary>());
        }

        /// <summary>递归收集页面节点并下传继承属性。Recursively collects pages and propagates inherited attributes.</summary>
        /// <param name="node">当前节点。The current node.</param>
        /// <param name="reference">当前节点引用。The current node reference.</param>
        /// <param name="inheritedResources">继承资源。Inherited resources.</param>
        /// <param name="inheritedMediaBox">继承媒体框。Inherited media box.</param>
        /// <param name="inheritedRotation">继承旋转。Inherited rotation.</param>
        /// <param name="visited">已访问节点。Visited nodes.</param>
        private void CollectPageNodes(PdfDictionary node, PdfReference reference, PdfDictionary inheritedResources,
            PdfObject inheritedMediaBox, int inheritedRotation, HashSet<PdfDictionary> visited)
        {
            if (node == null || visited.Contains(node)) return;
            visited.Add(node);

            PdfDictionary resources = Resolve(node.Get("Resources")) as PdfDictionary ?? inheritedResources;
            PdfObject mediaBox = node.Get("MediaBox") ?? inheritedMediaBox;
            PdfNumber rotationNumber = Resolve(node.Get("Rotate")) as PdfNumber;
            int rotation = rotationNumber == null ? inheritedRotation : NormalizeRotation((int)rotationNumber.Value);

            PdfArray kids = Resolve(node.Get("Kids")) as PdfArray;
            if (kids != null)
            {
                for (int i = 0; i < kids.Count; i++)
                {
                    PdfReference kidReference = kids[i] as PdfReference;
                    PdfDictionary kid = Resolve(kids[i]) as PdfDictionary;
                    if (kid != null) CollectPageNodes(kid, kidReference, resources, mediaBox, rotation, visited);
                }
                return;
            }

            PageNode page = new PageNode();
            page.Dictionary = node;
            page.Reference = reference;
            page.Resources = resources;
            page.MediaBox = mediaBox;
            page.Rotation = rotation;
            _pageNodes.Add(page);
        }

        /// <summary>把旋转角规整到 90 度倍数。Normalizes a rotation to a multiple of 90 degrees.</summary>
        /// <param name="rotation">原始角度。The raw angle.</param>
        /// <returns>0、90、180 或 270。One of 0, 90, 180 or 270.</returns>
        private static int NormalizeRotation(int rotation)
        {
            int value = rotation % 360;
            if (value < 0) value += 360;
            if (value < 45 || value >= 315) return 0;
            if (value < 135) return 90;
            if (value < 225) return 180;
            return 270;
        }

        /// <summary>
        /// 把来源文档的一页（含资源、内容、注释）深拷贝到当前文档。
        /// Deep-copies one source page, including resources, content and annotations, into this document.
        /// </summary>
        /// <param name="source">来源文档。The source document.</param>
        /// <param name="sourceIndex">来源页下标。The source page index.</param>
        /// <returns>新页面引用。The new page reference.</returns>
        private PdfReference ImportPage(PdfFile source, int sourceIndex)
        {
            if (sourceIndex < 0 || sourceIndex >= source.PageCount) throw new ArgumentOutOfRangeException("sourceIndex");
            PageNode node = source.GetPageNode(sourceIndex);
            PdfReference sourcePageReference = node.Reference;
            PdfReference newReference = new PdfReference(AddObject(PdfNull.Instance), 0);
            Dictionary<int, PdfReference> mapping = new Dictionary<int, PdfReference>();
            if (sourcePageReference != null) mapping[sourcePageReference.Number] = newReference;
            PdfDictionary clone = CloneObject(source, node.Dictionary, mapping, sourcePageReference, newReference) as PdfDictionary;
            if (clone == null) throw new InvalidDataException("Unable to clone the source page.");
            SetObject(newReference.Number, clone);

            PdfReference pagesReference = Catalog.Get("Pages") as PdfReference;
            clone.Set("Parent", pagesReference);
            if (clone.Get("Resources") == null && node.Resources != null) clone.Set("Resources", CloneObject(source, node.Resources, mapping, sourcePageReference, newReference));
            if (clone.Get("MediaBox") == null && node.MediaBox != null) clone.Set("MediaBox", CloneObject(source, node.MediaBox, mapping, sourcePageReference, newReference));
            if (clone.Get("Rotate") == null && node.Rotation != 0) clone.SetNumber("Rotate", node.Rotation);

            PdfArray annotations = Resolve(clone.Get("Annots")) as PdfArray;
            if (annotations != null)
            {
                for (int i = 0; i < annotations.Count; i++)
                {
                    PdfDictionary annotation = Resolve(annotations[i]) as PdfDictionary;
                    if (annotation != null) annotation.Set("P", newReference);
                }
            }
            MarkModified(newReference.Number);
            return newReference;
        }

        /// <summary>
        /// 深拷贝对象图并重映射引用，阻断页面与注释之间的回指环。
        /// Deep-copies an object graph and remaps references, breaking page/annotation back-reference cycles.
        /// </summary>
        /// <param name="source">来源文档。The source document.</param>
        /// <param name="value">待拷贝对象。The object to copy.</param>
        /// <param name="mapping">对象号到新引用的映射。Map from source object number to new reference.</param>
        /// <param name="sourcePageReference">来源页面引用。The source page reference.</param>
        /// <param name="newPageReference">新页面引用，可为 null。The new page reference, or null.</param>
        /// <returns>拷贝出的对象。The copied object.</returns>
        private PdfObject CloneObject(PdfFile source, PdfObject value, Dictionary<int, PdfReference> mapping,
            PdfReference sourcePageReference, PdfReference newPageReference)
        {
            if (value == null) return null;
            PdfReference reference = value as PdfReference;
            if (reference != null)
            {
                if (sourcePageReference != null && reference.Number == sourcePageReference.Number)
                {
                    return newPageReference;
                }
                PdfReference existing;
                if (mapping.TryGetValue(reference.Number, out existing)) return existing;
                PdfObject target = source.LoadObject(reference.Number);
                if (target == null) return PdfNull.Instance;
                PdfReference placeholder = new PdfReference(AddObject(PdfNull.Instance), 0);
                mapping[reference.Number] = placeholder;
                PdfObject copied = CloneObject(source, target, mapping, sourcePageReference, newPageReference);
                SetObject(placeholder.Number, copied);
                return placeholder;
            }

            PdfString text = value as PdfString;
            if (text != null) return new PdfString((byte[])text.Bytes.Clone(), text.IsHex);

            PdfArray array = value as PdfArray;
            if (array != null)
            {
                PdfArray copy = new PdfArray();
                for (int i = 0; i < array.Count; i++)
                {
                    copy.Add(CloneObject(source, array[i], mapping, sourcePageReference, newPageReference));
                }
                return copy;
            }

            PdfDictionary dictionary = value as PdfDictionary;
            if (dictionary != null)
            {
                PdfDictionary copy = new PdfDictionary();
                foreach (KeyValuePair<string, PdfObject> pair in dictionary)
                {
                    if (pair.Key == "Parent" && pair.Value is PdfReference && ((PdfReference)pair.Value).Number == GetPagesReferenceNumber(source)) continue;
                    if (pair.Key == "P" && sourcePageReference != null && pair.Value is PdfReference && ((PdfReference)pair.Value).Number == sourcePageReference.Number) continue;
                    copy.Set(pair.Key, CloneObject(source, pair.Value, mapping, sourcePageReference, newPageReference));
                }
                return copy;
            }

            PdfStream stream = value as PdfStream;
            if (stream != null)
            {
                PdfDictionary dictionaryCopy = CloneObject(source, stream.Dictionary, mapping, sourcePageReference, newPageReference) as PdfDictionary;
                return new PdfStream(dictionaryCopy, (byte[])stream.Data.Clone(), false);
            }

            if (value is PdfNumber) return new PdfNumber(((PdfNumber)value).Value);
            if (value is PdfName) return new PdfName(((PdfName)value).Value);
            if (value is PdfBoolean) return ((PdfBoolean)value).Value ? PdfBoolean.True : PdfBoolean.False;
            if (value is PdfNull) return PdfNull.Instance;
            return PdfNull.Instance;
        }

        /// <summary>
        /// 取得来源文档页面树根的对象号，用于跳过 /Parent 回指。
        /// Gets the source page tree root object number, used to skip /Parent back-references.
        /// </summary>
        /// <param name="source">来源文档。The source document.</param>
        /// <returns>对象号；直接对象时返回 -1。The object number, or -1 for a direct object.</returns>
        private static int GetPagesReferenceNumber(PdfFile source)
        {
            PdfReference reference = source.Catalog.Get("Pages") as PdfReference;
            return reference == null ? -1 : reference.Number;
        }
    }
}
