namespace TabOrderHelper
{
    /// <summary>
    /// タブオーダー共通クラス
    /// </summary>
    internal static class Common
    {
        /// <summary>
        /// セパレーター
        /// </summary>
        public const char SEP = ':';
    }

    /// <summary>
    /// 階層タブインデックスインターフェース
    /// </summary>
    internal interface IHasHierarchicalTabIndices :
        System.Windows.Forms.IWin32Window,
        System.Collections.Generic.IEnumerable<int>,
        System.IComparable,
        System.IComparable<IHasHierarchicalTabIndices>
    {
        /// <summary>
        /// 階層タブインデックスをシーケンスで返す
        /// </summary>
        System.Collections.Generic.IEnumerable<int> HierarchicalTabIndices { get; }
    }

    /// <summary>
    /// ControlNotFoundException
    /// </summary>
    internal class ControlNotFoundException : System.Exception
    {
        private ControlNotFoundException()
        {
            // do nothing
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="message">メッセージ</param>
        public ControlNotFoundException(string message) : base(message)
        {
            // do nothing
        }
    }

    /// <summary>
    /// Win32APIアクセス用クラス
    /// </summary>
    internal static class PlatformInvoker
    {
        /// <summary>
        /// GetWindow関数のコマンド
        /// </summary>
        public enum GetWindowCmd
        {
            GW_HWNDFIRST = 0,
            GW_HWNDLAST = 1,
            GW_HWNDNEXT = 2,
            GW_HWNDPREV = 3,
            GW_OWNER = 4,
            GW_CHILD = 5,
            GW_ENABLEDPOPUP = 6
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern System.IntPtr GetWindow(System.IntPtr hwd, uint uCmd);
    }

    /// <summary>
    /// ソート順
    /// </summary>
    internal enum Sort
    {
        Asc,
        Desc
    }

    /// <summary>
    /// 階層タブインデックスのソートヘルパークラス
    /// </summary>
    internal class SortHelperOfHierarchicalTabIndices :
        System.Collections.Generic.IComparer<IHasHierarchicalTabIndices>
    {
        private int _toggle = 1;

        /// <summary>
        /// デフォルトコンストラクタ
        /// </summary>
        public SortHelperOfHierarchicalTabIndices()
        {
            // do nothing
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="sort">ソート順</param>
        public SortHelperOfHierarchicalTabIndices(Sort sort)
        {
            switch (sort)
            {
                case Sort.Asc:
                    break;
                case Sort.Desc:
                    _toggle = -1;
                    break;
                default:
                    _toggle = 1;
                    break;
            }
        }

        /// <summary>
        /// コンペア
        /// List.Sort()から呼び出される
        /// </summary>
        /// <param name="x">比較元のタブインデックス</param>
        /// <param name="y">比較先のタブインデックス</param>
        /// <returns>コンペア結果(0:x == y, 1:x > y, -1:x < y</returns>
        public int Compare(IHasHierarchicalTabIndices x, IHasHierarchicalTabIndices y)
        {
            using (var enumerator1 = x.GetEnumerator())
            using (var enumerator2 = y.GetEnumerator())
            {
                var e1 = enumerator1.MoveNext();
                var e2 = enumerator2.MoveNext();

                while (e1 && e2)
                {
                    var compare = enumerator1.Current.CompareTo(enumerator2.Current) * _toggle;
                    if (compare != 0)
                        return compare;

                    e1 = enumerator1.MoveNext();
                    e2 = enumerator2.MoveNext();
                }

                // 比較対象が無くなった

                if (!e1 && !e2)
                    // TabIndexの階層構造に全く同じ値が設定されていた場合はZオーダーで比較する
                    return CompareZOrder(x.Handle, y.Handle);

                if (!e1)
                    return -1 * _toggle;

                if (!e2)
                    return 1 * _toggle;
            }
            return 0;
        }

        /// <summary>
        /// Zオーダーのコンペア
        /// </summary>
        /// <param name="hwdx">比較元のウィンドウハンドル</param>
        /// <param name="hwdy">比較先のウィンドウハンドル</param>
        /// <returns>コンペア結果(0:x == y, 1:x > y, -1:x < y</returns>
        private int CompareZOrder(System.IntPtr hwdx, System.IntPtr hwdy)
        {
            var h = PlatformInvoker.GetWindow(hwdx, (uint)PlatformInvoker.GetWindowCmd.GW_HWNDNEXT);
            while (h != default(System.IntPtr))
            {
                if (h == hwdy)
                    return -1 * _toggle;

                h = PlatformInvoker.GetWindow(h, (uint)PlatformInvoker.GetWindowCmd.GW_HWNDNEXT);
            }

            h = PlatformInvoker.GetWindow(hwdx, (uint)PlatformInvoker.GetWindowCmd.GW_HWNDPREV);
            while (h != default(System.IntPtr))
            {
                if (h == hwdy)
                    return 1 * _toggle;

                h = PlatformInvoker.GetWindow(h, (uint)PlatformInvoker.GetWindowCmd.GW_HWNDPREV);
            }
            return 0;
        }
    }
}
