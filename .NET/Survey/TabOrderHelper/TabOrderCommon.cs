namespace TabOrderHelper
{
    internal static class Common
    {
        public const char SEP = ':';
    }

    /// <summary>
    /// 
    /// </summary>
    internal interface IHasHierarchicalTabIndices :
        System.Windows.Forms.IWin32Window,
        System.Collections.Generic.IEnumerable<int>,
        System.IComparable,
        System.IComparable<IHasHierarchicalTabIndices>
    {
        System.Collections.Generic.IEnumerable<int> HierarchicalTabIndices { get; }
    }

    /// <summary>
    /// 
    /// </summary>
    internal class ControlNotFoundException : System.Exception
    {
        private ControlNotFoundException()
        {
            // do nothing
        }

        public ControlNotFoundException(string message) : base(message)
        {
            // do nothing
        }
    }

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

    internal enum Sort
    {
        Asc,
        Desc
    }

    /// <summary>
    /// SortHelperOfHierarchicalTabIndices
    /// </summary>
    internal class SortHelperOfHierarchicalTabIndices :
        System.Collections.Generic.IComparer<IHasHierarchicalTabIndices>
    {
        private int _togleNum = 1;
        
        private SortHelperOfHierarchicalTabIndices()
        {
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="sort"></param>
        public SortHelperOfHierarchicalTabIndices(Sort sort)
        {
            switch (sort)
            {
                case Sort.Asc:
                    break;
                case Sort.Desc:
                    _togleNum = -1;
                    break;
                default:
                    _togleNum = 1;
                    break;
            }
        }

        public int Compare(IHasHierarchicalTabIndices x, IHasHierarchicalTabIndices y)
        {
            using (System.Collections.Generic.IEnumerator<int> enumerator1 = x.GetEnumerator())
            using (System.Collections.Generic.IEnumerator<int> enumerator2 = y.GetEnumerator())
            {
                bool e1 = enumerator1.MoveNext();
                bool e2 = enumerator2.MoveNext();

                while (e1 && e2)
                {
                    int compare = enumerator1.Current.CompareTo(enumerator2.Current) * _togleNum;
                    if (compare != 0)
                        return compare;

                    e1 = enumerator1.MoveNext();
                    e2 = enumerator2.MoveNext();
                }
                if (!e1 && !e2)
                    return CompareZOrder(x.Handle, y.Handle);
                if (!e1)
                    return -1 * _togleNum;
                if (!e2)
                    return 1 * _togleNum;
            }
            return 0;
        }

        private int CompareZOrder(System.IntPtr hwdx, System.IntPtr hwdy)
        {
            var h = PlatformInvoker.GetWindow(hwdx, (uint)PlatformInvoker.GetWindowCmd.GW_HWNDNEXT);
            while (h != default(System.IntPtr))
            {
                if (h == hwdy)
                    return -1 * _togleNum;

                h = PlatformInvoker.GetWindow(h, (uint)PlatformInvoker.GetWindowCmd.GW_HWNDNEXT);
            }

            h = PlatformInvoker.GetWindow(hwdx, (uint)PlatformInvoker.GetWindowCmd.GW_HWNDPREV);
            while (h != default(System.IntPtr))
            {
                if (h == hwdy)
                    return 1 * _togleNum;

                h = PlatformInvoker.GetWindow(h, (uint)PlatformInvoker.GetWindowCmd.GW_HWNDPREV);
            }
            return 0;
        }
    }
}
