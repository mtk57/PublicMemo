namespace TabOrderHelper
{
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
}
