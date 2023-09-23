namespace TabOrderHelper
{
    internal sealed class TabOrderModel : IHasHierarchicalTabIndices
    {
        private System.Collections.Generic.IEnumerable<int> _hierarchicalTabIndices;

        public System.Windows.Forms.Control PrevControl { get; set; }
        public System.Windows.Forms.Control Control { get; }
        public System.Windows.Forms.Control NextControl { get; set; }
        public string IndexString { get; }
        public int ParentLastIndex { get; }
        public int LastIndex { get; }

        /// <summary>
        /// 内部的にナンバリングした重複無しのタブインデックス値
        /// </summary>
        public int UniqueTabIndex { get; set; }

        public bool IsContainer { get; }
        public bool IsRadioButton { get; }
        public bool IsTabStop { get; set; }
        public System.IntPtr Handle { get; }

        private TabOrderModel()
        {
            // do nothing
        }

        public TabOrderModel(System.Windows.Forms.Control control)
        {
            _hierarchicalTabIndices = GetHierarchicalTabindices(control);

            PrevControl = null;
            Control = control;
            NextControl = null;
            IndexString = GetHierarchicalTabIndicesString(control);
            ParentLastIndex = GetPreviousNumber(IndexString);
            LastIndex = GetLastNumber(IndexString);
            UniqueTabIndex = -1;
            IsRadioButton = control is System.Windows.Forms.RadioButton;
            IsTabStop = false;
        }

        public override string ToString()
        {
            if (PrevControl == null)
            {
                return $"Name={Control.Name}\t" +
                       $"TabIndex={Control.TabIndex}\t" +
                       $"IndexString={IndexString}\t" +
                       $"ParentLastIndex={ParentLastIndex}\t" +
                       $"LastIndex={LastIndex}\t" +
                       $"UniqueTabIndex={UniqueTabIndex}\t" +
                       $"IsContainer={IsContainer}\t" +
                       $"IsRadioButton={IsRadioButton}\t" +
                       $"IsTabStop={IsTabStop}";
            }

            return $"Name={Control.Name}\t" +
                   $"PrevTabIndex={PrevControl.TabIndex}\t" +
                   $"TabIndex={Control.TabIndex}\t" +
                   $"NextTabIndex={NextControl.TabIndex}\t" +
                   $"IndexString={IndexString}\t" +
                   $"ParentLastIndex={ParentLastIndex}\t" +
                   $"LastIndex={LastIndex}\t" +
                   $"UniqueTabIndex={UniqueTabIndex}\t" +
                   $"IsContainer={IsContainer}\t" +
                   $"IsRadioButton={IsRadioButton}\t" +
                   $"IsTabStop={IsTabStop}";
        }

        public System.Collections.Generic.IEnumerable<int> HierarchicalTabIndices
        {
            get { return _hierarchicalTabIndices; }
        }

        public int CompareTo(object obj)
        {
            return CompareTo((IHasHierarchicalTabIndices)obj);
        }

        public int CompareTo(IHasHierarchicalTabIndices other)
        {
            return new SortHelperOfHierarchicalTabIndices().Compare(this, other);
        }

        public System.Collections.Generic.IEnumerator<int> GetEnumerator()
        {
            return this.HierarchicalTabIndices.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return (System.Collections.IEnumerator)this.HierarchicalTabIndices.GetEnumerator();
        }

        private System.Collections.Generic.IEnumerable<int> GetHierarchicalTabindices(System.Windows.Forms.Control control)
        {
            var s = new System.Collections.Generic.Stack<int>();
            s.Push(control.TabIndex);
            var parent = control.Parent;
            while (IsParent(parent))
            {
                s.Push(parent.TabIndex);
                parent = parent.Parent;
            }

            while (s.Count != 0)
                yield return s.Pop();
        }

        private string GetHierarchicalTabIndicesString(System.Windows.Forms.Control control)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var item in GetHierarchicalTabindices(control))
                sb.AppendFormat("{0}" + Common.SEP, item.ToString());
            return System.Text.RegularExpressions.Regex.Replace(sb.ToString(), Common.SEP + "$", "");
        }

        private bool IsParent(System.Windows.Forms.Control target)
        {
            if (target == null) return false;
            if (target is System.Windows.Forms.Form) return false;
            return true;
        }

        private int GetPreviousNumber(string indexString)
        {
            var numbers = indexString.Split(Common.SEP);
            var length = numbers.Length;
            var secondLastNumber = -1;  //コンテナに内包されていない場合
            if (length > 1)
            {
                int.TryParse(numbers[length - 2], out secondLastNumber);
            }
            return secondLastNumber;
        }

        private int GetLastNumber(string indexString)
        {
            var parts = indexString.Split(Common.SEP);
            var lastPart = parts[parts.Length - 1];
            return int.Parse(lastPart);
        }
    }
}
