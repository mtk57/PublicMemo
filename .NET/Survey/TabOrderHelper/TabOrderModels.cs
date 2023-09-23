namespace TabOrderHelper
{
    internal sealed class TabOrderModel : IHasHierarchicalTabIndices
    {
        private System.Collections.Generic.IEnumerable<int> _hierarchicalTabIndices;

        private System.Windows.Forms.Control _prevControl;
        private System.Windows.Forms.Control _control;
        private System.Windows.Forms.Control _nextControl;

        private string _indexString;
        private int _parentLastIndex;
        private int _lastIndex;
        private int _tabIndex;  // これが内部的にナンバリングした重複無しのタブインデックス値

        private bool _isContainer;
        private bool _isRadioButton;
        private bool _isTabStop;

        private TabOrderModel()
        {
            // do nothing
        }

        public TabOrderModel(System.Windows.Forms.Control control)
        {
            _hierarchicalTabIndices = GetHierarchicalTabindices(control);

            _prevControl = null;
            _control = control;
            _nextControl = null;
            _indexString = GetHierarchicalTabIndicesString(_control);
            _parentLastIndex = GetPreviousNumber(_indexString);
            _lastIndex = GetLastNumber(_indexString);
            _tabIndex = -1;
            _isRadioButton = _control is System.Windows.Forms.RadioButton;
            _isTabStop = false;
        }

        public System.Windows.Forms.Control PrevControl { get { return _prevControl; } set { _prevControl = value; } }
        public System.Windows.Forms.Control Control { get { return _control; } }
        public System.Windows.Forms.Control NextControl { get { return _nextControl; } set { _nextControl = value; } }
        public string IndexString { get { return _indexString; } }
        public int ParentLastIndex { get { return _parentLastIndex; } }
        public int LastIndex { get { return _lastIndex; } }
        public int TabIndex { get { return _lastIndex; } set { _tabIndex = value; } }
        public bool IsContainer { get { return _isContainer; } }
        public bool IsRadioButton { get { return _isRadioButton; } }
        public bool IsTabStop { get { return _isTabStop; } set { _isTabStop = value; } }
        public System.IntPtr Handle { get { return _control.Handle; } }

        public override string ToString()
        {
            if (_prevControl == null)
            {
                return $"Name={_control.Name}\t" +
                       $"TabIndex={_control.TabIndex}\t" +
                       $"IndexString={_indexString}\t" +
                       $"ParentLastIndex={_parentLastIndex}\t" +
                       $"LastIndex={_lastIndex}\t" +
                       $"TabIndex={_tabIndex}\t" +
                       $"IsContainer={_isContainer}\t" +
                       $"IsRadioButton={_isRadioButton}\t" +
                       $"IsTabStop={_isTabStop}";
            }

            return $"Name={_control.Name}\t" +
                   $"PrevTabIndex={_prevControl.TabIndex}\t" +
                   $"TabIndex={_control.TabIndex}\t" +
                   $"NextTabIndex={_nextControl.TabIndex}\t" +
                   $"IndexString={_indexString}\t" +
                   $"ParentLastIndex={_parentLastIndex}\t" +
                   $"LastIndex={_lastIndex}\t" +
                   $"TabIndex={_tabIndex}\t" +
                   $"IsContainer={_isContainer}\t" +
                   $"IsRadioButton={_isRadioButton}\t" +
                   $"IsTabStop={_isTabStop}";
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
