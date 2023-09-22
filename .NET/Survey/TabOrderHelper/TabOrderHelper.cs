//
// デザイナーの「表示/タブオーダー」のように、階層化されたタブインデックスをリストで管理する
//
// 参考:https://zecl.hatenablog.com/entry/20090226/p1
//
using System.Collections.Generic;
using System.Linq;

namespace TabOrderHelper
{
    public sealed class TabOrderHelper
    {
        private System.Collections.Generic.List<ControlModel> _controlModels;

        public TabOrderHelper(System.Windows.Forms.Control rootControl)
        {
            _controlModels = new System.Collections.Generic.List<ControlModel>();
            CreateControlModels(rootControl);
        }

        public System.Windows.Forms.Control GetNextControl(System.Windows.Forms.Control control, bool forward = true)
        {
            return GetNextControl(FindControl(control), forward);
        }

        private void CreateControlModels(System.Windows.Forms.Control rootControl)
        {
            foreach (var c in GetAllControls(rootControl))
            {
                var indexString = GetHierarchicalTabIndicesString(c);
                var isContainer = IsContainer(c);
                var isRadioButton = IsRadioButton(c);

                var model = new ControlModel(c, indexString, isContainer, isRadioButton);
                _controlModels.Add(model);

                System.Diagnostics.Debug.WriteLine(model.ToString());
            }

            HasDuplicateLastIndex();

            return;
        }

        private System.Collections.Generic.IEnumerable<System.Windows.Forms.Control> GetAllControls(System.Windows.Forms.Control rootControl)
        {
            foreach (System.Windows.Forms.Control c in rootControl.Controls)
            {
                yield return c;
                foreach (System.Windows.Forms.Control a in GetAllControls(c))
                    yield return a;
            }
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
                sb.AppendFormat("{0},", item.ToString());
            return System.Text.RegularExpressions.Regex.Replace(sb.ToString(), ",$", "");
        }

        private bool IsParent(System.Windows.Forms.Control target)
        {
            if (target == null) return false;
            if (target is System.Windows.Forms.Form) return false;
            return true;
        }

        private bool IsContainer(System.Windows.Forms.Control target)
        {
            if (target is System.Windows.Forms.Panel ||
                target is System.Windows.Forms.GroupBox ||
                target is System.Windows.Forms.TabControl)
                return true;
            return false;
        }

        private bool IsRadioButton(System.Windows.Forms.Control target)
        {
            if (target is System.Windows.Forms.RadioButton)
                return true;
            return false;
        }

        private bool HasDuplicateLastIndex()
        {
            System.Collections.Generic.HashSet<int> uniqueNumbers = new System.Collections.Generic.HashSet<int>();

            foreach (var m in _controlModels)
            {
                var number = m.LastIndex;

                if (uniqueNumbers.Contains(number))
                {
                    throw new DuplicateTabIndexException($"Duplicate tab index values. Info=[{m.ToString()}]");
                }
                uniqueNumbers.Add(number);
            }
            return false;
        }

        private ControlModel FindControl(System.Windows.Forms.Control control)
        {
            foreach (var m in _controlModels)
            {
                if (m.Control.Name == control.Name) return m;
            }

            throw new ControlNotFoundException($"The specified control name cannot be found. Name=[{control.Name}]");
        }

        private System.Windows.Forms.Control GetNextControl(ControlModel model, bool forward = true)
        {
            var lastIndex = model.LastIndex;

            var nextLastIndex = forward ? FindNextGreaterNumber(lastIndex) : FindNextLessNumber(lastIndex);

            return FindControlByLastIndex(nextLastIndex).Control;
        }

        private int FindNextGreaterNumber(int lastIndex)
        {
            var model = _controlModels.FirstOrDefault(x => x.LastIndex > lastIndex);
            if (model == null)
            {
                model = _controlModels.OrderBy(x => x.LastIndex).FirstOrDefault();
            }
            return model.LastIndex;
        }

        private int FindNextLessNumber(int lastIndex)
        {
            var model = _controlModels.FirstOrDefault(x => x.LastIndex < lastIndex);
            if (model == null)
            {
                model = _controlModels.OrderByDescending(x => x.LastIndex).FirstOrDefault();
            }
            return model.LastIndex;
        }

        private ControlModel FindControlByLastIndex(int lastIndex)
        {
            return _controlModels.First(m => m.LastIndex == lastIndex);
        }

        private class ControlModel
        {
            private System.Windows.Forms.Control _control;
            private string _indexString;
            private int _lastIndex;
            private bool _isContainer;
            private bool _isRadioButton;

            private ControlModel()
            {
                // do nothing
            }

            public ControlModel(System.Windows.Forms.Control control, string indexString, bool isContainer = false, bool isRadioButton = false)
            {
                _control = control;
                _indexString = indexString;
                _lastIndex = GetLastNumber(_indexString);
                _isContainer = isContainer;
                _isRadioButton = isRadioButton;
            }

            public System.Windows.Forms.Control Control { get { return _control; } }
            public string IndexString { get { return _indexString; } }
            public int LastIndex { get { return _lastIndex; } }
            public bool IsContainer { get { return _isContainer; } }
            public bool IsRadioButton { get { return _isRadioButton; } }

            public override string ToString()
            {
                return $"Name={_control.Name}, TabIndex={_control.TabIndex}, IndexString={_indexString}, LastIndex={_lastIndex}, IsContainer={_isContainer}, IsRadioButton={_isRadioButton}";
            }

            private int GetLastNumber(string indexString)
            {
                var parts = indexString.Split(',');
                var lastPart = parts[parts.Length - 1];
                return int.Parse(lastPart);
            }
        }

        public class DuplicateTabIndexException : System.Exception
        {
            private DuplicateTabIndexException()
            {
                // do nothing
            }

            public DuplicateTabIndexException(string message) : base(message)
            {
            }
        }

        public class ControlNotFoundException : System.Exception
        {
            private ControlNotFoundException()
            {
                // do nothing
            }

            public ControlNotFoundException(string message) : base(message)
            {
            }
        }
    }
}
