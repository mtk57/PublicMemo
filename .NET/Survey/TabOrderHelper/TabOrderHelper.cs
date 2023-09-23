//
// デザイナーの「表示/タブオーダー」のように、階層化されたタブインデックスをリストで管理する
//
// 参考:https://zecl.hatenablog.com/entry/20090226/p1
//      https://atmarkit.itmedia.co.jp/fdotnet/dotnettips/243winkeyproc/winkeyproc.html
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace TabOrderHelper
{
    public sealed class TabOrderHelper
    {
        private const char SEP = ':';

        private System.Collections.Generic.List<ControlModel> _controlModels;

        private TabOrderHelper()
        {
            // do nothing
        }

        public TabOrderHelper(System.Windows.Forms.Control rootControl)
        {
            _controlModels = new System.Collections.Generic.List<ControlModel>();
            CreateControlModels(rootControl);
        }

        public System.Windows.Forms.Control GetNextControl(System.Windows.Forms.Control control, bool forward = true)
        {
            // TODO:タブオーダーは基本的に静的な情報なので、辞書で持っておくとよいかも。

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
            }

            HasDuplicateLastIndex();

            SetTabStop();

            foreach (var c in _controlModels)
                System.Diagnostics.Debug.WriteLine(c.ToString());

            // TODO:タブオーダーは基本的に静的な情報なので、辞書で持っておくとよいかも。

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
                sb.AppendFormat("{0}" + SEP, item.ToString());
            return System.Text.RegularExpressions.Regex.Replace(sb.ToString(), SEP + "$", "");
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

        /// <summary>
        /// タブが止まるコントロールを設定する
        /// </summary>
        private void SetTabStop()
        {
            // まずはラジオボタンのインデックスの子と親の辞書を作成する
            // Key=タブインデックス
            // Value=Keyの親のタブインデックス
            //       (つまりラジオボタンを内包するコンテナコントロールのタブインデックス。
            //        コンテナに内包されていない場合は-1となる)
            // 例:
            //    Key Value
            //    ---------
            //    4   10
            //    3   10
            //    0   -1
            //    1   -1
            //    2    5
            var grpIndex = _controlModels.Where(x => !x.IsTabStop && 
                                                     !x.IsContainer && 
                                                     x.IsRadioButton)
                                        .ToDictionary(x => x.LastIndex, x => x.ParentLastIndex);

            // 辞書をKeyで昇順ソートする
            // 例:
            //    Key Value
            //    ---------
            //    0   -1
            //    1   -1
            //    2    5
            //    3   10
            //    4   10
            var grpIndexSortedKey = grpIndex.OrderBy(x => x.Key)
                                            .ToDictionary(x => x.Key, x => x.Value);

            // Valueの重複を削除する
            // 例:
            //    Key Value
            //    ---------
            //    0   -1
            //    2    5
            //    3   10
            var grpIndexDeletedValue = grpIndexSortedKey.GroupBy(x => x.Value)
                                                        .Select(x => x.First())
                                                        .ToDictionary(x => x.Key, x => x.Value);

            // タブストップを設定する
            _controlModels.Where(x => grpIndexDeletedValue.Any(
                                               kvp => kvp.Key == x.LastIndex && 
                                               kvp.Value == x.ParentLastIndex))
                          .ToList()
                          .ForEach(x => x.IsTabStop = true);

            _controlModels.Where(m => !m.IsContainer && !m.IsRadioButton)
                          .ToList()
                          .ForEach(m => m.IsTabStop = true);
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
            var model = _controlModels.OrderBy(x => x.LastIndex)
                                      .FirstOrDefault(x =>  x.LastIndex > lastIndex &&
                                                           !x.IsContainer &&
                                                            x.IsTabStop);
            if (model == null)
            {
                model = _controlModels.OrderBy(x => x.LastIndex)
                                      .FirstOrDefault(x => !x.IsContainer && 
                                                            x.IsTabStop);
            }
            return model.LastIndex;
        }

        private int FindNextLessNumber(int lastIndex)
        {
            var model = _controlModels.OrderByDescending(x => x.LastIndex)
                                      .FirstOrDefault(x =>  x.LastIndex < lastIndex &&
                                                           !x.IsContainer &&
                                                            x.IsTabStop);
            if (model == null)
            {
                model = _controlModels.OrderByDescending(x => x.LastIndex)
                                      .FirstOrDefault(x => !x.IsContainer &&
                                                            x.IsTabStop);
            }
            return model.LastIndex;
        }

        private ControlModel FindControlByLastIndex(int lastIndex)
        {
            return _controlModels.First(m => m.LastIndex == lastIndex);
        }

        private sealed class ControlModel
        {
            private System.Windows.Forms.Control _control;
            private string _indexString;
            private int _parentLastIndex;
            private int _lastIndex;
            private bool _isContainer;
            private bool _isRadioButton;
            private bool _isTabStop;

            private ControlModel()
            {
                // do nothing
            }

            public ControlModel(System.Windows.Forms.Control control, string indexString, bool isContainer = false, bool isRadioButton = false)
            {
                _control = control;
                _indexString = indexString;
                _parentLastIndex = GetPreviousNumber(_indexString);
                _lastIndex = GetLastNumber(_indexString);
                _isContainer = isContainer;
                _isRadioButton = isRadioButton;
                _isTabStop = false;
            }

            public System.Windows.Forms.Control Control { get { return _control; } }
            public string IndexString { get { return _indexString; } }
            public int ParentLastIndex { get { return _parentLastIndex; } }
            public int LastIndex { get { return _lastIndex; } }
            public bool IsContainer { get { return _isContainer; } }
            public bool IsRadioButton { get { return _isRadioButton; } }
            public bool IsTabStop { get { return _isTabStop; } set { _isTabStop = value; } }

            public override string ToString()
            {
                return $"Name={_control.Name}\t" +
                       $"TabIndex={_control.TabIndex}\t" +
                       $"IndexString={_indexString}\t" +
                       $"ParentLastIndex={_parentLastIndex}\t" +
                       $"LastIndex={_lastIndex}\t" +
                       $"IsContainer={_isContainer}\t" +
                       $"IsRadioButton={_isRadioButton}\t" +
                       $"IsTabStop={_isTabStop}";
            }

            private int GetPreviousNumber(string indexString)
            {
                var numbers = indexString.Split(SEP);
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
                var parts = indexString.Split(SEP);
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
                // do nothing
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
                // do nothing
            }
        }
    }
}
