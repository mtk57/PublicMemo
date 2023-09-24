namespace TabOrderHelper
{
    /// <summary>
    /// タブオーダーモデル
    /// </summary>
    internal sealed class TabOrderModel : IHasHierarchicalTabIndices
    {
        private System.Collections.Generic.IEnumerable<int> _hierarchicalTabIndices;

        /// <summary>
        /// 前のコントロールモデル
        /// </summary>
        public TabOrderModel PrevControl { get; set; }

        /// <summary>
        /// カレントコントロール
        /// </summary>
        public System.Windows.Forms.Control Control { get; }

        /// <summary>
        /// 次のコントロールモデル
        /// </summary>
        public TabOrderModel NextControl { get; set; }

        /// <summary>
        /// タブインデックス文字列
        /// 階層表記はで親子をデリミタで区切る
        /// </summary>
        public string IndexString { get; }

        /// <summary>
        /// 最後の階層の親のタブインデックス
        /// </summary>
        public int ParentLastIndex { get; }

        /// <summary>
        /// 最後の階層のタブインデックス
        /// 重複の可能性あり
        /// 重複している場合、Zオーダーで順序を決定する
        /// </summary>
        public int LastIndex { get; }

        /// <summary>
        /// 内部的にナンバリングした重複無しのタブインデックス
        /// </summary>
        public int? UniqueTabIndex { get; set; }
        
        /// <summary>
        /// コンテナ系コントロールか否か
        /// </summary>
        public bool IsContainer { get; }

        /// <summary>
        /// ラジオボタンコントロールか否か
        /// </summary>
        public bool IsRadioButton { get; }

        /// <summary>
        /// コントロールのウィンドウハンドル
        /// Zオーダー判定時に必要
        /// </summary>
        public System.IntPtr Handle { get { return Control.Handle; } }

        private TabOrderModel()
        {
            // do nothing
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="control">コントロール</param>
        public TabOrderModel(System.Windows.Forms.Control control)
        {
            // 階層タブインデックスを取得する
            _hierarchicalTabIndices = GetHierarchicalTabindices(control);

            PrevControl = null;
            Control = control;
            NextControl = null;
            IndexString = GetHierarchicalTabIndicesString(control);
            ParentLastIndex = GetPreviousNumber(IndexString);
            LastIndex = GetLastNumber(IndexString);
            UniqueTabIndex = null;
            IsRadioButton = control is System.Windows.Forms.RadioButton;
        }

        public override string ToString()
        {
            if (PrevControl == null)
            {
                return $"Name={Control.Name}\t" +
                       $"PrevUniqueTabIndex=\t" +
                       $"TabIndex={Control.TabIndex}\t" +
                       $"NextUniqueTabIndex=\t" +
                       $"IndexString={IndexString}\t" +
                       $"ParentLastIndex={ParentLastIndex}\t" +
                       $"LastIndex={LastIndex}\t" +
                       $"UniqueTabIndex={UniqueTabIndex}\t" +
                       $"IsContainer={IsContainer}\t" +
                       $"IsRadioButton={IsRadioButton}";
            }

            return $"Name={Control.Name}\t" +
                   $"PrevUniqueTabIndex={PrevControl.UniqueTabIndex}\t" +
                   $"TabIndex={Control.TabIndex}\t" +
                   $"NextUniqueTabIndex={NextControl.UniqueTabIndex}\t" +
                   $"IndexString={IndexString}\t" +
                   $"ParentLastIndex={ParentLastIndex}\t" +
                   $"LastIndex={LastIndex}\t" +
                   $"UniqueTabIndex={UniqueTabIndex}\t" +
                   $"IsContainer={IsContainer}\t" +
                   $"IsRadioButton={IsRadioButton}";
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

        /// <summary>
        /// 階層構造を持ったコントロールのタブインデックスシーケンスを返す
        /// </summary>
        /// <param name="control">コントロール</param>
        /// <returns>タブインデックスシーケンス</returns>
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

        /// <summary>
        /// 階層構造を持ったコントロールのタブインデックスを文字列で返す
        /// 例.
        ///   Form
        ///     GroupBox     0
        ///        Button1   1
        ///        TextBox1  2
        ///     Button2      3
        ///     
        ///    はそれぞれ以下が返る
        ///    Button1="0:1"
        ///    TextBox1="0:2"
        ///    Button2="3"
        /// </summary>
        /// <param name="control">コントロール</param>
        /// <returns>タブインデックス</returns>
        private string GetHierarchicalTabIndicesString(System.Windows.Forms.Control control)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var item in GetHierarchicalTabindices(control))
                sb.AppendFormat("{0}" + Common.SEP, item.ToString());
            return System.Text.RegularExpressions.Regex.Replace(sb.ToString(), Common.SEP + "$", "");
        }

        /// <summary>
        /// 対象コントロールが親コントロールか否か
        /// </summary>
        /// <param name="target">対象コントロール</param>
        /// <returns>True:親コントロール, False:親コントロールではない</returns>
        private bool IsParent(System.Windows.Forms.Control target)
        {
            if (target == null) return false;
            if (target is System.Windows.Forms.Form) return false;
            return true;
        }

        /// <summary>
        /// タブインデックス文字列の最後の階層の1つ上を返す
        /// 例1:"1:2:3"の場合2が返る
        /// 例2:"3"の場合-1が返る
        /// </summary>
        /// <param name="indexString">タブインデックス文字列</param>
        /// <returns>最後の階層の1つ上の値</returns>
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

        /// <summary>
        /// タブインデックス文字列の最後の階層を返す
        /// 例1:"1:2:3"の場合3が返る
        /// </summary>
        /// <param name="indexString">タブインデックス文字列</param>
        /// <returns>最後の階層の値</returns>
        private int GetLastNumber(string indexString)
        {
            var parts = indexString.Split(Common.SEP);
            var lastPart = parts[parts.Length - 1];
            return int.Parse(lastPart);
        }
    }
}
