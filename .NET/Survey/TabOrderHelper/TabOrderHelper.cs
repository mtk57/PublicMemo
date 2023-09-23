//
// デザイナーの「表示/タブオーダー」のように、階層化されたタブインデックスをリストで管理する
//
// 参考:https://zecl.hatenablog.com/entry/20090226/p1
//      https://atmarkit.itmedia.co.jp/fdotnet/dotnettips/243winkeyproc/winkeyproc.html
//
using System.Linq;

namespace TabOrderHelper
{
    /// <summary>
    /// タブオーダーヘルパークラス
    /// 
    /// [本クラスを使用する場合の注意点]
    /// 1.コンテナ系コントロールは以下のみ対応する。
    ///   Panel
    ///   GroupBox
    ///
    /// </summary>
    public sealed class TabOrderHelper
    {
        private System.Collections.Generic.List<TabOrderModel> _modelList;
        private System.Collections.Generic.Dictionary<int, TabOrderModel> _modelDict;

        private TabOrderHelper()
        {
            // do nothing
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="form">フォーム</param>
        public TabOrderHelper(System.Windows.Forms.Control form)
        {
            _modelList = new System.Collections.Generic.List<TabOrderModel>();
            _modelDict = new System.Collections.Generic.Dictionary<int, TabOrderModel>();

            CreateModelList(form);
            CreateModelDict();

#if DEBUG
            foreach (var c in _modelList)
                System.Diagnostics.Debug.WriteLine(c.ToString());
#endif
        }

        /// <summary>
        /// カレントコントロールの次(もしくは前)のコントロールを返す
        /// </summary>
        /// <param name="control">カレントコントロール</param>
        /// <param name="forward">True:次のコントロール、False:前のコントロール</param>
        /// <returns>コントロール</returns>
        public System.Windows.Forms.Control GetNextControl(System.Windows.Forms.Control control, bool forward = true)
        {
            var tabIndex = control.TabIndex;
            return forward ? _modelDict[tabIndex].NextControl : _modelDict[tabIndex].PrevControl;
        }

        private void CreateModelList(System.Windows.Forms.Control rootControl)
        {
            foreach (var item in GetAllControls(rootControl))
            {
                if (IsContainer(item)) continue;

                var model = new TabOrderModel(item);
                _modelList.Add(model);
            }

            UpdateTabIndex();
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

        private bool IsContainer(System.Windows.Forms.Control target)
        {
            if (target is System.Windows.Forms.Panel ||
                target is System.Windows.Forms.GroupBox)
                return true;
            return false;
        }

        /// <summary>
        /// 内部的にナンバリングした重複無しのタブインデックス値を設定する
        /// </summary>
        private void UpdateTabIndex()
        {
            _modelList.Sort(new SortHelperOfHierarchicalTabIndices(Sort.Asc));

            var index = 0;
            int? groupIndex = null;

            for (var i=0; i< _modelList.Count; i++)
            {
                var model = _modelList[i];

                if (!model.IsRadioButton)
                {
                    model.UniqueTabIndex = index++;
                    continue;
                }

                // ラジオボタンの場合は同グループの最初のコントロールをタブオーダーの対象とする
                if (groupIndex == null || groupIndex != model.ParentLastIndex)
                {
                    model.UniqueTabIndex = index++;
                    groupIndex = model.ParentLastIndex;
                }
            }
        }

        private void SetPrevNextControl()
        {
            foreach (var x in _modelList)
            {
                x.NextControl = GetNextControl(x, true);
                x.PrevControl = GetNextControl(x, false);
            }
        }

        private System.Windows.Forms.Control GetNextControl(TabOrderModel model, bool forward = true)
        {
            var lastIndex = model.LastIndex;

            return forward ? GetNextGreaterTabIndexControl(lastIndex) : GetPrevLessTabIndexControl(lastIndex);
        }

        private System.Windows.Forms.Control GetNextGreaterTabIndexControl(int lastIndex)
        {
            var model = _modelList.OrderBy(x => x.LastIndex)
                                      .FirstOrDefault(x => x.LastIndex > lastIndex &&
                                                          !x.IsContainer &&
                                                           x.IsTabStop);
            if (model == null)
            {
                model = _modelList.OrderBy(x => x.LastIndex)
                                      .FirstOrDefault(x => !x.IsContainer &&
                                                            x.IsTabStop);
            }
            return model.Control;
        }

        private System.Windows.Forms.Control GetPrevLessTabIndexControl(int lastIndex)
        {
            var model = _modelList.OrderByDescending(x => x.LastIndex)
                                      .FirstOrDefault(x => x.LastIndex < lastIndex &&
                                                          !x.IsContainer &&
                                                           x.IsTabStop);
            if (model == null)
            {
                model = _modelList.OrderByDescending(x => x.LastIndex)
                                      .FirstOrDefault(x => !x.IsContainer &&
                                                            x.IsTabStop);
            }
            return model.Control;
        }

        private void CreateModelDict()
        {
            _modelDict = _modelList.OrderBy(x => x.LastIndex)
                                       .ToDictionary(x => x.LastIndex, x => x);
        }
    }
}
