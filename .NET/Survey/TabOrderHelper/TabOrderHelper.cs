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
        private System.Collections.Generic.Dictionary<string, TabOrderModel> _modelDict;

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
            Update(form);
        }

        /// <summary>
        /// カレントコントロールの次(もしくは前)のコントロールを返す
        /// </summary>
        /// <param name="control">カレントコントロール</param>
        /// <param name="forward">True:次のコントロール、False:前のコントロール</param>
        /// <returns>コントロール</returns>
        public System.Windows.Forms.Control GetNextControl(System.Windows.Forms.Control control, bool forward = true)
        {
            var name = control.Name;
            var nextControl = forward ? _modelDict[name].NextControl.Control : _modelDict[name].PrevControl.Control;

            if (nextControl.Visible)
            {
                return nextControl;
            }

            // 非表示の場合フォーカスしないので表示されているコントロールを探す
            var nextName = nextControl.Name;

            foreach (var c in _modelList)
            {
                if (forward)
                {
                    if (_modelDict[nextName].NextControl.Control.Visible)
                        return _modelDict[nextName].NextControl.Control;
                    nextName = _modelDict[nextName].NextControl.Control.Name;
                }
                else
                {
                    if (_modelDict[nextName].PrevControl.Control.Visible)
                        return _modelDict[nextName].PrevControl.Control;
                    nextName = _modelDict[nextName].PrevControl.Control.Name;
                }
            }

            // 全て非表示なのでアクティブコントロールを返す
            return control;
        }

        public void Update(System.Windows.Forms.Control form)
        {
            _modelList = new System.Collections.Generic.List<TabOrderModel>();
            _modelDict = new System.Collections.Generic.Dictionary<string, TabOrderModel>();

            CreateModelList(form);
            CreateModelDict();

#if DEBUG
            foreach (var c in _modelList)
                System.Diagnostics.Debug.WriteLine(c.ToString());
#endif
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

            UpdatePrevNextControl();
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

        private void UpdatePrevNextControl()
        {
            //foreach (var c in _modelList)
            //    System.Diagnostics.Debug.WriteLine("BEFORE\t" + c.ToString());

            foreach (var model in _modelList)
            {
                if (model.UniqueTabIndex >= 0)
                {
                    // ユニークタブインデックスが設定済の場合は、シンプルに次(or前)のユニークタブインデックスのコントロールを設定する
                    UpdatePrevNextControlByModel(model);
                }
            }

            foreach (var model in _modelList)
            {
                if (model.UniqueTabIndex == null && model.IsRadioButton)
                {
                    // ユニークタブインデックスが未設定のラジオボタンの場合は、同じユニークタブインデックスのコントロールを設定する
                    UpdatePrevNextControlByModelForRadioButton(model);
                }
            }

            //foreach (var c in _modelList)
            //    System.Diagnostics.Debug.WriteLine("AFTER\t" + c.ToString());

        }

        private void UpdatePrevNextControlByModel(TabOrderModel model)
        {
            for (var i=0; i<2; i++) // 2はforward=True/Falseを表す
            {
                var forward = (i == 0) ? true : false;
                var targetIndex = forward ? model.UniqueTabIndex + 1 : model.UniqueTabIndex - 1;

                TabOrderModel updateModel = null;
                TabOrderModel foundModel = null;

                foundModel = _modelList.FirstOrDefault(x => x.UniqueTabIndex == targetIndex);

                if (foundModel == null)
                {
                    // ターゲットが見つからないのでリストの先頭or末尾から有効な値を取得する

                    if (forward)
                        // Nextの場合は昇順ソートして先頭から検索
                        foundModel = _modelList.OrderBy(x => x.UniqueTabIndex)
                                               .FirstOrDefault(x => x.UniqueTabIndex >= 0);
                    else
                        // Prevの場合は降順ソートして先頭から検索
                        foundModel = _modelList.OrderByDescending(x => x.UniqueTabIndex)
                                               .FirstOrDefault(x => x.UniqueTabIndex >= 0);
   
                    if (foundModel == null)
                        // 有効な値が見つからない
                        throw new ControlNotFoundException($"Next or Preview Control not found. Info=[{model}]");
                }

                updateModel = new TabOrderModel(foundModel.Control);
                updateModel.UniqueTabIndex = foundModel.UniqueTabIndex;

                if (forward)
                    model.NextControl = updateModel;
                else
                    model.PrevControl = updateModel;
            }
        }

        private void UpdatePrevNextControlByModelForRadioButton(TabOrderModel model)
        {
            var enableRadioButton = _modelList.FirstOrDefault(x => x.UniqueTabIndex >= 0 &&
                                                                   x.ParentLastIndex == model.ParentLastIndex && 
                                                                   x.IsRadioButton);
            if (enableRadioButton == null)
                throw new ControlNotFoundException($"Next or Preview Control not found. Info=[{model}]");

            model.NextControl = enableRadioButton.NextControl;
            model.PrevControl = enableRadioButton.PrevControl;
        }

        private void CreateModelDict()
        {
            _modelDict = _modelList.ToDictionary(x => x.Control.Name, x => x);
        }
    }
}
