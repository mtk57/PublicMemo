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
    /// [使用例]
    /// public partial class Form1 : Form
    /// {
    ///     private TabOrderHelper _helper = null;
    ///
    ///     private void Form1_Load(object sender, EventArgs e)
    ///     {
    ///         _helper = new TabOrderHelper(this);
    ///     }
    ///
    ///     protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    ///     {
    ///         var activeControl = this.ActiveControl;
    ///
    ///         if (keyData == Keys.Tab)
    ///         {
    ///             var nextControl = _helper.GetNextControl(activeControl, true);
    ///             nextControl.Focus();
    ///             return true;
    ///         }
    ///         else if (keyData == (Keys.Shift | Keys.Tab))
    ///         {
    ///             var prevControl = _helper.GetNextControl(activeControl, false);
    ///             prevControl.Focus();
    ///             return true;
    ///         }
    ///         return base.ProcessCmdKey(ref msg, keyData);
    ///     }
    /// }
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

            if (nextControl.Visible && nextControl.Enabled)
            {
                return nextControl;
            }

            // 非表示 or 非活性の場合フォーカスしないのでフォーカスできるコントロールを探す
            var nextName = nextControl.Name;

            foreach (var c in _modelList)
            {
                if (forward)
                {
                    if (_modelDict[nextName].NextControl.Control.Visible ||
                        _modelDict[nextName].NextControl.Control.Enabled)
                        return _modelDict[nextName].NextControl.Control;
                    nextName = _modelDict[nextName].NextControl.Control.Name;
                }
                else
                {
                    if (_modelDict[nextName].PrevControl.Control.Visible ||
                        _modelDict[nextName].PrevControl.Control.Enabled)
                        return _modelDict[nextName].PrevControl.Control;
                    nextName = _modelDict[nextName].PrevControl.Control.Name;
                }
            }

            // 全て非表示・非活性なのでアクティブコントロールを返す
            return control;
        }

        /// <summary>
        /// 内部情報を更新する
        /// </summary>
        /// <param name="form">フォーム</param>
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

        /// <summary>
        /// モデルリストを作成する
        /// </summary>
        /// <param name="rootControl">ルートコントロール</param>
        private void CreateModelList(System.Windows.Forms.Control rootControl)
        {
            // ルートコントロール配下の全コントロールを調べる
            foreach (var item in GetAllControls(rootControl))
            {
                // コンテナ系はフォーカスが当たらないので無視
                if (IsContainer(item))
                    continue;

                var model = new TabOrderModel(item);
                _modelList.Add(model);
            }

            // 内部的にナンバリングした重複無しのタブインデックス値を設定する
            UpdateUniqueTabIndex();

            // 前後のコントロールを設定する
            UpdatePrevNextControl();
        }

        /// <summary>
        /// ルートコントロール配下の全コントロールの一覧を返す
        /// </summary>
        /// <param name="rootControl">ルートコントロール</param>
        /// <returns>全コントロールの一覧</returns>
        private System.Collections.Generic.IEnumerable<System.Windows.Forms.Control> GetAllControls(System.Windows.Forms.Control rootControl)
        {
            foreach (System.Windows.Forms.Control c in rootControl.Controls)
            {
                yield return c;
                foreach (System.Windows.Forms.Control a in GetAllControls(c))
                    yield return a;
            }
        }

        /// <summary>
        /// 対象コントロールがコンテナ系かどうかを返す
        /// </summary>
        /// <param name="target">対象コントロール</param>
        /// <returns>True:コンテナ系, False:コンテナ系以外</returns>
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
        private void UpdateUniqueTabIndex()
        {
            _modelList.Sort(new SortHelperOfHierarchicalTabIndices(Sort.Asc));

            var index = 0;
            int? groupIndex = null;

            for (var i=0; i< _modelList.Count; i++)
            {
                var model = _modelList[i];

                if (!model.IsRadioButton && !model.IsUserControlChild)
                {
                    // ラジオボタン以外かつユーザーコントロールの子供以外は無条件に設定
                    model.UniqueTabIndex = index++;
                    continue;
                }

                // ラジオボタンの場合は同グループの最初のコントロールをタブオーダーの対象とする
                if (model.IsRadioButton && (groupIndex == null || groupIndex != model.ParentLastIndex))
                {
                    model.UniqueTabIndex = index++;
                    groupIndex = model.ParentLastIndex;
                }
            }
        }

        /// <summary>
        /// 前後のコントロールを設定する
        /// </summary>
        private void UpdatePrevNextControl()
        {
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
        }

        /// <summary>
        /// 次(or前)のユニークタブインデックスのコントロールを設定する
        /// </summary>
        /// <param name="model">モデル</param>
        /// <exception cref="ControlNotFoundException"></exception>
        private void UpdatePrevNextControlByModel(TabOrderModel model)
        {
            for (var i=0; i<2; i++) // 2はforward=True/Falseを表す
            {
                var forward = (i == 0) ? true : false;
                var targetIndex = forward ? model.UniqueTabIndex + 1 : model.UniqueTabIndex - 1;

                TabOrderModel updateModel = null;
                TabOrderModel foundModel = null;

                // モデルリストからターゲットと一致するインデックスを探す
                foundModel = _modelList.FirstOrDefault(x => x.UniqueTabIndex == targetIndex);

                if (foundModel == null)
                {
                    // ターゲットが見つからないのでリストの先頭or末尾からインデックスを取得する

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

        /// <summary>
        /// ラジオボタンの場合は、同じユニークタブインデックスのコントロールを設定する
        /// </summary>
        /// <param name="model">モデル</param>
        /// <exception cref="ControlNotFoundException"></exception>
        private void UpdatePrevNextControlByModelForRadioButton(TabOrderModel model)
        {
            // モデルリストから条件に合致するモデルを探す
            var enableRadioButton = _modelList.FirstOrDefault(x => x.UniqueTabIndex >= 0 &&
                                                                   x.ParentLastIndex == model.ParentLastIndex && 
                                                                   x.IsRadioButton);
            if (enableRadioButton == null)
                throw new ControlNotFoundException($"Next or Preview Control not found. Info=[{model}]");

            model.NextControl = enableRadioButton.NextControl;
            model.PrevControl = enableRadioButton.PrevControl;
        }

        /// <summary>
        /// モデル辞書を作成する
        /// </summary>
        private void CreateModelDict()
        {
            _modelDict = _modelList.Where(x => x.UniqueTabIndex >= 0)
                                   .ToDictionary(x => x.Control.Name, x => x);
        }
    }
}
