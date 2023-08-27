using CustomTextBox;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace Test00ControlAttributeTest
{
    /*
    https://learn.microsoft.com/ja-jp/dotnet/desktop/winforms/controls/attributes-in-windows-forms-controls?view=netframeworkdesktop-4.8
    コントロールおよびコンポーネントのプロパティの属性
    -----
    属性	                        説明
    ------------------------------------------------------------------------------------------------------------------------------------
    AmbientValueAttribute	                        親要素から継承された値を持つことができる
                                                    プロパティに渡す値を指定し、そのプロパティが別のソースから値を取得するようにします。 これは "アンビエンス" と呼ばれています。
    BrowsableAttribute	                            プロパティまたはイベントが [プロパティ] ウィンドウに表示されるかどうかを指定します。
    CategoryAttribute	                            PropertyGrid コントロールが Categorized モードに設定されているときに、コントロールに表示するプロパティまたはイベントを分類するカテゴリの名前を指定します。
    DefaultValueAttribute	                        プロパティの既定値を指定します。
    DescriptionAttribute	                        プロパティまたはイベントの説明文を指定します。
    DesignerSerializationVisibilityAttribute        デザイン時にコンポーネントのプロパティをシリアル化するときに使用する永続化の種類を指定します。
    DesignOnlyAttribute                             プロパティを設定できるのがデザイン時だけかどうかを指定します。
    DisplayNameAttribute	                        引数を受け取らないプロパティ、イベント、または public void メソッドの表示名を指定します。
    EditorAttribute	                                プロパティの変更に使用するエディターを指定します。
    EditorBrowsableAttribute	                    プロパティまたはメソッドをエディターで表示できるかどうかを指定します。
    HelpKeywordAttribute	                        クラスまたはメンバーのコンテキスト キーワードを指定します。
    LocalizableAttribute	                        プロパティをローカライズする必要があるかどうかを指定します。
    PasswordPropertyTextAttribute	                アスタリスクなどの文字で、オブジェクトのテキスト表記を隠すように指示します。
    ReadOnlyAttribute	                            デザイン時に、この属性がバインドされるプロパティが読み取り専用か読み取り/書き込み可能かを指定します。
    RefreshPropertiesAttribute	                    関連付けられているプロパティ値が変更されたときに、プロパティ グリッドが更新されるように指定します。
    TypeConverterAttribute	                        この属性が関連付けられているオブジェクトのコンバーターとして使用する型を指定します。
    */

    /*
     <参考> ChatGPT3.5の回答
 
    コントロールを継承したクラス（カスタムコントロール）とユーザーコントロールでは、BrowsableAttributeの動作に違いがあります。

    この動作の違いは、コントロールのデザイン時の振る舞いに関連しています。以下に詳細を説明します。
    　　カスタムコントロールの場合：
　　　　　カスタムコントロールは、通常、独自のプロパティやイベントを持つことがあります。
    　　　カスタムコントロールは、デザイン時にプロパティウィンドウに表示されることが期待されています。
    　　　そのため、カスタムコントロールでは、BrowsableAttributeを明示的に指定しなくても、
    　　　独自のプロパティがプロパティウィンドウに表示されるようになっています。
　　　　ユーザーコントロールの場合：
　　　　　ユーザーコントロールは、他のコントロールを組み合わせて作成されるものであり、通常は独自のプロパティを持ちません。
    　　　ユーザーコントロールでは、デザイン時にプロパティウィンドウに表示される必要がない場合があります。
    　　　そのため、ユーザーコントロールでは、BrowsableAttributeを明示的に指定しない限り、
    　　　独自のプロパティはデフォルトでプロパティウィンドウに表示されません。
    */

    /// <summary>
    /// コントロールの属性テスト
    /// 
    /// プロパティウィンドウに反映されない場合はVisualStudioをいったん終了させてみるとよい。
    /// 
    /// <参考>
    /// http://dobon.net/vb/dotnet/control/propertygrid.html
    /// https://learn.microsoft.com/ja-jp/dotnet/desktop/winforms/controls/attributes-in-windows-forms-controls?view=netframeworkdesktop-4.8
    /// </summary>
    public partial class Test00TextBoxEx : Test00CustomTextBox
    {
        private int _testDefaultValue1;
        private Color _testDefaultValue2;
        private string _testDescription;
        private Color _testCategory;
        private bool _testBrowsable;
        private int _testReadOnly;
        private bool _testDesignerSerializationVisibility_Content;
        private bool _testDesignerSerializationVisibility_Hidden;
        private bool _testDesignerSerializationVisibility_Visible;
        private Color _testAmbientValue = Color.White;

        public Test00TextBoxEx()
        {
            InitializeComponent();
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
        }

        /// <summary>
        /// プロパティがデフォルト値（規定値）でないときに、値が太字で表示されます。
        /// プロパティがデフォルト値でないときだけ値が太字で表示するには、
        /// DefaultValueAttributeを使用して、プロパティのデフォルト値を決めておきます。
        /// </summary>
        [DefaultValue(0)]
        public int TestDefaultValue1
        {
            get { return _testDefaultValue1; }
            set { _testDefaultValue1 = value; }
        }

        /// <summary>
        /// 一つの方法としては、TypeConverterによって指定した型に変換できる文字列を
        /// DefaultValueAttributeコンストラクタに指定する方法があります。
        /// この文字列には、例えばColor型の赤であれば"Red"、Size型の(10, 20) であれば"10, 20"
        /// といった文字列が使えます。
        /// </summary>
        [DefaultValue(typeof(System.Drawing.Color), "Red")]
        public System.Drawing.Color TestDefaultValue2
        {
            get { return _testDefaultValue2; }
            set { _testDefaultValue2 = value; }
        }

        /// <summary>
        /// PropertyGridコントロールの説明ペインに、選択されているプロパティの説明を表示するには、
        /// DescriptionAttributeを使用します。
        /// </summary>
        [Description("ここにStringValueの説明を書きます。")]
        public string TestDescription
        {
            get { return _testDescription; }
            set { _testDescription = value; }
        }

        /// <summary>
        /// PropertyGridコントロールではプロパティを項目（カテゴリ)別に表示できます。
        /// 項目別に表示したとき、プロパティはデフォルトで「その他」に分類されますが、
        /// CategoryAttributeにより、プロパティの項目を指定することができます。
        /// </summary>
        [Category("表示")]
        public System.Drawing.Color TestCategory
        {
            get { return _testCategory; }
            set { _testCategory = value; }
        }

        /// <summary>
        /// PropertyGridコントロールに表示したくないプロパティには、
        /// Falseを指定したBrowsableAttributeを適用します。
        /// →デフォルト:true
        /// </summary>
        [Browsable(false)]
        public bool TestBrowsable
        {
            get { return _testBrowsable; }
            set { _testBrowsable = value; }
        }

        /// <summary>
        /// プロパティの値をユーザーが編集できないようにするには、
        /// Trueを指定したReadOnlyAttributeを使用します。
        /// /// →デフォルト:false
        /// </summary>
        [ReadOnly(true)]
        public int TestReadOnly
        {
            get { return _testReadOnly; }
            set { _testReadOnly = value; }
        }

        /// <summary>
        /// Designer.csに書き込まれない。
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public bool TestDesignerSerializationVisibility_Content
        {
            get { return _testDesignerSerializationVisibility_Content; }
            set { _testDesignerSerializationVisibility_Content = value; }
        }

        /// <summary>
        /// Designer.csに書き込まれない。
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool TestDesignerSerializationVisibility_Hidden
        {
            get { return _testDesignerSerializationVisibility_Hidden; }
            set { _testDesignerSerializationVisibility_Hidden = value; }
        }

        /// <summary>
        /// Designer.csに書き込まれる
        /// デフォルト:Visible
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public bool TestDesignerSerializationVisibility_Visible
        {
            get { return _testDesignerSerializationVisibility_Visible; }
            set { _testDesignerSerializationVisibility_Visible = value; }
        }

        /// <summary>
        /// プロパティまたはパラメータが親要素から継承された値を持つことができます。
        /// 親要素が値を提供しない場合、指定した既定値が使用されます。
        /// 
        /// <例>
        ///   [AmbientValue("Default")]
        ///   public string MyProperty { get; set; }
        /// 
        /// 上記の例では、MyPropertyプロパティにAmbientValueAttributeが適用されています。
        /// この属性により、MyPropertyが親要素から値を継承することができます。
        /// もし親要素が値を提供しない場合、"Default"という既定値が使用されます。
        /// 
        /// AmbientValueAttributeとDefaultValueAttributeを同時に指定した場合、次のような動作になります。
        ///   ・親要素から値が継承される場合：AmbientValueAttributeに指定された値が使用されます。
        ///   ・親要素から値が継承されない場合：DefaultValueAttributeに指定された値が使用されます。
        /// つまり、AmbientValueAttributeが優先されます。
        /// 親要素から値が継承される場合でも、AmbientValueAttributeに指定された値が使用されます。
        /// ただし、親要素から値が継承されない場合には、DefaultValueAttributeに指定された値が使用されます。
        /// 
        /// <参考>
        /// https://learn.microsoft.com/ja-jp/dotnet/desktop/winforms/controls/how-to-apply-attributes-in-windows-forms-controls?view=netframeworkdesktop-4.8
        /// </summary>
        [AmbientValue(typeof(Color), "Empty")]
        [Category("Appearance")]
        [DefaultValue(typeof(Color), "White")]
        [Description("The color used for painting alert text.")]
        public Color TestAmbientValue
        {
            get
            {
                if (this._testAmbientValue == Color.Empty &&
                    this.Parent != null)
                {
                    return this.Parent.ForeColor;
                }

                return this._testAmbientValue;
            }

            set
            {
                this._testAmbientValue = value;
            }
        }

        // This method is used by designers to enable resetting the
        // property to its default value.
        public void ResetTestAmbientValue()
        {
            this.TestAmbientValue = Color.White;
        }

        // This method indicates to designers whether the property
        // value is different from the ambient value, in which case
        // the designer should persist the value.
        private bool ShouldSerializeTestAmbientValue()
        {
            return (this._testAmbientValue != Color.Empty);
        }
    }
}
