using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    /// <summary>
    /// コントロールが非活性でもツールチップを表示するテスト
    /// 
    /// <参考>
    /// https://stackoverflow.com/questions/1732140/displaying-tooltip-over-a-disabled-control
    /// </summary>
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            var td = new ToolTipOnDisabledControl();
            //td.SetToolTip(this.button2, "2" );
            td.SetToolTip(this.button2, this.toolTip1, "2");    //test
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            button1.Enabled = !button1.Enabled;

            if (!button1.Enabled)
            {
                var td = new ToolTipOnDisabledControl();
                //td.SetToolTip(this.button1, "1");
                td.SetToolTip(this.button1, this.toolTip1, "1");    //test
            }
            else
            {
                this.toolTip1.SetToolTip(this.button1, "Hello button1");//test
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //this.button1.Location = new System.Drawing.Point(143, 102);
            //this.button1.Size = new System.Drawing.Size(270, 91);

            //var rect = new Rectangle();
            //rect.X = button1.Location.X;
            //rect.Y = button1.Location.Y;
            //rect.Width = button1.Size.Width;
            //rect.Height = button1.Size.Height;

            //label2.Text = rect.ToString();

            //label12.Text = checkBox1.Bounds.ToString();
        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            return;

            label6.Text = "";

            label4.Text = e.Location.ToString();    // クライアント座標

            //// マウスの現在の位置を取得
            //Point mousePosition = Cursor.Position;      // スクリーン座標
            //label8.Text = mousePosition.ToString();

            //// フォーム上の座標に変換
            //Point formPosition = this.PointToClient(mousePosition); // クライアント座標
            //label10.Text = formPosition.ToString();

            //// 座標にある子コントロールを取得
            //Control control = this.GetChildAtPoint(formPosition); // GetChildAtPointは完全一致したときでないと取得できない。

            foreach (Control childControl in this.Controls)
            {
                if (childControl.Enabled == false)
                {
                    if ((e.Location.X >= childControl.Location.X &&
                         e.Location.X <= (childControl.Location.X + childControl.Size.Width))
                          &&
                        (e.Location.Y >= childControl.Location.Y &&
                         e.Location.Y <= (childControl.Location.Y + childControl.Size.Height)))
                    {
                        label6.Text = childControl.Name;
                        label6.Text = this.toolTip1.GetToolTip(childControl);
                        this.toolTip1.ShowAlways = true;
                        this.toolTip1.Show(label6.Text, childControl, childControl.Width / 2, childControl.Height / 2);
                        break;
                    }
                    //else
                    //{
                    //    label6.Text = "";
                    //    this.toolTip1.ShowAlways = false;
                    //    this.toolTip1.Hide(childControl);
                    //    break;
                    //}
                    //break;
                }
            }
        }

    }

    /// <summary>
    /// // Reference example
    ///  var td = new ToolTipOnDisabledControl();
    ///  this.checkEdit3.Enabled = false;
    ///  td.SetTooltip(this.checkEdit3, "tooltip for disabled");
    /// </summary>
    public class ToolTipOnDisabledControl
    {
        #region Fields and Properties

        private Control enabledParentControl;

        private bool isShown;

        public Control TargetControl { get; private set; }

        public string TooltipText { get; private set; }
        public ToolTip ToolTip { get; }
        #endregion

        #region Public Methods
        public ToolTipOnDisabledControl()
        {
            this.ToolTip = new ToolTip();
        }

        //public void SetToolTip(Control targetControl, string tooltipText = null)
        public void SetToolTip(Control targetControl, ToolTip tip,  string tooltipText = null)//test
        {
            this.TargetControl = targetControl;
            if (string.IsNullOrEmpty(tooltipText))
            {
                this.TooltipText = this.ToolTip.GetToolTip(targetControl);
            }
            else
            {
                this.TooltipText = tooltipText;

                tip.SetToolTip(targetControl, "");  // test
            }

            if (targetControl.Enabled)
            {
                this.enabledParentControl = null;
                this.isShown = false;
                this.ToolTip.SetToolTip(this.TargetControl, this.TooltipText);
                return;
            }

            this.enabledParentControl = targetControl.Parent;
            while (!this.enabledParentControl.Enabled && this.enabledParentControl.Parent != null)
            {
                this.enabledParentControl = this.enabledParentControl.Parent;
            }

            if (!this.enabledParentControl.Enabled)
            {
                throw new Exception("Failed to set tool tip because failed to find an enabled parent control.");
            }

            this.enabledParentControl.MouseMove += this.EnabledParentControl_MouseMove;
            this.TargetControl.EnabledChanged += this.TargetControl_EnabledChanged;
        }

        public void Reset()
        {
            if (this.TargetControl != null)
            {
                this.ToolTip.Hide(this.TargetControl);
                this.TargetControl.EnabledChanged -= this.TargetControl_EnabledChanged;
                this.TargetControl = null;
            }

            if (this.enabledParentControl != null)
            {
                this.enabledParentControl.MouseMove -= this.EnabledParentControl_MouseMove;
                this.enabledParentControl = null;
            }

            this.isShown = false;
        }
        #endregion

        #region Private Methods
        private void EnabledParentControl_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Location.X >= this.TargetControl.Left &&
                e.Location.X <= this.TargetControl.Right &&
                e.Location.Y >= this.TargetControl.Top &&
                e.Location.Y <= this.TargetControl.Bottom)
            {
                if (!this.isShown)
                {
                    this.ToolTip.Show(this.TooltipText, this.TargetControl, this.TargetControl.Width / 2, this.TargetControl.Height / 2, this.ToolTip.AutoPopDelay);
                    this.isShown = true;
                }
            }
            //else if(this.isShown)
            else
            {
                this.ToolTip.Hide(this.TargetControl);
                this.isShown = false;
            }
        }

        private void TargetControl_EnabledChanged(object sender, EventArgs e)
        {
            if (TargetControl.Enabled)
            {
                TargetControl.EnabledChanged -= TargetControl_EnabledChanged;
                enabledParentControl.MouseMove -= EnabledParentControl_MouseMove;
            }
        }
        #endregion
    }
}
