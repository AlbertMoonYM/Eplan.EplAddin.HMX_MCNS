using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using System.Windows.Forms;

namespace Eplan.EplAddin.HMX_MCNS
{
    public class TransparentPanel : Panel
    {

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x00000020; // WS_EX_LAYERED
                return cp;
            }
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            // 배경을 그리지 않음
        }

        protected override void OnResize(EventArgs eventargs)
        {
            if (Parent == null)
                return;
            Rectangle rc = new Rectangle(Location, Size);
            Parent.Invalidate(rc, true);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.ResumeLayout(false);

        }
    }
}
