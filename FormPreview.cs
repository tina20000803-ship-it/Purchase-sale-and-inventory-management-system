using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 銷貨管理系統
{
    public partial class FormPreview : Form
    {
        public FormPreview()
        {
            InitializeComponent();
        }
        public FormPreview(Image img)
        {
            InitializeComponent();
            pictureBox1.Image = img;
        }
    }
}
