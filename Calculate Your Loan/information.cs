using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace Calculate_Your_Loan
{
    public partial class information : Form
    {
        public information()
        {
            InitializeComponent();
            Guna.UI.Lib.GraphicsHelper.ShadowForm(this);
        }

        private void gunaPictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void gunaPictureBox2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label9_Click(object sender, EventArgs e)
        {
            try
            {

                string file = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "//Calculate Your Loan";
                if (!Directory.Exists(file))
                {
                    Directory.CreateDirectory(file);
                }
                string mainfile = file + "//MoreAbout.pdf";
                Process.Start(mainfile);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }

        }

        
    }
}
