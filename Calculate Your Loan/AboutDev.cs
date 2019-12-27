using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculate_Your_Loan
{
    public partial class AboutDev : Form
    {
        public AboutDev()
        {
            InitializeComponent();
            Guna.UI.Lib.GraphicsHelper.ShadowForm(this);
            System.Timers.Timer t = new System.Timers.Timer();
            t.Interval = 1;
            t.Elapsed += t_Elapsed;
            t.Start();
            
            

            
            
           
        }

        void t_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            label2.Top -= 1;
            if (label2.Top < 0 - label2.Height)
            {
                label2.Top = panel2.Height;
            }
        }

        private void gunaLinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("http://www.itworldinnovate.com");

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

       public void gunaPictureBox2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

      
    }
}
