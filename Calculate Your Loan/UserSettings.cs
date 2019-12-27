using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Calculate_Your_Loan.Properties;
namespace Calculate_Your_Loan
{
    public partial class UserSettings : Form
    {
        Settings st = new Settings();

        public UserSettings()
        {
            InitializeComponent();
            Guna.UI.Lib.GraphicsHelper.ShadowForm(this);
            
        }

        private void UserSettings_Load(object sender, EventArgs e)
        {
                if(st.IncludeComparison==true)
                {
                    ComparC.Checked = true;
                }
                if(st.IncludeImpact==true)
                {
                    impactC.Checked =true;
                }
              textBox1.Text=st.ChooseDirectory;
            if(st.IsusingDefault==true)
            {
                button3.Visible = false;
                textBox1.Visible = false;
                label2.Visible = false;
                SaveReports.Checked = true;
            }
            else if(st.IsusingDefault==false)
            {
                button3.Visible = true;
                textBox1.Visible = true;
                label2.Visible = true;
                SaveReports.Checked = false;

            }
            if(st.ShowGridLines==true)
            {
                checkBox1.Checked = true;

            }
            else
            {
                checkBox1.Checked = false;
            }
        }

       
        private void gunaPictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            folder.Description = "Choosing location";
            if(folder.ShowDialog()==DialogResult.OK)
            {
                textBox1.Text = folder.SelectedPath;
            }
        }

        private void SaveReports_CheckedChanged(object sender, EventArgs e)
        {
            if (SaveReports.Checked == true)
            {
                button3.Visible = false;
                textBox1.Visible = false;
                label2.Visible =false;
            }
            else if(SaveReports.Checked==false)
            {
                button3.Visible = true;
                textBox1.Visible = true;
                label2.Visible = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (SaveReports.Checked == true)
            {

                st.ChooseDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Calculate Your Loan\\Loan Reports";
                st.IsusingDefault = true;


            }
            else if (SaveReports.Checked == false)
            {

                st.ChooseDirectory = textBox1.Text;
                st.IsusingDefault = false;
            }

            st.Save();
            MessageBox.Show("Settings successfully saved\nRESTART APPLICATION TO APPLY CHANGES", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void impactC_CheckedChanged(object sender, EventArgs e)
        {
            if(impactC.Checked==true)
            {
                st.IncludeImpact =true;

            }
            else if(impactC.Checked==false)
            {
                st.IncludeImpact = false;
            }
        }

        private void ComparC_CheckedChanged(object sender, EventArgs e)
        {
            if(ComparC.Checked==true)
            {
                st.IncludeComparison = true;
            }
            else if(ComparC.Checked==false)
            {
                st.IncludeComparison = false;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(st.ChooseDirectory);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked==true)
            {
                st.ShowGridLines = true;
            }
            else if(checkBox1.Checked==false)
            {
                st.ShowGridLines = false;
            }
        }
    }
}
