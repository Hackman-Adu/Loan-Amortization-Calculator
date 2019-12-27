using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Guna.UI.WinForms;
using System.Text.RegularExpressions;
using System.Threading;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using Calculate_Your_Loan.Properties;

namespace Calculate_Your_Loan
{
    public partial class Form1 : Form
    {
        Settings st = new Settings();
        public double principal, rate, time, regularpayment, bebalance, endbalance, schedulepayment, interest, principalpaid;
        public double extrapayment, totalpayment;
        public double paymentsNumber, actualtotalinterest;
        double annualInterest, semiannualInterest, monthlyinterest, weeklyinterest, quartelyInterest, dailyinterest, biweeklyinterest, bimonthlyinterest, seminmonthlyinterest;
        double annualpay, semiannualpay, monthlypay, weeklypay, quartelypay, dailypay, biweeklypay, bimonthlypay, seminmonthlypay;
        double AnnualRepay, semiannualrepay, monthlyrepay, weeklyrepay, quaterlyrepay, dailyrepay, biweekrepay, bimonthlyrepay, semimonthlyrepay;
        DataTable dt;
        SaveFileDialog save;


        public Form1()
        {
            InitializeComponent();
            colors();
            this.Click += Form1_Click;
            gridviewcolumns(gunaDataGridView1);
            gunaComboBox1.SelectedIndex = 0;
            viewsummary();
            //Guna.UI.Lib.GraphicsHelper.ShadowForm(this);

        }

        public void Form1_Click(object sender, EventArgs e)
        {
            closingAnyOpened();

        }
        public void closingAnyOpened()
        {
            Form abt = Application.OpenForms["AboutDev"];
            if (abt != null)
            {
                abt.Close();
            }
        }



        void colors()
        {
            panel1.BackColor = ColorTranslator.FromHtml("#01466e");
            gunaPictureBox1.BackColor = ColorTranslator.FromHtml("#01466e");
            gunaPictureBox2.BackColor = ColorTranslator.FromHtml("#01466e");
            foreach (Control c in this.Controls)
            {
                if (c is Button)
                {
                    Button b = c as Button;
                    b.BackColor = Color.FromArgb(255, 128, 0);
                }
            }
            foreach (Control c in groupBox1.Controls)
            {
                if (c is GunaTextBox)
                {
                    GunaTextBox t = c as GunaTextBox;
                    t.FocusedBorderColor = ColorTranslator.FromHtml("#01466e");
                }
            }
        }
        void viewsummary()
        {
            foreach (Control c in groupBox2.Controls)
            {
                if (c is GunaTextBox)
                {
                    GunaTextBox t = c as GunaTextBox;
                    t.Visible = true;
                }
            }
        }
        void gridviewcolumns(DataGridView dv)
        {
            dt = new DataTable();
            string[] cols = new string[] { "Payment Date", "Balance(Start)", "Payment","Extra Payment","Total Payment", "Interest", "Principal", "Balance(End)" };
            foreach (string f in cols)
            {
                dt.Columns.Add(f);
            }
            dv.DataSource = dt;

            string[] comboitems = new string[] { "-----select your payment frequency-----", "Annual Payment", "Semi Annual Payment", "Monthly Payment", "Weekly Payment", "Quarterly Payment", "Daily Payment","Bi-weekly Payment","Bi-monthly Payment","Semi-monthly Payment" };
            foreach (string ff in comboitems)
            {
                gunaComboBox1.Items.Add(ff);
            }
        }

        



        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime payD = dateTimePicker1.Value.Date;
            paydate.Text = payD.ToShortDateString();
        }

        private void gunaComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy == true)
            {
                MessageBox.Show("CSV file exporting is still in progress/nCannot make changes to input", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (gunaComboBox1.SelectedIndex != 0)
            {
                dt.Columns[2].ColumnName = gunaComboBox1.Text;
                rateLabel.Text = gunaComboBox1.Text.Replace("Payment", "Rate");
            }
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(backgroundWorker1.IsBusy==true)
            {
                MessageBox.Show("CSV file exporting is still in progress", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            
            if(gunaComboBox1.SelectedIndex==0)
            {
                MessageBox.Show("Select payment frequency", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            foreach (Control c in groupBox1.Controls)
            {
                if (c is GunaTextBox)
                {
                    GunaTextBox t = c as GunaTextBox;
                   if(t.Name!="ExtraP")
                   {
                       if (string.IsNullOrWhiteSpace(t.Text))
                       {
                           MessageBox.Show("All fields are required", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                           return;
                       }
                   }
                }
            }
            if (Ltime.Text.Contains("."))
            {
                MessageBox.Show("Please enter integer for the loan term", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;

            }
            if(string.IsNullOrWhiteSpace(Lamt.Text))
            {
                MessageBox.Show("Please enter loan amount", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            try
            {

                   calculatingPayment();
                
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
           


        }
        /// <summary>
        /// This method is used for calculating the scheduled payment
        /// </summary>
        void calculatingPayment()
        {
            principal = double.Parse(Regex.Replace(Lamt.Text, "[A-Za-z $]", ""));

            if (gunaComboBox1.SelectedIndex == 1)
            {
                time = double.Parse(Ltime.Text);
                rate = double.Parse(Lrate.Text) / 100;
                paymentsNumber = 1;
            }
            else if (gunaComboBox1.SelectedIndex == 2)
            {
                rate = double.Parse(Lrate.Text) / 100 / 2;
                time = double.Parse(Ltime.Text) * 2;
                paymentsNumber = 2;
            }
            else if (gunaComboBox1.SelectedIndex == 3)
            {
                rate = double.Parse(Lrate.Text) / 100 / 12;
                time = double.Parse(Ltime.Text) * 12;
                paymentsNumber = 12;
            }
            else if (gunaComboBox1.SelectedIndex == 4)
            {
                rate = double.Parse(Lrate.Text) / 100 / 52;
                time = double.Parse(Ltime.Text) * 52;
                paymentsNumber = 52;
            }
            else if (gunaComboBox1.SelectedIndex == 5)
            {
                rate = double.Parse(Lrate.Text) / 100 / 4;
                time = double.Parse(Ltime.Text) * 4;
                paymentsNumber = 4;
            }
            else if (gunaComboBox1.SelectedIndex == 6)
            {
                rate = double.Parse(Lrate.Text) / 100 / 365;
                time = double.Parse(Ltime.Text) * 365;
                paymentsNumber = 365;
            }
            else if (gunaComboBox1.SelectedIndex == 7)
            {
                rate = double.Parse(Lrate.Text) / 100 /26;
                time = double.Parse(Ltime.Text) * 26;
                paymentsNumber = 26;
            }
            else if (gunaComboBox1.SelectedIndex == 8)
            {
                rate = double.Parse(Lrate.Text) / 100 / 6;
                time = double.Parse(Ltime.Text) * 6;
                paymentsNumber = 6;
            }
            else if (gunaComboBox1.SelectedIndex == 9)
            {
                rate = double.Parse(Lrate.Text) / 100 / 24;
                time = double.Parse(Ltime.Text) * 24;
                paymentsNumber = 24;
            }
            double a = Math.Pow((1 + rate), -time);
            double aa = (1 - a) / rate;
            double ans = principal / aa;
            regularpayment = ans;
            RegP.Text = ans.ToString("c");
            double actualtime=double.Parse(Ltime.Text);
            //Variable Declaration for the payment options
            double annTime = double.Parse(Ltime.Text);
            double annrate  = double.Parse(Lrate.Text) / 100;
            double semiAtime = double.Parse(Ltime.Text)*2;
            double semiArate = double.Parse(Lrate.Text) / 100/2;
            double montlyTime = double.Parse(Ltime.Text) * 12;
            double monthlyrate = double.Parse(Lrate.Text) / 100 / 12;
            double weeklytime = double.Parse(Ltime.Text) * 52;
            double weeklyrate = double.Parse(Lrate.Text) / 100 / 52;
            double quartertime = double.Parse(Ltime.Text) * 4;
            double quarterate= double.Parse(Lrate.Text) / 100 / 4;
            double dailytime = double.Parse(Ltime.Text) * 365;
            double dailyrate = double.Parse(Lrate.Text) / 100 / 365;
            double biweeklytime = double.Parse(Ltime.Text) * 26;
            double biweeklyrate = double.Parse(Lrate.Text) / 100 / 26;
            double bimonthlytime = double.Parse(Ltime.Text) * 6;
            double bimonthlyrate = double.Parse(Lrate.Text) / 100 / 6;
            double semimontlytime = double.Parse(Ltime.Text) * 24;
            double seminmonthlyrate = double.Parse(Lrate.Text) / 100 / 24;
            actualtotalinterest = (paymentsNumber * regularpayment * actualtime) - principal;
            //Total interest for various options
            annualInterest = (1 * specificPay(annTime,annrate) * actualtime) - principal;
            semiannualInterest = (2 * specificPay(semiAtime, semiArate) * actualtime) - principal;
            monthlyinterest = (12 * specificPay(montlyTime, monthlyrate) * actualtime) - principal;
            weeklyinterest = (52 * specificPay(weeklytime,weeklyrate) * actualtime) - principal;
            quartelyInterest = (4 * specificPay(quartertime,quarterate) * actualtime) - principal;
            dailyinterest = (365 * specificPay(dailytime,dailyrate) * actualtime) - principal;
            biweeklyinterest = (26 * specificPay(biweeklytime,biweeklyrate) * actualtime) - principal;
            bimonthlyinterest = (6 * specificPay(bimonthlytime,bimonthlyrate) * actualtime) - principal;
            seminmonthlyinterest = (24 * specificPay(semimontlytime, seminmonthlyrate) * actualtime) - principal;
            //Regular payments for various options
            annualpay = specificPay(annTime, annrate);
            semiannualpay = specificPay(semiAtime, semiArate);
            monthlypay = specificPay(montlyTime, monthlyrate);
            weeklypay = specificPay(weeklytime, weeklyrate);
            quartelypay = specificPay(quartertime, quarterate);
            dailypay = specificPay(dailytime, dailyrate);
            biweeklypay = specificPay(biweeklytime, biweeklyrate);
            bimonthlypay = specificPay(bimonthlytime, bimonthlyrate);
            seminmonthlypay = specificPay(semimontlytime, seminmonthlyrate);
            
            //Total repay for various payment options

            AnnualRepay = annualInterest + principal;
            semiannualrepay = semiannualInterest + principal;
            monthlyrepay = monthlyinterest + principal;
            weeklyrepay = weeklyinterest + principal;
            quaterlyrepay = quartelyInterest + principal;
            dailyrepay = dailyinterest + principal;
            biweekrepay = biweeklyinterest + principal;
            bimonthlyrepay = bimonthlyinterest + principal;
            semimonthlyrepay = seminmonthlyinterest + principal;


        }
        //Method for determining the regular payments for the various payment options
        public double specificPay(double spefictime,double specificrate)
        {
            double a = Math.Pow((1 + specificrate), -spefictime);
            double aa = (1 - a) / specificrate;
            double ans = principal / aa;
            return ans;
        }
        void AmortizationTable()
        {
            
            //Annual Payment Option
            if (gunaComboBox1.SelectedIndex == 1)
            {
                bebalance = principal;
                schedulepayment = regularpayment;    
                if(!string.IsNullOrWhiteSpace(ExtraP.Text))
                {
                    extrapayment =  double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
                }
                else
                {
                extrapayment = 0;
                }

            if (extrapayment == principal)
            {
                interest = rate * bebalance;
                totalpayment = extrapayment + interest;

            }
            else if (extrapayment > principal)
            {
                interest = rate * bebalance;
                totalpayment = bebalance + interest;
            }
            else
            {
                interest = rate * bebalance;
                totalpayment = schedulepayment + extrapayment;
            }
               
                DateTime stardate = DateTime.Parse(paydate.Text);                           
                principalpaid = totalpayment - interest;
                endbalance = bebalance - principalpaid;
                for (DateTime d = stardate; d<=stardate.AddYears(int.Parse(Ltime.Text)); d = d.AddYears(1))
                {
                    dt.Rows.Add(d.ToString("dd/MM/yyy"), bebalance.ToString("c"), schedulepayment.ToString("c"),extrapayment.ToString("c"),totalpayment.ToString("c"), interest.ToString("c"), principalpaid.ToString("c"), endbalance.ToString("c"));
                    bebalance = bebalance - principalpaid;

                    schedulepayment = regularpayment;
                    interest = rate * bebalance;
                    extrapayment = extrapayment - 0;
                    totalpayment = schedulepayment + extrapayment;
                    if (totalpayment > bebalance)
                    {
                        totalpayment = bebalance + interest;
                       

                    }
                    else if (extrapayment > bebalance)
                    {
                        totalpayment = extrapayment + interest;
                    }
                    if (bebalance < 1)
                    {
                        return;
                    }
                    principalpaid = totalpayment - interest;
                    endbalance = bebalance - principalpaid;



                }


            }
            else if (gunaComboBox1.SelectedIndex == 2)
            {
                bebalance = principal;
                schedulepayment = regularpayment;
                if (!string.IsNullOrWhiteSpace(ExtraP.Text))
                {
                    extrapayment = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
                }
                else
                {
                    extrapayment = 0;
                }
                if (extrapayment == principal)
                {
                    interest = rate * bebalance;
                    totalpayment = extrapayment + interest;

                }
                else if (extrapayment > principal)
                {
                    interest = rate * bebalance;
                    totalpayment = bebalance + interest;
                }
                else
                {
                    interest = rate * bebalance;
                    totalpayment = schedulepayment + extrapayment;
                }

                
                DateTime stardate = DateTime.Parse(paydate.Text);
              
                interest = rate * bebalance;
                principalpaid = totalpayment - interest;
                endbalance = bebalance - principalpaid;
                for (DateTime d = stardate; d <= stardate.AddYears(int.Parse(Ltime.Text)); d = d.AddMonths(6))
                {
                    dt.Rows.Add(d.ToString("dd/MM/yyy"), bebalance.ToString("c"), schedulepayment.ToString("c"),extrapayment.ToString("c"),totalpayment.ToString("c"), interest.ToString("c"), principalpaid.ToString("c"), endbalance.ToString("c"));
                    bebalance = bebalance - principalpaid;

                    schedulepayment = regularpayment;
                    interest = rate * bebalance;
                    extrapayment = extrapayment - 0;
                    totalpayment = schedulepayment + extrapayment;
                    if (totalpayment > bebalance)
                    {
                        totalpayment = bebalance + interest;


                    }
                    else if (extrapayment > bebalance)
                    {
                        totalpayment = extrapayment + interest;
                    }
                    if (bebalance < 1)
                    {
                        return;
                    }
                    principalpaid = totalpayment - interest;
                    endbalance = bebalance - principalpaid;



                }
            }
            else if (gunaComboBox1.SelectedIndex == 3)
            {
                
                    bebalance = principal;
                    schedulepayment = regularpayment;
                    if (!string.IsNullOrWhiteSpace(ExtraP.Text))
                    {
                        extrapayment = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
                    }
                    else
                    {
                        extrapayment = 0;
                    }
                    if (extrapayment == principal)
                    {
                        interest = rate * bebalance;
                        totalpayment = extrapayment + interest;

                    }
                    else if (extrapayment > principal)
                    {
                        interest = rate * bebalance;
                        totalpayment = bebalance + interest;
                    }
                    else
                    {
                        interest = rate * bebalance;
                        totalpayment = schedulepayment + extrapayment;
                    }

                    DateTime stardate = DateTime.Parse(paydate.Text);
                    interest = rate * bebalance;
                    principalpaid = totalpayment - interest;
                    endbalance = bebalance - principalpaid;
                    
                    for (DateTime d = stardate; d <= stardate.AddYears(int.Parse(Ltime.Text)); d = d.AddMonths(1))
                    {
                        dt.Rows.Add(d.ToString("dd/MM/yyy"), bebalance.ToString("c"), schedulepayment.ToString("c"), extrapayment.ToString("c"), totalpayment.ToString("c"), interest.ToString("c"), principalpaid.ToString("c"), endbalance.ToString("c"));
                        bebalance = bebalance - principalpaid;

                        schedulepayment = regularpayment;
                        interest = rate * bebalance;
                        extrapayment = extrapayment - 0;
                        totalpayment = schedulepayment + extrapayment;
                        if (totalpayment > bebalance)
                        {
                            totalpayment = bebalance + interest;


                        }
                        else if (extrapayment > bebalance)
                        {
                            totalpayment = extrapayment + interest;
                        }
                        if (bebalance < 1)
                        {
                            return;
                        }
                        principalpaid = totalpayment - interest;
                        endbalance = bebalance - principalpaid;



                    }
                
            }
            else if (gunaComboBox1.SelectedIndex == 4)
            {
                bebalance = principal;
                schedulepayment = regularpayment;
                if (!string.IsNullOrWhiteSpace(ExtraP.Text))
                {
                    extrapayment = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
                }
                else
                {
                    extrapayment = 0;
                }
                if (extrapayment == principal)
                {
                    interest = rate * bebalance;
                    totalpayment = extrapayment + interest;

                }
                else if (extrapayment > principal)
                {
                    interest = rate * bebalance;
                    totalpayment = bebalance + interest;
                }
                else
                {
                    interest = rate * bebalance;
                    totalpayment = schedulepayment + extrapayment;
                }
               
                DateTime stardate = DateTime.Parse(paydate.Text);
                interest = rate * bebalance;
                principalpaid = totalpayment - interest;
                endbalance = bebalance - principalpaid;
                for (DateTime d = stardate; d <= stardate.AddYears(int.Parse(Ltime.Text)); d = d.AddDays(7))
                {
                    dt.Rows.Add(d.ToString("dd/MM/yyy"), bebalance.ToString("c"), schedulepayment.ToString("c"), extrapayment.ToString("c"),totalpayment.ToString("c"),interest.ToString("c"), principalpaid.ToString("c"), endbalance.ToString("c"));
                    bebalance = bebalance - principalpaid;

                    schedulepayment = regularpayment;
                    interest = rate * bebalance;
                    extrapayment = extrapayment - 0;
                    totalpayment = schedulepayment + extrapayment;
                    if (totalpayment > bebalance)
                    {
                        totalpayment = bebalance + interest;


                    }
                    else if (extrapayment > bebalance)
                    {
                        totalpayment = extrapayment + interest;
                    }
                    if (bebalance < 1)
                    {
                        return;
                    }
                    principalpaid = totalpayment - interest;
                    endbalance = bebalance - principalpaid;



                }
            }
            else if (gunaComboBox1.SelectedIndex == 5)
            {
                bebalance = principal;
                schedulepayment = regularpayment;
                if (!string.IsNullOrWhiteSpace(ExtraP.Text))
                {
                    extrapayment = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
                }
                else
                {
                    extrapayment = 0;
                }
                if (extrapayment == principal)
                {
                    interest = rate * bebalance;
                    totalpayment = extrapayment + interest;

                }
                else if (extrapayment > principal)
                {
                    interest = rate * bebalance;
                    totalpayment = bebalance + interest;
                }
                else
                {
                    interest = rate * bebalance;
                    totalpayment = schedulepayment + extrapayment;
                }
               
                DateTime stardate = DateTime.Parse(paydate.Text);                
                interest = rate * bebalance;
                principalpaid = totalpayment - interest;
                endbalance = bebalance - principalpaid;
                for (DateTime d = stardate; d <= stardate.AddYears(int.Parse(Ltime.Text)); d = d.AddMonths(3))
                {
                    dt.Rows.Add(d.ToString("dd/MM/yyy"), bebalance.ToString("c"), schedulepayment.ToString("c"),extrapayment.ToString("c"),totalpayment.ToString("c"), interest.ToString("c"), principalpaid.ToString("c"), endbalance.ToString("c"));
                    bebalance = bebalance - principalpaid;

                    schedulepayment = regularpayment;
                    interest = rate * bebalance;
                    extrapayment = extrapayment - 0;
                    totalpayment = schedulepayment + extrapayment;
                    if (totalpayment > bebalance)
                    {
                        totalpayment = bebalance + interest;


                    }
                    else if (extrapayment > bebalance)
                    {
                        totalpayment = extrapayment + interest;
                    }
                    if (bebalance < 1)
                    {
                        return;
                    }
                    principalpaid = totalpayment - interest;
                    endbalance = bebalance - principalpaid;



                }
            }
            else if (gunaComboBox1.SelectedIndex == 6)
            {
                bebalance = principal;
                schedulepayment = regularpayment;
                if (!string.IsNullOrWhiteSpace(ExtraP.Text))
                {
                    extrapayment = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
                }
                else
                {
                    extrapayment = 0;
                }
                if (extrapayment == principal)
                {
                    interest = rate * bebalance;
                    totalpayment = extrapayment + interest;

                }
                else if (extrapayment > principal)
                {
                    interest = rate * bebalance;
                    totalpayment = bebalance + interest;
                }
                else
                {

                    interest = rate * bebalance;
                    totalpayment = schedulepayment + extrapayment;
                }
                
                DateTime stardate = DateTime.Parse(paydate.Text);             
                interest = rate * bebalance;
                principalpaid =totalpayment- interest;
                endbalance = bebalance - principalpaid;
                for (DateTime d = stardate; d <= stardate.AddYears(int.Parse(Ltime.Text)); d = d.AddDays(1))
                {
                    dt.Rows.Add(d.ToString("dd/MM/yyy"), bebalance.ToString(""), schedulepayment.ToString("c"),extrapayment.ToString("c"),totalpayment.ToString("c"), interest.ToString("c"), principalpaid.ToString("c"), endbalance.ToString("c"));
                    bebalance = bebalance - principalpaid;

                    schedulepayment = regularpayment;
                    interest = rate * bebalance;
                    extrapayment = extrapayment - 0;
                    totalpayment = schedulepayment + extrapayment;
                    if (totalpayment > bebalance)
                    {
                        totalpayment = bebalance + interest;


                    }
                    else if (extrapayment > bebalance)
                    {
                        totalpayment = extrapayment + interest;
                    }
                   
                    if (bebalance < 1)
                    {
                        return;
                    }
                    
                    principalpaid =totalpayment - interest;
                    endbalance = bebalance - principalpaid;



                }
            }
            else if (gunaComboBox1.SelectedIndex == 7)
            {
                bebalance = principal;
                schedulepayment = regularpayment;
                if (!string.IsNullOrWhiteSpace(ExtraP.Text))
                {
                    extrapayment = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
                }
                else
                {
                    extrapayment = 0;
                }
                if (extrapayment == principal)
                {
                    interest = rate * bebalance;
                    totalpayment = extrapayment + interest;

                }
                else if (extrapayment > principal)
                {
                    interest = rate * bebalance;
                    totalpayment = bebalance + interest;
                }
                else
                {

                    interest = rate * bebalance;
                    totalpayment = schedulepayment + extrapayment;
                }

                DateTime stardate = DateTime.Parse(paydate.Text);
                interest = rate * bebalance;
                principalpaid = totalpayment - interest;
                endbalance = bebalance - principalpaid;
                for (DateTime d = stardate; d <= stardate.AddYears(int.Parse(Ltime.Text)); d = d.AddDays(14))
                {
                    dt.Rows.Add(d.ToString("dd/MM/yyy"), bebalance.ToString("c"), schedulepayment.ToString("c"), extrapayment.ToString("c"), totalpayment.ToString("c"), interest.ToString("c"), principalpaid.ToString("c"), endbalance.ToString("c"));
                    bebalance = bebalance - principalpaid;

                    schedulepayment = regularpayment;
                    interest = rate * bebalance;
                    extrapayment = extrapayment - 0;
                    totalpayment = schedulepayment + extrapayment;
                    if (totalpayment > bebalance)
                    {
                        totalpayment = bebalance + interest;


                    }
                    else if (extrapayment > bebalance)
                    {
                        totalpayment = extrapayment + interest;
                    }

                    if (bebalance < 1)
                    {
                        return;
                    }

                    principalpaid = totalpayment - interest;
                    endbalance = bebalance - principalpaid;



                }
            }
            else if (gunaComboBox1.SelectedIndex == 8)
            {
                bebalance = principal;
                schedulepayment = regularpayment;
                if (!string.IsNullOrWhiteSpace(ExtraP.Text))
                {
                    extrapayment = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
                }
                else
                {
                    extrapayment = 0;
                }
                if (extrapayment == principal)
                {
                    interest = rate * bebalance;
                    totalpayment = extrapayment + interest;

                }
                else if (extrapayment > principal)
                {
                    interest = rate * bebalance;
                    totalpayment = bebalance + interest;
                }
                else
                {

                    interest = rate * bebalance;
                    totalpayment = schedulepayment + extrapayment;
                }

                DateTime stardate = DateTime.Parse(paydate.Text);
                interest = rate * bebalance;
                principalpaid = totalpayment - interest;
                endbalance = bebalance - principalpaid;
                for (DateTime d = stardate; d <= stardate.AddYears(int.Parse(Ltime.Text)); d = d.AddMonths(2))
                {
                    dt.Rows.Add(d.ToString("dd/MM/yyy"), bebalance.ToString("c"), schedulepayment.ToString("c"), extrapayment.ToString("c"), totalpayment.ToString("c"), interest.ToString("c"), principalpaid.ToString("c"), endbalance.ToString("c"));
                    bebalance = bebalance - principalpaid;

                    schedulepayment = regularpayment;
                    interest = rate * bebalance;
                    extrapayment = extrapayment - 0;
                    totalpayment = schedulepayment + extrapayment;
                    if (totalpayment > bebalance)
                    {
                        totalpayment = bebalance + interest;


                    }
                    else if (extrapayment > bebalance)
                    {
                        totalpayment = extrapayment + interest;
                    }

                    if (bebalance < 1)
                    {
                        return;
                    }

                    principalpaid = totalpayment - interest;
                    endbalance = bebalance - principalpaid;



                }
            }
            else if (gunaComboBox1.SelectedIndex == 9)
            {
                bebalance = principal;
                schedulepayment = regularpayment;
                if (!string.IsNullOrWhiteSpace(ExtraP.Text))
                {
                    extrapayment = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
                }
                else
                {
                    extrapayment = 0;
                }
                if (extrapayment == principal)
                {
                    interest = rate * bebalance;
                    totalpayment = extrapayment + interest;

                }
                else if (extrapayment > principal)
                {
                    interest = rate * bebalance;
                    totalpayment = bebalance + interest;
                }
                else
                {

                    interest = rate * bebalance;
                    totalpayment = schedulepayment + extrapayment;
                }

                DateTime stardate = DateTime.Parse(paydate.Text);
                interest = rate * bebalance;
                principalpaid = totalpayment - interest;
                endbalance = bebalance - principalpaid;
                for (DateTime d = stardate; d <= stardate.AddYears(int.Parse(Ltime.Text)); d = d.AddDays(15))
                {
                    dt.Rows.Add(d.ToString("dd/MM/yyy"), bebalance.ToString("c"), schedulepayment.ToString("c"), extrapayment.ToString("c"), totalpayment.ToString("c"), interest.ToString("c"), principalpaid.ToString("c"), endbalance.ToString("c"));
                    bebalance = bebalance - principalpaid;

                    schedulepayment = regularpayment;
                    interest = rate * bebalance;
                    extrapayment = extrapayment - 0;
                    totalpayment = schedulepayment + extrapayment;
                    if (totalpayment > bebalance)
                    {
                        totalpayment = bebalance + interest;


                    }
                    else if (extrapayment > bebalance)
                    {
                        totalpayment = extrapayment + interest;
                    }

                    if (bebalance < 1)
                    {
                        return;
                    }

                    principalpaid = totalpayment - interest;
                    endbalance = bebalance - principalpaid;



                }
            }
        }

        

        void ShowingSummary1()
        {
            double totalI = 0;
            double totalpay = 0;            
            foreach (DataGridViewRow r in gunaDataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in r.Cells)
                {
                    if (cell.ColumnIndex == 5)
                    {
                        double interest = double.Parse(Regex.Replace(cell.Value.ToString(), "[A-Za-z $]", ""));
                        totalI += interest;

                    }
                }
            }
            foreach (DataGridViewRow r in gunaDataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in r.Cells)
                {
                    if (cell.ColumnIndex == 4)
                    {
                        double totalpayment1 = double.Parse(Regex.Replace(cell.Value.ToString(), "[A-Za-z $]", ""));
                        totalpay += totalpayment1;

                    }
                }
            }
            totalInterest.Text = totalI.ToString("c");
            gunaTextBox6.Text = totalpay.ToString("c");
            int i = gunaDataGridView1.Rows.Count - 1;
            payoff.Text = gunaDataGridView1.Rows[i].Cells[0].Value.ToString();           
            rT.Text = rate.ToString();

        }

      

        private void Form1_Load(object sender, EventArgs e)
        {   string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Calculate Your Loan\\Loan Reports";
            ToolTip tool = new ToolTip();
            tool.ToolTipTitle = Application.ProductName;
            tool.SetToolTip(button6, "Viewed Reports are automatically saved at" + Environment.NewLine + path);
            timer1.Start();
            timer2.Start();
            timer3.Start();
            foreach (Control c in groupBox2.Controls)
            {
                if (c is GunaTextBox)
                {
                    GunaTextBox t = c as GunaTextBox;
                    t.Visible = false;
                }
            }
            foreach(DataGridViewColumn colum in gunaDataGridView1.Columns)
            {
                colum.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            DateTime todaydate = DateTime.Now;
            paydate.Text = todaydate.ToShortDateString();
            gunaDataGridView1.ThemeStyle.HeaderStyle.BackColor = ColorTranslator.FromHtml("#01466e");
            gunaDataGridView1.ThemeStyle.RowsStyle.SelectionBackColor = ColorTranslator.FromHtml("#01466e");
            

            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy == true)
            {
                MessageBox.Show("CSV file exporting is still in progress", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            ClearAll();
        }
        

        private void ClearAll()
        {
            foreach(Control c in groupBox1.Controls)
            {
                if(c is GunaTextBox)
                {
                    GunaTextBox t = c as GunaTextBox;
                    t.Text = "";
                }
            }
            foreach (Control c in groupBox2.Controls)
            {
                if (c is GunaTextBox)
                {
                    GunaTextBox t = c as GunaTextBox;
                    t.Visible = false;
                }
            }
            paydate.Text = DateTime.Now.ToShortDateString();
            gunaComboBox1.SelectedIndex = 0;
            RegP.Text = "";           
            dt.Rows.Clear();
            Lamt.Clear();
            ExtraP.Clear();
            countingP.Text = "";
           
        }

        private void gunaPictureBox1_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy == true)
            {
                DialogResult r = MessageBox.Show("CSV file exporting is still in progress\nDo you want to exit the application anyway?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
                if(r==DialogResult.Yes)
                {
                    backgroundWorker1.CancelAsync();
                    Application.Exit();
                    
                }
          
            }
            else
            {
                Application.Exit();
            }
        }

        private void gunaPictureBox2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy == true)
            {
                MessageBox.Show("CSV file exporting is still in progress", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            CheckForIllegalCrossThreadCalls = false;
            if(gunaDataGridView1.Rows.Count!=0)
            {
                save= new SaveFileDialog();
                save.Title = "Exporting Output As";
                save.Filter = "CSV File|*.csv|PDF File|*.pdf";
                save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                save.FileName = "A " + Ltime.Text + " year loan with payment frequency[ "+gunaComboBox1.Text+" ]";
                if (save.ShowDialog() == DialogResult.OK)
                {
                    try
                    {

                        if(save.FilterIndex==1)
                        {
                            backgroundWorker1.RunWorkerAsync();
                        }
                        else if(save.FilterIndex==2)
                        {
                            ExportingScheduleOnlyAsPDF(save.FileName);
                            MessageBox.Show("PDF file successfully exported", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
              
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);

                    }
                }
            }
            else
            {
                
                MessageBox.Show("Nothing to save", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            
        }
        void exportingCSV(string csvfilename)
        {
            StreamWriter cw = new StreamWriter(csvfilename);

            string csv = string.Empty;
            foreach (DataGridViewColumn column in gunaDataGridView1.Columns)
            {
                csv = csv + column.HeaderText + ",";

            }
            //csv = csv.TrimEnd(',');
            csv += "\r\n";
            foreach (DataGridViewRow row in gunaDataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    csv += cell.Value.ToString().TrimEnd(',').Replace(",", "") + ",";
                }
                csv += "\r\n";
            }
            cw.Write(csv,false);
            cw.Close();
        }
        #region Method For Exporting To PDF
        void ExportingOutputPDF(string filename)
        {
           
            iTextSharp.text.Rectangle rec = new iTextSharp.text.Rectangle(iTextSharp.text.PageSize.A4.Rotate());
            Document dd = new Document(rec);
            FileStream fs = new FileStream(filename,FileMode.Create);
            PdfWriter wr = PdfWriter.GetInstance(dd, fs);
            dd.Open();
            //Writing the PDF Content
            //declaring font and colors
            iTextSharp.text.Font HeaderFont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 17, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font ContentFont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 10, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font TableHeaderFont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 10, iTextSharp.text.Font.NORMAL,BaseColor.WHITE);
            iTextSharp.text.Font SumFont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 13, iTextSharp.text.Font.NORMAL,BaseColor.WHITE);
            iTextSharp.text.Font SumFont1 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 11, iTextSharp.text.Font.BOLDITALIC); 
            iTextSharp.text.Font rtFont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 11, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font ContentFont1 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 12, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font footerfont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 9, iTextSharp.text.Font.ITALIC);
    

            double amt1 = 0;
            string name1 = gunaComboBox1.Text.Replace("Payment", "");
            if (string.IsNullOrWhiteSpace(ExtraP.Text))
            {
                amt1 = 0;
            }
            else
            {
                amt1 = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
            }            
            Chunk basicInformation = new Chunk("LOAN BASIC INFORMATION");
            basicInformation.Font = SumFont1;
            PdfPTable dt0 = new PdfPTable(2);
            dt0.HorizontalAlignment = Element.ALIGN_LEFT;
            dt0.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            dt0.SetWidths(new float[] { 100, 275 });
            dt0.TotalWidth = 750;
            dt0.DefaultCell.Padding=5;
            dt0.SpacingAfter = 10;
            dt0.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
            dt0.DefaultCell.FixedHeight = 25;
            dt0.DefaultCell.BorderWidth = 0;
            PdfPCell c = new PdfPCell(new Phrase(basicInformation));
            c.Colspan = 2;
            c.HorizontalAlignment = Element.ALIGN_LEFT;
            c.VerticalAlignment = Element.ALIGN_MIDDLE;
            c.BorderWidth = 0;
            c.FixedHeight = 30;
            c.PaddingBottom = 4;
            c.PaddingTop = 4;           
            dt0.AddCell(c);
            dt0.AddCell(new Phrase("Payment Frequency", ContentFont));
            dt0.AddCell(new Phrase(gunaComboBox1.Text, ContentFont));
            dt0.AddCell(new Phrase("Loan Amount",ContentFont));            
            double amt = double.Parse(Regex.Replace(Lamt.Text, "[A-Za-z $]", ""));
            dt0.AddCell(new Phrase(amt.ToString("c"), ContentFont));
            dt0.AddCell(new Phrase(name1+" Extra Payment", ContentFont));          
            dt0.AddCell(new Phrase(amt1.ToString("c"), ContentFont));
            dt0.AddCell(new Phrase("Annual Interest Rate", ContentFont));
            dt0.AddCell(new Phrase(Lrate.Text+" %", ContentFont));
            dt0.AddCell(new Phrase("Loan Term(yrs)", ContentFont));
            dt0.AddCell(new Phrase(Ltime.Text, ContentFont));
            dt0.AddCell(new Phrase("Initial Payment Date", ContentFont));
            dt0.AddCell(new Phrase(gunaDataGridView1.Rows[0].Cells[0].Value.ToString(), ContentFont));
            dd.Add(dt0);
            Chunk summary = new Chunk("LOAN SUMMARY");
            summary.Font = SumFont1;
            //Paragraph par2 = new Paragraph();
            //par2.Alignment = Element.ALIGN_LEFT;
            //par2.SpacingAfter = 5;
            //par2.Add(summary);
            //dd.Add(par2);

            PdfPTable dt1 = new PdfPTable(2);
            dt1.HorizontalAlignment = Element.ALIGN_LEFT;
            dt1.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            dt1.SetWidths(new float[] { 100,275 });
            dt1.TotalWidth = 750;
            dt1.SpacingAfter = 15 ;
            dt1.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
            dt1.DefaultCell.FixedHeight = 25;
            dt1.DefaultCell.Padding=5;
            dt1.DefaultCell.BorderWidth = 0;

            PdfPCell c3 = new PdfPCell(new Phrase(summary));
            c3.HorizontalAlignment = Element.ALIGN_LEFT;
            c3.VerticalAlignment = Element.ALIGN_MIDDLE;
            c3.BorderWidth = 0;
            c3.FixedHeight = 30;
            c3.Colspan = 2;
            c3.PaddingTop = 4;
            c3.PaddingBottom = 4;
            dt1.AddCell(c3);
            dt1.AddCell(new Phrase("Expected " +gunaComboBox1.Text, ContentFont));
            dt1.AddCell(new Phrase(RegP.Text, ContentFont));
            dt1.AddCell(new Phrase("Actual " + gunaComboBox1.Text, ContentFont));
            double pay1 = regularpayment;
            double pay2;
            if(string.IsNullOrWhiteSpace(ExtraP.Text))
            {
                pay2 = 0;
            }
            else
            {
               pay2= double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
            }
            double pay3 = pay1 + pay2;
            dt1.AddCell(new Phrase(pay3.ToString("c"), ContentFont));
            dt1.AddCell(new Phrase("Rate(per period)", ContentFont));
            double perperiodrate = double.Parse(rT.Text)*100;
            dt1.AddCell(new Phrase(perperiodrate.ToString("n3")+"%"+" (approximate figure)", ContentFont));
            dt1.AddCell(new Phrase("Total Repayment", ContentFont));
            dt1.AddCell(new Phrase(gunaTextBox6.Text, ContentFont));
            dt1.AddCell(new Phrase("Total Interest Paid", ContentFont));
            dt1.AddCell(new Phrase(totalInterest.Text, ContentFont));
            dt1.AddCell(new Phrase("Payoff Date", ContentFont));
            dt1.AddCell(new Phrase(payoff.Text, ContentFont));
            dt1.AddCell(new Phrase("Number of Payments", ContentFont));
            dt1.AddCell(new Phrase(gunaDataGridView1.Rows.Count.ToString()+" "+gunaComboBox1.Text+"s", ContentFont));
            dd.Add(dt1);
            string reportitle = string.Empty;
            double loanyear = double.Parse(Ltime.Text);
            reportitle = "IMPACT OF MAKING " + name1.ToUpper() + " EXTRA PAYMENT";
            Chunk rttitle = new Chunk(reportitle);
            rttitle.Font = rtFont;
            Paragraph rtp = new Paragraph(rttitle);
            rtp.Alignment = Element.ALIGN_LEFT;
            rtp.SpacingAfter = 5;
            if(st.IncludeImpact==true)
            {
                dd.Add(rtp);
            }
            double actualinterest = double.Parse(Regex.Replace(totalInterest.Text, "[A-Za-z $]", ""));
            double savingsA = actualtotalinterest - actualinterest;
            double h = 0;
            if (!string.IsNullOrWhiteSpace(ExtraP.Text))
            {
                h = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
            }
            double percentI =savingsA / actualtotalinterest * 100;
            Chunk savings = new Chunk("According to the calculations, it appears that by paying an extra amount of " + h.ToString("c") + ", you can save an amount of " + savingsA.ToString("c") + " in interest.This amount saved is the difference between your Original Total Interest of " + actualtotalinterest.ToString("c")+" ( This amount represents the total interest who would pay for making zero extra payment ) and your New Total Interest of " + actualinterest.ToString("c") + ". That means you can save " + percentI.ToString("n1") + "% on interest. In addition to the interest savings aspect, you can pay off the mortgage sooner than your mortgage term. For making extra payment,expected number of payments has reduced from "+time+" to "+gunaDataGridView1.Rows.Count.ToString()+". That means the higher extra payment you make, the earlier you can pay off the debt than the expected term");
            Chunk savings1 = new Chunk("No extra payment");
            savings1.Font = ContentFont1;
            savings.Font = ContentFont1;
            Paragraph readmore = new Paragraph();

            readmore.SpacingAfter = 25;
            if (string.IsNullOrWhiteSpace(ExtraP.Text))
            {
                readmore.Alignment = Element.ALIGN_JUSTIFIED;
                readmore.Add(savings1);
            }
            else
            {
                readmore.Alignment = Element.ALIGN_JUSTIFIED;
                readmore.Add(savings);
            }
           if(st.IncludeImpact==true)
           {
               dd.Add(readmore);
           }
            Chunk compTitle = new Chunk("COMPARISON OF PAYMENT FREQUENCY OPTIONS BASED ON REGULAR PAYMENT, TOTAL REPAYMENT AND TOTAL INTEREST");
            compTitle.Font = rtFont;
            Paragraph comparisonP = new Paragraph();
            comparisonP.Alignment = Element.ALIGN_LEFT;
            comparisonP.SpacingAfter = 10;
            comparisonP.Add(compTitle);
            double Specifiregularpay;
            if(string.IsNullOrWhiteSpace(ExtraP.Text))
            {
                Specifiregularpay = 0;
            }
            else
            {
              Specifiregularpay=  double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
            }
           if(st.IncludeComparison==true)
           {
               dd.Add(comparisonP);
           }
            PdfPTable comparisonTable = new PdfPTable(4);
            comparisonTable.WidthPercentage = 100;
            comparisonTable.HorizontalAlignment = Element.ALIGN_LEFT;
            comparisonTable.SpacingAfter = 3;
            comparisonTable.DefaultCell.Padding = 5;            
            comparisonTable.DefaultCell.BorderWidth = 0;
            string[] cheaders = new string[] { "Payment Frequency","Regular Payment" ,"Total Repayment","Total Interest" };
            foreach(string f in cheaders)
            {
                PdfPCell ccell = new PdfPCell(new Phrase(f,SumFont));
                ccell.HorizontalAlignment = Element.ALIGN_LEFT;
                ccell.FixedHeight = 30;
                ccell.Padding = 5;
                ccell.VerticalAlignment = Element.ALIGN_CENTER;
                ccell.BorderWidth = 0;
                ccell.BackgroundColor = new BaseColor(ColorTranslator.FromHtml("#01466e"));               
                comparisonTable.AddCell(ccell);
            }
            comparisonTable.HeaderRows = 1;
            comparisonTable.AddCell(new Phrase("Annual Payment Option",ContentFont));
            comparisonTable.AddCell(new Phrase(annualpay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(AnnualRepay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(annualInterest.ToString("c"), ContentFont));
           
            comparisonTable.AddCell(new Phrase("Semi - Annual Payment Option", ContentFont));
            comparisonTable.AddCell(new Phrase(semiannualpay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(semiannualrepay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(semiannualInterest.ToString("c"), ContentFont));
          
            comparisonTable.AddCell(new Phrase("Monthly Payment Option", ContentFont));
            comparisonTable.AddCell(new Phrase(monthlypay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(monthlyrepay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(monthlyinterest.ToString("c"), ContentFont));
        
            comparisonTable.AddCell(new Phrase("Weekly Payment Option", ContentFont));
            comparisonTable.AddCell(new Phrase(weeklypay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(weeklyrepay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(weeklyinterest.ToString("c"), ContentFont));
       
            comparisonTable.AddCell(new Phrase("Quarterly Payment Option", ContentFont));
            comparisonTable.AddCell(new Phrase(quartelypay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(quaterlyrepay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(quartelyInterest.ToString("c"), ContentFont));
          
            comparisonTable.AddCell(new Phrase("Daily Payment Option", ContentFont));
            comparisonTable.AddCell(new Phrase(dailypay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(dailyrepay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(dailyinterest.ToString("c"), ContentFont));
      
            comparisonTable.AddCell(new Phrase("Bi-Weekly Payment Option", ContentFont));
            comparisonTable.AddCell(new Phrase(biweeklypay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(biweekrepay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(biweeklyinterest.ToString("c"), ContentFont));
          
            comparisonTable.AddCell(new Phrase("Bi-Monthly Payment Option", ContentFont));
            comparisonTable.AddCell(new Phrase(bimonthlypay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(bimonthlyrepay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(bimonthlyinterest.ToString("c"), ContentFont));
        
            comparisonTable.AddCell(new Phrase("Semi-Monthly Payment Option", ContentFont));
            comparisonTable.AddCell(new Phrase(seminmonthlypay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(semimonthlyrepay.ToString("c"), ContentFont));
            comparisonTable.AddCell(new Phrase(seminmonthlyinterest.ToString("c"), ContentFont));
            
           if(st.IncludeComparison==true)
           {
               dd.Add(comparisonTable);
           }
            PdfPTable line = new PdfPTable(1);
            line.WidthPercentage = 100;
            line.HorizontalAlignment = Element.ALIGN_CENTER;          
            PdfPCell linecell = new PdfPCell();
            linecell.BorderWidth = 1f;
            linecell.HorizontalAlignment = Element.ALIGN_CENTER;
            line.AddCell(linecell);
            if(st.IncludeComparison==true)
            {
                dd.Add(line);
            }

            Paragraph footer1 = new Paragraph("Please these figures do not account for extra payment");
            footer1.Alignment = Element.ALIGN_LEFT;
            footer1.SpacingAfter = 30;
            footer1.Font = footerfont;
            if(st.IncludeComparison==true)
            {
                dd.Add(footer1);
            }
            

            Chunk summary1 = new Chunk("Amortization Schedule");
            summary1.Font = SumFont;
            //Paragraph par3 = new Paragraph();
            //par3.Alignment = Element.ALIGN_LEFT;
            //par3.SpacingAfter = 10;
            //par3.Add(summary1);
            //dd.Add(par3);
            PdfPTable dt3 = new PdfPTable(1);
            dt3.WidthPercentage = 100;
            dt3.HorizontalAlignment = Element.ALIGN_CENTER;   
            dt3.SpacingAfter = 10;
            PdfPCell mainheader = new PdfPCell(new Phrase(summary1));
            mainheader.FixedHeight = 30;
            mainheader.BorderWidth = 0;
            mainheader.BackgroundColor = new BaseColor(ColorTranslator.FromHtml("#01466e"));
            mainheader.HorizontalAlignment = Element.ALIGN_LEFT;
            mainheader.VerticalAlignment = Element.ALIGN_CENTER;
            mainheader.PaddingBottom = 4;
            mainheader.PaddingTop = 4;
            mainheader.PaddingLeft = 5;
            mainheader.Colspan = 1;
            dt3.AddCell(mainheader);
            dd.Add(dt3);

            PdfPTable dt2= new PdfPTable(8);
            dt2.WidthPercentage = 100;
            dt2.HorizontalAlignment = Element.ALIGN_CENTER;
            dt2.HeaderRows = 1;
            dt2.SpacingAfter = 10;
            
            dt.Columns[2].ColumnName = "Payment";
            

            foreach(DataGridViewColumn col in gunaDataGridView1.Columns)
            {
                PdfPCell cell=new PdfPCell(new Phrase(col.HeaderText,TableHeaderFont));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;                
                cell.BorderColor = BaseColor.WHITE;
                BaseColor b = new BaseColor(ColorTranslator.FromHtml("#01466e"));
                cell.BackgroundColor =b;
                cell.FixedHeight = 30;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.Padding = 5;
                dt2.AddCell(cell);

            }

            foreach(DataGridViewRow r in gunaDataGridView1.Rows)
            {
                foreach(DataGridViewCell cl in r.Cells)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(cl.Value.ToString(), ContentFont));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    if (st.ShowGridLines == true)
                    {
                        cell.BorderWidth = 0.5f;
                    }
                    else
                    {
                        cell.BorderWidth = 0;
                    }
                    cell.BorderColor = new BaseColor(64, 64, 64);
                    cell.FixedHeight = 24;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cell.Padding = 5;
                    dt2.AddCell(cell);
                }
            }

            dd.Add(dt2);
            Paragraph footer = new Paragraph("Please allow for slight rounding differences");
            footer.Alignment = Element.ALIGN_LEFT;
            footer.Font = footerfont;
            dd.Add(footer);
            dd.Close();

        }
        void ExportingScheduleOnlyAsPDF(string pdfilename)
        {
            iTextSharp.text.Rectangle rec = new iTextSharp.text.Rectangle(iTextSharp.text.PageSize.A4.Rotate());
            Document dd = new Document(rec);
            FileStream fs = new FileStream(pdfilename, FileMode.Create);
            PdfWriter wr = PdfWriter.GetInstance(dd, fs);
            dd.Open();
            iTextSharp.text.Font rtFont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 13, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font ContentFont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 12, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font ContentFont1 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 13, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font TableHeaderFont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 10, iTextSharp.text.Font.NORMAL, BaseColor.WHITE);
            iTextSharp.text.Font SumFont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.TIMES_ROMAN.ToString(), 13, iTextSharp.text.Font.NORMAL, BaseColor.WHITE);
            Chunk header = new Chunk("AMORTIZATION SCHEDULE");
            header.Font = rtFont;
            Paragraph headerP = new Paragraph();
            headerP.Alignment = Element.ALIGN_CENTER;
            headerP.SpacingAfter = 10;
            headerP.Add(header);
            dd.Add(headerP);
            PdfPTable summary = new PdfPTable(2);
            summary.TotalWidth = 750;
            summary.SetWidths(new float[] { 100, 257 });
            summary.HorizontalAlignment = Element.ALIGN_LEFT;
            summary.SpacingAfter = 10;            
            summary.DefaultCell.FixedHeight = 25;
            summary.DefaultCell.Padding = 5;
            summary.DefaultCell.BorderWidth = 0;

            summary.AddCell(new Phrase("Loan Amount", ContentFont1));
            summary.AddCell(new Phrase(Lamt.Text, ContentFont));
            summary.AddCell(new Phrase("Annual Interest Rate", ContentFont1));
            summary.AddCell(new Phrase(Lrate.Text+"%", ContentFont));
            summary.AddCell(new Phrase("Payment Frequency", ContentFont1));
            summary.AddCell(new Phrase(gunaComboBox1.Text, ContentFont));
            summary.AddCell(new Phrase("Number of Payments", ContentFont1));
            summary.AddCell(new Phrase(gunaDataGridView1.Rows.Count.ToString(), ContentFont));         
           summary.AddCell(new Phrase("Regular Payment", ContentFont1));
           

            double a1 = regularpayment;
            double a2;
            if(string.IsNullOrWhiteSpace(ExtraP.Text))
            {
                a2 = 0;
            }
            else
            {
                a2 = double.Parse(Regex.Replace(ExtraP.Text, "[A-Za-z $]", ""));
            }
            double a3 = a1 + a2;
            if(a2==0)
            {
                summary.AddCell(new Phrase(a3.ToString("c") + " (Including No Extra Payment)", ContentFont));
          
            }
            else
            {
                summary.AddCell(new Phrase(a3.ToString("c") + " (Including Extra Payment of " + a2.ToString("c")+")", ContentFont));         
            }

            dd.Add(summary);

            PdfPTable dt2 = new PdfPTable(8);
            dt2.WidthPercentage = 100;
            dt2.HorizontalAlignment = Element.ALIGN_CENTER;
            dt2.HeaderRows = 1;
            dt2.SpacingAfter = 10;
            dt.Columns[2].ColumnName = "Payment";
            foreach (DataGridViewColumn col in gunaDataGridView1.Columns)
            {
                PdfPCell cell = new PdfPCell(new Phrase(col.HeaderText, TableHeaderFont));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.BorderWidth = 0.5f;
                cell.BorderColor = BaseColor.WHITE;
                BaseColor b = new BaseColor(ColorTranslator.FromHtml("#01466e"));
                cell.BackgroundColor = b;
                cell.FixedHeight = 30;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.Padding = 5;
                dt2.AddCell(cell);
            }
            foreach (DataGridViewRow r in gunaDataGridView1.Rows)
            {
                foreach (DataGridViewCell cl in r.Cells)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(cl.Value.ToString(), ContentFont));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    if(st.ShowGridLines==true)
                    {
                        cell.BorderWidth = 0.5f;
                    }
                    else
                    {
                        cell.BorderWidth = 0;
                    }
                    cell.BorderColor = new BaseColor(64, 64, 64);
                    cell.FixedHeight = 24;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cell.Padding = 5;
                    dt2.AddCell(cell);
                }
            }


            dd.Add(dt2);
            Paragraph footer = new Paragraph("Please allow for slight rounding differences");
            footer.Alignment = Element.ALIGN_LEFT;
            footer.Font = ContentFont;
            dd.Add(footer);




           





            dd.Close();

        }
        #endregion
        private void button2_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy == true)
            {
                MessageBox.Show("CSV file exporting is still in progress", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            CheckForIllegalCrossThreadCalls = false;
            dt.Rows.Clear();
            try
            {
                if(!string.IsNullOrWhiteSpace(RegP.Text))
                {
                    AmortizationTable();
                    countingP.Text = "No.of payments: " + gunaDataGridView1.Rows.Count.ToString();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy == true)
            {
                MessageBox.Show("CSV file exporting is still in progress", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if(gunaDataGridView1.Rows.Count!=0)
            {
                ShowingSummary1();
                viewsummary();
            }
        }

        private void Lamt_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 46 && ch != 8)
            {
                e.Handled = true;
            }
            if(Lamt.Text.Contains("."))
            {
                if(ch=='.')
                {
                    e.Handled = true;
                }
            }
           
        }

        private void Lamt_Leave(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Lamt.Text))
            {
                double entered = double.Parse(Lamt.Text);
                Lamt.Text = entered.ToString("c");
            }
        }

        private void Lamt_Enter(object sender, EventArgs e)
        {
            if(!string.IsNullOrWhiteSpace(Lamt.Text))
            {
                Lamt.Clear();
            }
        }

        private void ExtraP_Leave(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(ExtraP.Text))
            {
                double entered = double.Parse(ExtraP.Text);
                ExtraP.Text = entered.ToString("c");
            }
        }

        private void ExtraP_Enter(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(ExtraP.Text))
            {
                ExtraP.Clear();
            }
        }

        private void ExtraP_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 46 && ch != 8)
            {
                e.Handled = true;
            }
            if (ExtraP.Text.Contains("."))
            {
                if (ch == '.')
                {
                    e.Handled = true;
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(gunaDataGridView1.Rows.Count!=0)
            {
                button1.Text = "Recalculate Regular Payment";
                button2.Text = "Re-Load Schedule";

            }
            else if(gunaDataGridView1.Rows.Count==0)
            {
                button1.Text = "Calculate Regular Payment";
                button2.Text = "Load Schedule";
            }
            
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (gunaDataGridView1.Rows.Count != 0&&rT.Visible==true)
            {
                button5.Text = "Update Summary";
            }

            else if (gunaDataGridView1.Rows.Count == 0)
            {
                button5.Text = "View Summary";
            }

        }

        private void backgroundWorker1_DoWork_1(object sender, DoWorkEventArgs e)
        {
            try
            {
                exportingCSV(save.FileName);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (!(e.Error == null))
                {
                    MessageBox.Show(e.Error.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                {
                    MessageBox.Show("Payment Schedule exported successfully", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (backgroundWorker1.IsBusy == true)
            {
                DialogResult r = MessageBox.Show("CSV file exporting is still in progress\nDo you want to exit the application anyway?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
                if (r == DialogResult.Yes)
                {
                    e.Cancel = false;
                }
                else
                {
                    e.Cancel =true;
                }

            }
        }

        private void gunaPictureBox3_Click(object sender, EventArgs e)
        {
            Form aboutfrom = Application.OpenForms["AboutDev"];
            if(aboutfrom==null)
            {
                AboutDev dev = new AboutDev();
                dev.Show();
            }
            else
            {
                aboutfrom.BringToFront();
            }
           
        }
        private void gunaPictureBox4_Click(object sender, EventArgs e)
        {
            information info = new information();
            info.ShowDialog();
        }

        private void gunaPictureBox5_Click(object sender, EventArgs e)
        {
            if(WindowState==FormWindowState.Maximized)
            {
                WindowState = FormWindowState.Normal;
                gunaPictureBox5.Image = Properties.Resources.max;
                toolTip1.SetToolTip(gunaPictureBox5, "Maximize");
            }
            else if(WindowState==FormWindowState.Normal)
            {
                WindowState = FormWindowState.Maximized;
                gunaPictureBox5.Image = Properties.Resources.normal;
                toolTip1.SetToolTip(gunaPictureBox5, "Restore");
            }
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
             if(WindowState==FormWindowState.Normal)
            {
                gunaPictureBox5.Image = Properties.Resources.max;
                
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy == true)
            {
                MessageBox.Show("CSV file exporting is still in progress", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if(gunaDataGridView1.Rows.Count!=0)
            {
                string path = "";
                if(st.IsusingDefault==true)
                {
                   path= Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Calculate Your Loan\\Loan Reports";
               
                }
                else
                {
                    path = st.ChooseDirectory;
                }
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string filename1 = string.Empty;
                if(string.IsNullOrWhiteSpace(ExtraP.Text))
                {
                  filename1=  path + "\\A " + Ltime.Text + " Year Loan [ " + Lamt.Text + " ]" + " With Payment Frequency [ " + gunaComboBox1.Text + " ]"+" With No Extra Payment" + ".pdf";
                }
                else
                {
                    filename1 = path + "\\A " + Ltime.Text + " Year Loan [ " + Lamt.Text + " ]" + " With Payment Frequency [ " + gunaComboBox1.Text + " ]" + " With Extra Payment of "+ExtraP.Text+".pdf";
             
                }
               
                try
                {
                    button1_Click(button6, e);
                    button2_Click(button6,e);

                    ShowingSummary1();
                    ExportingOutputPDF(filename1);
                    System.Diagnostics.Process.Start(filename1);
                   
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);

                }
            }
            else
            {
                MessageBox.Show("Nothing to view", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
       
            }
        }

        private void gunaPictureBox6_Click(object sender, EventArgs e)
        {
            UserSettings set = new UserSettings();
            set.ShowDialog();
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {
            closingAnyOpened();
        }

       





    }
}

