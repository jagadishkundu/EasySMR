using System;
using System.IO;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Threading;

namespace EasySMR
{
    public partial class Form1 : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\smrall.mdf; Integrated Security = True;Connect Timeout=30");

        public Form1()
        {
            Thread splash = new Thread(new ThreadStart(StartForm));
            splash.Start();
            Thread.Sleep(5000);
            InitializeComponent();
            splash.Abort();
            chart1.Series[0].IsVisibleInLegend = false;
            chart2.Series[0].IsVisibleInLegend = false;
            chart2.Titles.Clear();
            tableLayoutPanel12.Visible = false;
            tableLayoutPanel13.Visible = false;
            chart1.ChartAreas[0].AxisX.Minimum = 0;
            chart1.ChartAreas[0].AxisX.Maximum = 100;
        }

        public void StartForm()
        {
            Application.Run(new frmSplashScreen());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count < 2)
            {
                MessageBox.Show("Please enter joint data", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            chart1.Series[0].IsVisibleInLegend = true;
            chart2.Series[0].IsVisibleInLegend = true;
            tableLayoutPanel12.Visible = true;
            tableLayoutPanel13.Visible = true;

            con.Open();
            string sqlTrunc = "TRUNCATE TABLE " + "smrtable";
            SqlCommand comd4 = new SqlCommand(sqlTrunc, con);
            comd4.CommandType = CommandType.Text;
            comd4.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTrunc1 = "TRUNCATE TABLE " + "smrp";
            SqlCommand comd1 = new SqlCommand(sqlTrunc1, con);
            comd1.CommandType = CommandType.Text;
            comd1.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTrunc2 = "TRUNCATE TABLE " + "smrt";
            SqlCommand comd2 = new SqlCommand(sqlTrunc2, con);
            comd2.CommandType = CommandType.Text;
            comd2.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTrunc3 = "TRUNCATE TABLE " + "smrw";
            SqlCommand comd3 = new SqlCommand(sqlTrunc3, con);
            comd3.CommandType = CommandType.Text;
            comd3.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTrunca = "TRUNCATE TABLE " + "pclass";
            SqlCommand comda = new SqlCommand(sqlTrunca, con);
            comda.CommandType = CommandType.Text;
            comda.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTruncb = "TRUNCATE TABLE " + "wclass";
            SqlCommand comdb = new SqlCommand(sqlTruncb, con);
            comdb.CommandType = CommandType.Text;
            comdb.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTruncc = "TRUNCATE TABLE " + "tclass";
            SqlCommand comdc = new SqlCommand(sqlTruncc, con);
            comdc.CommandType = CommandType.Text;
            comdc.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTruncd = "TRUNCATE TABLE " + "aclass";
            SqlCommand comdd = new SqlCommand(sqlTruncd, con);
            comdd.CommandType = CommandType.Text;
            comdd.ExecuteNonQuery();
            con.Close();

            do
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    try
                    {
                        dataGridView1.Rows.Remove(row);
                    }
                    catch (Exception) { }
                }
            } while (dataGridView1.Rows.Count > 1);

            do
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    try
                    {
                        dataGridView2.Rows.Remove(row);
                    }
                    catch (Exception) { }
                }
            } while (dataGridView2.Rows.Count > 1);

            double rmrb = Convert.ToDouble(numericUpDown5.Text); // basic RMR
            double psi = Convert.ToDouble(numericUpDown1.Text); // SLope Angle
            double sd = Convert.ToDouble(numericUpDown2.Text); //Slope Direction
            double phi = Convert.ToDouble(numericUpDown3.Text); //Angle of friction
            double psi2 = Convert.ToDouble(numericUpDown3.Text); // Slope angle another variable
            double f4 = Convert.ToDouble(numericUpDown4.Text); // Adustment for excavation
            double pfc = 0; //planar failure count
            double wfc = 0; //wedge failure count
            double ftc = 0; // topple failure count
            double maxda = 0;
            double pll = Convert.ToDouble(numericUpDown6.Text); // Planar lateral limit
            double tll = Convert.ToDouble(numericUpDown7.Text); // topple lateral limit

            int rc = dataGridView3.Rows.Count - 1;  // Row count for planes
            int rfwc = (rc * (rc - 1)) / 2; // rows for wedges count
            double[] a = new double[dataGridView3.Rows.Count]; //
            double[] b = new double[dataGridView3.Rows.Count]; //
            double[] daw = new double[rfwc]; //
            double[] dap = new double[dataGridView3.Rows.Count]; //
            double[] dat = new double[dataGridView3.Rows.Count]; //
            double[] wsmr = new double[rfwc]; // Single array to store wedge smr value
            double[] psmr = new double[dataGridView3.Rows.Count]; // Single array to store planar smr value
            double[] tsmr = new double[dataGridView3.Rows.Count]; // Single array to store topple smr value

            double[] waf = new double[rfwc];    // Adjustment factor for all wedges
            double[] paf = new double[rc];       // Adjustment factors for all planes
            double[] taf = new double[rc];      // Adjustment factor for all topples
            int tsmrc = rfwc + (2*dataGridView3.Rows.Count) - 2;     // count for total number of failure elements
            double[] asmr = new double[tsmrc]; // Single array to store all smr value

            int wc = 0;     // Wedge count
            int pc = 0;    //plane count
            int tc = 0;    //topple count

            for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
                a[i] = (Math.PI / 180) * (Convert.ToInt32(dataGridView3.Rows[i].Cells[2].Value) - 90);     // Dip direction
                b[i] = (Math.PI / 180) * Convert.ToInt32(dataGridView3.Rows[i].Cells[1].Value); // Dip amount
            }

            for (int i = 0; i < dataGridView3.Rows.Count - 2; i++)
            {
                for (int k = i + 1; k < dataGridView3.Rows.Count - 1; k++)
                {
                    double T1 = (Math.Cos(b[i]) * Math.Sin(a[k]) * Math.Sin(b[k])) - (Math.Sin(a[i]) * Math.Sin(b[i]) * Math.Cos(b[k]));
                    double T2 = (Math.Cos(b[i]) * Math.Cos(a[k]) * Math.Sin(b[k])) - (Math.Cos(a[i]) * Math.Sin(b[i]) * Math.Cos(b[k]));
                    double T3 = (Math.Sin(a[i]) * Math.Sin(b[i]) * Math.Sin(b[k]) * Math.Cos(a[k])) - (Math.Cos(a[i]) * Math.Sin(b[i]) * Math.Sin(a[k]) * Math.Sin(b[k]));

                    double p = Math.Asin(Math.Abs(T3) / Math.Sqrt(Math.Pow(T1, 2) + Math.Pow(T2, 2) + Math.Pow(T3, 2)));
                    double plunge = p * (180 / Math.PI);
                    double t = Math.Atan(T1 / T2);
                    double trend = t * (180 / Math.PI);
                    daw[wc] = 0;
                    wsmr[wc] = rmrb;     
                    psmr[pc] = rmrb;       
                    tsmr[tc] = rmrb;       


                    if (T1 > 0 && T2 < 0 && T3 < 0 || T1 < 0 && T2 > 0 && T3 > 0)
                        trend = 180 + trend;

                    if (T1 > 0 && T2 > 0 && T3 > 0 || T1 < 0 && T2 < 0 && T3 < 0)
                        trend = 180 + trend;

                    if (T1 > 0 && T2 < 0 && T3 > 0 || T1 < 0 && T2 > 0 && T3 < 0)
                        trend = 360 + trend;

                    double wdiff = Math.Abs(sd - trend); // difference between slope direction and intersection trend
                    if (wdiff > 180)
                        wdiff = 360 - wdiff;
                    double apdip = (180 / Math.PI) * (Math.Atan(Math.Cos((Math.PI / 180) * wdiff) * Math.Tan((Math.PI / 180) * psi))); // apparent dip

                    double osd1 = 0;        
                    double osd2 = 0;        
                    double ot1 = 0;         
                    double ot2 = 0;         

                    double ddj1 = (180 / Math.PI) * a[i] + 90;      
                    double ddj2 = (180 / Math.PI) * a[k] + 90;      

                    double dif1 = ddj1 - sd;        
                    if (dif1 < -180)
                        dif1 = dif1 + 360;
                    else if (dif1 > 180)
                        dif1 = dif1 - 360;
                    if (dif1 < 0)
                        osd1 = -1;
                    else osd1 = 1;

                    double dif2 = ddj2 - sd;
                    if (dif2 < -180)
                        dif2 = dif2 + 360;
                    else if (dif2 > 180)
                        dif2 = dif2 - 360;
                    if (dif2 < 0)
                        osd2 = -1;
                    else osd2 = 1;

                    double dif3 = ddj1 - trend;
                    if (dif3 < -180)
                        dif3 = dif3 + 360;
                    else if (dif3 > 180)
                        dif3 = dif3 - 360;
                    if (dif3 < 0)
                        ot1 = -1;
                    else ot1 = 1;

                    double dif4 = ddj2 - trend;
                    if (dif4 < -180)
                        dif4 = dif4 + 360;
                    else if (dif4 > 180)
                        dif4 = dif4 - 360;
                    if (dif4 < 0)
                        ot2 = -1;
                    else ot2 = 1;

                    double dipi = (180 / Math.PI) * b[i];
                    double dipk = (180 / Math.PI) * b[k];
                    double dipdi = Convert.ToInt32(dataGridView3.Rows[i].Cells[2].Value);
                    double dipdk = Convert.ToInt32(dataGridView3.Rows[k].Cells[2].Value);

                    double diffi = Math.Abs(dipdi - sd);
                    if (diffi > 180)
                        diffi = 360 - diffi;

                    double diffk = Math.Abs(dipdk - sd);
                    if (diffk > 180)
                        diffk = 360 - diffk;

                    double F1spi = 0.64 - 0.006 * (180 / Math.PI) * Math.Atan(0.1 * (diffi - 17));
                    double F2spi = 0.5625 + 0.0051282 * (180 / Math.PI) * Math.Atan(0.17 * dipi - 5);
                    double F3spi = -30 + 0.33333333 * (180 / Math.PI) * Math.Atan(dipi - psi);

                    double F1w = 0.64 - 0.006 * (180 / Math.PI) * Math.Atan(0.1 * (wdiff - 17));
                    double F2w = 0.5625 + 0.0051282 * (180 / Math.PI) * Math.Atan(0.17 * plunge - 5);
                    double F3w = -30 + 0.33333333 * (180 / Math.PI) * Math.Atan(plunge - apdip);

                    double F1spk = 0.64 - 0.006 * (180 / Math.PI) * Math.Atan(0.1 * (diffk - 17));
                    double F2spk = 0.5625 + 0.0051282 * (180 / Math.PI) * Math.Atan(0.17 * dipk - 5);
                    double F3spk = -30 + 0.33333333 * (180 / Math.PI) * Math.Atan(dipk - psi);

                    if (wdiff < 90 && apdip >= plunge && plunge >= phi)
                    {
                        wfc = wfc + 1;

                        if (osd1 != ot1)         
                        {
                            double planaraf = F1spi * F2spi * F3spi;    // plaar adjustment factor
                            double wedgeaf = F1w * F2w * F3w;    //wedge adjustment factor
                            if (planaraf < wedgeaf)
                            {
                                waf[wc] = planaraf;
                                daw[wc] = planaraf + f4;
                            }
                            else
                            {
                                waf[wc] = wedgeaf;
                                daw[wc] = wedgeaf + f4;
                            }
                            wsmr[wc] = rmrb + daw[wc];
                            dataGridView2.Rows.Add("", i + 1, k + 1, trend, plunge, "Primary single", i + 1, dipdi, wsmr[wc]);

                            if (daw[wc] < maxda)
                                maxda = daw[wc];
                            wc++;
                        }

                        else if (osd2 != ot2)        
                        {
                            double planaraf = F1spk * F2spk * F3spk;    
                            double wedgeaf = F1w * F2w * F3w;    
                            if (planaraf < wedgeaf)
                            {
                                waf[wc] = planaraf;
                                daw[wc] = planaraf + f4;
                            }
                            else
                            {
                                waf[wc] = wedgeaf;
                                daw[wc] = wedgeaf + f4;
                            }

                            wsmr[wc] = rmrb + daw[wc];
                            dataGridView2.Rows.Add("", i + 1, k + 1, Math.Round(trend,2), Math.Round(plunge,2), "Primary single", k + 1, dipdk, Math.Round(wsmr[wc]));

                            if (daw[wc] < maxda)
                                maxda = daw[wc];
                            wc++;
                        }

                        else     
                        {
                            waf[wc] = F1w * F2w * F3w;
                            daw[wc] = F1w * F2w * F3w + f4;
                            wsmr[wc] = rmrb + daw[wc];
                            dataGridView2.Rows.Add("", i + 1, k + 1, Math.Round(trend,2), Math.Round(plunge,2), "Primary double", "both", Math.Round(trend,2), Math.Round(wsmr[wc],2));

                            if (daw[wc] < maxda)
                                maxda = daw[wc];
                            wc++;
                        }

                    }
                    else
                    {
                        if (osd1 != ot1)         
                        {
                            double planaraf = F1spi * F2spi * F3spi;    
                            double wedgeaf = F1w * F2w * F3w;    
                            if (planaraf < wedgeaf)
                            {
                                waf[wc] = planaraf;
                                daw[wc] = planaraf + f4;
                            }
                            else
                            {
                                waf[wc] = wedgeaf;
                                daw[wc] = wedgeaf + f4;
                            }
                            wsmr[wc] = rmrb + daw[wc];
                        }

                        else if (osd2 != ot2)        
                        {
                            double planaraf = F1spk * F2spk * F3spk;    
                            double wedgeaf = F1w * F2w * F3w;    
                            if (planaraf < wedgeaf)
                            {
                                waf[wc] = planaraf;
                                daw[wc] = planaraf + f4;
                            }
                            else
                            {
                                waf[wc] = wedgeaf;
                                daw[wc] = wedgeaf + f4;
                            }

                            wsmr[wc] = rmrb + daw[wc];
                        }

                        dataGridView2.Rows.Add("", i + 1, k + 1, Math.Round(trend,2), Math.Round(plunge,2), "No", "NA", "", Math.Round(wsmr[wc],2));
                        wc++;
                    }

                }

            }

            for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
                dataGridView1.Rows.Add();
                double dip = Convert.ToInt32(dataGridView3.Rows[i].Cells[1].Value);      
                double dipd = Convert.ToInt32(dataGridView3.Rows[i].Cells[2].Value);
                double diff = Math.Abs(dipd - sd);
                dap[i] = 0;     
                dat[i] = 0;      

                if (diff > 180)
                    diff = 360 - diff;

                double apdipi = (180 / Math.PI) * (Math.Atan(Math.Cos((Math.PI / 180) * diff) * Math.Tan((Math.PI / 180) * psi)));

                double F1p = 0.64 - 0.006 * (180 / Math.PI) * Math.Atan(0.1 * (diff - 17));
                double F2p = 0.5625 + 0.0051282 * (180 / Math.PI) * Math.Atan(0.17 * dip - 5);
                double F3p = -30 + 0.33333333 * (180 / Math.PI) * Math.Atan(dip - apdipi);

                double F1t = 0.64 - 0.006 * (180 / Math.PI) * Math.Atan(0.1 * (Math.Abs(diff - 180) - 17));
                double F2t = 0.5625 + 0.0051282 * (180 / Math.PI) * Math.Atan(0.17 * dip - 5);
                double F3t = -13 - 0.142857 * (180 / Math.PI) * Math.Atan(dip + apdipi - 120);

                if (diff <= pll && phi <= dip && dip < apdipi)
                {
                    pfc = pfc + 1;
                    paf[i] = F1p * F2p * F3p;
                    dap[i] = F1p * F2p * F3p + f4;
                    psmr[i] = rmrb + dap[i];

                    dataGridView1.Rows[i].Cells[1].Value = "Yes";
                    dataGridView1.Rows[i].Cells[2].Value = Math.Round(psmr[i],2);

                    if (dap[i] < maxda) 
                        maxda = dap[i];
                }
                else
                {
                    paf[i] = F1p * F2p * F3p;
                    dap[i] = F1p * F2p * F3p + f4;
                    psmr[i] = rmrb + dap[i];
                    dataGridView1.Rows[i].Cells[1].Value = "No";
                    dataGridView1.Rows[i].Cells[2].Value = Math.Round(psmr[i],2);
                }

                diff = Math.Abs(dipd - sd);
                diff = Math.Abs(180 - diff);
                apdipi = (180 / Math.PI) * (Math.Atan(Math.Cos((Math.PI / 180) * diff) * Math.Tan((Math.PI / 180) * psi)));

                if (diff <= tll && (90 - apdipi) + phi < dip)
                {
                    ftc = ftc + 1;
                    taf[i] = F1t * F2t * F3t;
                    dat[i] = F1t * F2t * F3t + f4;
                    tsmr[i] = rmrb + dat[i];
                    dataGridView1.Rows[i].Cells[3].Value = "Yes";
                    dataGridView1.Rows[i].Cells[4].Value = Math.Round(tsmr[i],2);

                    if (dat[i] < maxda)
                        maxda = dat[i];
                }
                else
                {
                    taf[i] = F1t * F2t * F3t;
                    dat[i] = F1t * F2t * F3t + f4;
                    tsmr[i] = rmrb + dat[i];
                    dataGridView1.Rows[i].Cells[3].Value = "No";
                    dataGridView1.Rows[i].Cells[4].Value = Math.Round(tsmr[i],2);

                }

            }
            label13.Text = Convert.ToString(pfc);
            label14.Text = Convert.ToString(wfc);
            label15.Text = Convert.ToString(ftc);
            label12.Text = Convert.ToString(pfc + wfc + ftc);
            double fsmr = rmrb + maxda;
            fsmr = Math.Round(fsmr, 1);         

            con.Open();
            int cls = 6;
            for (double j = 0; j <= 80; j = j + 20)
            {
                int fc = 0;
                int fcp = 0;
                int fct = 0;


                for (wc = 0; wc < rfwc; wc++)
                {
                    double wvalue = wsmr[wc];
                    if (wvalue < 0)
                        wvalue = 0;
                    if (wvalue >= j && wvalue < j + 20)
                    {
                        fc = fc + 1;
                    }

                }

                for (pc = 0; pc < dataGridView3.Rows.Count - 1; pc++)
                {
                    double pvalue = psmr[pc];
                    if (pvalue < 0)
                        pvalue = 0;
                    if (pvalue >= j && pvalue < j + 20)
                    {
                        fcp = fcp + 1;
                    }
                }

                for (pc = 0; pc < dataGridView3.Rows.Count - 1; pc++)
                {
                    double tvalue = tsmr[pc];
                    if (tvalue < 0)
                        tvalue = 0;

                    if (tvalue >= j && tvalue < j + 20)
                    {
                        fct = fct + 1;
                    }
                }

                double afc = fc + fcp + fct;
                cls--;
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "insert into wclass values(" + cls + "," + fc + "); insert into pclass values(" + cls + "," + fcp + "); insert into tclass values(" + cls + "," + fct + "); insert into aclass values(" + cls + "," + afc + ")";
                cmd.ExecuteNonQuery();

            }
            con.Close();


            con.Open();

            for (double j = 0; j <= 95; j = j + 5)
            {
                int fc = 0;
                int fcp = 0;
                int fct = 0;
                int smrcount = 0;

                for (wc = 0; wc < rfwc; wc++)
                {
                    asmr[smrcount] = wsmr[wc];

                    double wvalue = wsmr[wc];
                    if (wvalue < 0)
                        wvalue = 0;
                    if (wvalue >= j && wvalue < j + 5)
                    {
                        fc = fc + 1;
                    }
                    smrcount++;
                }

                for (pc = 0; pc < dataGridView3.Rows.Count - 1; pc++)
                {
                    asmr[smrcount] = psmr[pc];

                    double pvalue = psmr[pc];
                    if (pvalue < 0)
                        pvalue = 0;
                    if (pvalue >= j && pvalue < j + 5)
                    {
                        fcp = fcp + 1;
                    }
                    smrcount++;
                }

                for (pc = 0; pc < dataGridView3.Rows.Count - 1; pc++)
                {
                    asmr[smrcount] = tsmr[pc];

                    double tvalue = tsmr[pc];
                    if (tvalue < 0)
                        tvalue = 0;

                    if (tvalue >= j && tvalue < j + 5)
                    {
                        fct = fct + 1;
                    }
                    smrcount++;
                }

                double afc = fc + fcp + fct;
                double midp = j + 2.5;
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "insert into smrw values(" + midp + "," + fc + "); insert into smrp values(" + midp + "," + fcp + "); insert into smrt values(" + midp + "," + fct + "); insert into smrtable values(" + midp + "," + afc + ")";
                cmd.ExecuteNonQuery();
            }
            con.Close();
                        
            Allsmr();
            Allpie();
            chart2.Titles.Clear();
            chart2.Titles.Add("All Failure mode");
            chart1.Series[0].LegendText = "All Failure";

            for (int i =0; i< tsmrc; i++)

                for(int j = i+1; j < tsmrc; j++)
                {
                    double temp;
                    if (asmr[i] > asmr[j])
                    {
                        temp = asmr[i];
                        asmr[i] = asmr[j];
                        asmr[j] = temp;
                    }                       

                }

            double percentage = Convert.ToDouble(numericUpDown8.Text);
            double nsmr = tsmrc * (percentage / 100);
            nsmr = Math.Round(nsmr);
            if (nsmr < 1)
                nsmr = 1;

            double sum = 0;
            for (int i =0; i < nsmr; i++)
            {
                sum = sum + asmr[i];
            }

            double averagesmr = sum / nsmr;
            averagesmr = Math.Round(averagesmr, 1);
            label11.Text = Convert.ToString(averagesmr);

            double orv = 0;
            double swsmr = 0;
            double spsmr = 0;
            double stsmr = 0;

            double swaf = 0;
            double spaf = 0;
            double staf = 0; 

            double okv = 0; 
            double wkv = 0; 
            double pkv = 0; 
            double tkv = 0; 

            for (int i = 0; i < rfwc; i++)
            {
                swsmr = swsmr + wsmr[i];
                swaf = swaf + waf[i];
            }

            for (int i = 0; i < rc; i++)
            {
                spsmr = spsmr + psmr[i];
                stsmr = stsmr + tsmr[i];
                spaf = spaf + paf[i];
                staf = staf + taf[i];
            }

            orv = 100 - ((swsmr + spsmr + stsmr) / (rfwc + (2 * rc)));
            pkv = (spaf / (rc * 60)) * -100;
            tkv = (staf / (rc * 60)) * -100;
            wkv = (swaf / (rfwc * 60)) * -100;
            okv = ((spaf + staf + swaf) / ((rfwc + rc) * 60)) * -100;

            orv = Math.Round(orv, 1); // Overall Susceptibility
            pkv = Math.Round(pkv, 1); // Planar kinematic susceptibility
            tkv = Math.Round(tkv, 1); // topple kinematic susceptibility
            wkv = Math.Round(wkv, 1); // wedge kinematic susceptibility
            okv = Math.Round(okv, 1); // overall kinematic susceptibility

            label27.Text = Convert.ToString(orv) + "%";
            label28.Text = Convert.ToString(pkv) + "%";
            label29.Text = Convert.ToString(wkv) + "%";
            label30.Text = Convert.ToString(tkv) + "%";
            label31.Text = Convert.ToString(okv) + "%";
        }

        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;

        string excel_name = string.Empty;
        String Dip, Direction;
        public void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "File Excel|*.xlsx;*.xls";

            DialogResult re = fd.ShowDialog();
            excel_name = fd.SafeFileName;
            if (re == DialogResult.OK)
            {
                string fileName = fd.FileName;
                textBox11.Text = fileName;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(textBox11.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                comboBox1.Items.Clear();
                foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
                {
                    comboBox1.Items.Add(sheet.Name);
                }

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();
            }


        }
        public void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.dataGridView3.DataSource = null;
            this.dataGridView3.Rows.Clear();

            string selected = this.comboBox1.GetItemText(this.comboBox1.SelectedItem);
            ReadExcelFile(selected, textBox11.Text);
        }

        void ReadExcelFile(string sheetName, string path)
        {

            using (OleDbConnection conn = new OleDbConnection())
            {
                DataSet dt = new DataSet();
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                if (fileExtension == ".xls")
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                if (fileExtension == ".xlsx")
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                    }
                }

                int rCnt;
                int rw = 0;

                rw = dt.Tables[0].Rows.Count;

                for (rCnt = 1; rCnt <= rw; rCnt++)
                {
                    Dip = dt.Tables[0].Rows[rCnt - 1].ItemArray[0].ToString();
                    Direction = dt.Tables[0].Rows[rCnt - 1].ItemArray[1].ToString();

                    dataGridView3.Rows.Add("", Dip, Direction);
                }
            }
        }


        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
        }

        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView2.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
        }

        private void dataGridView3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;

            if (!char.IsDigit(ch) && ch != 8 && ch != 46 && ch != 13)
            {
                e.Handled = true;
            }
        }

        private void dataGridView3_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            try
            {
                e.Control.KeyPress +=
                new KeyPressEventHandler(dataGridView3_KeyPress);
            }

            catch (Exception)
            {
            }
        }

        private void dataGridView3_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView3.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Planarchart();
            Planarpie();
            chart1.Series[0].LegendText = "Planar";
            chart2.Titles.Clear();
            chart2.Titles.Add("Planar Failure mode");           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Topplechart();
            Topplepie();
            chart1.Series[0].LegendText = "Topple";
            chart2.Titles.Clear();
            chart2.Titles.Add("Topple Failure mode");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Wedgechart();
            Wedgepie();
            chart1.Series[0].LegendText = "Wedge";
            chart2.Titles.Clear();
            chart2.Titles.Add("Wedge Failure mode");
        }

        public void Planarchart()
        {
            foreach (var Series1 in chart1.Series)
            {
                Series1.Points.Clear();
            }
            con.Open();
            SqlCommand cmdata = con.CreateCommand();
            cmdata.CommandType = CommandType.Text;
            cmdata.CommandText = "select* from smrp";
            SqlDataReader myreader;
            myreader = cmdata.ExecuteReader();

            while (myreader.Read())
            {
                this.chart1.Series["Series1"].Points.AddXY(myreader["Mid Point"], myreader["Frequency"]);
            }
            myreader.Close();
            con.Close();
        }

        public void Topplechart()
        {
            foreach (var Series1 in chart1.Series)
            {
                Series1.Points.Clear();
            }
            con.Open();
            SqlCommand cmdata = con.CreateCommand();
            cmdata.CommandType = CommandType.Text;
            cmdata.CommandText = "select* from smrt";
            SqlDataReader myreader;
            myreader = cmdata.ExecuteReader();

            while (myreader.Read())
            {
                this.chart1.Series["Series1"].Points.AddXY(myreader["Mid Point"], myreader["Frequency"]);
            }
            myreader.Close();
            con.Close();
        }

        public void Wedgechart()
        {
            foreach (var Series1 in chart1.Series)
            {
                Series1.Points.Clear();
            }
            con.Open();
            SqlCommand cmdata = con.CreateCommand();
            cmdata.CommandType = CommandType.Text;
            cmdata.CommandText = "select* from smrw";
            SqlDataReader myreader;
            myreader = cmdata.ExecuteReader();

            while (myreader.Read())
            {
                this.chart1.Series["Series1"].Points.AddXY(myreader["Mid Point"], myreader["Frequency"]);
            }
            myreader.Close();
            con.Close();
        }

        public void Allsmr()
        {
            foreach (var Series1 in chart1.Series)
            {
                Series1.Points.Clear();
            }
            con.Open();
            SqlCommand cmdata = con.CreateCommand();
            cmdata.CommandType = CommandType.Text;
            cmdata.CommandText = "select* from smrtable";
            SqlDataReader myreader;
            myreader = cmdata.ExecuteReader();

            while (myreader.Read())
            {
                this.chart1.Series["Series1"].Points.AddXY(myreader["Mid Point"], myreader["Frequency"]);
            }
            myreader.Close();
            con.Close();
        }

        public void Allpie()
        {
            foreach (var Series1 in chart2.Series)
            {
                Series1.Points.Clear();
            }
            con.Open();
            SqlCommand cmdata = con.CreateCommand();
            cmdata.CommandType = CommandType.Text;
            cmdata.CommandText = "select* from aclass";
            SqlDataReader myreader;
            myreader = cmdata.ExecuteReader();

            while (myreader.Read())
            {
                this.chart2.Series["Series1"].Points.AddXY(myreader["Class"], myreader["Frequency"]);
            }
            myreader.Close();
            con.Close();
        }

        public void Planarpie()
        {
            foreach (var Series1 in chart2.Series)
            {
                Series1.Points.Clear();
            }
            con.Open();
            SqlCommand cmdata = con.CreateCommand();
            cmdata.CommandType = CommandType.Text;
            cmdata.CommandText = "select* from pclass";
            SqlDataReader myreader;
            myreader = cmdata.ExecuteReader();

            while (myreader.Read())
            {
                this.chart2.Series["Series1"].Points.AddXY(myreader["Class"], myreader["Frequency"]);
            }
            myreader.Close();
            con.Close();
        }

        public void Wedgepie()
        {
            foreach (var Series1 in chart2.Series)
            {
                Series1.Points.Clear();
            }
            con.Open();
            SqlCommand cmdata = con.CreateCommand();
            cmdata.CommandType = CommandType.Text;
            cmdata.CommandText = "select* from wclass";
            SqlDataReader myreader;
            myreader = cmdata.ExecuteReader();

            while (myreader.Read())
            {
                this.chart2.Series["Series1"].Points.AddXY(myreader["Class"], myreader["Frequency"]);
            }
            myreader.Close();
            con.Close();
        }

        public void Topplepie()
        {
            foreach (var Series1 in chart2.Series)
            {
                Series1.Points.Clear();
            }
            con.Open();
            SqlCommand cmdata = con.CreateCommand();
            cmdata.CommandType = CommandType.Text;
            cmdata.CommandText = "select* from tclass";
            SqlDataReader myreader;
            myreader = cmdata.ExecuteReader();

            while (myreader.Read())
            {
                this.chart2.Series["Series1"].Points.AddXY(myreader["Class"], myreader["Frequency"]);
            }
            myreader.Close();
            con.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            export();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            con.Open();
            string sqlTrunc = "TRUNCATE TABLE " + "smrtable";
            SqlCommand comd4 = new SqlCommand(sqlTrunc, con);
            comd4.CommandType = CommandType.Text;
            comd4.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTrunc1 = "TRUNCATE TABLE " + "smrp";
            SqlCommand comd1 = new SqlCommand(sqlTrunc1, con);
            comd1.CommandType = CommandType.Text;
            comd1.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTrunc2 = "TRUNCATE TABLE " + "smrt";
            SqlCommand comd2 = new SqlCommand(sqlTrunc2, con);
            comd2.CommandType = CommandType.Text;
            comd2.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTrunc3 = "TRUNCATE TABLE " + "smrw";
            SqlCommand comd3 = new SqlCommand(sqlTrunc3, con);
            comd3.CommandType = CommandType.Text;
            comd3.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTrunca = "TRUNCATE TABLE " + "pclass";
            SqlCommand comda = new SqlCommand(sqlTrunca, con);
            comda.CommandType = CommandType.Text;
            comda.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTruncb = "TRUNCATE TABLE " + "wclass";
            SqlCommand comdb = new SqlCommand(sqlTruncb, con);
            comdb.CommandType = CommandType.Text;
            comdb.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTruncc = "TRUNCATE TABLE " + "tclass";
            SqlCommand comdc = new SqlCommand(sqlTruncc, con);
            comdc.CommandType = CommandType.Text;
            comdc.ExecuteNonQuery();
            con.Close();

            con.Open();
            string sqlTruncd = "TRUNCATE TABLE " + "aclass";
            SqlCommand comdd = new SqlCommand(sqlTruncd, con);
            comdd.CommandType = CommandType.Text;
            comdd.ExecuteNonQuery();
            con.Close();

            do
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    try
                    {
                        dataGridView1.Rows.Remove(row);
                    }
                    catch (Exception) { }
                }
            } while (dataGridView1.Rows.Count > 1);

            do
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    try
                    {
                        dataGridView2.Rows.Remove(row);
                    }
                    catch (Exception) { }
                }
            } while (dataGridView2.Rows.Count > 1);

            do
            {
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    try
                    {
                        dataGridView3.Rows.Remove(row);
                    }
                    catch (Exception) { }
                }
            } while (dataGridView3.Rows.Count > 1);

            numericUpDown1.Value = 0;
            numericUpDown2.Value = 0;
            numericUpDown3.Value = 0;
            numericUpDown4.Value = 0;
            numericUpDown5.Value = 0;
            numericUpDown6.Value = 20;
            numericUpDown7.Value = 10;
            numericUpDown8.Value = 20;

            textBox1.Text = "";
            label11.Text = "...";
            label12.Text = "...";
            label13.Text = "...";
            label14.Text = "...";
            label15.Text = "...";
            label27.Text = "...";
            label28.Text = "...";
            label29.Text = "...";
            label30.Text = "...";
            label31.Text = "...";

            chart1.Series[0].IsVisibleInLegend = false;
            chart2.Series[0].IsVisibleInLegend = false;
            chart2.Titles.Clear();
            tableLayoutPanel12.Visible = false;
            tableLayoutPanel13.Visible = false;

            Allsmr();
            Allpie();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            var SavefileDialog1 = new SaveFileDialog();
            SavefileDialog1.FileName = "FrequencyGraph_";
            SavefileDialog1.DefaultExt = ".jpeg";

            if (SavefileDialog1.ShowDialog() == DialogResult.OK)
            {
                chart1.SaveImage(SavefileDialog1.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            var SavefileDialog2 = new SaveFileDialog();
            SavefileDialog2.FileName = "FrequencyGraph_";
            SavefileDialog2.DefaultExt = ".png";

            if (SavefileDialog2.ShowDialog() == DialogResult.OK)
            {
                chart1.SaveImage(SavefileDialog2.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            var SavefileDialog3 = new SaveFileDialog();
            SavefileDialog3.FileName = "FrequencyGraph_";
            SavefileDialog3.DefaultExt = ".tiff";

            if (SavefileDialog3.ShowDialog() == DialogResult.OK)
            {
                chart1.SaveImage(SavefileDialog3.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Tiff);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            var SavefileDialog4 = new SaveFileDialog();
            SavefileDialog4.FileName = "FrequencyGraph_";
            string filelocation = SavefileDialog4.FileName;
            SavefileDialog4.DefaultExt = ".emf";

            if (SavefileDialog4.ShowDialog() == DialogResult.OK)
            {
                chart1.SaveImage(SavefileDialog4.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Emf);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            var SavefileDialog5 = new SaveFileDialog();
            SavefileDialog5.FileName = "PieChart_";
            SavefileDialog5.DefaultExt = ".jpeg";
            
            if (SavefileDialog5.ShowDialog() == DialogResult.OK)
            {
                chart2.SaveImage(SavefileDialog5.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            var SavefileDialog6 = new SaveFileDialog();
            SavefileDialog6.FileName = "PieChart_";
            SavefileDialog6.DefaultExt = ".png";

            if (SavefileDialog6.ShowDialog() == DialogResult.OK)
            {
                chart2.SaveImage(SavefileDialog6.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            var SavefileDialog7 = new SaveFileDialog();
            SavefileDialog7.FileName = "PieChart_";
            SavefileDialog7.DefaultExt = ".tiff";

            if (SavefileDialog7.ShowDialog() == DialogResult.OK)
            {
                chart2.SaveImage(SavefileDialog7.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Tiff);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            var SavefileDialog8 = new SaveFileDialog();
            SavefileDialog8.FileName = "PieChart_";
            SavefileDialog8.DefaultExt = ".emf";

            if (SavefileDialog8.ShowDialog() == DialogResult.OK)
            {
                chart2.SaveImage(SavefileDialog8.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Emf);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        void export()
        {
            Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Excel._Worksheet worksheet = null;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "SMRDetail";

            int columncount1 = dataGridView1.Columns.Count + 1;
            int columncount2 = dataGridView2.Columns.Count + 1;

            worksheet.Cells[1, 1] = "Location";
            worksheet.Cells[1, 3] = Convert.ToString(textBox1.Text);
            worksheet.Cells[2, 1] = "Basic RMR";
            worksheet.Cells[3, 1] = "Slope angle";
            worksheet.Cells[4, 1] = "Slope direction";
            worksheet.Cells[5, 1] = "Angle of friction";
            worksheet.Cells[6, 1] = "Adjustment for excavation";
            worksheet.Cells[7, 1] = "planar lateral limit";
            worksheet.Cells[2, 4] = Convert.ToString(numericUpDown5.Value);
            worksheet.Cells[3, 4] = Convert.ToString(numericUpDown1.Value);
            worksheet.Cells[4, 4] = Convert.ToString(numericUpDown2.Value);
            worksheet.Cells[5, 4] = Convert.ToString(numericUpDown3.Value);
            worksheet.Cells[6, 4] = Convert.ToString(numericUpDown4.Value);
            worksheet.Cells[7, 4] = Convert.ToString(numericUpDown6.Value);
            worksheet.Cells[2, 6] = "Average SMR of Lowest";
            worksheet.Cells[2, 9] = Convert.ToString(numericUpDown8.Value);
            worksheet.Cells[2, 10] = "% data is";
            worksheet.Cells[2, 11] = Convert.ToString(label11.Text);
            worksheet.Cells[3, 6] = "Total failure count";
            worksheet.Cells[4, 6] = "Planar";
            worksheet.Cells[5, 6] = "Wedge";
            worksheet.Cells[6, 6] = "Topple";
            worksheet.Cells[7, 6] = "Topple lateral limit";
            worksheet.Cells[3, 8] = Convert.ToString(label12.Text);
            worksheet.Cells[4, 8] = Convert.ToString(label13.Text);
            worksheet.Cells[5, 8] = Convert.ToString(label14.Text);
            worksheet.Cells[6, 8] = Convert.ToString(label15.Text);
            worksheet.Cells[7, 8] = Convert.ToString(numericUpDown7.Value);
            worksheet.Cells[3, 10] = "Slope mass susceptibility";
            worksheet.Cells[4, 10] = "Planar kinematic susceptibility";
            worksheet.Cells[5, 10] = "Wedge kinematic susceptibility";
            worksheet.Cells[6, 10] = "Topple kinematic susceptibility";
            worksheet.Cells[7, 10] = "Overall kinematic susceptibility";
            worksheet.Cells[3, 14] = Convert.ToString(label27.Text);
            worksheet.Cells[4, 14] = Convert.ToString(label28.Text);
            worksheet.Cells[5, 14] = Convert.ToString(label29.Text);
            worksheet.Cells[6, 14] = Convert.ToString(label30.Text);
            worksheet.Cells[7, 14] = Convert.ToString(label31.Text);
            worksheet.Cells[8, 2] = "Planar";
            worksheet.Cells[8, 4] = "Topple";
            worksheet.Cells[8, 7] = "Wedge";

            for (int i = 1; i < columncount1; i++)
            {
                worksheet.Cells[9, i] = dataGridView1.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < columncount1-1; j++)
                {
                    worksheet.Cells[i + 10, j + 1] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                }
            }

            for (int i = columncount1; i < columncount1 + columncount2 -1; i++)
            {
                worksheet.Cells[9, i+1] = dataGridView2.Columns[i - columncount1].HeaderText;
            }

            for (int i = 0; i < dataGridView2.Rows.Count-1; i++)
            {
                for (int j = columncount1; j < columncount1 + columncount2-1; j++)
                {
                    worksheet.Cells[i + 10, j + 1] = Convert.ToString(dataGridView2.Rows[i].Cells[j - columncount1].Value);
                }
            }

            var SavefileDialog = new SaveFileDialog();
            SavefileDialog.FileName = "EasySMR_";
            SavefileDialog.DefaultExt = ".xlsx";

            if (SavefileDialog.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(SavefileDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();
            }

        }
    }
}

