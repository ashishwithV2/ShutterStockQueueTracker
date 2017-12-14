using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Reflection;

namespace ShutterStockQueueTracker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
         string StopRecords = "";
        //Start & Stop Button
     
           
         private  void button2_Click(object sender, EventArgs e)
        {

            if (webBrowser1.DocumentText != "")
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(comboBox1.Text, "^[0-9]*$"))
                {
                    MessageBox.Show("Invalid time duration?");
                }
                else
                {
                   
                    if (button2.Text == "Start")
                    {
                        button2.Text = "Stop";
                        button2.BackColor = Color.Tan;
                        StopRecords = "";
                        //timer1.Tick += new EventHandler(timer1_Tick);
                        timer1.Interval = Convert.ToInt32(comboBox1.Text) * 60 * 1000;
                        timer1.Enabled = true;
                      
                    }
                    else
                    {
                        timer1.Enabled = false;
                        StopRecords = "Stop";
                        button2.Text = "Start";
                        button2.BackColor = Color.Transparent;
                    }
                }
            }
            else
            {
                MessageBox.Show("Please Wait.. Webbrowser is not fully loaded.");
            }
          
            


        }
        // Form Load Event
        private void Form1_Load(object sender, EventArgs e)
        {
            
            this.AutoSize = true;
            combostyle();
            WebBroswerBind();
            string name = "ssqt.ico";
            var exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            var imgDir = Path.Combine(exeDir, name);
            System.Drawing.Icon ico = new System.Drawing.Icon(imgDir);
            this.Icon = ico;

        }
        // Combobox binding method
        private void combostyle()
        {
            comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
           
            comboBox1.Items.Clear();
            comboBox1.SelectedText = "30";
            comboBox1.Items.Add(1);
            comboBox1.Items.Add(2);
            comboBox1.Items.Add(3);
            comboBox1.Items.Add(4);
            comboBox1.Items.Add(5);
            comboBox1.Items.Add(10);
            comboBox1.Items.Add(15);
            comboBox1.Items.Add(20);
            comboBox1.Items.Add(25);
            comboBox1.Items.Add(30);
            comboBox1.Items.Add(35);
            comboBox1.Items.Add(40);
            comboBox1.Items.Add(45);
            comboBox1.Items.Add(50);
            comboBox1.Items.Add(55);
            comboBox1.Items.Add(60);
           
        }
        // Binding webbrowser method 
        private void WebBroswerBind()
        {
            var appName = System.IO.Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
            Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION",
                  appName, 11000, Microsoft.Win32.RegistryValueKind.DWord);
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                webBrowser1.AllowNavigation = true;
                webBrowser1.ScriptErrorsSuppressed = true;
              
                webBrowser1.Navigate(textBox1.Text);
            }
            else
            {
                MessageBox.Show("Please Enter URL!!");
            }
        }
        // Go Button Code here 
        private void button1_Click(object sender, EventArgs e)
        {
            WebBroswerBind();
        }
        // request with google speedsheet and bind the records
        private string GetValues(string photo, string largecollection, string photoip, string illustrationip, string vectorip, string bigstockphoto)
        {
            
          //  var script_url = "https://script.google.com/macros/s/AKfycbxEEAvSi51zKzeVfrqHBt0siUpykrmbcGbUuvM2n_5LP4c1hsI/exec"; // old jitendra shutterstocl google Excel sheet click
            var script_url = "https://script.google.com/macros/s/AKfycbzU8UC-MBtITMNq9RiNXes2SWLoeGfVe6DRLSgH5a9_ux7Qtss/exec";
            
            int _min = 1000;
            int _max = 99999999;
            Random _rdm = new Random();
            int id1 = _rdm.Next(_min, _max);

            var url = script_url + "?callback=ctrlq&id=" + id1 + "&photo=" + photo + "&largecollection=" + largecollection + "&photoip=" + photoip + "&illustrationip=" + illustrationip + "&vectorip=" + vectorip + "&bigstockphoto=" + bigstockphoto + "&action=insert";


            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
            request.Method = "GET";
            String test = String.Empty;
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                Stream dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                test = reader.ReadToEnd();
                reader.Close();
                dataStream.Close();
            }
            return test;
        }

        // Fetch
        private void Fetchrecords()
        {
            string photo = "";
            string largecollection = "";
            string photoip = "";
            string illustrationip = "";
            string vectorip = "";
            string bigstockphoto = "";
            this.webBrowser1.Navigate(textBox1.Text);
            // fetching data from webbrowser control
            string rec = webBrowser1.DocumentText;
            


            var HtmlDoc = new HtmlAgilityPack.HtmlDocument();
            HtmlDoc.LoadHtml(rec);
          
            string phto = "";
            string phto1 = "";
            string phto2 = "";
            string phto3 = "";
            string phto4 = "";
            string phto5 = "";

          
            foreach (var row in HtmlDoc.DocumentNode.SelectNodes("//table[@id = 'queues_table']/tbody/tr"))
            {
               Application.DoEvents();
                int PhotoTotal = 0;
                foreach (var cell in row.SelectNodes("td"))
                {
                    string td = Convert.ToString(cell.InnerText);

                    if ((td.Trim().ToLower()) == "photo")
                    {
                        PhotoTotal++;

                        phto = "photo";



                    }

                    else
                    {
                        if (PhotoTotal == 1 || PhotoTotal == 2)
                        {
                            PhotoTotal++;
                        }
                        else
                        {
                            break;
                        }

                    }

                    if (phto == "photo" && PhotoTotal == 3)
                    {
                        photo = Convert.ToString(td);
                    }
                    
                }

            }

            // photo
            foreach (var row in HtmlDoc.DocumentNode.SelectNodes("//table[@id = 'queues_table']/tbody/tr"))
            {
                Application.DoEvents();
                int PhotoTotal = 0;
                foreach (var cell in row.SelectNodes("td"))
                {
                    string td = Convert.ToString(cell.InnerText);

                    //large collection
                    if ((td.Trim().ToLower()) == "large collection")
                    {
                        PhotoTotal++;
                        phto1 = "large collection";

                    }

                    else
                    {
                        if (PhotoTotal == 1 || PhotoTotal == 2)
                        {
                            PhotoTotal++;
                        }
                        else
                        {
                            break;
                        }

                    }

                    if (phto1 == "large collection" && PhotoTotal == 3)
                    {
                        largecollection = Convert.ToString(td);
                    }
                   
                }
            }


            // photo
            foreach (var row in HtmlDoc.DocumentNode.SelectNodes("//table[@id = 'queues_table']/tbody/tr"))
            {
               Application.DoEvents();
                int PhotoTotal = 0;
                foreach (var cell in row.SelectNodes("td"))
                {
                    string td = Convert.ToString(cell.InnerText);

                    //photo ip
                    if ((td.Trim().ToLower()) == "photo ip")
                    {
                        PhotoTotal++;
                        phto2 = "photo ip";
                    }

                    else
                    {
                        if (PhotoTotal == 1 || PhotoTotal == 2)
                        {
                            PhotoTotal++;
                        }
                        else
                        {
                            break;
                        }

                    }

                    if (phto2 == "photo ip" && PhotoTotal == 3)
                    {
                        photoip = Convert.ToString(td);
                    }
                }
            }


            // illustration ip
            foreach (var row in HtmlDoc.DocumentNode.SelectNodes("//table[@id = 'queues_table']/tbody/tr"))
            {
               Application.DoEvents();
                int PhotoTotal = 0;
                foreach (var cell in row.SelectNodes("td"))
                {
                    string td = Convert.ToString(cell.InnerText);
                    //illustration ip
                    if ((td.Trim().ToLower()) == "illustration ip")
                    {
                        PhotoTotal++;

                        phto3 = "illustration ip";

                    }

                    else
                    {
                        if (PhotoTotal == 1 || PhotoTotal == 2)
                        {
                            PhotoTotal++;
                        }
                        else
                        {
                            break;
                        }

                    }

                    if (phto3 == "illustration ip" && PhotoTotal == 3)
                    {
                        illustrationip = Convert.ToString(td);
                    }
                }
            }

            //  //Vector Ip
            foreach (var row in HtmlDoc.DocumentNode.SelectNodes("//table[@id = 'queues_table']/tbody/tr"))
            {
                Application.DoEvents();
                int PhotoTotal = 0;
                foreach (var cell in row.SelectNodes("td"))
                {
                    
                    string td = Convert.ToString(cell.InnerText);

                    //Vector Ip
                    if ((td.Trim().ToLower()) == "vector ip")
                    {
                        PhotoTotal++;

                        phto4 = "vector ip";



                    }

                    else
                    {
                        if (PhotoTotal == 1 || PhotoTotal == 2)
                        {
                            PhotoTotal++;
                        }
                        else
                        {
                            break;
                        }

                    }

                    if (phto4 == "vector ip" && PhotoTotal == 3)
                    {
                        vectorip = Convert.ToString(td);
                    }

                }
            }

            //  //Vector Ip
            foreach (var row in HtmlDoc.DocumentNode.SelectNodes("//table[@id = 'queues_table']/tbody/tr"))
            {
                Application.DoEvents();
                int PhotoTotal = 0;
                foreach (var cell in row.SelectNodes("td"))
                {
                
                    string td = Convert.ToString(cell.InnerText);

                    //Bigstock Photo
                    if ((td.Trim().ToLower()) == "bigstock photo")
                    {
                        PhotoTotal++;

                        phto5 = "bigstock photo";
                    }

                    else
                    {
                        if (PhotoTotal == 1 || PhotoTotal == 2)
                        {
                            PhotoTotal++;
                        }
                        else
                        {
                            break;
                        }

                    }

                    if (phto5 == "bigstock photo" && PhotoTotal == 3)
                    {
                        bigstockphoto= Convert.ToString(td);
                    }

                }
            }
    
            // getting data from different source and sent back to google speedsheet
            GetValues(photo, largecollection, photoip, illustrationip, vectorip, bigstockphoto);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
               //if (StopRecords == "Stop")
               // {
               //     timer.Stop();
               //     timer.Enabled = false;
               // }
               // else
               // {
               //     Fetchrecords();
               // }

                Fetchrecords();

            }
            catch (Exception ex)
            {
                textBox2.Text = (DateTime.Now.ToString() + " " + ex.Message);
               
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }
        private void Form1_Resize(object sender, System.EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
            {
                 this.Hide();
            }
          
            
        }



    }
}
