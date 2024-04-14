using SAPbobsCOM;
using System;
using System.Windows.Forms;


namespace _1st_week
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public TextBox getUser
        {
            get { return textBox3; }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                string[] values = {  "dst_MSSQL",
                                 "dst_DB_2",
                                 "dst_SYBASE",
                                 "dst_MSSQL2005",
                                 "dst_MAXDB",
                                 "dst_MSSQL2008",
                                 "dst_MSSQL2012",
                                 "dst_MSSQL2014",
                                 "dst_HANADB",
                                 "dst_MSSQL2016",
                                 "dst_MSSQL2017",
                                 "dst_MSSQL2019"};

                comboBox1.Items.AddRange(values);

                comboBox1.SelectedIndex = 11;
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void toolStripContainer1_ContentPanel_Load(object sender, EventArgs e)
        {

        }

        private void toolStripContainer1_TopToolStripPanel_Click(object sender, EventArgs e)
        {

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                DiApi.oCompany = new Company();

                DiApi.oCompany.SLDServer = textBox1.Text + ":40000";

                string hanaServer = "NDB@" + textBox1.Text + ":30013";

                DiApi.oCompany.Server = comboBox1.SelectedItem.ToString() == "dst_HANADB" ? hanaServer : textBox1.Text;

                if (comboBox1.SelectedItem.ToString() != null)
                {
                    string val = comboBox1.SelectedItem.ToString();
                    DiApi.oCompany.DbServerType = (BoDataServerTypes)Enum.Parse(typeof(BoDataServerTypes), val);
                }

                DiApi.oCompany.CompanyDB = textBox2.Text;
                DiApi.oCompany.UserName = textBox3.Text;
                DiApi.oCompany.Password = textBox4.Text;
                DiApi.oCompany.language = BoSuppLangs.ln_Portuguese_Br;

                int res = DiApi.oCompany.Connect();

                if (res != 0)
                {

                    throw new Exception($"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                }

                MessageBox.Show("A conexão com o SAP B1 foi realizada com sucesso!");

                Forms.Menu menu = new Forms.Menu();

                this.Hide();

                menu.Location = this.Location;

                menu.Show();    
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
    }
}
