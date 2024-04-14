using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _1st_week.Forms
{
    public partial class Menu : Form
    {
        public Menu()
        {
            InitializeComponent();
            this.Load += load;
            ToolTip toolTip1 = new ToolTip();
            toolTip1.SetToolTip(button5, "Pesquisar");
            ToolTip toolTip2 = new ToolTip();
            toolTip2.SetToolTip(button6, "Adicionar");
            ToolTip toolTip3 = new ToolTip();
            toolTip3.SetToolTip(button7, "Primeiro registro");
            ToolTip toolTip4 = new ToolTip();
            toolTip4.SetToolTip(button8, "Próximo registro");
            ToolTip toolTip5 = new ToolTip();
            toolTip5.SetToolTip(button9, "Último registro");
            ToolTip toolTip6 = new ToolTip();
            toolTip6.SetToolTip(button10, "Registro anterior");

            this.FormClosing += FormClose;
        }

        private bool encerrar = true;

    
        private void load(object sender, EventArgs e)
        {
            try
            {
                SAPbobsCOM.CompanyService companyService = DiApi.oCompany.GetCompanyService();

                SAPbobsCOM.AdminInfo empresa = companyService.GetAdminInfo();

                Form1 form1 = new Form1();

                SAPbobsCOM.Recordset rs = DiApi.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = $@"SELECT ""U_NAME"" FROM OUSR WHERE ""USER_CODE"" = '{form1.getUser.Text}'", user = "";

                rs.DoQuery(query);

                if (rs.RecordCount == 1)
                {
                    user = rs.Fields.Item("U_NAME").Value.ToString();
                }

                textBox1.Text = empresa.CompanyName + " | " + user;
          
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void FormClose(object sender, FormClosingEventArgs e)
        {
            if (encerrar)
            {
                Application.Exit();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Forms.PN pN = new Forms.PN();

            pN.ShowDialog();
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Menu_Load(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void encerrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DiApi.oCompany.Disconnect();
            encerrar = false;
            this.Close();
            Form1 form = new Form1();

            form.Show();
        }

        private void encerrarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DiApi.oCompany.Disconnect();
            encerrar = false;

            this.Close();

            Form1 form = new Form1();

            form.Close();

            Application.Exit();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            textBox3.Text = DateTime.Now.ToString("dd/MM/yyyy - HH:mm");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Forms.ITEM item = new Forms.ITEM();

            item.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Forms.PV pv = new Forms.PV();

            pv.ShowDialog();
        }

      
        private void button11_Click(object sender, EventArgs e)
        {
            Forms.Esboco esboco = new Forms.Esboco();

            esboco.ShowDialog();  
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
               
                if (!DiApi.oCompany.InTransaction)
                {
                    DiApi.oCompany.StartTransaction();

                    MessageBox.Show("Transação iniciada com sucesso!");
                }
                else
                {
                    MessageBox.Show("A transação já foi iniciada, necessário concluir a transação antes de inicar outra!");
                }
                    
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {

                if (DiApi.oCompany.InTransaction)
                {
                    DiApi.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                    MessageBox.Show("Commit da transação realizado com sucesso!");
                }
                else
                {
                    MessageBox.Show("Não existe transação em andamento, necessário iniciar uma transação antes de realizar o commit!");
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {

                if (DiApi.oCompany.InTransaction)
                {
                    DiApi.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);

                    MessageBox.Show("Roll back da transação realizado com sucesso!");
                }
                else
                {
                    MessageBox.Show("Não existe transação em andamento, necessário iniciar uma transação antes de realizar o roll back!");
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Forms.TablesMD tablesMD = new Forms.TablesMD();
            tablesMD.ShowDialog();
        }
    }
}
