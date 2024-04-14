using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace _1st_week.Forms
{
    public partial class ITEM : Form
    {
        private List<object> listObj = new List<Object>();

        SAPbobsCOM.Recordset rs = DiApi.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

        public ITEM()
        {
            InitializeComponent();

            try
            {
                this.FormClosing += FormClose;

                List<KeyValuePair<int, string>> items = new List<KeyValuePair<int, string>>();
                List<KeyValuePair<int, string>> items2 = new List<KeyValuePair<int, string>>();


                string SeqITEM = $@"SELECT ""Series"",""SeriesName"" FROM NNM1 WHERE ""ObjectCode"" = '4' ORDER BY ""Series""",
                       grupos = $@"select ""ItmsGrpCod"",""ItmsGrpNam"" from OITB";

                rs.DoQuery(SeqITEM);

                if (rs.RecordCount > 0)
                {
                    for (int i = 1; i <= rs.RecordCount; i++)
                    {
                        items.Add(new KeyValuePair<int, string>((int)rs.Fields.Item("Series").Value, (string)rs.Fields.Item("SeriesName").Value));

                        if (i < rs.RecordCount)
                        {
                            rs.MoveNext();
                        }
                    }

                    comboBox2.DataSource = new BindingSource(items, null);
                    comboBox2.DisplayMember = "Value";
                    comboBox2.ValueMember = "Key";
                }

                rs.DoQuery(grupos);


                if (rs.RecordCount > 0)
                {
                    for (int i = 1; i <= rs.RecordCount; i++)
                    {
                        items2.Add(new KeyValuePair<int, string>((int)rs.Fields.Item("ItmsGrpCod").Value, (string)rs.Fields.Item("ItmsGrpNam").Value));

                        if (i < rs.RecordCount)
                        {
                            rs.MoveNext();
                        }
                    }

                    comboBox3.DataSource = new BindingSource(items2, null);
                    comboBox3.DisplayMember = "Value";
                    comboBox3.ValueMember = "Key";
                }

                listObj.Add(rs);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            radioButton2.Checked = true;
        }


        private void FormClose(object sender, FormClosingEventArgs e)
        {
            try
            {
                foreach (object item in listObj)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                }

                Menu menu = new Menu();

                menu.Enabled = true;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }


        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                SAPbobsCOM.Items oItems = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oItems);

                oItems.Series = (int)comboBox2.SelectedValue;
                oItems.ItemCode = textBox3.Text;
                oItems.ItemName = textBox2.Text;
                oItems.ItemsGroupCode = (int)comboBox3.SelectedValue;
                oItems.InventoryItem = checkBox1.Checked ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;
                oItems.SalesItem = checkBox2.Checked ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;
                oItems.PurchaseItem = checkBox3.Checked ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;
                if (radioButton1.Checked)
                {
                    oItems.ItemClass = ItemClassEnum.itcService;
                }
                else
                {
                    oItems.ItemClass = ItemClassEnum.itcMaterial;
                }
                oItems.MaterialType = (BoMaterialTypes)Enum.Parse(typeof(BoMaterialTypes), textBox1.Text);

                string codeNCM = $@"select ""AbsEntry"" from ONCM where ""NcmCode"" = '{textBox4.Text}'";

                rs.DoQuery(codeNCM);

                if (rs.RecordCount == 1)
                {
                    oItems.NCMCode = (int)rs.Fields.Item("AbsEntry").Value;
                }

                oItems.ValidRemarks = textBox5.Text;

                oItems.PurchaseUnitLength = double.Parse(textBox8.Text.Replace(".", "").Replace(",", "."),CultureInfo.InvariantCulture);
                oItems.SalesUnitLength = double.Parse(textBox8.Text.Replace(".", "").Replace(",", "."), CultureInfo.InvariantCulture);
                oItems.PurchaseUnitWidth = double.Parse(textBox7.Text.Replace(".", "").Replace(",", "."), CultureInfo.InvariantCulture);
                oItems.SalesUnitWidth = double.Parse(textBox7.Text.Replace(".", "").Replace(",", "."), CultureInfo.InvariantCulture);
                oItems.PurchaseUnitHeight = double.Parse(textBox6.Text.Replace(".", "").Replace(",", "."), CultureInfo.InvariantCulture);
                oItems.SalesUnitHeight = double.Parse(textBox6.Text.Replace(".", "").Replace(",", "."), CultureInfo.InvariantCulture);
                oItems.PurchaseUnitWeight = double.Parse(textBox9.Text.Replace(".", "").Replace(",", "."), CultureInfo.InvariantCulture);
                oItems.SalesUnitWeight = double.Parse(textBox9.Text.Replace(".", "").Replace(",", "."), CultureInfo.InvariantCulture);

                int res = oItems.Add();

                if (res != 0)
                    throw new Exception($"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");

                MessageBox.Show("Item cadastrado com sucesso!");

                listObj.Add(rs);
                listObj.Add(oItems);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
