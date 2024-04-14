using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace _1st_week.Forms
{
    public partial class PV : Form
    {
        private List<Object> listObj = new List<Object>();

        SAPbobsCOM.Recordset rs = DiApi.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

        public PV()
        {
            InitializeComponent();

            try
            {
                this.FormClosing += FormClose;

                List<KeyValuePair<string, string>> items = new List<KeyValuePair<string, string>>();

                string query = $@"SELECT ""CardCode"", ""CardName"" FROM OCRD WHERE ""CardType"" != 'S' AND ""validFor"" = 'Y' ORDER BY ""CardName""";

                rs.DoQuery(query);

                if (rs.RecordCount > 0)
                {
                    for (int i = 1; i <= rs.RecordCount; i++)
                    {
                        items.Add(new KeyValuePair<string, string>((string)rs.Fields.Item("CardCode").Value, (string)rs.Fields.Item("CardName").Value));

                        if (i < rs.RecordCount)
                        {
                            rs.MoveNext();
                        }
                    }

                    comboBox1.DataSource = new BindingSource(items, null);
                    comboBox1.DisplayMember = "Value";
                    comboBox1.ValueMember = "Key";
                }


                listObj.Add(rs);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

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

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }


        private void button1_Click(object sender, EventArgs e)
        {

            try
            {

                List<KeyValuePair<int, int>> items = new List<KeyValuePair<int, int>>();

                SAPbobsCOM.Documents oPV = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oOrders);

                oPV.CardCode = comboBox1.SelectedValue.ToString();
                oPV.BPL_IDAssignedToInvoice = int.Parse(textBox10.Text);
                oPV.Comments = richTextBox1.Text;
                oPV.DocDate = DateTime.ParseExact(textBox11.Text, "dd/MM/yyyy", null);
                oPV.DocDueDate = DateTime.ParseExact(textBox12.Text, "dd/MM/yyyy", null);
                oPV.TaxDate = DateTime.ParseExact(textBox13.Text, "dd/MM/yyyy", null);

                int cnt = 0;

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row != null && !row.IsNewRow)
                    {
                        if (row.Cells["dataGridViewTextBoxColumn5"].Value != null)
                        {
                            if (cnt > 0)
                            {
                                oPV.Lines.Add();
                            }

                            object c1 = row.Cells["dataGridViewTextBoxColumn1"].Value, c2 = row.Cells["dataGridViewTextBoxColumn2"].Value, c3 = row.Cells["dataGridViewTextBoxColumn3"].Value,
                                   c5 = row.Cells["dataGridViewTextBoxColumn5"].Value, c6 = row.Cells["dataGridViewTextBoxColumn6"].Value, c7 = row.Cells["Column7"].Value,
                                   c8 = row.Cells["Column8"].Value, c9 = row.Cells["Column9"].Value;

                            oPV.Lines.BaseType = c1 != null ? int.Parse(c1.ToString()) : -1;
                            if (c2 != null)
                                oPV.Lines.BaseEntry = int.Parse(c2.ToString());
                            if (c9 != null)
                                oPV.Lines.BaseLine = int.Parse(c9.ToString());
                            if (c3 != null)
                            {
                                items.Add(new KeyValuePair<int, int>(int.Parse(c3.ToString()), cnt));
                            }
                            oPV.Lines.ItemCode = c5 != null ? c5.ToString() : "";
                            oPV.Lines.Quantity = c6 != null ? double.Parse(c6.ToString()) : 0;
                            oPV.Lines.UnitPrice = c7 != null ? double.Parse(c7.ToString()) : 0;
                            oPV.Lines.TaxCode = c8 != null ? c8.ToString() : "";
                            cnt++;
                        }
                    }
                }

                int res = oPV.Add();

                if (res != 0)
                {
                    throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                }

                MessageBox.Show("Pedido de venda gerado com sucesso!");

                listObj.Add(oPV);

                if (items.Count > 0)
                {
                    int docentry = int.Parse(DiApi.oCompany.GetNewObjectKey());

                    SAPbobsCOM.Documents odln = null;
                    SAPbobsCOM.Documents oinv = null;
                    SAPbobsCOM.Documents odpi = null;

                    foreach (var item in items)
                    {
                        int doc = item.Key, lineOdln = 0, lineOinv = 0, lineOdpi = 0;

                        switch (doc)
                        {
                            case 15:
                                if (lineOdln == 0)
                                {
                                    odln = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
                                    odln.CardCode = oPV.CardCode;
                                    odln.BPL_IDAssignedToInvoice = oPV.BPL_IDAssignedToInvoice;
                                    odln.Comments = oPV.Comments;
                                    odln.DocDate = oPV.DocDate;
                                    odln.DocDueDate = oPV.DocDueDate;
                                    odln.TaxDate = oPV.TaxDate;
                                }

                                if (lineOdln > 0)
                                {
                                    odln.Lines.Add();
                                }

                                odln.Lines.BaseType = (int)BoObjectTypes.oOrders;
                                odln.Lines.BaseEntry = docentry;
                                odln.Lines.BaseLine = item.Value;
                                lineOdln++;
                                break;

                            case 13:
                                if (lineOinv == 0)
                                {
                                    oinv = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oInvoices);
                                    oinv.CardCode = oPV.CardCode;
                                    oinv.BPL_IDAssignedToInvoice = oPV.BPL_IDAssignedToInvoice;
                                    oinv.Comments = oPV.Comments;
                                    oinv.DocDate = oPV.DocDate;
                                    oinv.DocDueDate = oPV.DocDueDate;
                                    oinv.TaxDate = oPV.TaxDate;
                                }

                                if (lineOinv > 0)
                                {
                                    oinv.Lines.Add();
                                }

                                oinv.Lines.BaseType = (int)BoObjectTypes.oOrders;
                                oinv.Lines.BaseEntry = docentry;
                                oinv.Lines.BaseLine = item.Value;
                                lineOinv++;
                                break;

                            case 203:
                                if (lineOdpi == 0)
                                {
                                    odpi = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oDownPayments);
                                    odpi.CardCode = oPV.CardCode;
                                    odpi.BPL_IDAssignedToInvoice = oPV.BPL_IDAssignedToInvoice;
                                    odpi.DownPaymentPercentage = 100;
                                    odpi.DownPaymentType = DownPaymentTypeEnum.dptInvoice;
                                    odpi.Comments = oPV.Comments;
                                    odpi.DocDate = oPV.DocDate;
                                    odpi.DocDueDate = oPV.DocDueDate;
                                    odpi.TaxDate = oPV.TaxDate;
                                }

                                if (lineOdpi > 0)
                                {
                                    odpi.Lines.Add();
                                }

                                odpi.Lines.BaseType = (int)BoObjectTypes.oOrders;
                                odpi.Lines.BaseEntry = docentry;
                                odpi.Lines.BaseLine = item.Value;
                                lineOdpi++;
                                break;
                        }

                    }

                    if (odln != null)
                    {
                        res = odln.Add();
                        if (res != 0)
                        {
                            throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                        }

                        MessageBox.Show("Entrega gerada com sucesso!");

                        listObj.Add(odln);
                    }

                    if (oinv != null)
                    {
                        res = oinv.Add();
                        if (res != 0)
                        {
                            throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                        }

                        MessageBox.Show("Nota de saída gerada com sucesso!");

                        listObj.Add(oinv);
                    }

                    if (odpi != null)
                    {
                        res = odpi.Add();
                        if (res != 0)
                        {
                            throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                        }

                        MessageBox.Show("Adiantamento gerado com sucesso!");

                        listObj.Add(odpi);
                    }

                }

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

        private void PV_Load(object sender, EventArgs e)
        {


        }

        private void button4_Click(object sender, EventArgs e)
        {
            List<KeyValuePair<int, int>> items = new List<KeyValuePair<int, int>>();

            SAPbobsCOM.Documents oPV = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oDrafts);

            oPV.CardCode = comboBox1.SelectedValue.ToString();
            oPV.BPL_IDAssignedToInvoice = int.Parse(textBox10.Text);
            oPV.Comments = richTextBox1.Text;
            oPV.DocDate = DateTime.ParseExact(textBox11.Text, "dd/MM/yyyy", null);
            oPV.DocDueDate = DateTime.ParseExact(textBox12.Text, "dd/MM/yyyy", null);
            oPV.TaxDate = DateTime.ParseExact(textBox13.Text, "dd/MM/yyyy", null);

            int cnt = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row != null && !row.IsNewRow)
                {
                    if (row.Cells["dataGridViewTextBoxColumn5"].Value != null)
                    {
                        if (cnt > 0)
                        {
                            oPV.Lines.Add();
                        }

                        object c1 = row.Cells["dataGridViewTextBoxColumn1"].Value, c2 = row.Cells["dataGridViewTextBoxColumn2"].Value, c3 = row.Cells["dataGridViewTextBoxColumn3"].Value,
                               c5 = row.Cells["dataGridViewTextBoxColumn5"].Value, c6 = row.Cells["dataGridViewTextBoxColumn6"].Value, c7 = row.Cells["Column7"].Value,
                               c8 = row.Cells["Column8"].Value;

                        oPV.Lines.BaseType = c1 != null ? int.Parse(c1.ToString()) : -1;
                        if (c2 != null)
                            oPV.Lines.BaseEntry = int.Parse(c2.ToString());
                        if (c3 != null)
                        {
                            items.Add(new KeyValuePair<int, int>(int.Parse(c3.ToString()), cnt));
                        }
                        oPV.Lines.ItemCode = c5 != null ? c5.ToString() : "";
                        oPV.Lines.Quantity = c6 != null ? double.Parse(c6.ToString()) : 0;
                        oPV.Lines.Price = c7 != null ? double.Parse(c7.ToString()) : 0;
                        oPV.Lines.TaxCode = c8 != null ? c8.ToString() : "";
                        cnt++;
                    }
                }
            }

            int res = oPV.Add();

            if (res != 0)
            {
                throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
            }

            MessageBox.Show("Esboço do pedido de venda gerado com sucesso!");

            listObj.Add(oPV);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SAPbobsCOM.Documents oPV = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oOrders);

                if (!oPV.GetByKey(int.Parse(textBox15.Text)))
                    throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");

                int res;

                if (comboBox2.SelectedItem.ToString() == "Fechar")
                {
                    res = oPV.Close();
                    if (res != 0)
                    {
                        throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                    }

                    MessageBox.Show("Pedido de venda fechado com sucesso!");

                    return;
                }

                if (comboBox2.SelectedItem.ToString() == "Cancelar")
                {
                    res = oPV.Cancel();
                    if (res != 0)
                    {
                        throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                    }

                    MessageBox.Show("Pedido de venda cancelado com sucesso!");

                    return;
                }

                oPV.Comments = richTextBox1.Text;
                oPV.DocDate = DateTime.ParseExact(textBox11.Text, "dd/MM/yyyy", null);
                oPV.DocDueDate = DateTime.ParseExact(textBox12.Text, "dd/MM/yyyy", null);
                oPV.TaxDate = DateTime.ParseExact(textBox13.Text, "dd/MM/yyyy", null);

                res = oPV.Update();

                if (res != 0)
                {
                    throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                }

                MessageBox.Show("Pedido de venda atualizado com sucesso!");

                listObj.Add(oPV);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
    }
}
