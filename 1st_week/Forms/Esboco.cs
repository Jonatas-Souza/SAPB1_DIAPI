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
    public partial class Esboco : Form
    {
        private List<Object> listObj = new List<Object>();

        public Esboco()
        {
            InitializeComponent();
            this.FormClosing += FormClose;
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

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                SAPbobsCOM.Recordset rs = DiApi.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                string dataIni = textBox1.Text != null && textBox1.Text != "" ? textBox1.Text : "19000101",
                       dataFin = textBox2.Text != null && textBox2.Text != "" ? textBox2.Text : "29991231", 
                       status = comboBox1.SelectedItem.ToString(),

                query = $@"select
		                    ""DocEntry"",
	                    case 
		                    when ""ObjType"" = 23 then 'Cotação de vendas'
		                    when ""ObjType"" = 17 then 'Pedido de vendas'
		                    when ""ObjType"" = 15 then 'Entrega'
		                    when ""ObjType"" = 16 then 'Nota de saída'
		                    when ""ObjType"" = 16 then 'Devolução de entrega'
		                    when ""ObjType"" = 203 then 'Adiantamento de cliente'
		                    when ""ObjType"" = 13 then 'Nota de saída'
		                    when ""ObjType"" = 14 then 'Devolução de nota de saída'
		                    when ""ObjType"" = 1470000113 then 'Solicitação de compra'
		                    when ""ObjType"" = 540000006 then 'Oferta de compra'
		                    when ""ObjType"" = 22 then 'Pedido de compra'
		                    when ""ObjType"" = 21 then 'Devolução de mercadorias'
		                    when ""ObjType"" = 204 then 'Adiantamento de fornecedor'
		                    when ""ObjType"" = 18 then 'Nota de entrada'
		                    when ""ObjType"" = 19 then 'Devolução de nota de entrada'
		                    end as ""Tipo"",
		                    ""CardCode"",
		                    ""CardName"",
		                    ""TaxDate"",
		                    ""DocTotal"",
		                    ""Comments"",
		                    case when ""DocStatus"" = 'O' then 'Pendente'
		                    else 'Executado' end ""Status""
                    from
	                    ODRF 

                    where
 
	                    CAST(""TaxDate"" as date) BETWEEN '{dataIni}' and '{dataFin}'
	                    and ""DocStatus"" =  (case when '{status}' = 'Todos' then ""DocStatus"" when '{status}' = 'Pendentes' then'O' when '{status}' = 'Executados' then 'C' end)";

                rs.DoQuery(query);

                if (rs.RecordCount > 0)
                {
                    dataGridView1.Rows.Clear();

                    for (int i = 1; i <= rs.RecordCount; i++)
                    {
                        int rowIndex = dataGridView1.Rows.Add();

                        dataGridView1.Rows[rowIndex].Cells["Column2"].Value = rs.Fields.Item(0).Value;
                        dataGridView1.Rows[rowIndex].Cells["Column1"].Value = rs.Fields.Item(1).Value;
                        dataGridView1.Rows[rowIndex].Cells["Column3"].Value = rs.Fields.Item(2).Value;
                        dataGridView1.Rows[rowIndex].Cells["Column4"].Value = rs.Fields.Item(3).Value;
                        dataGridView1.Rows[rowIndex].Cells["Column5"].Value = rs.Fields.Item(4).Value;
                        dataGridView1.Rows[rowIndex].Cells["Column6"].Value = rs.Fields.Item(5).Value;
                        dataGridView1.Rows[rowIndex].Cells["Column7"].Value = rs.Fields.Item(6).Value;
                        dataGridView1.Rows[rowIndex].Cells["Column8"].Value = rs.Fields.Item(7).Value;

                        if (i < rs.RecordCount)
                        {
                            rs.MoveNext();
                        }

                    }
 
                }

                listObj.Add(rs);

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row = dataGridView1.SelectedRows[0];

                    int docentry = int.Parse(row.Cells["Column2"].Value.ToString());

                    SAPbobsCOM.Documents oDraft = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oDrafts);

                    if (!oDraft.GetByKey(docentry))
                    {
                        throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                    }

                    int res = oDraft.SaveDraftToDocument();

                    if (res != 0)
                        throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");

                    MessageBox.Show($"O esboço {docentry} foi executado com sucesso!");

                    listObj.Add(oDraft);
                }
                else
                {
                    MessageBox.Show("Primeiro selecione a linha do esboço que deseja executar!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row = dataGridView1.SelectedRows[0];

                    int docentry = int.Parse(row.Cells["Column2"].Value.ToString());

                    SAPbobsCOM.Documents oDraft = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oDrafts);

                    if (!oDraft.GetByKey(docentry))
                    {
                        throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                    }

                    int res = oDraft.Remove();

                    if (res != 0)
                        throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");

                    MessageBox.Show($"O esboço {docentry} foi removido com sucesso!");

                    listObj.Add(oDraft);
                }
                else
                {
                    MessageBox.Show("Primeiro selecione a linha do esboço que deseja remover!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
    }
}
