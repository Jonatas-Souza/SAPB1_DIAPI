using _1st_week.Models;
using Newtonsoft.Json;
using RestSharp;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Runtime.ConstrainedExecution;
using System.Security.Policy;
using System.Text.Json.Serialization;
using System.Windows.Forms;

namespace _1st_week.Forms
{
    public partial class PN : Form
    {
        private List<Object> listObj = new List<Object>();

        SAPbobsCOM.Recordset rs = DiApi.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

        public PN()
        {
            InitializeComponent();

            this.FormClosing += FormClose;

        }


        private void FormClose(object sender, FormClosingEventArgs e)
        {
            try
            {
                foreach (object PN in listObj)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(PN);
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
                SAPbobsCOM.BusinessPartners oBP = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                oBP.Series = (int)comboBox2.SelectedValue;
                if (oBP.Series <= 2)
                {
                    oBP.CardCode = textBox3.Text;
                }
                if (comboBox1.SelectedItem != null && comboBox1.SelectedItem.ToString() == "Fornecedor")
                {
                    oBP.CardType = BoCardTypes.cSupplier;
                }
                else if (comboBox1.SelectedItem != null && comboBox1.SelectedItem.ToString() == "Cliente")
                {
                    oBP.CardType = BoCardTypes.cCustomer;
                }
                else if (comboBox1.SelectedItem != null && comboBox1.SelectedItem.ToString() == "Lead")
                {
                    oBP.CardType = BoCardTypes.cLid;
                }
                oBP.CardName = textBox2.Text;
                oBP.GroupCode = (int)comboBox3.SelectedValue;
                oBP.Phone1 = textBox5.Text;
                oBP.Phone2 = textBox6.Text;
                oBP.EmailAddress = textBox7.Text;
                oBP.Website = textBox8.Text;

                int cnt = 0;

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row != null && !row.IsNewRow)
                    {
                        DataGridViewCell cell = row.Cells["Column1"];

                        if (cell.Value.ToString() != "" && cell.Value != null)
                        {
                            if (cnt > 0)
                            {
                                oBP.ContactEmployees.Add();
                            }
                            oBP.ContactEmployees.Name = cell.Value.ToString();
                            object c2 = row.Cells["Column2"].Value, c3 = row.Cells["Column3"].Value, c4 = row.Cells["Column4"].Value,
                                   c5 = row.Cells["Column5"].Value, c6 = row.Cells["Column6"].Value;
                            oBP.ContactEmployees.FirstName = c2 != null ? c2.ToString() : "";
                            oBP.ContactEmployees.MiddleName = c3 != null ? c3.ToString() : "";
                            oBP.ContactEmployees.LastName = c4 != null ? c4.ToString() : "";
                            oBP.ContactEmployees.E_Mail = c5 != null ? c5.ToString() : "";
                            oBP.ContactEmployees.Phone1 = c6 != null ? c6.ToString() : "";
                            cnt++;
                        }
                    }
                }

                int ct = textBox1.Text.Replace(".", "").Replace("/", "").Replace("-", "").Length;

                if (ct == 11)
                {
                    oBP.FiscalTaxID.TaxId4 = textBox1.Text;
                }
                else if (ct == 14)
                {
                    oBP.FiscalTaxID.TaxId0 = textBox1.Text;
                }

                oBP.FiscalTaxID.TaxId1 = textBox4.Text;

                int cnt2 = 0;

                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row != null && !row.IsNewRow)
                    {
                        if (row.Cells["ComboBoxColumn1"].Value != null && row.Cells["TextBoxColumn9"].Value != null)
                        {
                            if (cnt2 > 0)
                            {
                                oBP.Addresses.Add();
                            }

                            object c3 = row.Cells["TextBoxColumn1"].Value, c4 = row.Cells["TextBoxColumn2"].Value, c5 = row.Cells["TextBoxColumn3"].Value,
                                   c6 = row.Cells["TextBoxColumn4"].Value, c7 = row.Cells["TextBoxColumn5"].Value, c8 = row.Cells["TextBoxColumn6"].Value,
                                   c9 = row.Cells["TextBoxColumn7"].Value, c10 = row.Cells["TextBoxColumn8"].Value, c11 = row.Cells["TextBoxColumn10"].Value;
                            if (row.Cells["ComboBoxColumn1"].Value.ToString().Trim() == "Cobrança")
                            {
                                oBP.Addresses.AddressType = BoAddressType.bo_BillTo;
                            }
                            else
                            {
                                oBP.Addresses.AddressType = BoAddressType.bo_ShipTo;
                            }
                            oBP.Addresses.AddressName = row.Cells["TextBoxColumn9"].Value.ToString();
                            oBP.Addresses.TypeOfAddress = c3 != null ? c3.ToString() : "";
                            oBP.Addresses.Street = c4 != null ? c4.ToString() : "";
                            oBP.Addresses.StreetNo = c11 != null ? c11.ToString() : "";
                            oBP.Addresses.BuildingFloorRoom = c5 != null ? c5.ToString() : "";
                            oBP.Addresses.ZipCode = c6 != null ? c6.ToString() : "";
                            oBP.Addresses.Block = c7 != null ? c7.ToString() : "";
                            oBP.Addresses.Country = c8 != null ? c8.ToString() : "";
                            oBP.Addresses.State = c9 != null ? c9.ToString() : "";
                            oBP.Addresses.City = c10 != null ? c10.ToString() : "";
                            if (oBP.CardType != BoCardTypes.cSupplier && oBP.Addresses.AddressType == BoAddressType.bo_ShipTo)
                            {
                                oBP.FiscalTaxID.Add();
                                oBP.FiscalTaxID.Address = row.Cells["TextBoxColumn9"].Value.ToString();

                                if (ct == 11)
                                {
                                    oBP.FiscalTaxID.TaxId4 = textBox1.Text;
                                }
                                else if (ct == 14)
                                {
                                    oBP.FiscalTaxID.TaxId0 = textBox1.Text;
                                }

                                oBP.FiscalTaxID.TaxId1 = textBox4.Text;
                            }
                            cnt2++;
                        }
                    }
                }

                listObj.Add(oBP);

                int res = oBP.Add();

                if (res != 0)
                    throw new Exception($@"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");

                MessageBox.Show("Parceiro de negócio adicionado com sucesso!");
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
                List<KeyValuePair<int, string>> PNs = new List<KeyValuePair<int, string>>();
                List<KeyValuePair<int, string>> PNs2 = new List<KeyValuePair<int, string>>();

                string tipo;

                switch (comboBox1.SelectedItem.ToString())
                {
                    case "Cliente":
                        tipo = "C";
                        break;
                    case "Fornecedor":
                        tipo = "S";
                        break;
                    default:
                        tipo = "C";
                        break;
                }

                string SeqPN = $@"SELECT ""Series"",""SeriesName"" FROM NNM1 WHERE ""ObjectCode"" = '2' AND ""DocSubType"" = '{tipo}' ORDER BY ""Series""",
                       grupos = $@"select ""GroupCode"",""GroupName"" from OCRG WHERE ""GroupType"" = '{tipo}'";

                rs.DoQuery(SeqPN);

                if (rs.RecordCount > 0)
                {
                    for (int i = 1; i <= rs.RecordCount; i++)
                    {
                        PNs.Add(new KeyValuePair<int, string>((int)rs.Fields.Item("Series").Value, (string)rs.Fields.Item("SeriesName").Value));

                        if (i < rs.RecordCount)
                        {
                            rs.MoveNext();
                        }
                    }

                    comboBox2.DataSource = new BindingSource(PNs, null);
                    comboBox2.DisplayMember = "Value";
                    comboBox2.ValueMember = "Key";
                }

                rs.DoQuery(grupos);


                if (rs.RecordCount > 0)
                {
                    for (int i = 1; i <= rs.RecordCount; i++)
                    {
                        PNs2.Add(new KeyValuePair<int, string>((int)rs.Fields.Item("GroupCode").Value, (string)rs.Fields.Item("GroupName").Value));

                        if (i < rs.RecordCount)
                        {
                            rs.MoveNext();
                        }
                    }

                    comboBox3.DataSource = new BindingSource(PNs2, null);
                    comboBox3.DisplayMember = "Value";
                    comboBox3.ValueMember = "Key";
                }

                listObj.Add(rs);
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
                if (comboBox2.SelectedValue.ToString().IndexOf("[") == -1)
                {

                    int id = (int)comboBox2.SelectedValue;
                    string desc = comboBox2.Text;

                    if (desc != "Manual")
                    {

                        rs.DoQuery($@"SELECT ""BeginStr"" + LEFT('00000',5-LEN(""NextNumber"")) +  CAST(""NextNumber"" AS nvarchar) FROM NNM1 WHERE ""ObjectCode"" = '2' AND ""Series"" = {id}");

                        textBox3.Text = (string)rs.Fields.Item(0).Value;
                    }
                    else
                    {
                        textBox3.Text = "";
                    }

                    listObj.Add(rs);

                }
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

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (!String.IsNullOrEmpty(textBox9.Text))
                {
                    string url = $"https://viacep.com.br/ws/{textBox9.Text}/json/";

                    getCep(url);

                }
                else {

                    MessageBox.Show("Informe um CEP para realizar a busca!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void getCep(string cep)
        {
            try
            {
                var client = new RestClient(cep);
                var request = new RestRequest();

                RestResponse response = client.Execute(request);

                if (response.IsSuccessful)
                {
                   CEP end = JsonConvert.DeserializeObject<CEP>(response.Content);

                    if (!String.IsNullOrEmpty(end.erro) && end.erro == "true")
                    {
                        throw new Exception("CEP inválido ou inexistente!");
                    } 

                    if (dataGridView2.Rows.Count == 0 || !dataGridView2.Rows[dataGridView2.Rows.Count - 1].IsNewRow)
                    {
                        dataGridView2.Rows.Add();
                    }

                    int rowIndex = dataGridView2.Rows.Count - 1;

                    DataGridViewRow row = dataGridView2.Rows[rowIndex];

                    row.Cells["TextBoxColumn2"].Value = end.logradouro;
                    row.Cells["TextBoxColumn3"].Value = end.complemento;
                    row.Cells["TextBoxColumn4"].Value = end.cep;
                    row.Cells["TextBoxColumn5"].Value = end.bairro;
                    row.Cells["TextBoxColumn6"].Value = "BR";
                    row.Cells["TextBoxColumn7"].Value = end.uf;
                    row.Cells["TextBoxColumn8"].Value = end.localidade;

                }
                else
                {
                    throw new Exception("CEP inválido ou inexistente!");
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
           
        }
    }
}
