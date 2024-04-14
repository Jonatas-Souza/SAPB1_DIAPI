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
    public partial class TablesMD : Form
    {
        //List<object> listObj = new List<object>();


        public TablesMD()
        {
            SAPbobsCOM.Recordset rs = DiApi.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                InitializeComponent();

                this.FormClosing += FormClose;

                foreach (SAPbobsCOM.BoUTBTableType item in Enum.GetValues(typeof(BoUTBTableType)))
                {
                    comboBox1.Items.Add(item);
                }

                foreach (SAPbobsCOM.BoFieldTypes item in Enum.GetValues(typeof(BoFieldTypes)))
                {
                    comboBox2.Items.Add(item);
                }

                foreach (SAPbobsCOM.BoFldSubTypes item in Enum.GetValues(typeof(BoFldSubTypes)))
                {
                    comboBox3.Items.Add(item);
                }

                foreach (SAPbobsCOM.UDFLinkedSystemObjectTypesEnum item in Enum.GetValues(typeof(UDFLinkedSystemObjectTypesEnum)))
                {
                    comboBox4.Items.Add(item);
                }

                
                rs.DoQuery(@"select ""TableName"" from OUTB");

                if (rs.RecordCount > 0)
                {
                    for (int i = 1; i < rs.RecordCount; i++)
                    {
                        comboBox5.Items.Add(rs.Fields.Item(0).Value);

                        if (i < rs.RecordCount)
                            rs.MoveNext();

                    }
                }

                //listObj.Add(rs);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }

        }

        private void FormClose(object sender, FormClosingEventArgs e)
        {
            try
            {
                //foreach (object item in listObj)
                //{
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                //}

                Menu menu = new Menu();

                menu.Enabled = true;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


        }


        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SAPbobsCOM.UserTablesMD tablesMD = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oUserTables);

            try
            {
                

                tablesMD.TableName = textBox10.Text;
                tablesMD.TableDescription = textBox1.Text;
                tablesMD.TableType = (BoUTBTableType)comboBox1.SelectedItem;

                int ret = tablesMD.Add();

                if (ret != 0)
                {
                    throw new Exception($"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                }

                MessageBox.Show("Tabela de usuário criada com sucesso!");

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tablesMD);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            SAPbobsCOM.UserFieldsMD fieldsMD = DiApi.oCompany.GetBusinessObject(BoObjectTypes.oUserFields);

            try
            {
                fieldsMD.TableName = textBox4.Text;
                fieldsMD.Name = textBox3.Text;
                fieldsMD.Description = textBox2.Text;
                fieldsMD.Size = int.Parse(textBox6.Text);
                fieldsMD.EditSize = fieldsMD.Size;
                fieldsMD.Type = (BoFieldTypes)comboBox2.SelectedItem;
                if (comboBox3.SelectedItem != null)
                {
                    fieldsMD.SubType = (BoFldSubTypes)comboBox3.SelectedItem;
                }
                if (comboBox4.SelectedItem != null)
                {
                    fieldsMD.LinkedSystemObject = (UDFLinkedSystemObjectTypesEnum)comboBox4.SelectedItem;
                }
                if (comboBox5.SelectedItem != null)
                {
                    fieldsMD.LinkedTable = comboBox5.SelectedItem.ToString();
                }

                if (dataGridView1.RowCount > 0)
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row != null && !row.IsNewRow)
                        {
                            if (row.Cells["Column1"].Value != null && row.Cells["Column2"].Value != null)
                            {
                                fieldsMD.ValidValues.Value = row.Cells["Column1"].Value.ToString();
                                fieldsMD.ValidValues.Description = row.Cells["Column2"].Value.ToString();
                                fieldsMD.ValidValues.Add();
                            }
                        }
                    }
                }

                fieldsMD.DefaultValue = textBox5.Text;

                if (checkBox1.Checked)
                {
                    fieldsMD.Mandatory = BoYesNoEnum.tYES;
                }
                else
                {
                    fieldsMD.Mandatory = BoYesNoEnum.tNO;
                }

                int ret = fieldsMD.Add();

                if (ret != 0)
                {
                    throw new Exception($"({DiApi.oCompany.GetLastErrorCode()}) - {DiApi.oCompany.GetLastErrorDescription()}");
                }

                MessageBox.Show("Campo de usuário criado com sucesso!");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(fieldsMD);
            }
        }
    }
}
