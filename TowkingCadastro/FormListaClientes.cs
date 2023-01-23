using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TowkingCadastro
{
    public partial class FormListaClientes : Form
    {
        BindingList<Cliente> clientes;      
        public FormListaClientes()
        {
            InitializeComponent();
            this.clientes = Form1.clientes;
            PopulateDatGridView();           
        }

        private void PopulateDatGridView()
        {
            dataGridViewClientes.DataSource = clientes;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                FileInfo fileInfo = new FileInfo(Form1.path);
                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;
                
                for (int x = 2; x <= rows+1; x++)
                {
                    worksheet.DeleteRow(2);
                }
                package.Save();

                rows = 1;

                foreach(Cliente cliente in clientes)
                {
                    worksheet.Cells[rows + 1, 1].Value = cliente.Nome;
                    worksheet.Cells[rows + 1, 2].Value = cliente.Sobrenome;
                    worksheet.Cells[rows + 1, 3].Value = cliente.Nascimento;
                    worksheet.Cells[rows + 1, 4].Value = cliente.Telefone;
                    worksheet.Cells[rows + 1, 5].Value = cliente.Celular;
                    worksheet.Cells[rows + 1, 6].Value = cliente.CPF;
                    worksheet.Cells[rows + 1, 7].Value = cliente.Email;
                    worksheet.Cells[rows + 1, 8].Value = cliente.TipoVeiculo;
                    worksheet.Cells[rows + 1, 9].Value = cliente.ModeloVeiculo;
                    worksheet.Cells[rows + 1, 10].Value = cliente.Rua;
                    worksheet.Cells[rows + 1, 11].Value = cliente.Estado;
                    worksheet.Cells[rows + 1, 12].Value = cliente.Cidade;
                    worksheet.Cells[rows + 1, 13].Value = cliente.CEP;
                    worksheet.Cells[rows + 1, 14].Value = cliente.Bairro;
                    worksheet.Cells[rows + 1, 15].Value = cliente.isempresa ? "true" : "false";
                    rows++;
                }
               
                package.Save();
                MessageBox.Show("Salvo com sucesso!", "Sucesso!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Falha ao salvar as informações!\n{ex.Message}\n{ex.StackTrace}", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }           

        } 

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                var row = dataGridViewClientes.SelectedRows[0];
                string cpf = row.Cells[5].Value.ToString();
                int index = -1;
                for (int x = 0; x < clientes.Count; x++)
                {
                    if (clientes[x].CPF == cpf)
                        index = x;
                }
                clientes.RemoveAt(index);              
                dataGridViewClientes.Refresh();
                PopulateDatGridView();
                dataGridViewClientes.Refresh();
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Falha ao editar as informações!\n{ex.Message}\n{ex.StackTrace}", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
        }
    }
}
