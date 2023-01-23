using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TowkingCadastro
{
    public partial class Form1 : Form
    {
        public static BindingList<Cliente> clientes;
        public static string path = $@"{Directory.GetCurrentDirectory()}\clientes.xlsx";
        public Form1()
        {
            InitializeComponent();
            clientes = new BindingList<Cliente>();
            textBoxNascimento.PlaceholderText = DateTime.Now.ToString("yyyy-MM-dd");
            LoadClients();
        }

        private bool validateFields()
        {
            if (textBoxNome.Text == "" ||
                textBoxSobrenome.Text == "" ||
                textBoxEmail.Text == "" ||
                textBoxBairro.Text == "" ||
                textBoxCelular.Text == "" ||
                textBoxCep.Text == "" ||
                textBoxCpf.Text == "" ||
                textBoxModeloVeiculo.Text == "" ||
                textBoxRua.Text == "" ||
                textBoxTelefone.Text == "" ||
                textBoxCidade.Text == "" ||
                textBoxEstado.Text == "" ||
                textBoxTipoVeiculo.Text == "")
            {
                return false;
            }
            else
            {
                return true;
            }

        }

        private void LoadClients()
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                FileInfo fileInfo = new FileInfo(path);
                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;

                for (int x = 2, y = 1; x <= rows; y = 1, x++)
                {
                    if (worksheet.Cells[x, 1].Value == null)
                        break;
                    Cliente clienteTemp = new Cliente();
                   
                    clienteTemp.Nome = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Sobrenome = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Nascimento = DateTime.FromOADate(double.Parse(worksheet.Cells[x, y].Value.ToString())).ToString("yyyy-MM-dd"); y++;
                    clienteTemp.Telefone = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Celular = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.CPF = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Email = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.TipoVeiculo = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.ModeloVeiculo = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Rua = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Estado = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Cidade = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.CEP = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Bairro = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.isempresa = worksheet.Cells[x, y].Value.ToString() == "true" ? true : false; y++;
                    clientes.Add(clienteTemp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Falha ao carregar as informações!\n{ex.Message}\n{ex.StackTrace}", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (validateFields())
            {
                try
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    FileInfo fileInfo = new FileInfo(path);
                    ExcelPackage package = new ExcelPackage(fileInfo);
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    Cliente novoCliente = new Cliente();
                    novoCliente.Nome = textBoxNome.Text;
                    novoCliente.Sobrenome = textBoxSobrenome.Text;
                    novoCliente.Nascimento = textBoxNascimento.Text;
                    novoCliente.Telefone = textBoxTelefone.Text;
                    novoCliente.Celular = textBoxCelular.Text;
                    novoCliente.CPF = textBoxCpf.Text;
                    novoCliente.Email = textBoxEmail.Text;
                    novoCliente.TipoVeiculo = textBoxTipoVeiculo.Text;
                    novoCliente.ModeloVeiculo = textBoxModeloVeiculo.Text;
                    novoCliente.Rua = textBoxRua.Text;
                    novoCliente.Cidade = textBoxCidade.Text;
                    novoCliente.Estado = textBoxEstado.Text;
                    novoCliente.Bairro = textBoxBairro.Text;
                    novoCliente.CEP = textBoxCep.Text;
                    novoCliente.isempresa = checkBox1.Checked ? true : false;

                    clientes.Add(novoCliente);

                    int rows = worksheet.Dimension.Rows;
                    worksheet.Cells[rows + 1, 1].Value = novoCliente.Nome;
                    worksheet.Cells[rows + 1, 2].Value = novoCliente.Sobrenome;
                    worksheet.Cells[rows + 1, 3].Value = novoCliente.Nascimento;
                    worksheet.Cells[rows + 1, 4].Value = novoCliente.Telefone;
                    worksheet.Cells[rows + 1, 5].Value = novoCliente.Celular;
                    worksheet.Cells[rows + 1, 6].Value = novoCliente.CPF;
                    worksheet.Cells[rows + 1, 7].Value = novoCliente.Email;
                    worksheet.Cells[rows + 1, 8].Value = novoCliente.TipoVeiculo;
                    worksheet.Cells[rows + 1, 9].Value = novoCliente.ModeloVeiculo;
                    worksheet.Cells[rows + 1, 10].Value = novoCliente.Rua;
                    worksheet.Cells[rows + 1, 11].Value = novoCliente.Estado;
                    worksheet.Cells[rows + 1, 12].Value = novoCliente.Cidade;
                    worksheet.Cells[rows + 1, 13].Value = novoCliente.CEP;
                    worksheet.Cells[rows + 1, 14].Value = novoCliente.Bairro;                    
                    worksheet.Cells[rows + 1, 15].Value = novoCliente.isempresa ? "true" : "false";
                    //worksheet.Cells[rows + 1, 16].Value = "0015e000003VvtNAAS";

                    package.Save();

                    MessageBox.Show("Cadastrado com sucesso!", "Sucesso!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Falha ao salvar as informações!\n{ex.Message}\n{ex.StackTrace}", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Por favor preencha todas as informações!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void buttonRead_Click(object sender, EventArgs e)
        {
            FormListaClientes frmLista = new FormListaClientes();
            frmLista.Show();
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string path = $@"{Directory.GetCurrentDirectory()}\dados_clientes.xlsx";

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                FileInfo fileInfo = new FileInfo(path);
                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;

                for (int x = 2, y = 1; x < rows; y = 1, x++)
                {
                    Cliente clienteTemp = new Cliente();
                    if (worksheet.Cells[x, y].Value == null)
                        break;

                    clienteTemp.Nome = worksheet.Cells[x, y].Value?.ToString(); y++;
                    clienteTemp.Sobrenome = worksheet.Cells[x, y].Value?.ToString(); y++;
                    clienteTemp.Nascimento = DateTime.FromOADate(double.Parse(worksheet.Cells[x, y].Value.ToString())).ToString("yyyy-MM-dd"); ; y++;
                    clienteTemp.Telefone = worksheet.Cells[x, y].Value?.ToString(); y++;
                    clienteTemp.Celular = worksheet.Cells[x, y].Value?.ToString(); y++;
                    clienteTemp.CPF = worksheet.Cells[x, y].Value?.ToString(); y++;
                    if (ExistsCPF(clienteTemp.CPF))
                        continue;
                    clienteTemp.Email = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.TipoVeiculo = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.ModeloVeiculo = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Rua = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Estado = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Cidade = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.CEP = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.Bairro = worksheet.Cells[x, y].Value.ToString(); y++;
                    clienteTemp.isempresa = worksheet.Cells[x, y].Value.ToString() == "true" ? true : false; y++;
                    //y++;  
                    clientes.Add(clienteTemp);
                }
                MessageBox.Show("Adicionado com sucesso!", "Sucesso!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Falha ao carregar as informações!\n{ex.Message}\n{ex.StackTrace}", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
        }

        private bool ExistsCPF(string CPF)
        {
            foreach(Cliente cliente in clientes)
            {
                if (cliente.CPF == CPF)                
                    return true; 
            }
            return false;            
        }
    
    }
}
