using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xceed.Words.NET;

namespace Gerador_de_autos
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string caminho = @"base.docx";
        public MainWindow()
        {
            InitializeComponent();

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            using(var documento = DocX.Load(caminho))
            {
                documento.ReplaceText("#nome", txNome.Text);
                documento.ReplaceText("#numeroInfra", txNumeroInfra.Text);
                documento.ReplaceText("#data", txData.Text);
                documento.ReplaceText("#cnpjcpf", txCNPJCPF.Text);
                documento.ReplaceText("#endereco", txEndereco.Text);
                documento.ReplaceText("#municipio", txMunicipio.Text);
                documento.ReplaceText("#uf", txUf.Text);
                documento.ReplaceText("#area", txArea.Text);
                documento.ReplaceText("#hora", txHorario.Text);
                documento.ReplaceText("#coordenada", txCoordenada.Text);
                documento.ReplaceText("#lat", txLat.Text);
                documento.ReplaceText("#long", txLong.Text);
                documento.ReplaceText("#endCor", txEnderecoCorrespondencias.Text);
                documento.ReplaceText("#Cormunicipio", txMunicipioCorrespondencias.Text);
                documento.ReplaceText("#cep", txCEP.Text);
                documento.ReplaceText("#corUf", txUfCorrespondencias.Text);
                documento.ReplaceText("#telefone", txTelefone.Text);
                documento.ReplaceText("#aDes", txAreaAfetadaDes.Text);
                documento.ReplaceText("#aExp", txAreaAfetadaExp.Text);
                documento.ReplaceText("#NumRelatDesmate", txNumRelatDesmate.Text);
                documento.ReplaceText("#NumRelatExp", txNumRelatExp.Text);
                documento.ReplaceText("#nomFantasia", txNomeFant.Text);
                documento.ReplaceText("#atividade", txAtividade.Text);
                documento.ReplaceText("#endCor", txEnderecoCorrespondencias.Text);
                documento.ReplaceText("#representante", txEnderecoCorrespondencias.Text);
                documento.ReplaceText("#base", txBase.Text);




                string nome = @"Auto - " + txNome.Text + ".docx";
                documento.SaveAs(nome);
            }
        }

        private void formatarCpfCnpj(object sender, RoutedEventArgs e)
        {
            string dados = txCNPJCPF.Text;
            Console.WriteLine("cpf");

            if (dados.Length <= 11)
            {
                if (dados.Length == 11)
                {
                    txCNPJCPF.Text = cpf(txCNPJCPF.Text);
                }
                else if (dados.Length == 14)
                {
                    txCNPJCPF.Text = cnpj(txCNPJCPF.Text);
                }
            }


        }
        public string cpf(string cpf)
        {
            return Convert.ToUInt64(cpf).ToString(@"000\.000\.000\-00");
        }
        public string cnpj(string cnpj)
        {
            return Convert.ToUInt64(cnpj).ToString(@"00\.000\.000\/0000\-00");
        }

        private void formatarData(object sender, RoutedEventArgs e)
        {
            string tamanho = txData.Text;
            if(tamanho.Length == 7)
                txData.Text = data(txData.Text);
        }
        public string data(string cnpj)
        {
            return Convert.ToUInt64(cnpj).ToString(@"00\/00\/0000");
        }
    }
}
