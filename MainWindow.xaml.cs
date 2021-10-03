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
using Microsoft.WindowsAPICodePack.Dialogs;
using System.IO;

namespace Gerador_de_autos
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        String LocalModelos;
        public MainWindow()
        {

            InitializeComponent();

        }

        private void btModelos_Copy_Click(object sender, RoutedEventArgs e)
        {

        }

        //Desativa as checkBox, verifica os arquivos na pasta selecionada, ativa as checkBox para cada arquivo encontrado.
        private void buscarModelos(object sender, RoutedEventArgs e)
        {
            cbInfracao.IsEnabled = false;
            cbInspecao.IsEnabled = false;
            cbNotificacao.IsEnabled = false;
            cbApreensao.IsEnabled = false;
            cbDeposito.IsEnabled = false;
            cbEmbargo.IsEnabled = false;

            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\Users";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                DirectoryInfo diretorio = new DirectoryInfo(dialog.FileName);
                FileInfo[] Arquivos = diretorio.GetFiles("*.*");
                LocalModelos = dialog.FileName;
                lbCaminhoLocal.Content = LocalModelos;

                foreach (FileInfo fileinfo in Arquivos)
                {
                    if (fileinfo.Name == "Auto de Infração.docx")
                    {
                        cbInfracao.IsEnabled = true;
                        lbModelos.Visibility = Visibility.Collapsed;
                    }
                    else if (fileinfo.Name == "Auto de Inspeção.docx")
                    {
                        cbInspecao.IsEnabled = true;
                        lbModelos.Visibility = Visibility.Collapsed;
                    }
                    else if (fileinfo.Name == "Notificação.docx")
                    {
                        cbNotificacao.IsEnabled = true;
                        lbModelos.Visibility = Visibility.Collapsed;
                    }
                    else if (fileinfo.Name == "Termo de Apreensão.docx")
                    {
                        cbApreensao.IsEnabled = true;
                        lbModelos.Visibility = Visibility.Collapsed;
                    }
                    else if (fileinfo.Name == "Termo de Depósito.doc")
                    {
                        cbDeposito.IsEnabled = true;
                        lbModelos.Visibility = Visibility.Collapsed;
                    }
                    else if (fileinfo.Name == "Termo de Embargo.docx")
                    {
                        cbEmbargo.IsEnabled = true;
                        lbModelos.Visibility = Visibility.Collapsed;
                    }
                   
                    
                }
                if (cbEmbargo.IsEnabled == false && cbDeposito.IsEnabled == false && cbApreensao.IsEnabled == false && cbNotificacao.IsEnabled == false && cbInspecao.IsEnabled == false && cbInfracao.IsEnabled == false)
                {
                    lbModelos.Visibility = Visibility.Visible;
                }
            }

        }
        //Função que ativa ou desativa as Tab.
        private void visibleComponent(TabItem labol, bool visivel)
        {
            if (visivel)
            {
                labol.Visibility = Visibility.Visible;
                tbDados.Visibility = Visibility.Visible;
                tbGerar.Visibility = Visibility.Visible;
            }
            else
            {
                labol.Visibility = Visibility.Collapsed;
            }
            
        }

        private void CheckBoxClick(object sender, RoutedEventArgs e)
        {
            //Ativa ou desativa as TabItem de acordo com as checkBox selecionadas.
            visibleComponent(tbApreensao, cbApreensao.IsChecked.Value);
            visibleComponent(tbInspecao, cbInspecao.IsChecked.Value);
            visibleComponent(tbInfracao, cbInfracao.IsChecked.Value);
            visibleComponent(tbEmbargo, cbEmbargo.IsChecked.Value);
            visibleComponent(tbDeposito, cbDeposito.IsChecked.Value);
            visibleComponent(tbNotificacao, cbNotificacao.IsChecked.Value);

            if (cbApreensao.IsChecked == false && cbDeposito.IsChecked == false && cbEmbargo.IsChecked == false && cbInfracao.IsChecked == false && cbInspecao.IsChecked == false && cbNotificacao.IsChecked == false)
            {
                tbGerar.Visibility = Visibility.Collapsed;
                tbDados.Visibility = Visibility.Collapsed;
            }
        }

        private void gerarLaudos(object sender, RoutedEventArgs e)
        {
            if (cbModoLivre.IsChecked.Value)
            {
                CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                dialog.InitialDirectory = "C:\\Users";
                dialog.IsFolderPicker = true;
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    DirectoryInfo diretorio = new DirectoryInfo(dialog.FileName);
                    FileInfo[] Arquivos = diretorio.GetFiles("*.*");
                    LocalModelos = dialog.FileName;
                }
                    
            }
            else
            {

            }
        }

        private void buscarSalvar(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\Users";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                LocalModelos = dialog.FileName;
                lbCaminhoSalvar.Content = LocalModelos;
                btGerarLaudos.IsEnabled = true;
                lbOndeSalvar.Visibility = Visibility.Collapsed;
            }
            else
            {
                btGerarLaudos.IsEnabled = false;
                lbOndeSalvar.Visibility = Visibility.Visible;
            }
        }

        private void usarLocalModelos(object sender, RoutedEventArgs e)
        {
            if (cbMesmoLocal.IsChecked.Value)
            {
                btSalvar1.IsEnabled = false;
                lbCaminhoSalvar.Content = LocalModelos;
                btGerarLaudos.IsEnabled = true;
                lbOndeSalvar.Visibility = Visibility.Collapsed;
            }
            else
            {
                btSalvar1.IsEnabled = true;
                btGerarLaudos.IsEnabled = false;
                lbCaminhoSalvar.Content = ".";
                lbOndeSalvar.Visibility = Visibility.Visible;
            }
        }

      

        private void mudarLivre(bool valor)
        {

            cbMesmoLocal.IsEnabled = !valor;
            btModelos.IsEnabled = !valor;
            visibleComponent(tbApreensao, valor);
            visibleComponent(tbInspecao, valor);
            visibleComponent(tbInfracao, valor);
            visibleComponent(tbEmbargo, valor);
            visibleComponent(tbDeposito, valor);
            visibleComponent(tbNotificacao, valor);
            
            if (valor)
            {
                lbSelecionarLivre.Visibility = Visibility.Visible;
                lbModelos.Visibility = Visibility.Collapsed;
            }
            else
            {
                tbDados.Visibility = Visibility.Collapsed;
                tbGerar.Visibility = Visibility.Collapsed;
                lbSelecionarLivre.Visibility = Visibility.Collapsed;
                lbModelos.Visibility = Visibility.Visible;
            }
        }

        private void checkLivre(object sender, RoutedEventArgs e)
        {
            mudarLivre(cbModoLivre.IsChecked.Value);
        }

        private void clickInformacao(object sender, MouseButtonEventArgs e)
        {
            Console.WriteLine("clicou na foto");
        }

        private void calculoMulta(object sender, DependencyPropertyChangedEventArgs e)
        {

        }


        /*
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
*/
    }
}
