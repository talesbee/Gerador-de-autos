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
using Color = System.Drawing.Color;

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

        private void InicialConfig(bool status)
        {
            tbInfracao.Visibility = Visibility.Hidden;
            tbInspecao.Visibility = Visibility.Hidden;
            tbNotificacao.Visibility = Visibility.Hidden;
            tbApreensao.Visibility = Visibility.Hidden;
            tbDeposito.Visibility = Visibility.Hidden;
            tbEmbargo.Visibility = Visibility.Hidden;
            tbGerar.Visibility = Visibility.Hidden;
            tbDados.Visibility = Visibility.Hidden;

            cbInfracao.IsEnabled = status;
            cbInspecao.IsEnabled = status;
            cbNotificacao.IsEnabled = status;
            // cbApreensao.IsEnabled = status;
            //cbDeposito.IsEnabled = status;
            //cbEmbargo.IsEnabled = status;

            cbInfracao.IsChecked = false;
            cbInspecao.IsChecked = false;
            cbNotificacao.IsChecked = false;
            // cbApreensao.IsChecked = status;
            //cbDeposito.IsChecked = status;
            // cbEmbargo.IsChecked = status;


            if (!status)
            {
                lbModelos.Content = "Primeiro selecione a pasta onde estão os Modelos";
                lbModelos.Foreground = System.Windows.Media.Brushes.Red;
                lbModelos.Visibility = Visibility.Visible;
            }
            else
            {
                lbModelos.Visibility = Visibility.Hidden;
            }

        }

        private bool arquivosEncontrados(FileInfo[] arquivos)
        {
            var encontrados = 0;
            foreach (FileInfo fileinfo in arquivos)
            {
                if (fileinfo.Name == "Auto de Infração.docx")
                {
                    cbInfracao.IsEnabled = true;
                    encontrados++;
                }
                else if (fileinfo.Name == "Auto de Inspeção.docx")
                {
                    cbInspecao.IsEnabled = true;
                    encontrados++;
                }
                else if (fileinfo.Name == "Notificação.docx")
                {
                    cbNotificacao.IsEnabled = true;
                    encontrados++;
                }
                else if (fileinfo.Name == "Termo de Apreensão.docx")
                {
                    cbApreensao.IsEnabled = true;
                    encontrados++;
                }
                else if (fileinfo.Name == "Termo de Depósito.doc")
                {
                    cbDeposito.IsEnabled = true;
                    encontrados++;
                }
                else if (fileinfo.Name == "Termo de Embargo.docx")
                {
                    cbEmbargo.IsEnabled = true;
                    encontrados++;

                }

            }
            if (cbEmbargo.IsEnabled == false && cbDeposito.IsEnabled == false && cbApreensao.IsEnabled == false && cbNotificacao.IsEnabled == false && cbInspecao.IsEnabled == false && cbInfracao.IsEnabled == false)
            {
                lbModelos.Visibility = Visibility.Visible;
            }

            if (encontrados > 0)
            {
                lbModelos.Content = "Modelos Localizados!";
                lbModelos.Foreground = System.Windows.Media.Brushes.Green;
                return true;
            }
            else
            {
                lbModelos.Content = "Nenhum Modelo correspondente!";
                lbModelos.Foreground = System.Windows.Media.Brushes.Red;
                return false;
            }

        }

        //Desativa as checkBox, verifica os arquivos na pasta selecionada, ativa as checkBox para cada arquivo encontrado.
        private void buscarModelos(object sender, RoutedEventArgs e)
        {
            InicialConfig(false);


            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\Users";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                DirectoryInfo diretorio = new DirectoryInfo(dialog.FileName);
                FileInfo[] Arquivos = diretorio.GetFiles("*.*");
                LocalModelos = dialog.FileName;
                lbCaminhoLocal.Content = LocalModelos;

                if (!arquivosEncontrados(Arquivos))
                    lbModelos.Content = "";
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
                btSalvar.IsEnabled = false;
                lbCaminhoSalvar.Content = LocalModelos;
                btGerarLaudos.IsEnabled = true;
                lbOndeSalvar.Visibility = Visibility.Collapsed;
            }
            else
            {
                btSalvar.IsEnabled = true;
                btGerarLaudos.IsEnabled = false;
                lbCaminhoSalvar.Content = ".";
                lbOndeSalvar.Visibility = Visibility.Visible;
            }
        }


        private void mostrarLivre(bool valor)
        {
            var modo = Visibility;
            if (valor)
            {
                modo = Visibility.Visible;
            }
            else
            {
                modo = Visibility.Hidden;
            }

            lbGeralLivre.Visibility = modo;
            cbGeralLivre.Visibility = modo;

            lbInfraLivre.Visibility = modo;
            cbInfraLivre.Visibility = modo;

            lbInspLivre.Visibility = modo;
            cbInspLivre.Visibility = modo;

            lbNotLivre.Visibility = modo;
            cbNotLivre.Visibility = modo;

            lbApreLivre.Visibility = modo;
            cbApreLivre.Visibility = modo;

            lbDepLivre.Visibility = modo;
            cbDepLivre.Visibility = modo;

            lbEmbLivre.Visibility = modo;
            cbEmbLivre.Visibility = modo;
        }

        private void checkLivre(object sender, RoutedEventArgs e)
        {

            InicialConfig(cbModoLivre.IsChecked.Value);
            btModelos.IsEnabled = !cbModoLivre.IsChecked.Value;
            cbMesmoLocal.Visibility = Visibility.Collapsed;

            // mostrarLivre(cbModoLivre.IsChecked.Value);
        }

        private void clickInformacao(object sender, MouseButtonEventArgs e)
        {
            Console.WriteLine("clicou na foto");
        }

        private void calculoMulta(object sender, DependencyPropertyChangedEventArgs e)
        {

        }

        private void salvarDoc(DocX doc, string local, string laudo)
        {
            Console.WriteLine("Salvando!");
            doc.SaveAs(local + laudo);
        }

        private void replaceGeral(DocX doc)
        {
            doc.ReplaceText("#NomeAutuado", txNome.Text);
            doc.ReplaceText("#CNPJ_CPFAutuado", txCnpjCpf.Text);
            doc.ReplaceText("#DataDoAuto", dpData.Text);
            doc.ReplaceText("#horarioInfração", txHora.Text);
            doc.ReplaceText("#NomeDaMae", txFiliacao.Text);
            doc.ReplaceText("#Atividade", txAtividade.Text);
            doc.ReplaceText("#EnderecoEmpreendimento", txEnderecoComercial.Text);
            doc.ReplaceText("#MunicipioEmpreendimento", txMunicipio.Text);
            doc.ReplaceText("#UfEmpreendimento", txUf.Text);
            doc.ReplaceText("#AreaEmpreendimento", txAreaEmpreendimento.Text);
            doc.ReplaceText("#Coordenada", txCoordenada.Text);
            doc.ReplaceText("#Latitude ", txLat.Text);
            doc.ReplaceText("#Longitude", txLong.Text);
            doc.ReplaceText("#EnderecoCorrespondencia", txEnderecoCorrespondencia.Text);
            doc.ReplaceText("#MunicipioCorrespondencia", txMunicipioCorrespondencia.Text);
            doc.ReplaceText("#UfCorrespondencia", txUfCorrespondencia.Text);
            doc.ReplaceText("#CEPCorrespondencia", txCEP.Text);
            doc.ReplaceText("#Telefone", txTelefone.Text);
            doc.ReplaceText("#RepresentanteLegal", txNomeRepresentante.Text);
        }

        private void replaceInfra(DocX infra)
        {
            infra.ReplaceText("#nAutoInfracao", txInfraNumero.Text);
            infra.ReplaceText("#AreaDesmate", txInfraAreaDesmate.Text);
            infra.ReplaceText("#RelatorioTecnico", txInfraRelatorio.Text);
            infra.ReplaceText("#AreaExplocacaoSeletiva", txInfraAreaExploracao.Text);
            infra.ReplaceText("#Descricao", lbInfraOcorrencia.Text);
            infra.ReplaceText("#ValorMultaExtenso", txInfraMultaExtenso.Text);
            infra.ReplaceText("#DispositivosLegais", txInfraDispositivosInfri.Text);
            infra.ReplaceText("#DescicaoMulta", lbInfraDescriMulta.Text);
            infra.ReplaceText("#ValorMulta", txInfraMulta.Text);
        }

        private void replaceNot(DocX Not)
        {
            Not.ReplaceText("#nNotificacao", txNotNumero.Text);
            Not.ReplaceText("#protocoloNotificacao", txNotProcesso.Text);
            Not.ReplaceText("#Objetivo", lbNotObjetivo.Text);
            Not.ReplaceText("#TxtNotif", lbNotNotificacao.Text);
            Not.ReplaceText("#AreaFDesmate", txNotFlorDesmate.Text);
            Not.ReplaceText("#ReposicaoFloresta", txNotFlorRep.Text);
            Not.ReplaceText("#AreaCDesmate", txNotCerDesmate.Text);
            Not.ReplaceText("#ReposicaoCerrado", txNotCerRep.Text);
            Not.ReplaceText("#TotalHectareDesmate", txNotTDesmate.Text);
            Not.ReplaceText("#TotalReposicaoM3", txNotTRep.Text);
        }

        private void replaceInspe(DocX Insp)
        {
            Insp.ReplaceText("#nAutoinspecao", txInspNumero.Text);
            Insp.ReplaceText("#Objetivo", lbInspObj.Text);
            Insp.ReplaceText("#Constatações", lbInspConsta.Text);
        }

        private void replaceDocumento(string local)
        {
            DocX documento;
            string nome;
            if (cbInfracao.IsChecked.Value)
            {
                string arq = local + "\\Auto de Infração.docx";
                documento = DocX.Load(arq);
                replaceGeral(documento);
                replaceInfra(documento);
                Console.WriteLine(local);
                nome = @"Auto de Infração - " + txInfraNumero.Text + ".docx";
                arq = local + "\\";
                salvarDoc(documento, arq, nome);
            }
            if (cbNotificacao.IsChecked.Value)
            {
                string arq = local + "\\Notificação.docx";
                documento = DocX.Load(arq);
                replaceGeral(documento);
                replaceNot(documento);
                Console.WriteLine(local);
                nome = @"Notificação - " + txNotNumero.Text + ".docx";
                arq = local + "\\";
                salvarDoc(documento, arq, nome);
            }
            if (cbInspecao.IsChecked.Value)
            {
                string arq = local + "\\Auto de Inspeção.docx";
                documento = DocX.Load(arq);
                replaceGeral(documento);
                replaceInspe(documento);
                Console.WriteLine(local);
                nome = @"Auto de Inspeção - " + txInspNumero.Text + ".docx";
                arq = local + "\\";
                salvarDoc(documento, arq, nome);
            }

        }
        private void gerarLaudos(object sender, RoutedEventArgs e)
        {
            string local = lbCaminhoSalvar.Content.ToString();
            if (!local.Equals("."))
            {
                replaceDocumento(local);
            }
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
