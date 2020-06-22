using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.Security;
using System.Drawing;
using System.Threading;

namespace DataStack
{
    public partial class Form1 : Form
    {
        #region Variáveis
        FileHelpersArquivo ofileHelpers = new FileHelpersArquivo();
        string sArquivo;        
        string sMessagem;
        string[] excelFiles;
        bool bErro = false;
        bool resul = true;
        #endregion
        public Form1() => InitializeComponent();

        #region btnImport
        private void btnImport_Click(object sender, EventArgs e)
        {

            lbPronto.Visible = false;
            Progressbar.Value = 0;

            OpenFileDialog oAbreArq = new OpenFileDialog
            {
                Title = "Selecione o Arquivo",
                Multiselect = true
            };
            
          

            try
            {
                string sRecebeArquivo = RetornaOpenFile();
                excelFiles = sRecebeArquivo.Remove(sArquivo.Length - 1).Split(';');
                sArquivo = null;
            }
            catch (Exception ex)
            {
                bErro = true;
                sMessagem = ex.Message != "StartIndex não pode ser menor que zero.\r\nNome do parâmetro: startIndex" ? ex.Message : "Nenhum arquivo foi selecionado.";
                MessageBox.Show(sMessagem);
                string[] objeto = { "RetornaTrafego", sMessagem, "1" };
                Thread Plano = new Thread(ofileHelpers.reportLogging);
                Plano.Start(objeto);
            }            

            btnImportar.BackColor = Color.LightBlue;
            btnImportar.ForeColor = Color.DimGray;
            btnImportar.Text = "Carregando...";
            btnImportar.Enabled = false;
            ofileHelpers.ProgressoMudou += new EventHandler<ProgressChangedEventArgs>(AvisoDoProgresso);

            Progressbar.Visible = true;

            try
            {
                if (!bErro)
                    resul = ofileHelpers.RetornaDados(excelFiles);

                excelFiles = null;
                
            }
            catch (Exception ex)
            {
                sMessagem = ex.Message;
                MessageBox.Show(sMessagem);
                string[] objeto = { "btnImport_Click", sMessagem, "1" };
                Thread Plano = new Thread(ofileHelpers.reportLogging);
                Plano.Start(objeto);                
            }

            Progressbar.Visible = false;            
            lbPronto.Visible = true;
            if (!resul)
                MessageBox.Show("Os arquivos selecionados não possuem as planilhas referenciadas");

            Progressbar.Visible = false;
            btnImportar.BackColor = Color.SteelBlue;
            btnImportar.ForeColor = Color.White;
            btnImportar.Text = "Importar";
            btnImportar.Enabled = true;
            bErro = false;
            GC.Collect();

        }
        #endregion

        #region Aviso ProgressBar
        public void AvisoDoProgresso(object sender, ProgressChangedEventArgs e)
        {
            if (e != null)
            {
                if (Progressbar.InvokeRequired)
                {
                    this.Invoke(new EventHandler<ProgressChangedEventArgs>(AvisoDoProgresso), sender, e);
                    return;
                }
                Progressbar.Value = e.ProgressPercentage;
            }
        }
        #endregion

        #region Retorno Nome Arquivos selecionados
        public string RetornaOpenFile()
        {
            sArquivo = "";
            OpenFileDialog oAbreArq = new OpenFileDialog
            {
                Title = "Selecione o Arquivo",
                Multiselect = true
            };

            if (oAbreArq.ShowDialog() == DialogResult.OK)
            {
                foreach (var item in oAbreArq.FileNames)
                {
                    try
                    {
                        if (item != "System.String[]")
                        {
                            sArquivo += item.Replace("XLS", "xls") + ";";
                        }
                    }
                    catch (SecurityException ex)
                    {
                        // O usuário não possui permissões apropriadas para ler arquivos, descobrir caminhos etc.
                        MessageBox.Show("Erro de segurança. Entre em contato com seu administrador para obter detalhes.\n\n" +
                            "Erro: " + ex.Message + "\n\n" +
                            "Detalhes (envie ao suporte):\n\n" + ex.StackTrace
                        );
                    }
                }
            }
            return sArquivo;
        }
        #endregion
    }
}
