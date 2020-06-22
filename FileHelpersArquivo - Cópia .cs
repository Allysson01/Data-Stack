using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.IO;
using System.ComponentModel;
using System.Windows.Forms;
using NPOI.XSSF.UserModel;
using System.Threading;
using System.Threading.Tasks;

namespace DataStack
{

    public class FileHelpersArquivo
    {
        #region Variáveis       
        XSSFWorkbook Templatebook = null;
        HSSFWorkbook PlanilhaCorrente = null;
        HSSFWorkbook PlanilhaTrafego = null;
        FileStream arquivoXLS = null;
        FileStream Template = null;
        ISheet Planos = null;
        ISheet Trafego = null;
        ISheet DemServicos = null;
        string sPlanilhaA = "";
        int CountLinhasServico = 0;
        int CountLinhasFatura = 0;
        int NumberOfSheets = 0;
        #endregion

        #region Progressbar
        protected virtual void OnProgressoMudou(int porcentagem)
        {
            // Faz uma cópia do evento para uma variável local, para
            // evitar problemas com várias threads
            EventHandler<ProgressChangedEventArgs> evento = this.ProgressoMudou;

            if (evento != null)
            {
                // Cria uma instância do objeto que carrega a informação
                ProgressChangedEventArgs p = new ProgressChangedEventArgs(porcentagem, null);
                // Dispara o evento!
                evento(this, p);
            }
        }

        // Evento que permite receber notificações sobre o processo X
        public event EventHandler<ProgressChangedEventArgs> ProgressoMudou;
        #endregion

        #region RetornaDados
        public bool RetornaDados(string[] excelFiles)
        {

            //Recupera a planilha template
            Template = new FileStream(Path.Combine(Application.StartupPath, @"Template.xlsx"), FileMode.Open, FileAccess.ReadWrite);

            //Carrega o template
            Templatebook = new XSSFWorkbook(Template);

            //Carrega as planilhas do template
            Planos = Templatebook.GetSheet("Planos");
            Trafego = Templatebook.GetSheet("Tráfego");

            this.OnProgressoMudou(1);

            GC.Collect();

            for (int i = 0; i < excelFiles.Length; i++)
            {
                CarregaProgresso(i);

                try
                {
                    arquivoXLS = new FileStream(excelFiles[i], FileMode.Open, FileAccess.ReadWrite);

                    // recupera o Workbook do arquivo
                    PlanilhaCorrente = new HSSFWorkbook(arquivoXLS);

                    arquivoXLS.Close();

                    //Carrega as planilhas dos arquivos selecionados                    
                    DemServicos = PlanilhaCorrente.GetSheet("Demonstrativo dos Serviços");
                    NumberOfSheets = PlanilhaCorrente.NumberOfSheets;
                    PlanilhaTrafego = PlanilhaCorrente;

                    Thread t1 = new Thread(new ThreadStart(RetornaTrafego));
                    t1.Name = "trafego";
                    
                    t1.Start();                    

                    RetornaPlano();

                    Thread.Sleep(100);

                    CarregaProgresso(i);
                }
                catch (Exception ex)
                {
                    string Ex = ex.Message;
                }                
            }
            
            this.OnProgressoMudou(100);

            try
            {                
                //Endereça e nomeia o arquivo
                var localizacaoArquivo = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Planilhas";
                bool exists = Directory.Exists(localizacaoArquivo);
                if (!exists)
                    Directory.CreateDirectory(localizacaoArquivo);

                string nameArquivo = localizacaoArquivo + @"\Documento " + "(" + DateTime.Now.ToString("dd-MM-yyyy") + ") " + DateTime.Now.ToString("HH: mm").Replace(" ", "").Replace(" / ", ".").Replace(":", "h") + "NEW.xlsx";

                //Salva o Arquivo
                FileStream file = new FileStream(nameArquivo, FileMode.Create);

                Templatebook.Write(file);

                file.Close();

                //Executa o arquivo
                System.Diagnostics.Process.Start(nameArquivo);

            }
            catch (Exception ex)
            {
                string Ex = ex.Message;
            }
            finally
            {
                excelFiles = null;
                Planos = null;
                Trafego = null;
                Template = null;
                //Templatebook = null;
                arquivoXLS = null;
                GC.Collect();
            }
            return true;
        }
        #endregion

        void CarregaProgresso(int i)
        {
            try
            {
                var Icount = (i).ToString() + "5";
                this.OnProgressoMudou(Convert.ToInt32(Icount));
            }
            catch (Exception ex)
            {
                string Ex = ex.Message;
            }
        }

        void RetornaPlano()
        {
            #region (Carregamento da aba Planos )                  
            try
            {
                if (DemServicos != null && DemServicos.LastRowNum >= 3)
                {
                    var ServicosLinhaCount = (DemServicos.LastRowNum - 1);
                    var ServicosColunaCount = (DemServicos.GetRow(0).PhysicalNumberOfCells);

                    for (int p = 1; p < ServicosLinhaCount; p++)
                    {                        
                        Planos.CreateRow(CountLinhasServico + p);
                        try
                        {
                            Planos.GetRow(CountLinhasServico + p).CreateCell(0);

                            var Inicial = Convert.ToString(DemServicos.GetRow(p + 2).GetCell(9));
                            var Final = Convert.ToString(DemServicos.GetRow(p + 2).GetCell(10));

                            if (string.IsNullOrEmpty(Inicial))
                            {
                                var teste = "";
                            }

                            if (!String.IsNullOrEmpty(Inicial) && !String.IsNullOrEmpty(Final))
                            {
                                TimeSpan date = Convert.ToDateTime(DemServicos.GetRow(p + 2).GetCell(10).ToString()) - Convert.ToDateTime(DemServicos.GetRow(p + 2).GetCell(9).ToString());
                                var totalDias = (date.Days + 1).ToString();
                                Planos.GetRow(CountLinhasServico + p).GetCell(0).SetCellValue(totalDias);
                            }
                            else
                                Planos.GetRow(CountLinhasServico + p).GetCell(0).SetCellValue("*");
                        }
                        catch (Exception ex)
                        {
                            string Ex = ex.Message;
                        }

                        for (int f = 0; f < ServicosColunaCount; f++)
                        {
                            try
                            {
                                Planos.GetRow(CountLinhasServico + p).CreateCell(f + 1);
                                if (f > 8 && f < 11 && DemServicos.GetRow(p + 2).GetCell(f) != null && !string.IsNullOrEmpty(DemServicos.GetRow(p + 2).GetCell(f).ToString()))
                                {
                                    DateTime Data = Convert.ToDateTime(DemServicos.GetRow(p + 2).GetCell(f).ToString());
                                    Planos.GetRow(CountLinhasServico + p).GetCell(f + 1).SetCellValue(Data.ToString("dd/MM/yyyy"));
                                }
                                else
                                    Planos.GetRow(CountLinhasServico + p).GetCell(f + 1).SetCellValue(Convert.ToString(DemServicos.GetRow(p + 2).GetCell(f)));
                            }
                            catch (Exception ex)
                            {
                                string Ex = ex.Message;
                            }
                        }   
                    }                    
                }
            }
            catch (Exception ex)
            {
                string Ex = ex.Message;
            }
            finally
            {
                DemServicos = null;
                GC.Collect();
            }

            CountLinhasServico = Templatebook.GetSheet("Planos").LastRowNum;

            #endregion
        }

        void RetornaTrafego()
        {
            #region Carregamento da aba Tráfego

            try
            {
                for (int a = 0; a < NumberOfSheets; a++)
                {
                    sPlanilhaA = a > 0 ? ("Demonstrativo da Fatura " + (a).ToString()).ToString() : ("Demonstrativo da Fatura").ToString();
                 
                    for (int b = 0; b < NumberOfSheets; b++)
                    {
                        string sPlanilhaB = PlanilhaTrafego.GetSheetAt(b).SheetName.ToString();
                        if (sPlanilhaB == sPlanilhaA)
                        {
                            try
                            {
                                var DemFatura = PlanilhaTrafego.GetSheet(sPlanilhaA);
                                if (DemFatura != null && DemFatura.LastRowNum >= 3)
                                {
                                    var FaturalinhaCount = (DemFatura.LastRowNum - 1);
                                    var FaturaColunaCount = (DemFatura.GetRow(0).PhysicalNumberOfCells);
                                    CountLinhasFatura = Templatebook.GetSheet("Tráfego").LastRowNum;

                                    for (int t = 1; t < FaturalinhaCount; t++)
                                    {
                                        Trafego.CreateRow(CountLinhasFatura + t);
                                        try
                                        {
                                            Trafego.GetRow(CountLinhasFatura + t).CreateCell(0);
                                            var DataChamada = Convert.ToString(DemFatura.GetRow(t + 2).GetCell(6));

                                            if (!String.IsNullOrEmpty(DataChamada))
                                                Trafego.GetRow(CountLinhasFatura + t).GetCell(0).SetCellValue(DataChamada);
                                            else
                                                Trafego.GetRow(CountLinhasFatura + t).GetCell(0).SetCellValue("*");
                                        }
                                        catch (Exception ex)
                                        {
                                            string Ex = ex.Message;
                                        }

                                        for (int j = 0; j < FaturaColunaCount; j++)
                                        {
                                            Trafego.GetRow(CountLinhasFatura + t).CreateCell(j + 1);
                                            Trafego.GetRow(CountLinhasFatura + t).GetCell(j + 1).SetCellValue(Convert.ToString(DemFatura.GetRow(t + 2).GetCell(j)));
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                string Ex = ex.Message;
                            }
                            finally
                            {
                                GC.Collect();
                            }
                            break;
                        }
                        continue;
                    }                   
                }
            }
            catch (Exception ex)
            {
                string Ex = ex.Message;
            }

            CountLinhasFatura = Templatebook.GetSheet("Tráfego").LastRowNum;

            #endregion
        }
    }
}