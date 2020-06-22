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
        HSSFWorkbook PlanilhaTrafego = null;        
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
            using (FileStream Template = new FileStream(Path.Combine(Application.StartupPath, @"Template.xlsx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Carrega o template
                Templatebook = new XSSFWorkbook(Template);
            }

            //Pega o valor aproximado da quantide memória usada 
            //var memoriaAtual = GC.GetTotalMemory(true);

            //Carrega o template
            //Templatebook = new XSSFWorkbook(Template);

            //Carrega as planilhas do template
            Planos = Templatebook.GetSheet("Planos");
            Trafego = Templatebook.GetSheet("Tráfego");

            //Template = null;

            //Muda a barra de progresso
            CarregaProgresso(5);

            for (int i = 0; i < excelFiles.Length; i++)
            {
                CarregaProgresso(i);
                
                try
                {
                    HSSFWorkbook PlanilhaCorrente = null;

                    using (FileStream arquivoXLS = new FileStream(excelFiles[i], FileMode.Open, FileAccess.ReadWrite))
                    {
                        // recupera o Workbook do arquivo
                        PlanilhaCorrente = new HSSFWorkbook(arquivoXLS);
                    }

                    //arquivoXLS.Close();

                    //Carrega as planilhas dos arquivos selecionados                    
                    DemServicos = PlanilhaCorrente.GetSheet("Demonstrativo dos Serviços");
                    NumberOfSheets = PlanilhaCorrente.NumberOfSheets;
                    PlanilhaTrafego = PlanilhaCorrente; 

                    var task = Task.Factory.StartNew(() =>
                    {
                        Application.DoEvents();
                        Cursor.Current = Cursors.WaitCursor;
                        RetornaTrafego();
                        Cursor.Current = Cursors.Default;

                        Application.DoEvents();
                        Cursor.Current = Cursors.WaitCursor;
                        RetornaPlano();
                        Cursor.Current = Cursors.Default;

                    });

                    task.Wait();                    
                }
                catch (Exception ex)
                {
                    string Ex = ex.Message;
                    string[] objeto = { "RetornaDados", Ex, "1" };
                    Thread Plano = new Thread(reportLogging); 
                    Plano.Start(objeto);                    
                }
            }

            excelFiles = null;
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
                using (FileStream file = new FileStream(nameArquivo, FileMode.Create))
                {
                    Templatebook.Write(file);
                    Templatebook = null;
                    GC.Collect();
                }

                //Templatebook.Write(file);

                //file.Close();
                //Templatebook = null;

                //GC.Collect();

                //Executa o arquivo
                System.Diagnostics.Process.Start(nameArquivo);

            }
            catch (Exception ex)
            {
                string Ex = ex.Message;
                string[] objeto = { "RetornaDados", Ex, "1" };
                Thread Plano = new Thread(reportLogging);
                Plano.Start(objeto);                
            }
            finally
            {
                
                Planos = null;
                Trafego = null;
                CountLinhasFatura = 0;
                CountLinhasServico = 0;                
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
                string[] objeto = { "CarregaProgresso", Ex, "1" };
                Thread Plano = new Thread(reportLogging);
                Plano.Start(objeto);                
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
                            string[] objeto = { "RetornaPlano", Ex, "2" };
                            Thread Plano = new Thread(reportLogging);
                            Plano.Start(objeto);                            
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
                                string[] objeto = { "RetornaPlano", Ex, "3" };
                                Thread Plano = new Thread(reportLogging);
                                Plano.Start(objeto);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string Ex = ex.Message;
                string[] objeto = { "RetornaPlano", Ex, "1" };
                Thread Plano = new Thread(reportLogging);
                Plano.Start(objeto);
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
                                            string[] objeto = { "RetornaTrafego", Ex, "3" };
                                            Thread Plano = new Thread(reportLogging);
                                            Plano.Start(objeto);
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
                                string[] objeto = { "RetornaTrafego", Ex, "2" };
                                Thread Plano = new Thread(reportLogging);
                                Plano.Start(objeto);                               
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
                string[] objeto = { "RetornaTrafego", Ex, "1" };
                Thread Plano = new Thread(reportLogging);
                Plano.Start(objeto);
            }

            CountLinhasFatura = Templatebook.GetSheet("Tráfego").LastRowNum;
            
            #endregion
        }

        public void reportLogging(object Method)
        {
            string[] res = Method as string[];
            
            String path = @"c:\\LoggsReport\\";

            bool exists = Directory.Exists(path);
            if (!exists)
                Directory.CreateDirectory(path);


            string nameArq = path + @"\ (" + res[0].ToString() + ")" + DateTime.Now.ToString("dd-MM-yyyy_HH_mm") + ".txt";

            StreamWriter writer = new StreamWriter(nameArq);

            string Camp = res[0].ToString() + " - Erro: " + res[1].ToString() + " Try/catch: " + res[2].ToString();

            writer.WriteLine(Camp);

            writer.Close();
        }
    }
}