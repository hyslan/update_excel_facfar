using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Update_Excel_FACFar
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Iniciando processo.");
            RunExcel();
            Console.WriteLine("Processo finalizado.");
        }
        static void RunExcel()
        {
            // Criar uma instância do Excel
            Excel.Application excelApp = new Excel.Application();

            // Tornar o aplicativo Excel invisível
            excelApp.Visible = false;

            string mlg = "\\\\spo-leste60_fs\\FISCALIZAÇÃO\\FAC FAR\\BANCO DE DADOS ML\\MLG\\Relatório Combinado MLG.XLSX";
            string mln = "\\\\spo-leste60_fs\\FISCALIZAÇÃO\\FAC FAR\\BANCO DE DADOS ML\\MLN - Alto Tietê\\Relatório Combinado MLN.XLSX";
            string mlq = "\\\\spo-leste60_fs\\FISCALIZAÇÃO\\FAC FAR\\BANCO DE DADOS ML\\MLQ - Itaquera\\Relatório Combinado MLQ.XLSX";
            Excel.Workbook wb_mlg = excelApp.Workbooks.Open(mlg);
            Excel.Workbook wb_mln = excelApp.Workbooks.Open(mln);
            Excel.Workbook wb_mlq = excelApp.Workbooks.Open(mlq);

            wb_mlg.RefreshAll();
            wb_mln.RefreshAll();
            wb_mlq.RefreshAll();
            Console.WriteLine("Atualizando Query do Excel...");
            System.Threading.Thread.Sleep(45000);

            // Fecha e salva as alterações feitas no arquivo
            wb_mlg.Close(true);
            wb_mln.Close(true);
            wb_mlq.Close(true);
            excelApp.Quit();

            // Libera os objetos
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb_mlg);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb_mln);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb_mlq);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            // Executa a coleta de lixo para liberar a memória ocupada pelos objetos COM
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
