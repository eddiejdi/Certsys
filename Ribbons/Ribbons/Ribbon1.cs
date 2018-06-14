using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Ribbons;
//using Microsoft.Office.Tools.Excel;




namespace Ribbons
{
    public partial class Ribbon1
    {
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        
        private void BtnPontoCertsys_Click(object sender, RibbonControlEventArgs e)
        {
            // - Obtem planilha e Workbook ativa (verificar como escolher uma definida)
            Worksheet sheet = Globals.ThisAddIn.GetSheetActive();
            Workbook workbook = Globals.ThisAddIn.GetWorkbookActive();

            // sheetAtiva.Range["A1"].Value = "FUNCIONOU CACETE!!!!!!";

            // - Obtem as linhas preenchidas
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            // - Seleciona qual coluna para ser selecionada
            Excel.Range range = sheet.get_Range("A2", last);

            var VLimiteLinha = last.Row;
            string DiaHoje = DateTime.Now.Day.ToString();
            string DiaLinha, Projeto, HI1, HI2, HI3, HF1, HF2, HF3, Status;
            string User = workbook.Sheets["UserPass"].Range["B1"].Text;
            string Pass = workbook.Sheets["UserPass"].Range["B2"].Text;

            Range rng = workbook.Sheets["UserPass"].Range["A1:B2"];

            rng.CopyPicture(Excel.XlPictureAppearance.xlScreen,Excel.XlCopyPictureFormat.xlBitmap); 

            // - Loop para encontrar a linha correspondente ao dia
            for (int i = 2; i <= VLimiteLinha; i++)
            {
                if (sheet.Range["A" + i].Text != null)
                {
                    DiaLinha =  sheet.Range["A" + i].Text; // - Obtém o dia de Hoje
                    HI1 =       sheet.Range["B" + i].Text; // - Obtém hora inicial período 1
                    HI2 =       sheet.Range["D" + i].Text; // - Obtém hora inicial período 2
                    HI3 =       sheet.Range["F" + i].Text; // - Obtém hora inicial período 3
                    HF1 =       sheet.Range["C" + i].Text; // - Obtém hora final período 1
                    HF2 =       sheet.Range["E" + i].Text; // - Obtém hora final período 2
                    HF3 =       sheet.Range["G" + i].Text; // - Obtém hora final período 3
                    Projeto =   sheet.Range["I" + i].Text; // - Obtém Projeto Preenchido
                    Status  =   sheet.Range["K" + i].Text; // - Status do Registro

                    if (DiaLinha == DiaHoje || Status != "X")
                    {
                        WebRobots web = new WebRobots(User, Pass, Projeto, HI1, HI2, HI3, HF1, HF2, HF3);
                        // - Abre o Navegador para inserir dados no Ponto
                        var ret = web.WebPonto();

                        sheet.Range["K" + i].Value = "X";
                        return;

                    }
                }
            }

        }
    }
}
