using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
//using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
// - Cuidado que esses dois abaixo funcionam mas conflitam entre si
//using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
// -------------------------------------


// Tudo o que for relacionado á planilha (thsiworkbook, ActiveSheet e etc) tem que estar nesta Classe...
namespace Ribbons
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado por VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
        public Worksheet GetSheetActive()
        {
            //Worksheet activeSheet = (Worksheet)Application.ActiveSheet;
            return (Worksheet)Application.ActiveSheet;
        }
        public Workbook GetWorkbookActive()
        {
            //Worksheet activeSheet = (Worksheet)Application.ActiveSheet;
            return (Workbook)Application.ActiveWorkbook;
        }
    }
}
