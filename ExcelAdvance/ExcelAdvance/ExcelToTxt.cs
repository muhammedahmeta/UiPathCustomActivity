using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Activities;
using System.ComponentModel;
using Spire.Xls;

namespace ExcelAdvance
{
    public class ExcelToTxt : CodeActivity
    {
        [Category("Options")]
        public bool AddHeaders { get; set; }

        [Category("Input")]
        [DisplayName("ExcelPath")]
        [Description("Full path of Excel")]
        [RequiredArgument]
        public InArgument<string> ExcelPath { get; set; }

        [Category("Options")]
        [DisplayName("SplitChar")]
        [Description("Excel columns spilit as this string character")]
        [RequiredArgument]
        public InArgument<string> SplitChar { get; set; }

        [Category("Options")]
        [DisplayName("SheetIndex")]
        [Description("The index of sheet. For first page 0")]
        [RequiredArgument]
        public InArgument<int> SheetIndex { get; set; }

        [Category("Output")]
        [DisplayName("TxtPath")]
        [Description("The full path of the txt file to write to.")]
        [RequiredArgument]
        public InArgument<string> TxtPath { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            var excelpath = ExcelPath.Get(context);
            var pageindex = SheetIndex.Get(context);
            var txtpath = TxtPath.Get(context);
            var splitchar = SplitChar.Get(context);
            var addheaders = AddHeaders.GetType();

            Workbook wb = new Workbook();
            wb.LoadFromFile(@excelpath);
            Worksheet ws = wb.Worksheets[pageindex];
            DataTable dt = ws.ExportDataTable();

            //TXT VARSA İÇİNİ TEMİZLE
            if (File.Exists(txtpath))
            {
                File.WriteAllText(txtpath, "");
            }
            //BAŞLIK
            if (AddHeaders)
            {
                string[] columnNames = dt.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
                File.AppendAllText(@txtpath, string.Join(splitchar, columnNames) + Environment.NewLine);
            }
            //SATIRLAR
            foreach (DataRow dtR in dt.Rows)
            {
                File.AppendAllText(@txtpath, string.Join(splitchar, dtR.ItemArray) + Environment.NewLine);
            }
        }
    }
}
