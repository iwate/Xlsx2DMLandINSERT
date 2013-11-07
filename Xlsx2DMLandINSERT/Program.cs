using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xlsx2DMLandINSERT.SchemaModels;

namespace Xlsx2DMLandINSERT
{
    class Program
    {
        class App
        {
            HSSFWorkbook hssfworkbook;
            IList<TableModel> tables;
            public App()
            {
                InitializeWorkbook(@"C:\Users\b-yostan\Documents\SpecificTable.xls");
            }

            void InitializeWorkbook(string path)
            {
                //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
                //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    hssfworkbook = new HSSFWorkbook(file);
                }
            }
            public void Read()
            {
                tables = new List<TableModel>();
                for (int i = 0, len = hssfworkbook.NumberOfSheets; i < len; i++)
                {
                    var sheet = hssfworkbook.GetSheetAt(i);
                    var rows = sheet.GetRowEnumerator();
                    var table = new TableModel();
                    HSSFRow row;
                    rows.MoveNext();//dummy
                    rows.MoveNext();//1->2
                    
                    // TableModel.Name, TableModel.From を取得
                    row = (HSSFRow)rows.Current;
                    try{
                        var cell = row.GetCell(TableCellNo.Name);
                        table.Name = row.GetCell(TableCellNo.Name).StringCellValue;
                        table.From = row.GetCell(TableCellNo.From).StringCellValue;
                    }catch(Exception ex){
                        Console.WriteLine(ex);
                        return;
                    }
                    
                    // Shcema定義部までインクリメント
                    rows.MoveNext();//2->3
                    rows.MoveNext();//3->4
                    rows.MoveNext();//4->5
                    rows.MoveNext();//5->6
                    rows.MoveNext();//6->7

                    while (rows.MoveNext())
                    {
                        row = (HSSFRow)rows.Current;
                        if (row.GetCell(ColumnCellNo.Id) == null)
                        {
                            break;
                        }
                        var column = new ColumnModel();
                        try
                        {
                            column.Id = row.GetCell(ColumnCellNo.Id).StringCellValue;
                            column.Type = row.GetCell(ColumnCellNo.Type).StringCellValue;
                            var numCell = row.GetCell(ColumnCellNo.Number);
                            column.Number = numCell.CellType == CellType.NUMERIC ? numCell.NumericCellValue.ToString() : numCell.StringCellValue;
                            column.FK = row.GetCell(ColumnCellNo.FK).StringCellValue;
                            column.NN = row.GetCell(ColumnCellNo.NN).StringCellValue.Length > 0;
                            column.PK = row.GetCell(ColumnCellNo.PK).StringCellValue.Length > 0;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                            return;
                        }
                        table.Columns.Add(column);
                    }
                    tables.Add(table);
                }
            }
        }
        static void Main(string[] args)
        {
            (new App()).Read();
            Console.ReadLine();
        }
    }
}
