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
        static HSSFWorkbook hssfworkbook;
        static IList<TableModel> tables;
        static String ToDb = "gallery2";
        static String FromDb = "gallery2shop";
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                DumpHelp();
                Environment.Exit(1);
            }
            if (args[0] == "-c")
            {
                if (args.Length < 5)
                {
                    DumpHelp();
                    Environment.Exit(1); 
                }

                ToDb = args[3];
                FromDb = args[4];
                InitializeWorkbook(args[1]);
                Read();

                String schemaOutput = args[2] + (args[2].Last() == '\\' ? "CreateTables.sql" : "\\CreateTables.sql");
                StreamWriter schemaSW = new StreamWriter(schemaOutput, false, Encoding.GetEncoding("Shift-JIS"));
                schemaSW.Write(Schema());
                schemaSW.Close();

                String insertSqlOutput = args[2] + (args[2].Last() == '\\' ? "InsertData.sql" : "\\InsertData.sql");
                StreamWriter insertSqlSW = new StreamWriter(insertSqlOutput, false, Encoding.GetEncoding("Shift-JIS"));
                insertSqlSW.Write(InsertSQL());
                insertSqlSW.Close();
            }
            else
            {
                DumpHelp();
            }
        }

        static void InitializeWorkbook(string path)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(file);
            }
        }
        static public void Read()
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
                    table.Name = row.GetCell(TableCellNo.Name).StringCellValue;
                    var cell = row.GetCell(TableCellNo.From);
                    table.From = cell != null ? row.GetCell(TableCellNo.From).ToString() : "";
                }catch(Exception ex){
                    Console.WriteLine(ex);
                    return;
                }
                    
                // Shcema定義部までインクリメント
                rows.MoveNext();//2->3
                rows.MoveNext();//3->4
                rows.MoveNext();//4->5
                rows.MoveNext();//5->6

                while (rows.MoveNext())
                {
                    row = (HSSFRow)rows.Current;
                    var cell = row.GetCell(ColumnCellNo.Id);
                    if (cell == null || cell.CellType == CellType.BLANK)
                    {
                        continue;
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
                        var fromCell = row.GetCell(ColumnCellNo.FromId);
                        column.FromId = (fromCell !=null /*&& fromCell.CellType == CellType.STRING*/) ? fromCell.ToString() : "";
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
        static public String Schema()
        {
            var schema = String.Format("USE [{0}]\nGO\n\n\n", ToDb);
            foreach(var t in tables)
            {
                schema += t.Schema();
            }

            foreach (var t in tables)
            {
                schema += t.ForeginKeySchema();
            }

            return schema;
        }
        static public String InsertSQL()
        {
            var sql = String.Format("USE [{0}]\nGO\n\n\n", ToDb);
            foreach (var t in tables)
            {
                sql += t.InsertSQL(FromDb);
            }
            return sql;
        }
        static void DumpHelp()
        {
            Console.Write("Usage: Xlsx2DMLandINSERT.exe -c [input file] [output path] [destination db name] [src db name]");
        }
    }
}
