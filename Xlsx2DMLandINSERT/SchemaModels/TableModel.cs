using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Xlsx2DMLandINSERT.SchemaModels
{
    static class TableCellNo
    {
        static public int Name = 2;
        static public int From = 9;
    }
    class TableModel
    {
        public String Name { get; set; }
        public String From { get; set; }
        public IList<ColumnModel> Columns {get; set;}
        public TableModel()
        {
            this.Columns = new List<ColumnModel>();
        }
        public String Schema()
        {
            String schema = "";

            schema += String.Format("CREATE TABLE [dbo].[{0}](\n", Name);
            schema += String.Join(",\n", Columns.Select(c => 
                c.Id == "Id" ? 
                    String.Format(
                        "\t[{0}] [{1}] IDENTITY(1,1) {2}",
                        c.Id,
                        c.Type,
                        c.NN ? "NOT NULL" : ""
                    ) : 
                    String.Format(
                        "\t[{0}] [{1}]{2} {3} ",
                        c.Id,
                        c.Type,
                        c.Number != "" ? "(" + c.Number + ")" : "",
                        c.NN ? "NOT NULL" : ""
                    )
            ).ToArray());
            schema += String.Format("\n\tCONSTRAINT [{0}_Id] PRIMARY KEY CLUSTERED(\n", Name);
            schema += "\t\t[Id] ASC\n";
            schema += "\t)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]";
            schema += ") ON [PRIMARY]\n";
            schema += "GO\n";

            return schema;
        }

        public String ForeginKeySchema()
        {
            return String.Join("",Columns.Select(c =>{
                if (Regex.IsMatch(c.FK, @"\S+"))
                {
                    return String.Format(
                       "ALTER TABLE [{0}] ADD CONSTRAINT FK_{0}_{1} FOREIGN KEY ({1}) REFERENCES {2}\nGO\n",
                       Name,
                       c.Id,
                       c.FK
                   );
                }
                return "";
            }).ToArray());
        }

        public String InsertSQL(String fromDb)
        {
            var sql = "";

            if (From != "")
            {
                sql += String.Format("SET IDENTITY_INSERT {0} ON\n", Name);
                sql += String.Format("INSERT INTO {0} (\n", Name);
                sql += String.Join(",\n",Columns.Select(c => c.Id).ToArray());
                sql += "\n)\n";

                sql += "SELECT\n";
                sql += String.Join(",\n", Columns.Select(c => String.Format("{0} AS {1}", c.FromId == "" ? "NULL" : c.FromId, c.Id)).ToArray());
                sql += String.Format("\nFROM [{0}]..{1}\n", fromDb, From);
                sql += String.Format("SET IDENTITY_INSERT {0} OFF\n", Name);
                sql += "GO\n\n\n";
            }
            return sql;
        }
    }
}
