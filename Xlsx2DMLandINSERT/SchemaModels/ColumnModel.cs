using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xlsx2DMLandINSERT.SchemaModels
{
    static class ColumnCellNo
    {
        static public int Id = 2;
        static public int Type = 3;
        static public int Number = 4;
        static public int FK = 5;
        static public int NN = 6;
        static public int PK = 7;
        static public int FromId = 9;
    }
    class ColumnModel
    {
        public String Id { get; set; }
        public String Type { get; set; }
        public String Number { get; set; }
        public String FK { get; set; }
        public Boolean NN { get; set; }
        public Boolean PK { get; set; }
        public String FromId { get; set; }
    }
}
