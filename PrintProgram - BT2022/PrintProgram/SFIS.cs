using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintProgram
{
    /// <summary>
    /// SFIS 類別用於記錄工單相關資訊。
    /// </summary>
    public class SFIS
    {
        public string itemNo { get; set; }
        public string productid { get; set; }
        public string powercord { get; set; }
        public string biosVer { get; set; }
        public string BIOSCS { get; set; }
        public string ProductID_MF { get; set; }
        public string Ecn { get; set; }
    }
}
