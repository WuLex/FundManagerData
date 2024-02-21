using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonExcel.Models
{
    public class FundManagerList
    {
        /// <summary>
        /// 
        /// </summary>
        public List<List<string>> data { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public int record { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public int pages { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public int curpage { get; set; }
    }
}
