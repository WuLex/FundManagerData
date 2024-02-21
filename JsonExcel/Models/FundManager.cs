using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonExcel
{
    /// <summary>
    /// 已废弃
    /// </summary>

    public class FundManager
    {
        /// <summary>
        /// 经理ID
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string ManagerName { get; set; }

        /// <summary>
        /// 公司ID
        /// </summary>
        public string CompanyId { get; set; }

        /// <summary>
        /// 所属公司
        /// </summary>
        public string Company { get; set; }

        /// <summary>
        /// 基金代码
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 基金名称
        /// </summary>
        public string FundName { get; set; }

        /// <summary>
        /// 累计从业时间
        /// </summary>
        public string Totaldays { get; set; }

        /// <summary>
        /// 现任基金最佳回报
        /// </summary>
        public string BestProfit { get; set; }

        public string BestFundCode { get; internal set; }
        public string BestFundName { get; set; }
        public string TotalMoney { get; set; }
    }
}