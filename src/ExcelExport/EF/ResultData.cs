using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelExport.EF
{
    public class ResultData<T>
    {
        public int Code { get; set; }

        public string Msg { get; set; }

        public T Data { get; set; }
    }
}