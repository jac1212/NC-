using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;

namespace NCvoucher
{
    class ArrayList
    {
        public static JArray getArrayList(string path)
        {
            JArray arrList = new JArray();
            //读取Excel
            DataTable dt = NPOIHelper.ReadExcelToDataTable(0, path, null, true);

            if (dt != null & dt.Rows.Count > 0)
            {
                foreach (System.Data.DataRow dr in dt.Rows)
                {
                    JObject obj = new JObject();
                    foreach (System.Data.DataColumn dc in dt.Columns)
                    {
                        try
                        {
                            obj.Add(dc.ColumnName, dr[dc.ColumnName].ToString());
                        }
                        catch { }
                    }
                    if (obj["序号"].ToString() != "")
                        arrList.Add(obj);
                }
            }
            return arrList;
        }
    }
}
