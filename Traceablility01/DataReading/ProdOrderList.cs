using Dapper;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Traceablility01.Data;

namespace Traceablility01.DataReading
{
    public class ProdOrderList
    {
        public List<ProdOrder>? ProdOrderItems;

        private string sqlCommand = "with p1 as(\r\n SELECT seven.prodregistration.feature,\r\n       seven.prodregistration.batch,\r\n       (SELECT TOP 1 prodorder.orderno\r\n        FROM   seven.prodregistration AS PR WITH (nolock)\r\n               JOIN mes.prodresource WITH (nolock)\r\n                 ON prodregistration.orderid = prodresource.id\r\n                    AND prodregistration.ordertypeid = prodresource.typeid\r\n               JOIN mes.prodorder\r\n                 ON prodresource.orderid = prodorder.id\r\n                    AND prodresource.ordertypeid = prodorder.typeid\r\n        WHERE  PR.id = prodregistration.id)                 [Zlecenie],\r\n\t\t       ((SELECT TOP 1 documentno\r\n         FROM   seven.document WITH (nolock)\r\n         WHERE  id = prodregistration.documentid))          [Dokument]\r\nFROM   seven.prodregistration\r\nWHERE  1 = 1\r\n       AND ( ( ( seven.prodregistration.created >= 'X1X' )\r\n               AND ( seven.prodregistration.created <= 'X2X'\r\n                   ) )\r\n             AND ( seven.prodregistration.iscanceled = 0 )\r\n             AND ( seven.prodregistration.parentid <> 0 ) )\r\n       AND (( prodregistration.destwarehouseid NOT IN ( 1, 7 )\r\n              AND prodregistration.srcwarehouseid NOT IN ( 1, 7 ) ))\r\nUNION\r\nSELECT seven.prodregistration.feature,\r\n       seven.prodregistration.batch,\r\n       (SELECT TOP 1 prodorder.orderno\r\n        FROM   seven.prodregistration AS PR WITH (nolock)\r\n               JOIN mes.prodresource WITH (nolock)\r\n                 ON prodregistration.orderid = prodresource.id\r\n                    AND prodregistration.ordertypeid = prodresource.typeid\r\n               JOIN mes.prodorder\r\n                 ON prodresource.orderid = prodorder.id\r\n                    AND prodresource.ordertypeid = prodorder.typeid\r\n        WHERE  PR.id = prodregistration.id)                 [Zlecenie],\r\n\t\t       ((SELECT TOP 1 documentno\r\n         FROM   seven.document WITH (nolock)\r\n         WHERE  id = prodregistration.documentid))          [Dokument]\r\nFROM   seven.prodregistration\r\nWHERE  1 = 1\r\n       AND ( ( ( seven.prodregistration.created >= 'X1X' )\r\n               AND ( seven.prodregistration.created <= 'X2X'\r\n                   ) )\r\n             AND ( seven.prodregistration.iscanceled = 0 ) \r\n)\r\n       AND (( prodregistration.destwarehouseid NOT IN ( 1, 7 )\r\n              AND prodregistration.srcwarehouseid NOT IN ( 1, 7 ) ))\r\n)\r\n\r\nselect distinct p1.zlecenie as [ProdLabel]\r\n\t\t\t  , replace(left(p1.[zlecenie],charindex('/',p1.[zlecenie])-1),'-',' ') as [ProdBatch]\r\nfrom p1\r\n where p1.[feature] in( \r\n\t\t\t\t\t\tselect p1.Batch \r\n\t\t\t\t\t\tfrom p1\r\n\t\t\t\t\t\twhere p1.[Dokument]='X3X'\r\n\t\t\t\t\t)\r\nand p1.[Zlecenie] is not null";
        private void fillList(string sql)
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnVal("MESXL_PRD")))
            {
                ProdOrderItems = connection.Query<ProdOrder>($"{sql}").ToList();
            };
        }

        public ProdOrderList(string date01, string date02, string documentPZ)
        {
            sqlCommand = sqlCommand.Replace("X1X", date01);
            sqlCommand = sqlCommand.Replace("X2X", date02);
            sqlCommand = sqlCommand.Replace("X3X", documentPZ);
            fillList(sqlCommand);
        }
    }
}
