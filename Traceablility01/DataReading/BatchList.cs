using Dapper;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Data;
using Traceablility01.Data;
using System.Linq;

namespace Traceablility01.DataReading
{
    public class BatchList
    {

        public List<BatchData> BatchData = new List<BatchData>();

        private string sqlCommand = "select distinct p1.Batch\r\nfrom (\r\nSELECT top 5000 ProdRegistration.Batch as [Batch]\r\nFROM mes.ProdOrder\r\nJOIN mes.ProdResource WITH (nolock) ON ProdResource.OrderId=ProdOrder.id\r\n                    AND ProdResource.OrderTypeId=ProdOrder.TypeId\r\nJOIN seven.ProdRegistration WITH (nolock) ON prodregistration.orderid=ProdResource.id\r\n                    AND ProdRegistration.OrderTypeId=ProdResource.TypeId\r\nwhere ProdRegistration.Batch is not null and ProdRegistration.Batch not in('')\r\norder by ProdRegistration.Created desc, ProdRegistration.Batch desc\r\n) p1";

        private void addBatchData()
        {
            using (IDbConnection connection = new SqlConnection(Helper.CnnVal("MESXL_PRD")))
            {
                BatchData = connection.Query<BatchData>($"{sqlCommand}").ToList();
            };
        }

        public BatchList()
        {
            addBatchData();
        }

    }
}
