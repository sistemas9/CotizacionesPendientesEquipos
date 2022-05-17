using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace CotizacionesPendientesEquipos
{
  class databaseAyt03
  {
    IDataReader rd = null;
    int rowsAfected = 0;
    public SqlConnection conn;

    public databaseAyt03()
    {
      //var AppSettings = new ConfigurationBuilder()
      //        .SetBasePath(Directory.GetCurrentDirectory())
      //        .AddJsonFile("appsettings.json")
      //        .Build();
      var AppSettings = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(File.ReadAllText(@"appsettings.json"));

      String conectionString = AppSettings["DB_AYT03"];
      conn = new SqlConnection(conectionString);
      conn.Close();
    }

    public IDataReader query(String queryStr)
    {
      SqlCommand command = new SqlCommand(queryStr, conn);
      command.CommandTimeout = 90;
      if (conn.State == ConnectionState.Closed)
        conn.Open();
      try
      {
        rd = command.ExecuteReader();
      }
      catch (Exception ex)
      {
        Console.WriteLine(ex.ToString());
      }
      return rd;
    }

    public int queryInsert(String queryStr)
    {
      SqlCommand command = new SqlCommand(queryStr, conn);
      if (conn.State == ConnectionState.Closed)
        conn.Open();
      try
      {
        rowsAfected = command.ExecuteNonQuery();
      }
      catch (Exception ex)
      {
        Console.WriteLine(ex.ToString());
        return -2;
      }
      return rowsAfected;
    }
  }
}
