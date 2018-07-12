using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Eon.DataService.Client.Core.API;
using Eon.DataService.Client.WCF.Factories;
using System.Reflection;
using System.Configuration;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Globalization;

namespace DSApiTest
{
    public class Program
    {
        static void Main(string[] args)
        {
            Methods m = new Methods();
            m.InitializeDataServiceClient();

            CultureInfo culture = new CultureInfo("fr-FR");
            string TODAY = DateTime.Today.ToString("MM-dd-yyyy");
            DateTime today = DateTime.Now;
            DateTime answer = today.AddDays(-7);
            string TODAY_L7 = answer.ToString("MM-dd-yyyy");

            TODAY = "11/04/2017";
            TODAY_L7 = "10/29/2017";
            string query = "";

            #region COAL
            string type = "COAL";     
            
            query = @"return lim.GetRecords({'PA0002140.0.0'},{'Index'},date.Parse('" + TODAY_L7 + "'),date.Parse ('" + TODAY + "'),'Days');";
            System.Console.WriteLine(query);
            System.Data.DataTable dt = m.DSApiQueriesGetData(query);

            var csv = new StringBuilder();
            csv.AppendLine(@"IMPORT TYPE,TYPE,MARKET,COMP,DATE,MATURITY,PFORMAT,PSET,SETTLE,PRICE");
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DateTime day = DateTime.ParseExact(dt.Rows[i][0].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                if (dt.Rows[i][1].ToString() != "")
                {
                    csv.AppendLine("PRICE,DAILY,COAL,API2," + day.ToString("dd/MM/yyyy") + ",0CAL,*,SETTLE,F," + dt.Rows[i][1].ToString());
                }                    
                
            }
            
            query = @"return lim.GetRecords({'PA0002141.0.0'},{'Index'},date.Parse('" + TODAY_L7 + "'),date.Parse ('" + TODAY + "'),'Days');";
            System.Console.WriteLine(query);
            dt = m.DSApiQueriesGetData(query);
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DateTime day = DateTime.ParseExact(dt.Rows[i][0].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);                
                if (dt.Rows[i][1].ToString() != "")
                {
                    csv.AppendLine(@"PRICE,DAILY,COAL,API4," + day.ToString("dd/MM/yyyy") + @",0CAL,*,SETTLE,F," + dt.Rows[i][1].ToString());
                }
                
            }
            
            File.WriteAllText(@"LIMPRC_"+DateTime.Now.ToString("yyyyMMddHHmmss")+"_"+type+".csv", csv.ToString());
            #endregion

            #region FX
            type = "FX";
            System.Console.WriteLine(query);
            query = "return lim.GetRecords({'ECBUSD'},{'Spot'},date.Parse('" + TODAY_L7 + "'),date.Parse ('" + TODAY + "'),'Days');";
            dt = m.DSApiQueriesGetData(query);

            csv = new StringBuilder();
            csv.AppendLine(@"IMPORT TYPE,TYPE,MARKET,COMP,DATE,MATURITY,PFORMAT,PSET,SETTLE,PRICE");
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DateTime day = DateTime.ParseExact(dt.Rows[i][0].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                
                if (dt.Rows[i][1].ToString() != "")
                {
                    string s2 = dt.Rows[i][1].ToString();
                    csv.AppendLine("PRICE,DAILY,FXRINV,USD," + day.ToString("dd/MM/yyyy") + ",0CAL,*,SETTLE,F," + s2);
                    var v = 1 / Convert.ToDouble(dt.Rows[i][1].ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-us"));
                    csv.AppendLine("PRICE,DAILY,FX,USD," + day.ToString("dd/MM/yyyy") + ",0CAL,*,SETTLE,F," + v.ToString(System.Globalization.CultureInfo.GetCultureInfo("en-us")));                
                }
            }


            File.WriteAllText(@"LIMPRC_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + type + ".csv", csv.ToString());
            #endregion

            #region CO2            
            type = "CO2";
            System.Console.WriteLine(query);
            query = "return lim.GetRecords({'ECX.CFI'},{'Close'},date.Parse('" + TODAY_L7 + "'),date.Parse ('" + TODAY + "'),'Days');";
            dt = m.DSApiQueriesGetData(query);

            csv = new StringBuilder();
            csv.AppendLine(@"IMPORT TYPE,TYPE,MARKET,COMP,DATE,MATURITY,PFORMAT,PSET,SETTLE,PRICE");
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DateTime day = DateTime.ParseExact(dt.Rows[i][0].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                
                if (dt.Rows[i][1].ToString() != "")
                {
                    string s2 = dt.Rows[i][1].ToString();
                    csv.AppendLine("PRICE,DAILY,BLUNXT,EUA," + day.ToString("dd/MM/yyyy") + ",0CAL,*,SETTLE,F," + s2);
                }
            }
            query = "return lim.GetRecords({'ECX.CER'},{'Close'},date.Parse('" + TODAY_L7 + "'),date.Parse ('" + TODAY + "'),'Days');";
            System.Console.WriteLine(query);             
            dt = m.DSApiQueriesGetData(query);
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DateTime day = DateTime.ParseExact(dt.Rows[i][0].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                if (dt.Rows[i][1].ToString() != "")
                {
                    string s2 = dt.Rows[i][1].ToString();
                    csv.AppendLine("PRICE,DAILY,BLUNXT,CER," + day.ToString("dd/MM/yyyy") + ",0CAL,*,SETTLE,F," + s2);
                }
            }

            File.WriteAllText(@"LIMPRC_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + type + ".csv", csv.ToString());
            #endregion

            #region EPEX
            type = "EPEX";
            csv = new StringBuilder();
            string dst = "";
            string dstdate = "";
            dt = m.DSApiQueriesGetData("return lim.GetRecords({'PNXT.DAHOURLY'},{'IndexDST'},date.Parse('" + TODAY_L7 + "'),date.Parse ('" + TODAY + "'),'Hours');");
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                string s2 = dt.Rows[i][1].ToString() == "" ? "0" : dt.Rows[i][1].ToString();
                dstdate = dt.Rows[i][0].ToString().Split(' ')[0];
                dst = ",,,,,,,,,B,2400,2459," + s2;
            }


            csv.AppendLine(@"IMPORT TYPE,TYPE,MARKET,COMP,DATE,MATURITY,PFORMAT,PSET,SETTLE,D,H,E,PRICE");
            string prev = "";
            dt = m.DSApiQueriesGetData("return lim.GetRecords({'PNXT.DAHOURLY'},{'Index'},date.Parse('" + TODAY_L7 + "'),date.Parse ('" + TODAY + "'),'Hours');");
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                string s2 = dt.Rows[i][1].ToString() == "" ? "0" : dt.Rows[i][1].ToString();
                string mdate = dt.Rows[i][0].ToString().Split(' ')[0];
                string mh = (dt.Rows[i][0].ToString().Split(' ')[1]).Split(':')[0];
                DateTime day = DateTime.ParseExact(mdate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                if (mdate != prev)
                    csv.AppendLine("PRICE,HOURLY,PWNEXT,HOURLY," + day.ToString("dd/MM/yyyy") + ",0CAL,@,SETTLE,N,B," + mh + "00," + mh + "59," + s2);
                else
                    csv.AppendLine(",,,,,,,,,B," + mh + "00," + mh + "59," + s2);

                if (dstdate == mdate)
                {
                    csv.AppendLine(dst);
                    dstdate = "";
                }
                

                prev = mdate;
            }
            File.WriteAllText(@"LIMPRC_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + type + ".csv", csv.ToString());
            #endregion

            #region EUA/CER
            type = "EUA";

            csv = new StringBuilder();
            csv.AppendLine(@"IMPORT TYPE,TYPE,MARKET,COMP,DATE,MATURITY,PFORMAT,PSET,SETTLE,PRICE");

            string[] fc = { "ECX.CFI_2018H", "ECX.CFI_2018M", "ECX.CFI_2018U", "ECX.CFI_2018Z", "ECX.CFI_2019H", "ECX.CFI_2019M", "ECX.CFI_2019U", "ECX.CFI_2019Z", "ECX.CFI_2020H", "ECX.CFI_2020M", "ECX.CFI_2020U", "ECX.CFI_2020Z", "ECX.CFI_2021H", "ECX.CFI_2021Z", "ECX.CFI_2022Z", "ECX.CFI_2023Z", "ECX.CFI_2024Z", "ECX.CER_2018H", "ECX.CER_2018M", "ECX.CER_2018U", "ECX.CER_2018Z", "ECX.CER_2019H", "ECX.CER_2019M", "ECX.CER_2019U", "ECX.CER_2019Z", "ECX.CER_2020H", "ECX.CER_2020Z" };

            for (int j =0 ; j < fc.Length; j++)
            {
                query = "return lim.GetRecords({'" + fc[j] + "'},{'Close'},date.Parse('" + TODAY + "'),date.Parse('" + TODAY + "'),'Days');";
                System.Console.WriteLine(query);   
                dt = m.DSApiQueriesGetData(query);
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    DateTime day = DateTime.ParseExact(dt.Rows[i][0].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                    
                    if (dt.Rows[i][1].ToString() != "")
                    {
                        string s2 = dt.Rows[i][1].ToString();
                        csv.AppendLine("PRICE,HOURLY,ICE,EUA," + day.ToString("dd/MM/yyyy") + ",MATURITY,*,MARKET,F," + s2);
                    }                    
                }
            }
            File.WriteAllText(@"LIMPRC_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + type + ".csv", csv.ToString());

            #endregion 

            #region COAL FWD
            type = "COALFWD";
            csv = new StringBuilder();
            csv.AppendLine(@"IMPORT TYPE,TYPE,MARKET,COMP,DATE,MATURITY,PFORMAT,PSET,SETTLE,PRICE");

            fc = new string[] { "TFS.COAL.TFS_API_2_2018K", "TFS.COAL.TFS_API_2_2018K", "TFS.COAL.TFS_API_2_2018M", "TFS.COAL.TFS_API_2_2018M", "TFS.COAL.TFS_API_2_2018N", "TFS.COAL.TFS_API_2_2018N", "TFS.COAL.TFS_API_2_2019Q1", "TFS.COAL.TFS_API_2_2019Q1", "TFS.COAL.TFS_API_2_2019Q2", "TFS.COAL.TFS_API_2_2019Q2", "TFS.COAL.TFS_API_2_2018Q3", "TFS.COAL.TFS_API_2_2018Q3", "TFS.COAL.TFS_API_2_2018Q4", "TFS.COAL.TFS_API_2_2018Q4", "TFS.COAL.TFS_API_2_2019CAL", "TFS.COAL.TFS_API_2_2019CAL", "TFS.COAL.TFS_API_2_2020CAL", "TFS.COAL.TFS_API_2_2020CAL", "TFS.COAL.TFS_API_2_2021CAL", "TFS.COAL.TFS_API_2_2021CAL" };//, "TFS.COAL.TFS_API_2_2022CAL", "TFS.COAL.TFS_API_2_2022CAL" };
            string []fctype = {"High","Low","High","Low","High","Low","High","Low","High","Low","High","Low","High","Low","High","Low","High","Low","High","Low","High","Low"};
            for (int j = 0; j < fc.Length; j++)
            {
                query = "return lim.GetRecords({'" + fc[j] + "'},{'" + fctype[j] + "'},date.Parse('" + TODAY + "'),date.Parse('" + TODAY + "'),'Days');";
                System.Console.WriteLine(query);
                dt = m.DSApiQueriesGetData(query);
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    DateTime day = DateTime.ParseExact(dt.Rows[i][0].ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                    if (dt.Rows[i][1].ToString() != "")
                    {
                        string s2 = dt.Rows[i][1].ToString();
                        string t = "BID";
                        if (fctype[j] == "High")
                            t = "ASK";
                        csv.AppendLine("PRICE,HOURLY,COAL,API2," + day.ToString("dd/MM/yyyy") + ",MATURITY,*,"+t+",F," + s2);
                    }
                }
            }
            File.WriteAllText(@"LIMPRC_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + type + ".csv", csv.ToString());

            #endregion 


        }

        static void usage()
        {
            Console.WriteLine("Param : <request> <filename>");
        }

        public class Methods
        {
            private IDataServiceClientApi dataServiceClient { get; set; }

            public void InitializeDataServiceClient()
            {
                string dssource = ConfigurationManager.AppSettings["HostURL"].ToString();
                var factory = new WCFDataServiceClientAPIFactory();
                //dataServiceClient = factory.CreateDataServiceClient(System.Configuration.ConfigurationManager.AppSettings["DefaultEnvironment"], Environment.UserDomainName + "\\" + Environment.UserName);
                dataServiceClient = factory.CreateDataServiceClient(System.Configuration.ConfigurationManager.AppSettings["DefaultEnvironment"], "coucoucz" );
            }

            public System.Data.DataTable DSApiQueriesGetData(string sampleQuery)
            {

                try
                {
                    // Passing the query to DataServices ExecuteQuery() method which will return a DataReader object
                    using (var data = dataServiceClient.ExecuteQuery(sampleQuery))
                    {
                        return ReaderToTableConverter(data);                        
                    }

                }

                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
                return null;
            }

            public void DSApiQueries(string sampleQuery, string filename)
            {              

                try
                {
                    // Passing the query to DataServices ExecuteQuery() method which will return a DataReader object
                    using (var data = dataServiceClient.ExecuteQuery(sampleQuery))
                    {
                        //Converting DataReader to a dataTable
                        using (var limdata = ReaderToTableConverter(data))
                        {
                            //Writing the data to excel
                            WriteDataToExcel(limdata, filename);
                        }

                    }

                }

                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
                
            }

            public void DSApiQueries()
            {
                Console.WriteLine("1) Querying Multiple LIM Symbols for a date range using GetRecords");
                Console.WriteLine("2) Querying Single LIM Symbols for a date range using GetRecords");
                Console.WriteLine("3) Querying Single LIM Symbols for a date range using GetSingleRecords");
                Console.WriteLine("4) Executing queries using ExecuteMIMQuery");
                Console.WriteLine("5) FELIX");
                Console.WriteLine("Enter Option(1, 2, 3 or 4)...");
                var option = Console.ReadLine();

                List<Entity> limCurveLst = null;
                var sampleQuery = string.Empty;
                switch (option)
                {
                    case "1": // Querying Multiple LIM Symbols for a date range
                        {
                            sampleQuery = "return lim.GetRecords({'APX.DAHOURLY','HUPX.POWER.SPOT'},{'Index','Close'},date.Parse('10/21/17'),date.Parse('10/27/17'),'Hours');";                     
                            break;
                        }
                    case "2": // Querying Single LIM Symbols for a date range using GetRecords
                        {
                            sampleQuery = "return lim.GetRecords({'APX.DAHOURLY'},{'Index'},date.Parse('10/25/17'),date.Parse('10/27/17'),'Hours');";
                            break;
                        }
                    case "3": // Querying Single LIM Symbols for a date range using GetSingleRecords
                        {
                            sampleQuery = "return lim.GetSingleRecord('APX.DAILYAVG','Index',date.Parse('10/26/17'),date.Parse('10/27/17'),'Days');";
                            break;
                        }
                    case "4": //Executing queries using ExecuteMIMQuery
                        {
                            sampleQuery = "return lim.ExecuteMIMQuery('SHOW val:The 15 minute nearest_integer(PriceMwh of EPEX.INTRA.DAY.AUCTION.15MIN.DE, 0.01) repeated for the entire day when date is from 10/21/2017 to 10/27/2017');";
                            break;
                        }
                    case "5": //Executing queries using ExecuteMIMQuery
                       {
                            sampleQuery = "return lim.GetRecords({'PNXT.DAHOURLY'},{'Index'},date.Parse('02/20/18'),date.Parse('02/26/18'),'Hours');";
                            break;                        
                       }

                    case "6":
                        {
                            sampleQuery = "return lim.GetRecords({'PA0002141.0.0'},{'Index'},date.Parse('02/20/18'),date.Parse ('02/20/18'),'Days');";
                            break;

                        
                        }                      
                }

                try
                {                    
                    // Passing the query to DataServices ExecuteQuery() method which will return a DataReader object
                    using (var data = dataServiceClient.ExecuteQuery(sampleQuery))
                    {
                        //Converting DataReader to a dataTable
                        using (var limdata = ReaderToTableConverter(data))
                        {
                            //Writing the data to excel
                            WriteDataToExcel(limdata);
                        }

                    }

                }

                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }

                Console.ReadLine();

            }

            /// <summary>
            /// Write Data to Excel File
            /// </summary>
            public void WriteDataToExcel(System.Data.DataTable table, string filname="")
            {
                Application excelApp = new Application();
                excelApp.Workbooks.Add();
                Worksheet workSheet = excelApp.ActiveSheet;

                string newPath = ConfigurationManager.AppSettings[Constants.OutputPath];
                if (filname == "")
                    newPath = newPath.Replace("TimeStamp", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                else
                    newPath = filname;

                try
                {
                    if (table == null || table.Columns.Count == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");

                    
                    
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        workSheet.Cells[1, (i + 1)] = table.Columns[i].ColumnName;
                    }

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        for (int j = 0; j < table.Columns.Count; j++)
                        {
                            workSheet.Cells[(i + 2), (j + 1)] = table.Rows[i][j];
                        }
                    }

                    if (newPath != null && newPath != "")
                    {
                        try
                        {
                            workSheet.SaveAs(newPath);
                            excelApp.Quit();
                            Console.WriteLine("File Saved to " + newPath);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                + ex.Message);
                        }
                    }
                    else
                    {
                        excelApp.Quit();
                        
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("ExportToExcel: \n" + ex.Message);
                }
                finally
                {
                    excelApp.Quit();
                   
                }
            }                                      

            /// <summary>
            /// Method to Convert Reader to DataTable
            /// </summary>
            /// <param name="drCurveData">DataReader</param>
            /// <returns>DataTable</returns>
            public System.Data.DataTable ReaderToTableConverter(IDataReader drCurveData)
            {
                System.Data.DataTable table = new System.Data.DataTable();

                //Create the coloumn entries
                for (var i = 0; i < drCurveData.FieldCount; ++i)
                {
                    var column = table.Columns.Add(drCurveData.GetName(i));
                    column.DataType = drCurveData.GetFieldType(i);
                }
                try
                {
                    do
                    {
                        while (drCurveData.Read())
                        {
                            var row = table.NewRow();
                            for (var i = 0; i < drCurveData.FieldCount; i++)
                            {
                                row[i] = drCurveData.GetValue(i);
                            }
                            table.Rows.Add(row);
                        }
                    } while (drCurveData.NextResult()); //move to the next batch
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    //drCurveData = table.CreateDataReader();
                }

                return table;
            }

            /// <summary>
            /// Method to convert Datatable to List
            /// </summary>
            /// <param name="dtTable">DataTable object</param>
            /// <returns>LIST</returns>
            public List<Entity> DataTabletoListConverter(System.Data.DataTable dtTable)
            {
                List<Entity> limSymbolList = new List<Entity>();
                int columnCount = 1;

                while (columnCount < dtTable.Columns.Count)
                {
                    foreach (DataRow row in dtTable.Rows)
                    {
                        var values = row.ItemArray;

                        if (values[0] != null && !string.IsNullOrEmpty(values[0].ToString()))
                        {
                            var limSymbolEntity = new Entity()
                            {
                                Date = values[0].ToString(),
                                Value = values[columnCount].ToString(),

                            };

                            limSymbolList.Add(limSymbolEntity);
                        }
                    }

                    columnCount++;
                }
                return limSymbolList;
            }

        }
    }
}
