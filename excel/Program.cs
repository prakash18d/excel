using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable sheetExcelReport = datatable("C:\\Users\\prakashdubey1\\Downloads\\sheet.xlsx", "sheet");
            DataTable shsheetExcelCompareReportt = datatable("C:\\Users\\prakashdubey1\\Downloads\\sheet1.xlsx", "sheet1");


            List<string> tempdata = sheetExcelReport.AsEnumerable().Select(s => s.Field<string>("Content")).ToList();
            List<string> tempdata1 = shsheetExcelCompareReportt.AsEnumerable().Select(s => s.Field<string>("Content")).ToList();

            var diffs = tempdata.Union(tempdata1).Except(tempdata.Intersect(tempdata1));
            #region Pushing Data to excel
            //foreach (DataRow row in sheet1.Rows)
            //{

            //}
            //DataTable sheet2 = datatable("C:\\Users\\deepakverma01\\Desktop\\Befor Change Questionaire Report CR.xlsx", "Sheet2");
            //DataTable table = new DataTable();
            //table.Columns.Add("FirstName", typeof(string));
            //table.Columns.Add("Sheet1Number", typeof(int));
            //table.Columns.Add("Sheet2Number", typeof(int));
            //string test1 = null;
            //string test2 = null;
            //foreach (DataColumn column in sheet1.Columns)
            //{
            //    string a = column.ColumnName;
            //    foreach (DataColumn column1 in sheet2.Columns)
            //    {
            //        string b = column.ColumnName;
            //        if (a == b)
            //        {
            //            List<object> lst1 = (from d in sheet1.AsEnumerable() select d.Field<object>(a)).ToList();
            //            List<object> lst2 = (from d in sheet2.AsEnumerable() select d.Field<object>(a)).ToList();
            //            for (int i = 0; i < lst1.Count; i++)
            //            {
            //                for (int j = i; j < lst2.Count; j++)
            //                {
            //                    if ((i == j) && (!lst1.SequenceEqual(lst2)))
            //                    {
            //                        if (lst1[i].ToString() != lst2[j].ToString())
            //                        {
            //                            test1 = lst1[i].ToString();
            //                            test2 = lst2[i].ToString();
            //                            table.Rows.Add(a, test1, test2);
            //                        }


            //                    }
            //                }
            //            }


            //        }

            //    }
            //}

            //DataView dv = new DataView(table);
            //table = dv.ToTable(true);

            //var lines = new List<string>();

            //string[] columnNames = table.Columns.Cast<DataColumn>().
            //                                  Select(column => column.ColumnName).
            //                                  ToArray();

            //var header = string.Join(",", columnNames);
            //lines.Add(header);

            //var valueLines = table.AsEnumerable()
            //                   .Select(row => string.Join(",", row.ItemArray));
            //lines.AddRange(valueLines);

            //File.WriteAllLines("C:\\Users\\prakashdubey1\\Downloads\\final.csv", lines);
            //Console.WriteLine("success");
            //Console.ReadKey();
            #endregion

        }
        public static DataTable datatable(string path, string sheetName)
        {
            DataTable dt = new DataTable();
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;

                        try
                        {
                            da.Fill(dt);
                        }
                        catch
                        {
                            Console.WriteLine("hello Worls");
                        }
                        
                        return dt;
                    }
                }
            }
        }

    }
}