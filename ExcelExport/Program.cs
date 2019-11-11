using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace ExcelExport
{
    class Program
    {
        static void Main(string[] args)
        {
            //MemoryStream ms = new MemoryStream();
            //TextWriter tw = new StreamWriter(ms);

            StreamWriter sw = new StreamWriter("D:\\Test.sql");
            
            int col = 7;
            int row = 3;
            int minDistance = 0;
            int maxDistance = 150;
            int minWeight = 0;
            int maxWeight = 1;
            string cost = "";
            string measuringUnit = "1";
            string countryId = "736";
            string isAverage = "false";
            string active = "true";
            string createDate = (DateTime.Now).ToString();
            string modifiedBy = "1";
            int c = 2;
            int[] distanceArr = new int[7] {300, 600, 1000, 1400, 1800, 182222222,150};

            string sql = "";
            string insertSql = "INSERT INTO GlobalShippingCost (MethodId,MinWeight,MaxWeight,MinDistance,MaxDistance,Cost,MeasuringUnit,CountryId,IsAverage,Active,CreateDate,LastModified,ModifiedBy) VALUES(1," + (minWeight).ToString() + "," + (maxWeight).ToString() + "," + (minDistance).ToString() + "," + (maxDistance).ToString() + ",";



            Excel excel = new Excel(@"D:\Test.xlsx", 1);

            for (int i = 1; i < row; i++)
            {
                for (int j = 0; j <= col; j++)
                {
                    if (j < col)
                    {
                        cost = (excel.ReadCell(i, c)).ToString();

                        sql += insertSql + cost + "," + measuringUnit + "," + countryId + "," + "'" + isAverage + "'" + "," + "'" + active + "'" + "," + "'" + createDate + "'" + "," + "'" + createDate + "'" + "," + modifiedBy + ")";
                        minDistance = maxDistance + 1;
                        maxDistance = distanceArr[c - 2];
                        c++;
                    }
                    else
                    {
                        sql +=
                            "INSERT INTO GlobalShippingCost(MethodId,MinWeight,MaxWeight,MinDistance,MaxDistance,Cost,MeasuringUnit,CountryId,IsAverage,Active,CreateDate,LastModified,ModifiedBy)  VALUES(1,"+(minWeight).ToString()+","+(maxWeight).ToString()+" , 0, 0, 7.71, 1, 736, 'true', 'true', '"+createDate+ "', '" + createDate + "', 1)";
                    }
                    
                }

                minWeight = maxWeight;
                maxWeight = maxWeight + 1;
                c = 2;
            }
            sw.WriteLine(sql);
            sw.Close();


        }
    }
}
