using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace Write_CSV_to_Excel
{
    class Program
    {
        #region Fields
        /// <summary>
        /// Path of CSV File
        /// </summary>
        public static string pathCSV = @"C:\Users\faranam\Desktop\Exam\09 - Excel export\data.csv";

        /// <summary>
        /// output Excel File Path 
        /// </summary>
        public static string outputPath = @"C:\Users\faranam\Desktop\Exam\09 - Excel export\empty.xlsx";

        /// <summary>
        /// Data Table of CSV File
        /// </summary>
        public static System.Data.DataTable datatable = new System.Data.DataTable(tableName: "Excel Data");
        #endregion

        #region Read to File Function
        /// <summary>
        /// Read CSV file and Return Data as DataTable
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static System.Data.DataTable ReadFile(string path)
        {
            var rowsArray = File.ReadAllLines(path);
            var rowArray = new string[6];
            var column = new string[6];
            column = rowsArray[0].Split(',');
            foreach (var item in column)
            {
                datatable.Columns.Add(item);
            }

            for (int i = 1; i < rowsArray.Length; i++)
            {
                rowArray = rowsArray[i].Split(',');
                DataRow row;
                row = datatable.NewRow();
                for (int j = 0; j < 6; j++)
                {
                    row[column[j]] = rowArray[j];
                }
                datatable.Rows.Add(row: row);
            }

            return datatable;
        }
        #endregion

        #region Insert Data To Excel Function
        /// <summary>
        /// Write DataTable to Excel File
        /// </summary>
        /// <param name="input"></param>
        public static void InsertDataToExcel(System.Data.DataTable input)
        {
            input.ExportToExcel(excelFilePath: outputPath);
        }
        #endregion

        #region Main Function
        /// <summary>
        /// Main Function
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            InsertDataToExcel(ReadFile(pathCSV));

            Console.ReadKey();
        }
        #endregion
    }
}


