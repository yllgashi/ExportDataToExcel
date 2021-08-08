using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Reflection;

namespace ExportDataToExcel
{
    public class Export
    {
        public static void ToExcel<T>(List<T> data)
        {
            // create excel sheet
            Application excel = CreateExcelSheet();

            // write excel header (properties)
            WriteExcelHeaderFromList(excel, data);

            // write data
            WriteExcelDataFromList(excel, data);

            // excel fix
            excel.Columns.AutoFit();
            excel.Visible = true;
        }


        public static void ToExcel(SqlDataReader dataReader)
        {
            // create excel sheet
            Application excel = CreateExcelSheet();

            // write excel header (properties)
            WriteExcelHeaderFromDataReader(excel, dataReader);

            // write data
            WriteExcelDataFromDataReader(excel, dataReader);

            // excel fix
            excel.Columns.AutoFit();
            excel.Visible = true;
        }

        #region Export from SqlDataReader
        private static void WriteExcelHeaderFromDataReader(Application excel, SqlDataReader dataReader)
        {
            for (int i = 1; i < dataReader.FieldCount + 1; i++)
            {
                excel.Cells[1, i] = dataReader.GetName(i - 1);
                string a = dataReader.GetName(i - 1);
            }
        }

        private static void WriteExcelDataFromDataReader(Application excel, SqlDataReader dataReader)
        {
            for (int i = 2; dataReader.Read(); i++)
            {
                for (int j = 1; j < dataReader.FieldCount + 1; j++)
                {
                    excel.Cells[i, j] = dataReader.GetValue(j - 1).ToString();
                    string a = dataReader.GetValue(j - 1).ToString();
                }
            }
        }
        #endregion

        #region Export from List
        private static void WriteExcelHeaderFromList<T>(Application excel, List<T> data)
        {
            List<string> props = GetPropNames(data[0]);

            for (int i = 1; i < props.Count + 1; i++)
            {
                excel.Cells[1, i] = props[i - 1];
            }
        }

        private static void WriteExcelDataFromList<T>(Application excel, List<T> data)
        {

            for (int i = 2; i < data.Count + 2; i++)
            {
                int j = 1;

                foreach (PropertyInfo prop in data[0].GetType().GetProperties())
                {

                    excel.Cells[i, j] = prop.GetValue(data[i - 2], null).ToString();
                    j++;
                }
                j = 0;
            }
        }
        #endregion

        #region Common
        private static Application CreateExcelSheet()
        {
            Application excel = new Application();
            excel.Application.Workbooks.Add(Type.Missing);

            return excel;
        }

        private static List<string> GetPropNames(object obj)
        {
            List<string> propsList = new List<string>();

            foreach (PropertyInfo prop in obj.GetType().GetProperties())
            {
                propsList.Add(prop.Name);
            }

            return propsList;
        }
        #endregion

    }
}