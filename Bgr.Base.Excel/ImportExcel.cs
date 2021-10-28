using Bgr.Base.Excel.Contracts;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;

namespace Bgr.Base.Excel
{
    public class ImportExcel : IImportExcel
    {
        private DataSet _dataSet;
        public Action<string> Log { get; set; }
        public ImportExcel()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        public DataTable Import(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return ExportToDataTable(reader);
                }
            }
        }
        public DataTable Import(Stream file)
        {
            using (var reader = ExcelReaderFactory.CreateReader(file))
            {
                return ExportToDataTable(reader);
            }

        }

        private  DataTable ExportToDataTable(IExcelDataReader reader)
        {
            _dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });
            return _dataSet.Tables[0];
        }

        /// <summary>
        /// Sheet number starts at 0
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetNumber"></param>
        /// <returns></returns>
        public IList<T> GetSheet<T>(int sheetNumber)
        {
            return DataTableService.ConvertDataTable<T>(_dataSet.Tables[sheetNumber]);
        }
        public IList<T> GetSheet<T>(string sheetName)
        {
            return DataTableService.ConvertDataTable<T>(_dataSet.Tables[sheetName]);
        }
        public DataTable GetSheet(string sheetName)
        {
            return _dataSet.Tables[sheetName];
        }
        /// <summary>
        /// Sheet number starts at 0
        /// </summary>
        /// <param name="numeroHoja"></param>
        /// <returns></returns>
        public DataTable GetSheet(int sheetNumber)
        {
            return _dataSet.Tables[sheetNumber];
        }
        
        public void Import(string filePath, bool useHeaderRow = true)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                Import(stream, useHeaderRow);
            }
        }
 
        public void Import(byte[] book, bool useHeaderRow = true)
        {
            Import(new MemoryStream(book), useHeaderRow);
        }
        
        public void Import(Stream streamBook, bool useHeaderRow = true)
        {
            using (var reader = ExcelReaderFactory.CreateReader(streamBook))
            {
                _dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = useHeaderRow
                    }
                });
            }
        }
    
        public IList<T> Import<T>(byte[] book)
        {
            Import(book);
            return GetSheet<T>(0);
        }
    
        public IList<T> Import<T>(string filePath)
        {
            Import(filePath);
            return GetSheet<T>(0);
        }

        public IList<T> Import<T>(byte[] book, Func<DataRow, T> Map)
        {
            Import(book);
            return DataTableService.ConvertDataTable<T>(_dataSet.Tables[0], Map);
        }
       
        public IList<T> Import<T>(string filePath, Func<DataRow, T> Map)
        {
            Import(filePath);
            return DataTableService.ConvertDataTable<T>(_dataSet.Tables[0], Map);
        }
      
        public IList<T> GetSheet<T>(int sheetNumber, Func<DataRow, T> Map)
        {
            return DataTableService.ConvertDataTable<T>(_dataSet.Tables[sheetNumber], Map);
        }
        
        public IList<T> GetSheet<T>(string sheetName, Func<DataRow, T> Map)
        {
            return DataTableService.ConvertDataTable<T>(_dataSet.Tables[sheetName], Map);
        }
    }

    internal static class DataTableService
    {
        public static List<T> ConvertDataTable<T>(DataTable dt, Func<DataRow, T> Map)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = Map(row);
                data.Add(item);
            }
            return data;
        }

        public static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }

        private static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                    {
                        pro.SetValue(obj, dr[column.ColumnName], null);
                    }
                }
            }
            return obj;
        }

    }
}
