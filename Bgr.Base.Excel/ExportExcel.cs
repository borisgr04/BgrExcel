using Bgr.Base.Excel.Contracts;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;

namespace Bgr.Base.Excel
{
    /// <summary>
    /// Export data to Excel
    /// </summary>
    public class ExportExcel : IExportExcel
    {
        public Action<string> Log { get; set; }
        private DataSet _dataSet;
        public ExportExcel()
        {
            _dataSet = new DataSet();
        }
        /// <summary>
        /// Clean the data to export. 
        /// </summary>
        public void Clear()
        {
            _dataSet.Clear();
            _dataSet = new DataSet();
        }
        /// <summary>
        /// Add sheet to export
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="sheetName"></param>
        public void AddData<T>(IList<T> data, string sheetName = "Hoja1")
        {
            var dataTable= ConvertToDataTable<T>(data, sheetName);
            _dataSet.Tables.Add(dataTable);
        }
        /// <summary>
        /// Export to file
        /// </summary>
        /// <param name="filePath"></param>
        public void ExportFile(string filePath)
        {
            var bytes = ExportBytes();
            File.WriteAllBytes(filePath, bytes);
        }
        /// <summary>
        /// Export to bytes
        /// </summary>
        /// <returns>Book in bytes</returns>
        public byte[] ExportBytes()
        {
            var memory = Export();
            return memory.ToArray();
        }

        /// <summary>
        /// Export to MemoryStream
        /// </summary>
        /// <returns>Book in bytes</returns>
        public MemoryStream Export()
        {
            if (_dataSet.Tables.Count == 0)
            {
                throw new NoDataFoundException("No hay datos para exportar");
            }
            return ExportDataSetToExcel(_dataSet);
        }

        private MemoryStream ExportDataSetToExcel(DataSet ds)
        {
            MemoryStream memoryStream = new MemoryStream();
            using (XLWorkbook wb = new XLWorkbook())
            {
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    wb.Worksheets.Add(ds.Tables[i], ds.Tables[i].TableName);
                    Log?.Invoke("Exportó " + ds.Tables[i].TableName);

                }
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;
                wb.SaveAs(memoryStream);
            }
            return memoryStream;
        }
        private DataTable ConvertToDataTable<T>(IList<T> data, string tableName = "Hoja1")
        {
            DataTable table = null;
            if (data != null)
            {
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
                table = new DataTable();

                foreach (PropertyDescriptor prop in properties)
                {
                    var title=GetColumnName(prop);
                    table.Columns.Add(title, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);

                }

                foreach (T item in data)
                {
                    DataRow row = table.NewRow();
                    foreach (PropertyDescriptor prop in properties)
                    {
                        var columnName = GetColumnName(prop);
                        row[columnName] = prop.GetValue(item) ?? DBNull.Value;
                    }
                    table.Rows.Add(row);
                }

                table.TableName = tableName;
            }
            return table;
        }
        private string GetColumnName(PropertyDescriptor prop) 
        {
            return prop.Attributes.OfType<ColumExcelAttribute>().Any() ? prop.Attributes.OfType<ColumExcelAttribute>().First().Title : prop.Name;
        }
    }


    [Serializable]
    public class NoDataFoundException : Exception
    {
        public NoDataFoundException() { }
        public NoDataFoundException(string message) : base(message) { }
        public NoDataFoundException(string message, Exception inner) : base(message, inner) { }
        protected NoDataFoundException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
