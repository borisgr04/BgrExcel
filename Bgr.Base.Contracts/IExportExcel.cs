using System;
using System.Collections.Generic;
using System.IO;

namespace Bgr.Base.Excel.Contracts
{
    public interface IExportExcel
    {
        Action<string> Log { get; set; }

        void AddData<T>(IList<T> data, string sheetName = "Hoja1");
        void Clear();
        MemoryStream Export();
        byte[] ExportBytes();
        void ExportFile(string filePath);
    }

    public class ColumExcelAttribute : Attribute
    {
        public ColumExcelAttribute(string title)
        {
            Title = title;
        }
        public string Title { get; set; }
    }
}