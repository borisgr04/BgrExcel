using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace Bgr.Base.Excel.Contracts
{
    public interface IImportExcel
    {
        Action<string> Log { get; set; }

        void Import(Stream streamBook, bool useHeaderRow = true);
        void Import(byte[] book, bool useHeaderRow = true);
        void Import(string filePath, bool useHeaderRow = true);


        IList<T> GetSheet<T>(string sheetName);
        IList<T> GetSheet<T>(int sheetNumber);
        
        IList<T> Import<T>(byte[] book);
        IList<T> Import<T>(string filePath);

        IList<T> Import<T>(byte[] book, Func<DataRow, T> Map);
        IList<T> Import<T>(string filePath, Func<DataRow, T> Map);
        IList<T> GetSheet<T>(int sheetNumber, Func<DataRow, T> Map);
        IList<T> GetSheet<T>(string sheetName, Func<DataRow, T> Map);
    }
}
