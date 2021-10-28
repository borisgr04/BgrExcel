using Bgr.Base.Contracts;
using Bgr.Base.Excel;
using NUnit.Framework;
using System;
using System.Data;

namespace Tests
{
    [TestFixture]
    class TestImportExcel
    {
        public string ObtenerPath(string nameFile) => AppDomain.CurrentDomain.BaseDirectory + nameFile;


        [Test]
        public void ImportarToDataTable()
        {
            var importar = new ImportExcel();
            importar.Import(ObtenerPath("Test.xlsx"), true);
            var data = importar.GetSheet(0);
            Assert.AreEqual(2677, data.Rows.Count);
        }

        [Test]
        public void ImportarToList()
        {
            var importar = new ImportExcel();
            importar.Import(ObtenerPath("Test.xlsx"), true);
            var data = importar.GetSheet<Data>(0);
            Console.WriteLine(data.Count);
            Assert.AreEqual(2677, data.Count);

        }

        //[Test]
        //public void ImportarSinHeaderToList()
        //{
        //    IImportExcel importar = new ImportExcel();
        //    importar.Import(ObtenerPath("Test.xlsx"),false);
        //    var data = importar.GetSheet(0);
        //    Console.WriteLine(data.Rows.Count);
        //    Assert.AreEqual(2678, data.Rows.Count);
        //}

        [Test]
        public void ImportarSinHeaderToListMap()
        {
            IImportExcel importar = new ImportExcel();
            importar.Import(ObtenerPath("Test.xlsx"));
            var data = importar.GetSheet(0, Mapper);
            Console.WriteLine(data.Count);
            Assert.AreEqual(2677, data.Count);
        }


        [Test]
        public void ImportarMap()
        {
            IImportExcel importar = new ImportExcel();
            importar.Import(ObtenerPath("Test.xlsx"));
            var data = importar.GetSheet(0, Mapper);
            Console.WriteLine(data.Count);
            Assert.AreEqual(2677, data.Count);
        }


        [Test]
        public void ImportarConHeaderDirectoToList()
        {
            IImportExcel importar = new ImportExcel();
            importar.Import(ObtenerPath("Test.xlsx"));
            var data = importar.GetSheet<DataPersonal>("Hoja3");
            Console.WriteLine(data.Count);
            Assert.AreEqual(1, data.Count);
        }

        [Test]
        public void ImportarConMapperDirectoToList()
        {
            IImportExcel importar = new ImportExcel();
            var data = importar.Import<Data>(ObtenerPath("Test.xlsx"), Mapper);
            Console.WriteLine(data.Count);
            Assert.AreEqual(2677, data.Count);
        }


        [Test]
        public void ImportarConMapperLamda()
        {
            IImportExcel importar = new ImportExcel();
            var data = importar.Import<Data>(ObtenerPath("Test.xlsx"),
                (DataRow dataRow) =>
                {
                    return new Data
                    {
                        Id = double.Parse(dataRow[0].ToString()),
                        Numero = double.Parse(dataRow[1].ToString()),
                        Cantidad = double.Parse(dataRow[1].ToString()),
                    };
                });
            Console.WriteLine(data.Count);
            Assert.AreEqual(2677, data.Count);
        }

        public Data Mapper(DataRow dataRow)
        {
            return new Data
            {
                Id = double.Parse(dataRow[0].ToString()),
                Numero = double.Parse(dataRow[1].ToString()),
                Cantidad = double.Parse(dataRow[1].ToString()),
            };
        }

    }
    class DataPersonal
    {
        public string Nombre { get; set; }
        public double Edad { get; set; }
        public DateTime Fecha { get; set; }
    }
    class Data
    {
        public double Id { get; set; }
        public double Numero { get; set; }
        public double Cantidad { get; set; }
    }
}
