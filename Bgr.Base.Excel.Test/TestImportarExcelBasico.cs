using Bgr.Base.Excel;
using NUnit.Framework;

namespace Tests
{
    [TestFixture]
    public class TestImportarExcelBasico
    {

        [Test]
        public void TestImportar()
        {
            var importar = new ImportExcel();
            var data = importar.Import(@".\ArchivosInicializacion\CodigosUnspsc.xlsx");
            Assert.Greater(data.Rows.Count, 0);
        }


    }

}