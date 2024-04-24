using FileConverter;

namespace FileConverterTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            WordHelper wordHelper = new WordHelper();

            string sofficePath = @"D:\Program Files\LibreOffice\program\soffice.exe";
            wordHelper.Init(sofficePath);

            var result = wordHelper.ConverterToPDF(@"C:\Users\ivesBao\Desktop\test.docx", @"C:\Users\ivesBao\Desktop\testpdf");
            Console.WriteLine(result);
        }
    }
}