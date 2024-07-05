using FileConverter;

namespace FileConverterTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            ConverterHelper convertHelper = new ConverterHelper();

            string sofficePath = @"D:\Program Files\LibreOffice\program\soffice.exe";
            convertHelper.Init(sofficePath);

            var result = convertHelper.ConverterToPDF(@"C:\Users\ivesBao\Desktop\test.docx", @"C:\Users\ivesBao\Desktop\testpdf");
            Console.WriteLine(result);
        }
    }
}