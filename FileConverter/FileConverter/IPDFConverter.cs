using System;
using System.Collections.Generic;
using System.Text;

namespace FileConverter
{
    public interface IPDFConverter
    {
        string ConverterToPDF(string inputFile, string outputFile);
    }
}
