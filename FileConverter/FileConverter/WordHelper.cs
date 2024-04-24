using System.IO;
using System.Linq;

namespace FileConverter
{
    public class WordHelper : IPDFConverter
    {
        /// <summary>
        /// 可执行程序路径
        /// </summary>
        public string SofficePath { get; set; }

        /// <summary>
        /// Converts the input file to PDF format and saves it in the output directory.
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFileDir">导出文件输出目录</param>
        /// <returns></returns>
        public string ConverterToPDF(string inputFile, string outputFileDir)
        {
            Check(inputFile, outputFileDir);

            var connecting = new ProcessCommandBase(SofficePath);

            connecting.AddParameter($"--invisible --convert-to pdf");
            connecting.AddParameter($"\"{inputFile}\"");
            connecting.AddParameter($"--outdir");
            connecting.AddParameter($"\"{outputFileDir}\"");

            connecting.Exec(true);


            return ResolveFinalPath(inputFile, outputFileDir);
        }

        public void Init(string exePath)
        {
            SofficePath = exePath;
        }

        public void Check(string inputFile, string outputFileDir)
        {
            if (!File.Exists(inputFile))
            {
                throw new System.Exception("Input file not found");
            }

            if (!File.Exists(SofficePath))
            {
                throw new System.Exception("Soffice exe not found");
            }

            var fileInfo = new FileInfo(inputFile);
            var lastIndex = fileInfo.Name.LastIndexOf(fileInfo.Extension);
            if (lastIndex == -1)
            {
                throw new System.Exception("Error origin file type");
            }
        }


        private string ResolveFinalPath(string inputFile, string outputFileDir)
        {
            Check(inputFile, outputFileDir);

            var fileInfo = new FileInfo(inputFile);
            var lastIndex = fileInfo.Name.LastIndexOf(fileInfo.Extension);
            var originClearName = fileInfo.Name.Remove(lastIndex);

            var files = Directory.EnumerateFiles(outputFileDir).ToList();
            foreach (var item in files)
            {
                var finalFileInfo = new FileInfo(item);
                if (finalFileInfo.Name.StartsWith(originClearName) && finalFileInfo.Name.ToLower().Contains(".pdf"))
                {
                    return item;
                }
            }

            return string.Empty;
        }
    }
}
