using BabyBoiProgramWithExtrraBoi.Infrastructure;

namespace BabyBoiProgramWithExtrraBoi
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string inputDirectory = args[0];
            string inputFileName = args[1];
            string outDirectory = args[2];
            string outputFileName = args[3];
            ConvertToExcel convertToExcel = new ConvertToExcel(Path.Combine(inputDirectory, inputFileName), Path.Combine(outDirectory, outputFileName));

            convertToExcel.LoadFile();
            convertToExcel.ExportToExcel();
        }
    }
}