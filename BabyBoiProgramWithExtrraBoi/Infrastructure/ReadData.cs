using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BabyBoiProgramWithExtrraBoi.Infrastructure
{
    internal class ReadData
    {
        public List<List<string>> Data { get; private set; }
        public List<string> RawData { get; private set; }

        public string FilePath { get; private set; }
        private const char keyDelimeter = '_';
        private const int spacesPerColumn = 4;

        public ReadData(string filePath)
        {
            FilePath = filePath;
            RawData = new List<string>();
            Data = new List<List<string>>();
        }

        /// <summary>
        /// Read the  file  in
        /// </summary>
        public void ReadFile()
        {
            using (StreamReader sr = new StreamReader(FilePath))
            {
                string line = "";
                while ((line = sr.ReadLine()) != null)
                {
                    RawData.Add(line);
                }
            }
        }

        /// <summary>
        /// Formats the data to be converted to excel
        /// </summary>
        /// <param name="delimeter"></param>
        public void ProcessData()
        {
            foreach (string line in RawData)
            {
                string word = "";
                char[] lineSplit = line.ToCharArray();
                List<string> dataLine = new List<string>();
                int tabsBeforeWord = 0;
                for (int i = 0; i < lineSplit.Length; i++)
                {
                    if ((!word.Equals("") && lineSplit[i].Equals(' ') && !lineSplit[i + 1].Equals(' ')) || !lineSplit[i].Equals(' '))
                    {
                        word += lineSplit[i];
                    }
                    else if (!word.Equals(""))
                    {
                        for (int k = 0; k < tabsBeforeWord / 4; k++)
                        {
                            dataLine.Add("");
                        }
                        dataLine.Add(word);
                        word = "";
                        tabsBeforeWord = 0;
                    }
                    else if (lineSplit[i].Equals(' '))
                    {
                        tabsBeforeWord++;
                    }
                }

                if (!word.Equals(""))
                {
                    for (int i = 0; i < tabsBeforeWord / 4; i++)
                    {
                        dataLine.Add("");
                    }
                    dataLine.Add(word);
                }

                Data.Add(dataLine);
            }

            GroupData();
        }

        /// <summary>
        /// Groups the data together inth the Data variable
        /// </summary>
        private void GroupData()
        {
            int numOfBlankLines = 0;
            bool isRecordsToGroup = false;
            Dictionary<string, List<List<string>>> groupedLines = new Dictionary<string, List<List<string>>>();
            foreach (List<string> line in Data)
            {
                if (line.Count > 0)
                {
                    if (!isRecordsToGroup)
                    {
                        if (numOfBlankLines == 3)
                        {
                            isRecordsToGroup = true;
                        }
                    }
                    else
                    {
                        string key = "";

                        //As per Shelly 'That shouldn't be blank. It was an error in creating the work order. Don't worrry about it.'
                        //Ignoreing lines that have blanks
                        if (line.Count == 7)
                        {
                            key = line[0] + keyDelimeter + line[5] + keyDelimeter + line[6];

                            if (!groupedLines.ContainsKey(key))
                            {
                                groupedLines.Add(key, new List<List<string>>() { line });
                            }
                            else
                            {
                                groupedLines[key].Add(line);
                            }
                        }
                    }
                    numOfBlankLines = 0;
                }
                else
                {
                    numOfBlankLines++;
                    if (isRecordsToGroup)
                    {
                        break;
                    }
                }
            }

            CombineData(groupedLines);
        }

        /// <summary>
        /// Combine all similiar entries
        /// </summary>
        /// <param name="dictionary"></param>
        private void CombineData(Dictionary<string, List<List<string>>> dictionary)
        {
            foreach (string key in dictionary.Keys)
            {
                if (dictionary[key].Count > 1)
                {
                    List<string> newLine = dictionary[key][0];
                    double total = 0;
                    int insertIndex = Data.FindIndex(x => dictionary[key][0].Equals(x));

                    foreach (List<string> line in dictionary[key])
                    {
                        total += Convert.ToDouble(line[4]);
                        Data.Remove(line);
                    }

                    newLine[4] = Convert.ToString(Math.Ceiling(total * 100) / 100);
                    Data.Insert(insertIndex, newLine);
                }
            }
        }
    }
}
