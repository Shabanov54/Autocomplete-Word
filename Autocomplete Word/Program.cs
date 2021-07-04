using System;
using Autocomplete_Word.FindWord;

namespace Autocomplete_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            DataExel dataExel = new DataExel();
            var data =  dataExel.ReadExeleName();

            DataInsertToWord dataInsertToWord = new DataInsertToWord(data);
            dataInsertToWord.FindFileAndCreate();
            dataInsertToWord.FindWord();
        }
    }
}
