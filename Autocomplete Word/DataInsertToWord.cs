using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Autocomplete_Word.FindWord;

namespace Autocomplete_Word
{
    internal class DataInsertToWord
    {
        private List<DataList> data;
        public DataInsertToWord(List<DataList> data)
        {
            this.data = data;
        }
        internal void FindFileAndCreate()
        {

            string wordBook = @"C:\Акт передачи OPENVPN\Акт передачи OPENVPN ().doc";
            string directory = @"C:\Акт передачи OPENVPN";
            FileInfo fileInfo = new FileInfo(directory);
            foreach (var item in data)
            {
                string nameFile = $"Акт передачи OPENVPN ({item.name}).doc";

                if (fileInfo.Name != nameFile)
                {
                    string newNamefile = $"Акт передачи OPENVPN ({item.name}).doc";
                    string newPath =  $@"C:\Акт передачи OPENVPN\ {newNamefile}";
                    FileInfo info = new FileInfo(newPath);
                    if (info.Exists ==false)
                    {
                        File.Copy(wordBook, newPath, true);
                    }
                }
            }
        }
        internal void FindWord()
        {
            WriteWordFilePass filePass = new WriteWordFilePass(data);
            filePass.FindPass();
            WriteWordFileName fileName = new WriteWordFileName(data);
            fileName.FindName();
            WriteWordFileLogin fileLogin = new WriteWordFileLogin(data);
            fileLogin.FindLogin();
            WriteWordFileDateTime dateTime = new WriteWordFileDateTime(data);
            dateTime.FindDateTime();
        }
    }
}
