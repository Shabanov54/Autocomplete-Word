﻿using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace Autocomplete_Word.FindWord
{
     internal class WriteWordFileName
    {
        private List<DataList> data;
        internal WriteWordFileName(List<DataList> data)
        {
            this.data = data;
        }
        internal void FindName()
        {
            string name = "%name%";
            string str = string.Empty;
            foreach (var item in data)
            {
                string newNamefile = $"Акт передачи OPENVPN ({item.name}).doc";
                object newPath = $@"C:\Акт передачи OPENVPN\ {newNamefile}";
                Word.Application app = new Word.Application();
                app.Documents.Open(ref newPath);
                Word.Find find = app.Selection.Find;

                find.Text = name;
                find.Replacement.Text = item.name;

                Object missing = Type.Missing;
                Object wrap = Word.WdFindWrap.wdFindContinue;
                Object replace = Word.WdReplace.wdReplaceAll;
                find.Execute(FindText: Type.Missing,
                    MatchCase: false,
                    MatchWholeWord: false,
                    MatchWildcards: false,
                    MatchSoundsLike: missing,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: wrap,
                    Format: false,
                    ReplaceWith: missing, Replace: replace);

                app.ActiveDocument.Save();
                app.ActiveDocument.Close();
            }
        }

    }
}
