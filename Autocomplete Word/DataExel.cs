using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Text;

namespace Autocomplete_Word
{
    internal class DataExel
    {
        List<DataList> data = new List<DataList>();
        internal List<DataList> ReadExeleName()
        {
            string exceleBook = @"C:\Акт передачи OPENVPN\Данные.xlsx";

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(exceleBook, FileMode.Open,FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet().Tables[0];
                    foreach (DataRow row in result.Rows)
                    {
                        if (!string.IsNullOrEmpty(row[4].ToString()))
                        {
                            DataList dataList = new DataList();
                            
                            dataList.id = string.IsNullOrEmpty(row[0].ToString()) ? null : row[0].ToString();
                            dataList.name = string.IsNullOrEmpty(row[1].ToString()) ? null : row[1].ToString();
                            dataList.login = string.IsNullOrEmpty(row[2].ToString()) ? null : row[2].ToString();
                            dataList.pass = string.IsNullOrEmpty(row[3].ToString()) ? null : row[3].ToString();
                            dataList.dateTime = string.IsNullOrEmpty(row[4].ToString()) ? null : row[4].ToString();
                            if (dataList.id != "id" )
                            {
                                if (dataList.dateTime != null)
                                {
                                    var parsedate = DateTime.Parse(dataList.dateTime);
                                    dataList.dateTime = parsedate.ToString("d");
                                }
                                data.Add(dataList);
                            }

                        }
                    }
                    return data;

                }
            }
        }
    }
}
