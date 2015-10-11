using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;



namespace azk_excel
{
    enum ColumnIndex
    {
        MARKET_COLUMN = 2,
        REFERENCE_COLUMN = 4,
        REFERENCE_NAME_COLUMN = 5,
        COLOR_COLUMN = 6,
        COLOR_CODE_COLUMN = 7,
        SEASON_COLUMN = 9,
        SIZE_COLUMN = 12,
        PRICE_COLUMN = 14,
        QTY_COLUMN = 15,
        NOTE_COLUMN = 16,
    };

    class OrderData
    {
        public string season_;
        public string market_;
        public string reference_;
        public string referenceName_;
        public string colorCode_;
        public string color_;
        public string size_;
        public int price_;
        public int qty_;
        public string note_;
    }
    class DataExcelManager
    {
        private string excelPath_;
        private string sheetName_;
        private Microsoft.Office.Interop.Excel.Application excelApp_;
        //private string[] keyList_;
        //private Hashtable keyTable_;

        private int getSheetIndex(Microsoft.Office.Interop.Excel.Workbook book, string sheetName)
        {
            int i = 0;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sh in book.Sheets)
            {
                if (sheetName == sh.Name)
                {
                    return i + 1;
                }
                i += 1;
            }
            return -1;
        }

        private Microsoft.Office.Interop.Excel.Worksheet getSheet()
        {
            Microsoft.Office.Interop.Excel.Workbook dataBook = excelApp_.Workbooks.Open(excelPath_, 0, true );
            int sheetId = getSheetIndex(dataBook, sheetName_);
            Microsoft.Office.Interop.Excel.Worksheet sheet = dataBook.Sheets[sheetId];

            return sheet;
        }

        /*
        private void loadKeyTable(Microsoft.Office.Interop.Excel.Worksheet sheet)
        {
            foreach (string key in keyList_)
            {
                for (int i = 1; i < 100; i++)
                {
                    Microsoft.Office.Interop.Excel.Range rgn = sheet.Cells[3, i];
                    dynamic val = rgn.Value2;
                    string title = Convert.ToString(val);

                    int indexof = title.IndexOf(key);
                    if (indexof != -1)
                    {
                        keyTable_.Add(key, i);
                        break;
                    }
                }
            }
        }
        */

        public DataExcelManager(Microsoft.Office.Interop.Excel.Application app, string path, string sheet)
        {
            excelApp_ = app;
            excelPath_ = path;
            sheetName_ = sheet;
            //keyList_ = new string[] { "Collection Retail", "Market", "Product reference", "Product Reference Name", "Product color code", "Product color", "Product size", "YEN", "Note" };
            //keyTable_ = new Hashtable();
        }

        /*
        public void Test()
        {
            loadKeyTable(getSheet());

            foreach (DictionaryEntry kv in keyTable_)
            {
                Console.WriteLine(kv.Key + ":" + kv.Value);
            }
        }
        */

        public List<OrderData> CreateList()
        {
            List<OrderData> ret = new List<OrderData>();

            Microsoft.Office.Interop.Excel.Worksheet sheet = getSheet();
            //loadKeyTable(sheet);

            int index = 4;
            while (true)
            {
                Microsoft.Office.Interop.Excel.Range rgn = sheet.Cells[index, 1];
                dynamic val = rgn.Value2;
                string tmp = Convert.ToString(val);
                if (tmp == null)
                {
                    break;
                }

                OrderData data = new OrderData();

                rgn = sheet.Cells[index, ColumnIndex.SEASON_COLUMN];
                val = rgn.Value2;
                data.season_ = Convert.ToString(val);

                rgn = sheet.Cells[index, ColumnIndex.MARKET_COLUMN];
                val = rgn.Value2;
                data.market_ = Convert.ToString(val);

                rgn = sheet.Cells[index, ColumnIndex.REFERENCE_COLUMN];
                val = rgn.Value2;
                data.reference_ = Convert.ToString(val);

                rgn = sheet.Cells[index, ColumnIndex.REFERENCE_NAME_COLUMN];
                val = rgn.Value2;
                data.referenceName_ = Convert.ToString(val);

                rgn = sheet.Cells[index, ColumnIndex.COLOR_CODE_COLUMN];
                val = rgn.Value2;
                data.colorCode_ = Convert.ToString(val);

                rgn = sheet.Cells[index, ColumnIndex.COLOR_COLUMN];
                val = rgn.Value2;
                data.color_ = Convert.ToString(val);

                rgn = sheet.Cells[index, ColumnIndex.SIZE_COLUMN];
                val = rgn.Value2;
                data.size_ = Convert.ToString(val);

                rgn = sheet.Cells[index, ColumnIndex.PRICE_COLUMN];
                val = rgn.Value2;
                data.price_ = Convert.ToInt32(val);

                rgn = sheet.Cells[index, ColumnIndex.QTY_COLUMN];
                val = rgn.Value2;
                data.qty_ = Convert.ToInt32(val);

                rgn = sheet.Cells[index, ColumnIndex.NOTE_COLUMN];
                val = rgn.Value2;
                data.note_ = Convert.ToString(val);

                ret.Add(data);
                index++;
            }
            return ret;
        }
    }

    class Program
    {
        static Microsoft.Office.Interop.Excel.Application excelApp_;

        static void WriteExcel(string path, List<OrderData>dataList)
        {
            Microsoft.Office.Interop.Excel.Workbook outBook = excelApp_.Workbooks.Open(path);
            Microsoft.Office.Interop.Excel.Worksheet sheet = outBook.Sheets[1];

            int index = 0;

            foreach (OrderData data in dataList)
            {
                for (int i = 0; i < data.qty_; i++)
                {
                    int column = index / 6 + 1;
                    int top = (index % 6) * 9 + 1;

                    Microsoft.Office.Interop.Excel.Range range = sheet.Cells[top, column];
                    range.Value2 = "Christian Louboutin";

                    range = sheet.Cells[top + 1, column];
                    range.Value2 = data.season_ + " / " + data.market_;

                    range = sheet.Cells[top + 2, column];
                    range.Value2 = data.reference_;

                    range = sheet.Cells[top + 3, column];
                    range.Value2 = data.referenceName_;

                    range = sheet.Cells[top + 4, column];
                    range.Value2 = data.colorCode_ + " " + data.color_;

                    range = sheet.Cells[top + 5, column];
                    if (data.size_ == "TU")
                    {
                        range.Value2 = "SIZE:";
                    }
                    else
                    {
                        range.Value2 = "SIZE: " + data.size_;
                    }

                    if (data.price_ == 0)
                    {
                        range = sheet.Cells[top + 6, column];
                        range.Value2 = "N/A";
                        range = sheet.Cells[top + 7, column];
                        range.Value2 = "N/A";
                    }
                    else
                    {
                        range = sheet.Cells[top + 6, column];
                        range.Value2 = data.price_ * 1.08;
                        range = sheet.Cells[top + 7, column];
                        range.Value2 = data.price_;
                    }

                    range = sheet.Cells[top + 8, column];
                    range.Value2 = data.note_;

                    index++;
                }

            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Input Order excel path and return");
            string orderExcelPath = Console.ReadLine();

            //Console.WriteLine("Input Order excel sheet");
            string sheetName = "Press Order";

            excelApp_ = new Microsoft.Office.Interop.Excel.Application();
            excelApp_.Visible = false;

            List<OrderData> dataList;

            DataExcelManager dataManager = new DataExcelManager(excelApp_, orderExcelPath, sheetName);
            //dataManager.Test();
            dataList = dataManager.CreateList();

            Console.WriteLine("Output excel sheet");
            string outExcelPath = Console.ReadLine();

            WriteExcel(outExcelPath, dataList);

            excelApp_.Workbooks.Close();
            excelApp_.Quit();
            
        }
    }
}
