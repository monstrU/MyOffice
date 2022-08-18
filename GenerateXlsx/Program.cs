using NCT.MyOfficeSDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateXlsx();
        }

        static void CreateXlsx()
        {
            Application application = new Application();

            
            var doc = application.createDocument(DocumentType.Workbook);

            var pos = doc.getRange().getBegin();
            UInt32 rows = 50;
            UInt32 colunms = 30;
            
            var table = pos.insertTable(rows, colunms, "List");
            table.getCell("A1").setText("Sellers");
            table.getCell(new CellPosition(1, 0)).setText("January");
            table.getCell(new CellPosition(1, 1)).setText("February");
            table.getCell(new CellPosition(1, 2)).setText("March");
            table.getCell(new CellPosition(1, 3)).setText("March");


            table.getCell(new CellPosition(2, 0)).setText("100");
            table.getCell(new CellPosition(2, 1)).setText("500");
            table.getCell(new CellPosition(2, 2)).setText("200");
            table.getCell(new CellPosition(2, 3)).setText("400");

            //table.getCell(new CellPosition(4, 4)).setText("текст на русском");
            doc.saveAs("Example.xlsx");
            Console.WriteLine("Сгенериран файл  Example.xlsx");
        }
    }
}
