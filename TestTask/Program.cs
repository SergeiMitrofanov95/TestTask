using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestTask
{
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string, string> machine_tools = new Dictionary<string, string>();
            Dictionary<string, int> machine_busy = new Dictionary<string, int>();
            Dictionary<string, string> nomenclatures = new Dictionary<string, string>();
            Dictionary<string, int> parties = new Dictionary<string, int>();
            Dictionary<string, Dictionary<string, int>> times = new Dictionary<string, Dictionary<string, int>>();

            Excel.Application ObjExcel = new Excel.Application();
            Excel.Workbook ObjBook = ObjExcel.Workbooks.Open(@"C:\Users\123\Desktop\Тестовое задание\Тестовое задание\machine_tools.xlsx");
            Excel.Worksheet ObjSheet = (Excel.Worksheet)ObjBook.Sheets[1];
            var lastCell = ObjSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int lastRow = (int)lastCell.Row - 1;
            for (int i = 0; i < lastRow; i++)
            {
                machine_tools.Add(ObjSheet.Cells[i + 2, 1].Text.ToString(), ObjSheet.Cells[i + 2, 2].Text.ToString());
                machine_busy.Add(ObjSheet.Cells[i + 2, 1].Text.ToString(), 0);
            }
            ObjBook.Close(false);

            ObjBook = ObjExcel.Workbooks.Open(@"C:\Users\123\Desktop\Тестовое задание\Тестовое задание\nomenclatures.xlsx");
            ObjSheet = (Excel.Worksheet)ObjBook.Sheets[1];
            lastCell = ObjSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            lastRow = (int)lastCell.Row - 1;
            for (int i = 0; i < lastRow; i++)
            {
                nomenclatures.Add(ObjSheet.Cells[i + 2, 1].Text.ToString(), ObjSheet.Cells[i + 2, 2].Text.ToString());
            }
            ObjBook.Close(false);

            ObjBook = ObjExcel.Workbooks.Open(@"C:\Users\123\Desktop\Тестовое задание\Тестовое задание\parties.xlsx");
            ObjSheet = (Excel.Worksheet)ObjBook.Sheets[1];
            lastCell = ObjSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            lastRow = (int)lastCell.Row - 1;
            for (int i = 0; i < lastRow; i++)
            {
                if(parties.ContainsKey(ObjSheet.Cells[i + 2, 2].Text.ToString()))
                {
                    parties[ObjSheet.Cells[i + 2, 2].Text.ToString()]++;
                }
                else
                {
                    parties.Add(ObjSheet.Cells[i + 2, 2].Text.ToString(), 1);
                }
            }
            ObjBook.Close(false);
            ObjExcel.Quit();

            ObjBook = ObjExcel.Workbooks.Open(@"C:\Users\123\Desktop\Тестовое задание\Тестовое задание\times.xlsx");
            ObjSheet = (Excel.Worksheet)ObjBook.Sheets[1];
            lastCell = ObjSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            lastRow = (int)lastCell.Row - 1;
            for (int i = 0; i < lastRow; i++)
            {
                if (times.ContainsKey(ObjSheet.Cells[i + 2, 1].Text.ToString()))
                {
                    times[ObjSheet.Cells[i + 2, 1].Text.ToString()].Add(ObjSheet.Cells[i + 2, 2].Text.ToString(), Convert.ToInt32(ObjSheet.Cells[i + 2, 3].Text.ToString()));
                }
                else
                {
                    times.Add(ObjSheet.Cells[i + 2, 1].Text.ToString(), new Dictionary<string, int>());
                    times[ObjSheet.Cells[i + 2, 1].Text.ToString()].Add(ObjSheet.Cells[i + 2, 2].Text.ToString(), Convert.ToInt32(ObjSheet.Cells[i + 2, 3].Text.ToString()));
                }
            }
            ObjBook.Close(false);

            List<string> times_keys = new List<string>();
            times_keys.AddRange(times.Keys);
            foreach(string id_machine in times_keys)
            {
                times[id_machine] = times[id_machine].OrderBy(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);
            }
            times_keys.Clear();

            ObjBook = ObjExcel.Workbooks.Add();
            ObjSheet = (Excel.Worksheet)ObjBook.Sheets[1];
            int Volume_Parties = 0;
            int time = 0;
            foreach(int val in parties.Values)
            {
                Volume_Parties += val;
            }
            Console.WriteLine("Партия\tОборудование\tВремя начала\tВремя окончания");
            ObjSheet.Cells[1, 1].Value = "Партия";
            ObjSheet.Cells[1, 2].Value = "Оборудование";
            ObjSheet.Cells[1, 3].Value = "Время начала";
            ObjSheet.Cells[1, 4].Value = "Время окончания";
            int row = 2;
            while (Volume_Parties > 0)
            {
                foreach (string id_machine in times.Keys)
                {
                    if (machine_busy[id_machine] <= time)
                    {
                        foreach(string id_nomenclatures in times[id_machine].Keys)
                        {
                            if (parties[id_nomenclatures] > 0)
                            {
                                parties[id_nomenclatures]--;
                                Volume_Parties--;
                                machine_busy[id_machine] += times[id_machine][id_nomenclatures];
                                Console.WriteLine(nomenclatures[id_nomenclatures] + "\t" + machine_tools[id_machine] + "\t\t" + (machine_busy[id_machine] - times[id_machine][id_nomenclatures]) + "\t\t" + machine_busy[id_machine]);
                                ObjSheet.Cells[row, 1].Value = nomenclatures[id_nomenclatures];
                                ObjSheet.Cells[row, 2].Value = machine_tools[id_machine];
                                ObjSheet.Cells[row, 3].Value = (machine_busy[id_machine] - times[id_machine][id_nomenclatures]);
                                ObjSheet.Cells[row, 4].Value = machine_busy[id_machine];
                                row++;
                                break;
                            }
                        }
                    }
                }
                time += 10;
            }
            ObjBook.SaveAs(@"C:\Users\123\Desktop\Тестовое задание\Тестовое задание\Plan.xlsx");
            ObjBook.Close(true);
            ObjExcel.Quit();
            GC.Collect();
            Console.WriteLine("Нажмите кнопку для продолжения...");
            Console.ReadKey();
        }
    }
}
