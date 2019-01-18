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
            Dictionary<string, int> mp = new Dictionary<string, int>();
            Dictionary<string, string> nomenclatures = new Dictionary<string, string>();
            Dictionary<string, int> parties = new Dictionary<string, int>();
            Dictionary<string, Dictionary<string, int>> times = new Dictionary<string, Dictionary<string, int>>();
            List<string> plan_parties = new List<string>();
            List<string> plan_machine = new List<string>();
            List<int> plan_time_start = new List<int>();
            List<int> plan_time_end = new List<int>();
            //считывание файла "machine-tools"
            Excel.Application ObjExcel = new Excel.Application();
            Excel.Workbook ObjBook = ObjExcel.Workbooks.Open(@"C:\Users\123\Desktop\Тестовое задание\Тестовое задание\machine_tools.xlsx");
            Excel.Worksheet ObjSheet = (Excel.Worksheet)ObjBook.Sheets[1];
            var lastCell = ObjSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int lastRow = (int)lastCell.Row - 1;
            for (int i = 0; i < lastRow; i++)
            {
                machine_tools.Add(ObjSheet.Cells[i + 2, 1].Text.ToString(), ObjSheet.Cells[i + 2, 2].Text.ToString());
                machine_busy.Add(ObjSheet.Cells[i + 2, 1].Text.ToString(), 0);
                mp.Add(ObjSheet.Cells[i + 2, 1].Text.ToString(), 0);
            }
            ObjBook.Close(false);
            //считывание файла "nomenclatures"
            ObjBook = ObjExcel.Workbooks.Open(@"C:\Users\123\Desktop\Тестовое задание\Тестовое задание\nomenclatures.xlsx");
            ObjSheet = (Excel.Worksheet)ObjBook.Sheets[1];
            lastCell = ObjSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            lastRow = (int)lastCell.Row - 1;
            for (int i = 0; i < lastRow; i++)
            {
                nomenclatures.Add(ObjSheet.Cells[i + 2, 1].Text.ToString(), ObjSheet.Cells[i + 2, 2].Text.ToString());
            }
            ObjBook.Close(false);
            //считывание файла "parties"
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
            //считывание файла "times"
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
            //сортировка времени обработки материалов каждой из печей по возрастанию
            List<string> times_keys = new List<string>();
            times_keys.AddRange(times.Keys);
            foreach(string id_machine in times_keys)
            {
                times[id_machine] = times[id_machine].OrderBy(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);
            }
            times_keys.Clear();
            //проверка работоспособности машин
            bool log = true;
            if (machine_tools.Count > 0)
            {
                Console.WriteLine("Количество работоспособных машин {0}:", machine_tools.Count);
                foreach (string id_machine in machine_tools.Keys)
                {
                    Console.WriteLine(machine_tools[id_machine]);
                }
                Console.WriteLine("Если одна из машин вышла из строя, напишите ее название.");
                Console.WriteLine("Если все машины работоспособны, напишите \"Далее\".");
            }
            else
            {
                Console.WriteLine("Нет работоспособных машин!");
                Console.WriteLine("Нажмите кнопку для продолжения...");
                Console.ReadKey();
                ObjExcel.Quit();
                return;
            }
            while (log)
            {
                string con = Console.ReadLine();
                if (machine_tools.ContainsValue(con))
                {
                    foreach(string id_machine in machine_tools.Keys)
                    {
                        if (machine_tools[id_machine] == con)
                        {
                            machine_tools.Remove(id_machine);
                            machine_busy.Remove(id_machine);
                            mp.Remove(id_machine);
                            times.Remove(id_machine);
                            break;
                        }
                    }
                    if (machine_tools.Count > 0)
                    {
                        Console.WriteLine("Количество работоспособных машин {0}:", machine_tools.Count);
                        foreach (string id_machine2 in machine_tools.Keys)
                        {
                            Console.WriteLine(machine_tools[id_machine2]);
                        }
                        Console.WriteLine("Введите следующую команду.");
                    }
                    else
                    {
                        Console.WriteLine("Нет работоспособных машин!");
                        Console.WriteLine("Нажмите кнопку для продолжения...");
                        Console.ReadKey();
                        ObjExcel.Quit();
                        return;
                    }
                }
                else if (con == "Далее")
                {
                    break;
                }
                else
                {
                    Console.WriteLine("Команда введена не верно, повторите ввод.");
                }
            }
            //проверка партий
            foreach(string id_machine in times.Keys)
            {
                foreach(string id_nomenclatures in times[id_machine].Keys)
                {
                    if (!times_keys.Contains(id_nomenclatures))
                    {
                        times_keys.Add(id_nomenclatures);
                    }
                }
            }
            List<string> parties_keys = new List<string>();
            parties_keys.AddRange(parties.Keys);
            foreach (string id_nomenclatures in parties_keys)
            {
                if (!times_keys.Contains(id_nomenclatures))
                {
                    Console.WriteLine("{0} не обрабатывается действующим оборудованием,", nomenclatures[id_nomenclatures]);
                    Console.WriteLine("следовательно данный материал будет удален из очереди.");
                    parties.Remove(id_nomenclatures);
                }
            }
            times_keys.Clear();
            parties_keys.Clear();
            if (parties.Count > 0)
            {
                Console.WriteLine("Необходимо обработать следующие материалы:");
                foreach(string id_parties in parties.Keys)
                {
                    Console.WriteLine("{0} в количестве {1} партий.", nomenclatures[id_parties], parties[id_parties]);
                }
                Console.WriteLine("Если какой-то материал не надо обрабатывать,");
                Console.WriteLine("напишите \"Удалить [Название материала]\".");
                Console.WriteLine("Если нужно увеличить или уменьшить количество материала для обработки,");
                Console.WriteLine("напишите \"Изменить [Название материала] [Количество]\".");
                Console.WriteLine("Если есть необходимость обработать некоторый материал в первую очередь,");
                Console.WriteLine("напишите \"Обработать [Название материала]\".");
                Console.WriteLine("Для продолжения планирования, напишите \"Далее\".");
            }
            else
            {
                Console.WriteLine("Нет материалов для обработки!");
                Console.WriteLine("Нажмите кнопку для продолжения...");
                Console.ReadKey();
                ObjExcel.Quit();
                return;
            }
            while (log)
            {
                string[] con = Console.ReadLine().Split();
                if (con[0] == "Удалить")
                {
                    if (con.Length == 2)
                    {
                        bool log2 = true;
                        foreach (string id_nomenclatures in parties.Keys)
                        {
                            if (nomenclatures[id_nomenclatures] == con[1])
                            {
                                parties.Remove(id_nomenclatures);
                                log2 = false;
                                break;
                            }
                        }
                        if (log2)
                        {
                            Console.WriteLine("Неверно введено название материала, повторите ввод.");
                        }
                        else if (parties.Count > 0)
                        {
                            Console.WriteLine("Необходимо обработать следующие материалы:");
                            foreach (string id_parties in parties.Keys)
                            {
                                Console.WriteLine("{0} в количестве {1} партий.", nomenclatures[id_parties], parties[id_parties]);
                            }
                            Console.WriteLine("Введите следующую команду.");
                        }
                        else
                        {
                            Console.WriteLine("Нечего обрабатывать!");
                            Console.WriteLine("Нажмите кнопку для продолжения...");
                            Console.ReadKey();
                            ObjExcel.Quit();
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Неверный формат ввода команды, повторите ввод.");
                    }
                }
                else if (con[0] == "Изменить")
                {
                    if (con.Length == 3)
                    {
                        bool log2 = true;
                        bool log3 = false;
                        foreach (string id_nomenclatures in parties.Keys)
                        {
                            if (nomenclatures[id_nomenclatures] == con[1])
                            {
                                log2 = false;
                                int a;
                                try
                                {
                                    a = Convert.ToInt32(con[2]);
                                    if(a < 0 && Math.Abs(a) > parties[id_nomenclatures])
                                    {
                                        parties[id_nomenclatures] = 0;
                                    }
                                    else
                                    {
                                        parties[id_nomenclatures] += a;
                                    }
                                    log3 = true;
                                    break;
                                }
                                catch (FormatException)
                                {
                                    Console.WriteLine("Неверно введено количество. Введите целое число.");
                                    break;
                                }
                            }
                        }
                        if (log2)
                        {
                            Console.WriteLine("Неверно введено название материала, повторите ввод.");
                        }
                        else if(log3)
                        {
                            Console.WriteLine("Необходимо обработать следующие материалы:");
                            foreach (string id_parties in parties.Keys)
                            {
                                Console.WriteLine("{0} в количестве {1} партий.", nomenclatures[id_parties], parties[id_parties]);
                            }
                            Console.WriteLine("Введите следующую команду.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Неверный формат ввода команды, повторите ввод.");
                    }
                }
                else if (con[0] == "Обработать")
                {
                    if (con.Length == 2)
                    {
                        bool log2 = true;
                        foreach (string id_nomenclatures in parties.Keys)
                        {
                            if (nomenclatures[id_nomenclatures] == con[1])
                            {
                                Dictionary<string, int> times_2 = new Dictionary<string, int>();
                                foreach(string id_machine in times.Keys)
                                {
                                    if (times[id_machine].ContainsKey(id_nomenclatures))
                                    {
                                        times_2.Add(id_machine, times[id_machine][id_nomenclatures]);
                                    }
                                }
                                times_2 = times_2.OrderBy(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);
                                int time1 = 0;
                                while (parties[id_nomenclatures] > 0)
                                {
                                    foreach (string id_machine in times_2.Keys)
                                    {
                                        if (machine_busy[id_machine] <= time1)
                                        {
                                            if (parties[id_nomenclatures] > 0)
                                            {
                                                parties[id_nomenclatures]--;
                                                machine_busy[id_machine] += times[id_machine][id_nomenclatures];
                                                plan_parties.Add(id_nomenclatures);
                                                plan_machine.Add(id_machine);
                                                plan_time_start.Add(time1);
                                                plan_time_end.Add(machine_busy[id_machine]);
                                            }
                                        }
                                    }
                                    time1 += 10;
                                }
                                parties.Remove(id_nomenclatures);
                                foreach(string id_machine in times.Keys)
                                {
                                    times[id_machine].Remove(id_nomenclatures);
                                }
                                log2 = false;
                                break;
                            }
                        }
                        if (log2)
                        {
                            Console.WriteLine("Неверно введено название материала, повторите ввод.");
                        }
                        else if (parties.Count > 0)
                        {
                            Console.WriteLine("Необходимо обработать следующие материалы:");
                            foreach (string id_parties in parties.Keys)
                            {
                                Console.WriteLine("{0} в количестве {1} партий.", nomenclatures[id_parties], parties[id_parties]);
                            }
                            Console.WriteLine("Введите следующую команду.");
                        }
                        else
                        {
                            for (int index = 0; index < plan_machine.Count; ++index)
                            {
                                ++mp[plan_machine[index]];
                            }
                            //вывод плана в консоль
                            Console.WriteLine("Партия\tОборудование\tВремя начала\tВремя окончания");
                            for (int index = 0; index < plan_parties.Count; ++index)
                            {
                                Console.WriteLine(nomenclatures[plan_parties[index]] + "\t" + machine_tools[plan_machine[index]] + "\t\t" + plan_time_start[index] + "\t\t" + plan_time_end[index]);
                            }
                            foreach (string id in mp.Keys)
                            {
                                Console.WriteLine("Общее время работы {0}: {1}. За это время обработано {2} партий.", machine_tools[id], machine_busy[id], mp[id]);
                            }
                            //вывод плана в excel
                            ObjBook = ObjExcel.Workbooks.Add();
                            ObjSheet = (Excel.Worksheet)ObjBook.Sheets[1];
                            ObjSheet.Cells[1, 1].Value = "Партия";
                            ObjSheet.Cells[1, 2].Value = "Оборудование";
                            ObjSheet.Cells[1, 3].Value = "Время начала";
                            ObjSheet.Cells[1, 4].Value = "Время окончания";
                            for (int index = 0; index < plan_parties.Count; ++index)
                            {
                                ObjSheet.Cells[index + 2, 1].Value = nomenclatures[plan_parties[index]];
                                ObjSheet.Cells[index + 2, 2].Value = machine_tools[plan_machine[index]];
                                ObjSheet.Cells[index + 2, 3].Value = plan_time_start[index];
                                ObjSheet.Cells[index + 2, 4].Value = plan_time_end[index];
                            }
                            ObjSheet.Cells[1, 6].Value = "Общее время работы";
                            int row1 = 2;
                            foreach (string id in machine_busy.Keys)
                            {
                                ObjSheet.Cells[row1, 6].Value = machine_tools[id];
                                ObjSheet.Cells[row1, 7].Value = machine_busy[id];
                                ++row1;
                            }
                            ObjSheet.Cells[1, 9].Value = "Обработано партий";
                            row1 = 2;
                            foreach (string id in mp.Keys)
                            {
                                ObjSheet.Cells[row1, 9].Value = machine_tools[id];
                                ObjSheet.Cells[row1, 10].Value = mp[id];
                                ++row1;
                            }
                            ObjBook.SaveAs(@"C:\Users\123\Desktop\Тестовое задание\Тестовое задание\Plan.xlsx");
                            ObjBook.Close(true);
                            Console.WriteLine("Нечего обрабатывать!");
                            Console.WriteLine("Нажмите кнопку для продолжения...");
                            Console.ReadKey();
                            ObjExcel.Quit();
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Неверный формат ввода команды, повторите ввод.");
                    }
                }
                else if (con[0] == "Далее")
                {
                    break;
                }
                else
                {
                    Console.WriteLine("Команда введена не верно, повторите ввод.");
                }
            }
            //планирование
            int Volume_Parties = 0;
            int time = 0;
            foreach(int val in parties.Values)
            {
                Volume_Parties += val;
            }
            List<string> plan_parties2 = new List<string>();
            List<string> plan_machine2 = new List<string>();
            List<int> plan_time_start2 = new List<int>();
            List<int> plan_time_end2 = new List<int>();
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
                                plan_parties2.Add(id_nomenclatures);
                                plan_machine2.Add(id_machine);
                                plan_time_start2.Add(time);
                                plan_time_end2.Add(machine_busy[id_machine]);
                                break;
                            }
                        }
                    }
                }
                time += 10;
            }
            //оптимизация
            if (parties.Count > 1)
            {
                Dictionary<string, List<string>> plan_machine_parties = new Dictionary<string, List<string>>();
                foreach (string id_machine in machine_tools.Keys)
                {
                    plan_machine_parties.Add(id_machine, new List<string>());
                }
                for (int index = 0; index < plan_parties2.Count; ++index)
                {
                    plan_machine_parties[plan_machine2[index]].Add(plan_parties2[index]);
                }
                plan_machine2.Clear();
                plan_parties2.Clear();
                plan_time_end2.Clear();
                plan_time_start2.Clear();
                log = true;
                Dictionary<string, Dictionary<string, int>> times2 = new Dictionary<string, Dictionary<string, int>>();
                foreach (string id_machine in times.Keys)
                {
                    foreach (string id_nomenclatures in times[id_machine].Keys)
                    {
                        if (times2.ContainsKey(id_nomenclatures))
                        {
                            times2[id_nomenclatures].Add(id_machine, times[id_machine][id_nomenclatures]);
                        }
                        else
                        {
                            times2.Add(id_nomenclatures, new Dictionary<string, int>());
                            times2[id_nomenclatures].Add(id_machine, times[id_machine][id_nomenclatures]);
                        }
                    }
                }
                times_keys.AddRange(times2.Keys);
                foreach (string id_nomenclatures in times_keys)
                {
                    times2[id_nomenclatures] = times2[id_nomenclatures].OrderBy(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);
                }
                times_keys.Clear();
                while (log)
                {
                    string max = machine_busy.Aggregate((l, r) => l.Value > r.Value ? l : r).Key;
                    int old_time = machine_busy[max];
                    log = false;
                    foreach(string id_nomenclatures in times[max].Keys.Reverse())
                    {
                        if (plan_machine_parties[max].Contains(id_nomenclatures))
                        {
                            foreach (string id_machine in times2[id_nomenclatures].Keys)
                            {
                                if (machine_busy[id_machine] + times2[id_nomenclatures][id_machine] < old_time)
                                {
                                    machine_busy[max] -= times2[id_nomenclatures][max];
                                    machine_busy[id_machine] += times2[id_nomenclatures][id_machine];
                                    plan_machine_parties[max].Remove(id_nomenclatures);
                                    plan_machine_parties[id_machine].Add(id_nomenclatures);
                                    log = true;
                                    break;
                                }
                            }
                        }
                        if (log)
                        {
                            break;
                        }
                    }
                }
                if (plan_machine.Count > 0)
                {
                    foreach (string id_machine in machine_tools.Keys)
                    {
                        machine_busy[id_machine] = plan_time_end[plan_machine.LastIndexOf(id_machine)];
                    }
                }
                else
                {
                    foreach (string id_machine in machine_tools.Keys)
                    {
                        machine_busy[id_machine] = 0;
                    }
                }
                foreach(string id in plan_machine_parties.Keys)
                {
                    mp[id] = plan_machine_parties[id].Count;
                }
                time = 0;
                while (plan_machine_parties.Count > 0)
                {
                    foreach(string id_machine in machine_tools.Keys)
                    {
                        if (machine_busy[id_machine] <= time && plan_machine_parties.ContainsKey(id_machine))
                        {
                            plan_machine.Add(id_machine);
                            plan_parties.Add(plan_machine_parties[id_machine][0]);
                            plan_machine_parties[id_machine].RemoveAt(0);
                            if (plan_machine_parties[id_machine].Count == 0)
                            {
                                plan_machine_parties.Remove(id_machine);
                            }
                            plan_time_start.Add(time);
                            machine_busy[id_machine] += times[id_machine][plan_parties.Last()];
                            plan_time_end.Add(machine_busy[id_machine]);
                        }
                    }
                    time += 10;
                }
            }
            else
            {
                plan_machine.AddRange(plan_machine2);
                plan_parties.AddRange(plan_parties2);
                plan_time_end.AddRange(plan_time_end2);
                plan_time_start.AddRange(plan_time_start2);
                for(int index=0; index<plan_machine.Count; ++index)
                {
                    ++mp[plan_machine[index]];
                }
            }
            //вывод плана в консоль
            Console.WriteLine("Партия\tОборудование\tВремя начала\tВремя окончания");
            for(int index=0; index<plan_parties.Count; ++index)
            {
                Console.WriteLine(nomenclatures[plan_parties[index]] + "\t" + machine_tools[plan_machine[index]] + "\t\t" + plan_time_start[index] + "\t\t" + plan_time_end[index]);
            }
            foreach(string id in mp.Keys)
            {
                Console.WriteLine("Общее время работы {0}: {1}. За это время обработано {2} партий.", machine_tools[id], machine_busy[id], mp[id]);
            }
            //вывод плана в excel
            ObjBook = ObjExcel.Workbooks.Add();
            ObjSheet = (Excel.Worksheet)ObjBook.Sheets[1];
            ObjSheet.Cells[1, 1].Value = "Партия";
            ObjSheet.Cells[1, 2].Value = "Оборудование";
            ObjSheet.Cells[1, 3].Value = "Время начала";
            ObjSheet.Cells[1, 4].Value = "Время окончания";
            for (int index = 0; index < plan_parties.Count; ++index)
            {
                ObjSheet.Cells[index + 2, 1].Value = nomenclatures[plan_parties[index]];
                ObjSheet.Cells[index + 2, 2].Value = machine_tools[plan_machine[index]];
                ObjSheet.Cells[index + 2, 3].Value = plan_time_start[index];
                ObjSheet.Cells[index + 2, 4].Value = plan_time_end[index];
            }
            ObjSheet.Cells[1, 6].Value = "Общее время работы";
            int row = 2;
            foreach (string id in machine_busy.Keys)
            {
                ObjSheet.Cells[row, 6].Value = machine_tools[id];
                ObjSheet.Cells[row, 7].Value = machine_busy[id];
                ++row;
            }
            ObjSheet.Cells[1, 9].Value = "Обработано партий";
            row = 2;
            foreach (string id in mp.Keys)
            {
                ObjSheet.Cells[row, 9].Value = machine_tools[id];
                ObjSheet.Cells[row, 10].Value = mp[id];
                ++row;
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
