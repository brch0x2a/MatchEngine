using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MatchEngineV00
{
    class Matcher
    {
        public string ParosBook { get; set; }
        public string AreaBook { get; internal set; }

        public string Path { get; set; }

        private Mapeo mapeo;
        private List<Paro> paros;

        public Matcher() {
            mapeo = new Mapeo();
            paros = new List<Paro>();
        }


        public void  ReadParos() {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ParosBook);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            int[] parosIndex = { 1, 2, 6, 8, 12 };
            int parosIndexCount = parosIndex.Count();
          

            string linea = "";
            int dia = 0;
            int turno = 0;
            int codigo = 0;
            int minutos = 0;
            int hora = 0;

            string area = "S";//nomenclatura salsitas

            DateTime dateTime;




            Console.WriteLine("Read Paros...");
            for (int i = 2; i <= rowCount; i++)
            {
                if (xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 12].Value2 != null) {
                    if (xlRange.Cells[i, 1].Value2.ToString().Contains(area)) {
                        linea = xlRange.Cells[i, 1].Value2.ToString();//linea
                        codigo = Convert.ToInt32(xlRange.Cells[i, 2].Value2);//codigo

                        dateTime = DateTime.ParseExact(xlRange.Cells[i, 6].Value2.ToString(), "dd-MM-yyyy",
                                               System.Globalization.CultureInfo.InvariantCulture);

                        dia = (int)dateTime.DayOfWeek;//dia
                        turno = (int)xlRange.Cells[i, 8].Value2;//turno
                        minutos = (int)xlRange.Cells[i, 12].Value2;//minuto

                        int.TryParse(xlRange.Cells[i, 10].Value2.ToString().Split(':')[0], out hora);

                        if (turno == 3 && hora < 6) dia--;

                        paros.Add(new Paro { linea = linea, codigo = codigo, dia = dia, turno = turno, minutos = minutos });
                    }
                }
            }
            Console.WriteLine("Done!");



//            foreach (var item in paros){  Console.WriteLine(item.ToString());     }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }


        public void ReadArea() {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(AreaBook);



            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["S2"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = 480;
            int colCount = xlRange.Columns.Count;

            int data = 0; 

            List<Paro> paros = new List<Paro>();
           

            Console.WriteLine("Read Area...");

            for (int i = 19; i <= rowCount; i++)
            {
               if (xlRange.Cells[i, 5].Value2 != null) {
                    data = (int)xlRange.Cells[i, 5].Value2;

                    mapeo.list.Add(i);
                    //Console.WriteLine("[" + i + "]\t" + data);
                }

            }
            Console.WriteLine("Done!");

            //foreach (var item in mapeo.list){   Console.WriteLine("["+a+"]"+item.ToString());   a++;  }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

        }

        public void PreMatch() {
            int row, col;
            int calc = 0;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(AreaBook);

            List<Excel._Worksheet> hojas = new List<Excel._Worksheet>();

            for (int j = 1; j < 10; j++)
            {
                hojas.Add(xlWorkbook.Sheets["S" + j.ToString()]);

                //Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["S2"];
                string test = "S" + j.ToString();
                int aux = 0;
                int.TryParse(test.Split('S')[1], out aux);
                Excel.Range xlRange = hojas[aux - 1].UsedRange;

                Console.WriteLine("Prematch...");
                for (int i = 0; i < paros.Count; i++)
                {
                    if (paros[i].linea.Equals("S"+j.ToString()))
                    {
                        Console.Write(paros[i].ToString());
                        row = mapeo.list[paros[i].codigo - 1];
                        col = 5 + (paros[i].dia - 1) * 3 + paros[i].turno;

                        Console.Write("\t|__|\t[" + row + "][" + col + "]-->>" + paros[i].minutos + "\n");

                        if (xlRange.Cells[row, col].Value2 != null)
                        {
                            int.TryParse(xlRange.Cells[row, col].Value2.ToString(), out calc);
                            xlRange.Cells[row, col].Value2 = calc + paros[i].minutos;
                        }

                        xlRange.Cells[row, col].Value2 = paros[i].minutos;
                    }
                }
                Marshal.ReleaseComObject(xlRange);
            }
            Console.WriteLine("Done!");
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
           // Marshal.ReleaseComObject(xlRange);
            for (int h = 0; h < 8; h++)
            {
                Marshal.ReleaseComObject(hojas[h]);
            }
           

            xlApp.Visible = false;
            xlApp.UserControl = false;
            xlWorkbook.SaveAs(Path + "test505" + ".xlsm");


            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }



    }
}
