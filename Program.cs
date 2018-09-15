using System;
using Excel = Microsoft.Office.Interop.Excel;
namespace optzadacha3v2
{
    class Program
    {
        public class Account
        {
            public int i = 0;
            public double A { get; set; }
            public double B { get; set; }
            public double E { get; set; }
            public double X1 { get; set; }
            public double X2 { get; set; }
            public double F1 { get; set; }
            public double F2 { get; set; }

        }
        public double F(double lam, double x, double y, double z, double dx, double dy, double dz)
        {
            //изменить уравнение
            double x1 = x + dx * lam;
            double y1 = y - dy * lam;
            double z1 = z - dz * lam;

            double fr1 = Math.Pow(x1, 2) + 5 * Math.Pow(y1, 2) - 4 * x1 * y1 - 8 * y1 + z1 + 4 / z1 + 1;
            return (fr1);
        }
        public double x1b(double a, double b)
        {
            double fr1 = b - (b - a) / 1.618;
            return fr1;
        }

        public double x2b(double a, double b)
        {
            double fr1 = a + (b - a) / 1.618;
            return fr1;
        }
        public double F1(double x, double y, double z)
        {
            double fr1 = Math.Pow(x, 2) + 5 * Math.Pow(y, 2) - 4 * x * y -8*y+ z + 4 / z + 1;
            return (fr1);
        }
        public double kas(out double dx, out double dy, double x, double y, double z)
        {
            dx = (-1) * (2 * x - 4 * y);
            dy = (-1) * (10 * y - 4 * x - 8);
            double dz = (-1)*(1 - 4 / (Math.Pow(z, 2)));
            return (dz);
        }
        public int spusk(double x, double y, double z, double dx, double dy, double dz, int shet, out int sc)
        {
            Program program = new Program();
            double[] array = new double[1000];
            double[,] xarray = new double[1000, 3];
            
            int i = 0;
            double lam = 0.02;
            xarray[0, 0] = -1;
            xarray[0, 1] = -2;
            xarray[0, 2] = 0.5;
            //вычислим начальное значение
            array[0] = Math.Round((program.F1(xarray[0, 0], xarray[0, 1], xarray[0, 2])), 4);
          //  Console.WriteLine(array[0] + "fasgasdg");
            //сделаем первый шаг, для этого вычислим ху
            x = Math.Round((x + lam *i* dx), 4);
            y = Math.Round((y + lam * i * dy), 4);
            z = Math.Round((z + lam * i * dz), 4);
          //  Console.WriteLine(i);
         //   Console.WriteLine("спуск");
          //  Console.WriteLine(0.2 * i+"lam");
         //   Console.WriteLine(array[i]+"f");
            i++;
            shet++;
            //вычислим значение функции
         //   Console.WriteLine(i);
         //   Console.WriteLine("спк");
         //   Console.WriteLine(array[i-1]);
        //    Console.WriteLine(array[1]);
            array[1] = Math.Round((program.F1(x, y, z)), 4);
          //  Console.WriteLine(array[i]);
            while (array[i - 1] > array[i])
            {
                shet++;
                i++;
                xarray[i, 0] = Math.Round((xarray[i - 1, 0] + lam * i * dx), 4);
          //      Console.WriteLine("//");
        //        Console.WriteLine(xarray[i, 0]);
                xarray[i, 1] = Math.Round((xarray[i - 1, 1] + lam * i * dy), 4);
         //      Console.WriteLine(xarray[i, 1]);
                xarray[i, 2] = Math.Round((xarray[i - 1, 2] + lam * i * dz), 4);
        //        Console.WriteLine(xarray[i, 2]);
         //       Console.WriteLine("//");
                array[i] = Math.Round((program.F1(xarray[i, 0], xarray[i, 1], xarray[i, 2])), 4);
                //       Console.WriteLine(array[i]);
                Console.WriteLine(0.2 * i+"lam");
                Console.WriteLine(array[i]+"f");
            }
         //   Console.WriteLine("check");
         //   Console.WriteLine("спускв");
            sc = shet;
            Console.WriteLine("exit");
            return (i);
        }
        public double gold(int b, double lam, double x, double y, double z, double dx, double dy, double dz, int shet, out int sc)
        {
            Account account = new Account();
            Program program = new Program();
            int j = 0;//потом не забыть вернуть это счетчик
            account.E =  0.1;
            account.A = (b - 1) * lam;
            account.B = b * lam;
            j++;
            account.F1 = Math.Round(program.F(account.A, x, y, z, dx, dy, dz), 4);
            account.F2 = Math.Round(program.F(account.B, x, y, z, dx, dy, dz), 4);
            j++;
            account.X1 = Math.Round(program.x1b(account.A, account.B), 4); //левая точка
            account.X2 = Math.Round(program.x2b(account.A, account.B), 4); //правая точка
            account.F1 = Math.Round(program.F(account.X1, x, y, z, dx, dy, dz), 4);
            account.F2 = Math.Round(program.F(account.X2, x, y, z, dx, dy, dz), 4);
            account.i++; account.i++;
            j++;
            j++;
            shet++;
            account.i++;
            if (account.F1 >= account.F2)
            {
                account.A = account.X1;
            }
            else
            {
                account.B = account.X2;
            }
            while (account.B - account.A > account.E)
            {
                shet++;
                j++;
                account.i++;
                if (account.F1 <= account.F2)
                {
                    account.F2 = Math.Round(account.F1, 4);
                    account.X2 = Math.Round(account.X1, 4);
                    account.X1 = Math.Round(program.x1b(account.A, account.B), 4);
                    account.F1 = Math.Round(program.F(account.X1, x, y, z, dx, dy, dz), 4);
                }
                else
                {
                    account.F1 = Math.Round(account.F2, 4);
                    account.X1 = Math.Round(account.X2, 4);
                    account.X2 = Math.Round(program.x2b(account.A, account.B), 4);
                    account.F2 = Math.Round(program.F(account.X2, x, y, z, dx, dy, dz), 4);
                }
                if (account.F1 >= account.F2)
                {
                    account.A = account.X1;
                }
                else
                {
                    account.B = account.X2;
                }
            }
            account.X1 = (account.A + account.B) / 2;
            sc = shet;
            return (account.X1);
        }
        static void Main(string[] args)
        {
            Program program = new Program();
            var excelapp = new Excel.Application
            {
                Visible = true
            };
            excelapp.Workbooks.Open("C:\\Users\\ak647\\Desktop\\New folder (2)\\Excel.xlsx"); // Задаем начальное положение файла Exel, куда будут записываться результаты
            Excel.Worksheet workSheet = (Excel.Worksheet)excelapp.ActiveSheet;
            double e = 0.1;//начальный этап
            int i, k, j, n, g, shet;
            shet = 1;
            g = 2;
            n = 3;
            i = 1;
            k = 1;
            j = k;
            double[,] x = new double[10000, 3];
            double[,] y = new double[5, 3];
            double[,] z = new double[4, 3];
            double[,] d = new double[1, 3];
            x[i, 0] = -1;
            x[i, 1] = -2;
            x[i, 2] = 0.5;
            workSheet.Cells[g, "A"] = shet;
            workSheet.Cells[g, "B"] = x[i, 0];
            workSheet.Cells[g, "C"] = x[i, 1];
            workSheet.Cells[g, "D"] = x[i, 2];
            workSheet.Cells[g, "E"] = program.F1(x[i, 0], x[i, 1], x[i, 2]);
            shet++;
            g++;
            y[i, 0] = x[i, 0];
            y[i, 1] = x[i, 1];
            y[i, 2] = x[i, 2];//основной этап, шаг 1
            double lambda = 0.00000000001;
            while (true)
            {
                d[0, 2] = program.kas(out d[0, 0], out d[0, 1], y[i, 0], y[i, 1], y[i, 2]);
                Console.WriteLine(d[0, 0] + "dx");
                Console.WriteLine(d[0, 1] + "dy");
                Console.WriteLine(d[0, 2] + "dz");
                int b = program.spusk(x[i, 0], x[i, 1], x[i, 2], d[0, 0], d[0, 1], d[0, 2], shet, out shet);
                double lam = program.gold(b, lambda, x[i, 0], x[i, 1], x[i, 2], d[0, 0], d[0, 1], d[0, 2], shet, out shet);
                y[1, 0] = Math.Round(x[k, 0] + lam * d[0, 0], 4);
                Console.WriteLine("y1" + y[1, 0] + " x");
                y[1, 1] = Math.Round(x[k, 1] + lam * d[0, 1], 4);
                Console.WriteLine("y1" + y[1, 1] + " y");
                y[1, 2] = Math.Round(x[k, 2] + lam * d[0, 2], 4);//шаг 2
                Console.WriteLine("y1" + y[1, 2] + " z");
                //lambda = lambda - 0.006;
                for (int c = 1; c < 4; c++)
                {
                    d[0, 2] = program.kas(out d[0, 0], out d[0, 1], y[j, 0], y[j, 1], y[j, 2]);
                    Console.WriteLine(d[0, 0] + "dx");
                    Console.WriteLine(d[0, 1] + "dy");
                    Console.WriteLine(d[0, 2] + "dz");
                    b = program.spusk(y[j, 0], y[j, 1], y[j, 2], d[0, 0], d[0, 1], d[0, 2], shet, out shet);
                    Console.WriteLine(b + "колв");
                    lam = program.gold(b, lambda, y[j, 0], y[j, 1], y[j, 2], d[0, 0], d[0, 1], d[0, 2], shet, out shet);
                    Console.WriteLine(lam + "лямбда");
                    z[j, 0] = Math.Round(y[j, 0] + lam * d[0, 0], 4);
                    Console.WriteLine("z" + j + " " + y[j, 0] + " x");
                    z[j, 1] = Math.Round(y[j, 1] + lam * d[0, 1], 4);
                    Console.WriteLine("z" + j + " " + y[j, 1] + " y");
                    z[j, 2] = Math.Round(y[j, 2] + lam * d[0, 2], 4);//шаг 3
                    Console.WriteLine("z" + j + " " + y[j, 2] + " z");
                    d[0, 0] = Math.Round(z[j, 0] - y[j - 1, 0], 4);
                    d[0, 1] = Math.Round(z[j, 1] - y[j - 1, 1], 4);
                    d[0, 2] = Math.Round(z[j, 2] - y[j - 1, 2], 4);
                    Console.WriteLine(d[0, 0] + "dx");
                    Console.WriteLine(d[0, 1] + "dy");
                    Console.WriteLine(d[0, 2] + "dz");
                    b = program.spusk(z[j, 0], z[j, 1], z[j, 2], d[0, 0], d[0, 1], d[0, 2], shet, out shet);
                    lam = program.gold(b, 0.1, z[j, 0], z[j, 1], z[j, 2], d[0, 0], d[0, 1], d[0, 2], shet, out shet);
                    y[j + 1, 0] = Math.Round(z[j, 0] + lam * d[0, 0], 4);
                    y[j + 1, 1] = Math.Round(z[j, 1] + lam * d[0, 1], 4);
                    y[j + 1, 2] = Math.Round(z[j, 2] + lam * d[0, 2], 4);
                    j++;
                }//шаг 4
                x[k + 1, 0] = y[n + 1, 0];
                x[k + 1, 1] = y[n + 1, 1];
                x[k + 1, 2] = y[n + 1, 2];
                workSheet.Cells[g, "A"] = shet;
                workSheet.Cells[g, "B"] = x[k + 1, 0];
                workSheet.Cells[g, "C"] = x[k + 1, 1];
                workSheet.Cells[g, "D"] = x[k + 1, 2];
                workSheet.Cells[g, "E"] = program.F1(x[k + 1, 0], x[k + 1, 1], x[k + 1, 2]);
                g++;
                if (Math.Abs((x[k + 1, 0] + x[k + 1, 1] + x[k + 1, 2]) - (x[k, 0] + x[k, 1] + x[k, 2])) < e)
                {
                    break;
                }
                else
                {
                    k++;
                    j = 1;
                }
            }
            Console.ReadKey();
        }
    }
}