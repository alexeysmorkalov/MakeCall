using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Serialization;
using JulMar.Atapi;

namespace MakeCall
{
    public class MakeCallApp
    {
        #region ctor
        public MakeCallApp()
        {
        }
        #endregion

        #region methods
        // собственно это и есть основная программа        
        public void Run()
        {
            Console.Clear();
            try
            {
                Log(Settings.Interval.ToString() + "ms", Settings.InDir, Settings.OutDir);

                // Инициализируем TAPI
                Log("TAPI инициализируется...");
                if (TapiManager.Initialize())
                {

                    Log("TAPI OK");
                    LogLines();
                    // Главный цикл
                    ConsoleKeyInfo cki;
                    // функция WatchDir срабатывает каждую секунду
                    using (new Timer(WatchDir, null, 0, Settings.Interval))
                    {
                        do
                        {
                            cki = Console.ReadKey();
                        } while (cki.Key != ConsoleKey.Escape);
                    }
                    Log("TAPI Shutdown...");
                    TapiManager.Shutdown();
                    Log("Завершение работы");
                }
                else
                    Log("Ошибка инициализации TAPI");
            }
            catch (Exception e)
            {
                Log(e.Message);
            }
        }

        // обрабатываем новые файлы в каталоге, если таковые имеются
        void WatchDir(object state)
        {
            var files = Directory.EnumerateFiles(Settings.InDir);
            foreach (var file in files)
            {
                ProcessFile(Path.GetFullPath(file));
            }
        }

        void ProcessCall(TapiLine tapiLine, string toNo, string outFileNameSuccess, string outFileNameFail)
        {
            if (!tapiLine.IsOpen)
                tapiLine.Open(MediaModes.InteractiveVoice);
            try
            {
                try
                {
                    TapiCall call = tapiLine.MakeCall(toNo);
                    Log("OK  ->", toNo);
                    if (!File.Exists(outFileNameSuccess))
                        File.Move(outFileNameFail, outFileNameSuccess);
                }
                catch
                {
                    Log("Fail  ->", toNo);
                }
            }
            finally
            {
                tapiLine.Close();
            }
        }

        void LogLines()
        {
            var lines = TapiManager.Lines;
            var availLines = new StringBuilder(" Available TAPI lines: \n");

            foreach (var line in lines)
            {
                availLines.AppendFormat("ID:{0} ", line.Id);
                availLines.AppendFormat("Name:{0} ", line.Name);
                availLines.Append("Address: ");
                foreach (var addr in line.Addresses)
                {
                    availLines.AppendFormat("{0} CanAnswer:{1} CanDial:{2} ", addr.Address, addr.Capabilities.CallFeatures.CanAnswer, addr.Capabilities.CallFeatures.CanDial);
                }

                availLines.AppendFormat("IsOpen:{0} ", line.IsOpen);
                availLines.AppendFormat("IsValid:{0} \n", line.IsValid);
            }
            File.WriteAllText("MakeCall.Lines.log", availLines.ToString());

            // Short to screen
            availLines = new StringBuilder();

            foreach (var line in lines)
            {
                foreach (var addr in line.Addresses)
                {
                    if (addr.Address.Length > 0 && addr.Capabilities.CallFeatures.CanDial)
                        availLines.AppendFormat("{0},", addr.Address);
                }
            }
            if (availLines.Length > 0)
                Log(" Available TAPI addresses: " + availLines.ToString());
            else
                Log(" No available TAPI addresses");
        }

        void ProcessFile(string fileName)
        {
            try
            {
                //var tempFileName = Path.GetTempPath() + Path.DirectorySeparatorChar + Path.GetFileName(fileName);
                var outFileNameSuccess = Path.GetFullPath(Settings.OutDir) + Path.DirectorySeparatorChar + Path.GetFileName(fileName);
                var outFileNameFail = Path.ChangeExtension(outFileNameSuccess, "err");

                File.Move(fileName, outFileNameFail);
                var taskStrings = File.ReadAllLines(outFileNameFail);
                // В файле должно быть 2 строки
                if (taskStrings.Length == 2)
                {
                    var fromNo = taskStrings[0];
                    var toNo = taskStrings[1];
                    Log(Path.GetFileName(fileName), fromNo, toNo);

                    TapiLine tapiLine = TapiManager.Lines.First(l => l.Addresses.Length > 0 && l.Addresses[0].Address == fromNo);
                    if (tapiLine != null)
                        // Распараллеливаем задачу
                        Task.Factory.StartNew(() => ProcessCall(tapiLine, toNo, outFileNameSuccess, outFileNameFail));
                    else
                        Log("Не могу найти линию", fromNo);
                }
                else
                    Log("Неверный формат файла", Path.GetFileName(fileName));

            }
            catch (Exception e)
            {
                Log("Ошибка обработки файла " + fileName, e.Message);
            }

        }
        #endregion

        #region TapiManager
        static TapiManager _tapiManager = null;
        public TapiManager TapiManager
        {
            get
            {
                if (_tapiManager == null)
                    _tapiManager = new TapiManager("MakeCallApp", TapiVersion.V20);
                return _tapiManager;
            }
        }
        #endregion

        #region Log
        static void Log()
        {
            Console.WriteLine(DateTime.Now.ToShortTimeString());
        }

        static void Log(string info)
        {
            Console.WriteLine(DateTime.Now.ToShortTimeString() + " " + info);
        }

        static void Log(string info1, string info2)
        {
            Console.WriteLine(DateTime.Now.ToShortTimeString() + " " + info1 + " , " + info2);
        }

        static void Log(string info1, string info2, string info3)
        {
            Console.WriteLine(DateTime.Now.ToShortTimeString() + " " + info1 + ", " + info2 + " => " + info3);
        }
        #endregion

    }
}
