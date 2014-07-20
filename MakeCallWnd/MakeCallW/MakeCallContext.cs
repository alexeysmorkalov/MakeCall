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
using System.Windows.Forms;
using Timer = System.Threading.Timer;

namespace MakeCallW
{
    internal class MakeCallContext: ApplicationContext
    {
        #region ** fields
        NotifyIcon _notifyIcon;
        MenuItem _exitMenuItem;
        bool _tapiInitialized = false;
        Timer _watchTimer = null;
        string _lastError = "";
        string _lastUnsuccessCall = "";
        string _lastSuccessCall = "";
        bool _closeApplication = false;
        DateTime _startDateTime;

        #endregion

        public MakeCallContext()
        {

            SetupNotifyIcon();

            Application.ApplicationExit += Application_ApplicationExit;

            Run();
        }
        void Run()
        {
            _startDateTime = DateTime.Now;
            if (CanStart() && TapiManager.Initialize())
            {
                _tapiInitialized = true;
                _watchTimer = new Timer(WatchDir, null, 0, Settings.Interval);
            }
        }

        bool CanStart()
        {
            if (Settings.InDir == "")
            {
                _lastError = "Папка для входящих указана неверно";
                return false;
            }
            if (Settings.OutDir == "")
            {
                _lastError = "Папка для исходящих указана неверно";
                return false;
            }
            return true;
        }
        void WatchDir(object state)
        {
            if (_closeApplication)
            {
                Application.Exit();
                return;
            }
            var files = Directory.EnumerateFiles(Settings.InDir);
            foreach (var file in files)
            {
                ProcessFile(Path.GetFullPath(file));
            }
        }

        void ProcessFile(string fileName)
        {
            try
            {
                var taskStrings = File.ReadAllLines(fileName);
                if (taskStrings.Length == 1 && taskStrings[0] == "0")
                {
                    _closeApplication = true;
                    File.Delete(fileName);
                    return;
                }
                //var tempFileName = Path.GetTempPath() + Path.DirectorySeparatorChar + Path.GetFileName(fileName);
                var outFileNameSuccess = Path.GetFullPath(Settings.OutDir) + Path.DirectorySeparatorChar + Path.GetFileName(fileName);
                var outFileNameFail = Path.ChangeExtension(outFileNameSuccess, "err");

                if (File.Exists(outFileNameFail))
                    File.Delete(outFileNameFail);
                File.Move(fileName, outFileNameFail);


                // В файле должно быть 2 строки
                if (taskStrings.Length == 2)
                {
                    var fromNo = taskStrings[0];
                    var toNo = taskStrings[1];


                    if (TapiManager.Lines.Any(l => l.Addresses.Length > 0 && l.Addresses[0].Address == fromNo))
                    { 
                        TapiLine tapiLine = TapiManager.Lines.First(l => l.Addresses.Length > 0 && l.Addresses[0].Address == fromNo);
                        if (tapiLine != null)
                            // Распараллеливаем задачу
                            Task.Factory.StartNew(() => ProcessCall(tapiLine, fromNo, toNo, outFileNameSuccess, outFileNameFail));
                    }
                    else
                        _lastUnsuccessCall =  "Не найдена линия: " + fromNo + ">" + toNo;
                }
                else
                    _lastError = "Неверный формат файла: " + Path.GetFileName(fileName);

            }
            catch (Exception e)
            {
                _lastError = "Ошибка обработки файла " + Path.GetFileName(fileName) + " " + e.Message;
            }

        }

        void ProcessCall(TapiLine tapiLine, string fromNo, string toNo, string outFileNameSuccess, string outFileNameFail)
        {
            if (!tapiLine.IsOpen)
                tapiLine.Open(MediaModes.InteractiveVoice);
            try
            {
                try
                {
                    TapiCall call = tapiLine.MakeCall(toNo);
                    _lastSuccessCall = fromNo + ">" +  toNo;
                    if (File.Exists(outFileNameSuccess))
                        File.Delete(outFileNameSuccess);
                    File.Move(outFileNameFail, outFileNameSuccess);
                }
                catch (Exception e)
                {
                    _lastUnsuccessCall = fromNo + ">" + toNo + " " + e.Message; 
                }
            }
            finally
            {
                tapiLine.Close();
            }
        }

        void SetupNotifyIcon()
        {
            _notifyIcon = new NotifyIcon();
            _notifyIcon.Icon = MakeCallW.Properties.Resources.CallIco;

            _notifyIcon.BalloonTipTitle = "Статус MakeCall:";
            _notifyIcon.BalloonTipText = "Текст";
            _notifyIcon.BalloonTipIcon = ToolTipIcon.Info;

            _notifyIcon.MouseClick += _notifyIcon_MouseClick;

            _exitMenuItem = new MenuItem("Выход", new EventHandler(Exit));
            _exitMenuItem.DefaultItem = true;
            _notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { _exitMenuItem });

            _notifyIcon.Visible = true;
        }

        void Application_ApplicationExit(object sender, EventArgs e)
        {
            if (_tapiInitialized)
            {
                _watchTimer.Dispose();
                _watchTimer = null;
                
                TapiManager.Shutdown();
            }
            _notifyIcon.Visible = false;
        }

        void _notifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            SetupBaloon();
            _notifyIcon.ShowBalloonTip(10000);
        }

        void _notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            SetupBaloon();
            _notifyIcon.ShowBalloonTip(10000);
        }
        void Exit(object sender, EventArgs e)
        {

            Application.Exit();
        }

        void SetupBaloon()
        {
            _notifyIcon.BalloonTipTitle = "MakeCall запущен " + _startDateTime.ToString();
            StringBuilder sb = new StringBuilder();
            if (_tapiInitialized)
                sb.AppendLine("TAPI OK Линий: " + TapiManager.Lines.Length.ToString());
            else
            {
                sb.AppendLine("TAPI не запущен");
            }

            if (_lastSuccessCall != "")
                sb.AppendLine("Последний успешный звонок: " + _lastSuccessCall);

            if (_lastUnsuccessCall != "")
                sb.AppendLine(_lastUnsuccessCall);

            if (_lastError != "")
                sb.AppendLine("Последняя ошибка: " + _lastError);

            _notifyIcon.BalloonTipText = sb.ToString();

            if (!_tapiInitialized)
                _notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
            else if (_lastError != "" || _lastUnsuccessCall != "")
                _notifyIcon.BalloonTipIcon = ToolTipIcon.Warning;
            else
                _notifyIcon.BalloonTipIcon = ToolTipIcon.Info;

        }
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

    }
}
