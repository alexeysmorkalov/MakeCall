using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MakeCallW
{
    internal class MakeCallContext: ApplicationContext
    {
        #region ** fields
        NotifyIcon _notifyIcon;
        MenuItem _exitMenuItem;

        #endregion

        public MakeCallContext()
        {
            _notifyIcon = new NotifyIcon();
            _notifyIcon.Icon = MakeCallW.Properties.Resources.CallIco;
            
            _notifyIcon.BalloonTipTitle = "Статус MakeCall:";
            _notifyIcon.BalloonTipText = "Текст";
            _notifyIcon.BalloonTipIcon = ToolTipIcon.Info; 

            _notifyIcon.MouseClick += _notifyIcon_MouseClick;

            _exitMenuItem = new MenuItem("Выход", new EventHandler(Exit));
            _exitMenuItem.DefaultItem = true;
            _notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { _exitMenuItem});

            _notifyIcon.Visible = true;

            Application.ApplicationExit += Application_ApplicationExit;

            Run();
        }
        void Run()
        {
        }

        void Application_ApplicationExit(object sender, EventArgs e)
        {
            _notifyIcon.Visible = false;
        }

        void _notifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            _notifyIcon.ShowBalloonTip(6000);
        }

        void _notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            _notifyIcon.ShowBalloonTip(6000);
        }
        void Exit(object sender, EventArgs e)
        {

            Application.Exit();
        }
    }
}
