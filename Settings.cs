using System;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MakeCall
{
    public static class Settings
    {
     
        /// <summary>
        /// Папка  входящих задач
        /// </summary>
        public static string InDir
        {
            get
            {
                string s = ConfigurationManager.AppSettings["InDir"];
                if (s == "" || !Directory.Exists(s))
                    throw new ConfigurationErrorsException("Укажите корректную папку входящих задач InDir в файле MakeCall.exe.Config");
                return Path.GetFullPath(s);
            }
        }

        /// <summary>
        /// Папка  обработанных задач
        /// </summary>
        public static string OutDir
        {
            get
            {
                string s = ConfigurationManager.AppSettings["OutDir"];
                if (s == "" || !Directory.Exists(s))
                    throw new ConfigurationErrorsException("Укажите корректную папку обработанных задач OutDir в файле MakeCall.exe.Config");
                return Path.GetFullPath(s);
            }
        }

        /// <summary>
        /// Интервал, с которым программа проверяет наличие новых файлов
        /// </summary>
        public static int Interval
        {
            get
            {
                string s = ConfigurationManager.AppSettings["Interval"];
                int res = 1000;
                if (s != "") 
                    int.TryParse(s, out res);
                return res;
            }
        }

    }
}
