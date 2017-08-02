using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System.Xml;

namespace PatternDoc
{
    class ClassLOAccess
    {
        string inputFile = "";
        string outputFile = "";
        string substitutionData = "";

        public void ClassLOAcc(string iFile, string oFile, string sData)
        {
            inputFile = iFile;
            outputFile = oFile;
            substitutionData = sData;
        }


        public void acalc()
        {
            FileWaitInit();
            // Получаем копию менеджера сервисов
            object usm = Activator.CreateInstance(Type.GetTypeFromProgID("com.sun.star.ServiceManager"));
            // создаем копию рабочего стола. Помните, что рабочий стол только  один.
            object desk = Invoke(usm, "createInstance", "com.sun.star.frame.Desktop");
            PrintImpName(desk);
            object oDoc = Invoke(desk, "loadComponentFromURL", PathConverter(inputFile), "_blank", 0, new object[0]);

            int nColumn = 0;
            int nRow = 0;

            object oSheets = Invoke(oDoc, "getSheets", new object[0]);
            object oSheet = Invoke(oSheets, "getByIndex", new object[1] { 0});
            object oCell = Invoke(oSheet, "getCellByPosition", new object[2] { nColumn, nRow });

            FileWatch(150);

            //Invoke(oDoc, "storeToURL", PathConverter(outputFile), new object[0]); // new unoidl.com.sun.star.beans.PropertyValue[0]);

            File.Copy(inputFile, outputFile);

            System.Threading.Thread.Sleep(5000);

            try
            {
                Invoke(oDoc, "Close", true);
            }
            catch (Exception e) { Console.Write(e.Message); }

        }

        string[] s;

        public void dwriter()
        {
            FileWaitInit();
            // Получает копию менеджера сервисов
            object usm = Activator.CreateInstance(Type.GetTypeFromProgID("com.sun.star.ServiceManager"));
            // создает копию рабочего стола. Помнить! что рабочий стол только  один.
            object desk = Invoke(usm, "createInstance", "com.sun.star.frame.Desktop");
            PrintImpName(desk);

            //List<Aproperty> Bloadprop = new List<Aproperty>();
            //object[] loadProps = new object[1];
            //Bloadprop.Add(new Aproperty( "Hidden", true));
            //loadProps[0] = new Aproperty("Hidden", true);
            // Загружает новый документ 
            object oDoc = Invoke(desk, "loadComponentFromURL", PathConverter(inputFile), "_blank", 0, new object[0]); // { "Hidden",false });
            object TextStr = Invoke(oDoc, "getText", new object[0]);
            object oCursor = Invoke(TextStr, "createTextCursor", new object[0]);
            // Что насчет получения текущего компонента? Этот пример не получает аргументов
            object x = Invoke(oDoc, "supportsService", "com.sun.star.text.TextDocument");
            PrintImpName(oDoc);

            if (!(x is bool && (bool)x))
                return;

            object oBookmarks = Invoke(oDoc, "getBookmarks", new object[0]);
            x = Invoke(oBookmarks, "getCount", new object[0]);

            int nCount = (int)x;

            object booknm;
            object oMark;
            object oRng;

            string[] bokms = new string[nCount];

            // составляем список закладок
            for (int nk = 0; nk < nCount; nk++)
            {
                oMark = Invoke(oBookmarks, "getByIndex", (object)nk);
                booknm = Invoke(oMark, "getName", new object[0]);
                bokms[nk] = (string)booknm;
            }

            List<Abookmarks> BBokmarks = new List<Abookmarks>();

            XmlTextReader reader = new XmlTextReader(substitutionData);
            string bookname = "";
            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element: // Узел является элементом.
                        bookname = reader.Name;
                        break;
                    case XmlNodeType.Text: // вставить в список как элемент.
                        BBokmarks.Add(new Abookmarks(bookname, reader.Value));
                        break;
                }
            }

            object sss;
            string str;
            string[] crStr;
            //int len = 0;
            string[] tbnm;
            object itsTable;
            object[] parmas;
            char[] separate;
            int addrows = 0;
            bool istable = false;
            ATables atass;
            int currentRow = 0;
            int currentColumn = 0;
            string[] pos;

            List<ATables> Tablist = new List<ATables>();
            object oAllTables =Invoke( oDoc, "getTextTables", new object[0]);

            foreach (Abookmarks abkmrk in BBokmarks)
            {
                // ищем в списке полей карточки документа соответствие закладке в шаблоне 
                for (int nl = 0; nl < nCount; nl++)
                {
                    if (bokms[nl].ToLower().Contains("table"))
                    {
                        Console.WriteLine(bokms[nl]);
                    }
                    tbnm = bokms[nl].Split('$');
                    str = bokms[nl].ToLower();
                    if (tbnm[0].ToLower().CompareTo(abkmrk.Name.ToLower()) == 0)
                    {
                        //находим соответствие
                        s = new string[1];
                        s[0] = bokms[nl];
                        oMark = Invoke(oBookmarks, "getByName", (object)s[0]);
                        PrintImpName(oMark);
                        s[0] = abkmrk.Value;
                        //если это значение карточки документа
                        if ((!bokms[nl].Contains("\n")) && (!bokms[nl].ToLower().Contains("$table")))
                        {
                            oRng = Invoke(oMark, "getAnchor", new object[0]);
                            TextStr = Invoke(oRng, "getStart", new object[0]);
                            sss = Invoke(TextStr, "SetString", (object)s[0]);
                        }
                        // иначе это таблица
                        else
                        {
                            List<Abookmarks> SubBokmarks = new List<Abookmarks>();
                            separate = new char[1];
                            separate[0] = '\n';
                            crStr = abkmrk.Value.Split(separate);
                            int inn = 0;
                            foreach (string sg in crStr)
                            {
                                SubBokmarks.Add(new Abookmarks(tbnm[0] + inn.ToString(), sg));
                                inn++;
                            }
                            s = new string[1];
                            s[0] = bokms[nl];
                            oMark = Invoke(oBookmarks, "getByName", (object)s[0]);
                            PrintImpName(oMark);
                            s[0] = SubBokmarks[0].Value;
                            oRng = Invoke(oMark, "getAnchor", new object[0]);
                            TextStr = Invoke(oRng, "getStart", new object[0]);
                            sss = Invoke(TextStr, "SetString", (object)s[0]);

                            parmas = new object[1];
                            parmas[0] = (object)tbnm[1];
                            itsTable = Invoke(oAllTables, "getByName", parmas);

                            //вставляем текст в первый ряд
                            object xRows = Invoke(itsTable, "GetRows", new object[0]);
                            object xRowscount = Invoke(xRows, "GetCount", new object[0]);
                            object xColumns = Invoke(itsTable, "GetColumns", new object[0]);

                            //проверяем, есть ли добавленные строки в таблице
                            istable = false;
                            atass = null;
                            addrows = 0;

                            if (Tablist.Count > 0)
                            {
                                foreach (ATables atas in Tablist)
                                {
                                    if (atas.Name.ToLower().CompareTo(tbnm[1].ToLower()) == 0)
                                    {
                                        istable = true;
                                        atass = atas;
                                        break;
                                    }
                                }
                            }

                            if (!istable)
                            {
                                Tablist.Add(new ATables(tbnm[1], SubBokmarks.Count));
                                //вставляем строки
                                addrows = SubBokmarks.Count - 1;
                            }
                            else
                            {
                                if (atass.Count < SubBokmarks.Count)
                                {
                                    //добавляем недостающие строки
                                    addrows = SubBokmarks.Count - atass.Count;
                                    atass.Count = SubBokmarks.Count;
                                }
                            }

                            parmas = new object[2];
                            if (addrows > 0 && xRowscount!=null)
                            {
                                parmas[0] = xRowscount;
                                parmas[1] = (object)(addrows);
                                Invoke(xRows, "insertByIndex", parmas);
                            }
                            //в певую позицию
                            s = new string[1];
                            s[0] = bokms[nl];

                            pos = tbnm[2].Split('.');
                            currentColumn =Convert.ToInt32( pos[1],10)-1;
                            currentRow = Convert.ToInt32(pos[0], 10)-1;

                            for (int k = 1; k < SubBokmarks.Count; k++)
                            {
                                // перебираем ячейки
                                parmas = new object[2];
                                parmas[1] = currentRow + k;
                                parmas[0] = currentColumn ;
                                object oCell = Invoke(itsTable, "getCellByPosition", parmas);
                                
                                parmas = new object[1];
                                parmas[0] = crStr[k];
                                TextStr = Invoke(oCell, "getText", new object[0]);
                                Invoke(TextStr, "SetString", parmas);
                                
                            }
                        }
                    }
                }

            }

            FileWatch(150);

            try
            {
                Invoke(oDoc, "Close", true);
            }
            catch (Exception e) { Console.Write(e.Message); }
 
            File.Copy(inputFile , outputFile);
        }

        private void FileWatch(int sleep)
        {
            hfile.EnableRaisingEvents = true;

            while (!fileREnamed && !fileChanged)
            {
                System.Threading.Thread.Sleep(sleep);
            }
        }

        FileSystemWatcher hfile;
            private void FileWaitInit()
        {
            hfile = new FileSystemWatcher();
            string[] cess = inputFile.Split('\\');
            hfile.Path = inputFile.Substring(0, inputFile.Length - cess[cess.Length - 1].Length);
            hfile.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
           | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            hfile.Filter = cess[cess.Length - 1];
            hfile.Changed += Hfile_Changed;
            hfile.Disposed += Hfile_Disposed;
            hfile.Renamed += Hfile_Renamed;
        }

        

        string newFileName = "";
        string newFilepath = "";

        private void Hfile_Renamed(object sender, RenamedEventArgs e)
        {
            fileREnamed=true;
            newFileName = e.Name;
            newFilepath = e.FullPath;
            inputFile = newFilepath + newFileName;
        }

        private bool fileChanged = false;
        //private bool fileDisposed = false;
        private bool fileREnamed = false;

        private void Hfile_Disposed(object sender, EventArgs e)
        {
            //fileDisposed=true;
        }

        private void Hfile_Changed(object sender, FileSystemEventArgs e)
        {
            fileChanged=true;
        }

        private class Abookmarks
        {
            public string Name;
            public string Value;

            public Abookmarks(string aname, string avalue)
            {
                this.Name = aname;
                this.Value = avalue;
            }
            public override string ToString()
            {
                return this.Name;
            }
        }

        private class ATables
        {
            public string Name;
            private int Rowcount;

            public ATables(string aname, int avalue)
            {
                this.Name = aname;
                this.Rowcount = avalue;
            }
            public override string ToString()
            {
                return this.Name;
            }

            public int Count
            {
                get { return Rowcount; }
                set { Rowcount = value; }
            }
        }

        private string PathConverter(string file)
        {
            try
            {
                file = file.Replace(@"\", "/");

                return "file:///" + file;
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }

        object Invoke(object obj, string method, params object[] par)
        {
          return
            obj.GetType().InvokeMember(method, BindingFlags.InvokeMethod, null, obj, par);
        }

        void PrintImpName(object obj)
        {
            object x = Invoke(obj, "getImplementationName", new object[0]);
            System.Console.WriteLine(x.ToString());
        }
    }
}
