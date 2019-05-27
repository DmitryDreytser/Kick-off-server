//#define DEBUG

using BigFileRead;
using Compound;
using HTTPServer.Properties;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Mail;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using UserDef;
using Utilities.Exchange;
//using Microsoft.Office.Interop.Excel;

namespace BigFileRead
{
    class LogParser
    {

        public static string DecTo36(int num)
        {
            string alf = "0123456789ABСDEFGHIJKLMNOPQRSTUVWXYZ";
            int sys = 36;
            string result = null;
            List<int> cifr = new List<int>();
            while (num != 0)
            {
                cifr.Add(num % sys);
                num /= sys;
            }
            for (int i = cifr.Count - 1; i > -1; i--) result += alf[cifr[i]];
            return result;
        }

        public static Dictionary<string, string> logevents = new Dictionary<string, string>() {
            //Справочники
            {"RefOpen","Открыт"},
            {"RefWrite","Записан"},
            {"RefNew","Создан"},
            {"RefMarkDel","Помечен на удаление"},
            {"RefUnmarkDel","Снята пометка удаления"},
            {"RefDel","Удален"},
            {"RefGrpMove","Перенесен в другую группу"},
            {"RefAttrWrite","Значение реквизита записано"},
            {"RefAttrDel","Значение реквизита удалено"},
            //Документы
            {"DocNew","Cоздан"},
            {"DocOpen","Открыт"},
            {"DocWrite","Записан"},
            {"DocWriteNew","Записан новый"},
            {"DocNotWrite","Не записан"},
            {"DocPassed","Проведен"},
            {"DocBackPassed","Проведен задним числом"},
            {"DocNotPassed","Не проведен"},
            {"DocMakeNotPassed","Сделан не проведенным"},
            {"DocWriteAndPassed","Записан и проведен"},
            {"DocWriteAndRepassed","Записан и проведен задним числом"},
            {"DocWriteAndPostBfAP","Записан и проведен задним числом"},
            {"DocTimeChanged","Изменено время"},
            {"DocOperOn","Проводки включены"},
            {"DocOperOff","Проводки выключены"},
            {"DocMarkDel","Помечен на удаление"},
            {"DocUnmarkDel","Снята пометка удаления"},
            {"DocDel","Удален"},
            //События обмена
            {"DistBatchErr","Ошибка автообмена в пакетном режиме"},
            {"DistDnldBeg","Начата выгрузка изменений данных"},
            {"DistDnldSuc","Выгрузка изменений данных успешно завершена"},
            {"DistDnldFail","Выгрузка изменений данных не выполнена"},
            {"DistDnlErr","Ошибка выгрузки изменений данных"},
            {"DistUplBeg","Начата загрузка изменений данных"},
            {"DistUplSuc","Загрузка изменений данных успешно завершена"},
            {"DistUplFail","Загрузка изменений данных не выполнена"},
            {"DistUplErr","Ошибка загрузки изменений данных"},
            {"DistUplStatus","Загрузка изменений данных"},
            {"DistDnldPrimBeg","Первичная выгрузка периферийной ИБ"},
            {"DistDnldPrimSuc","Первичная выгрузка периферийной ИБ успешно завершена"},
            {"DistInit","Распределенная ИБ инициализирована"},
            {"DistPIBCreat","Создана периферийная ИБ"},
            {"DistAEParam","Изменены параметры автообмена"},
            //События изменения конфигурации
            {"RestructSaveMD","Запись измененной конфигурации"},
            {"RestructStart","Начало реструктуризации"},
            {"RestructCopy","Начато копирование результатов реструктуризации"},
            {"RestructAcptEnd","Реструктуризация завершена"},
            {"RestructStatus","Статус реструктуризации"},
            {"RestructAnalys","Анализ информации"},
            {"RestructStartWarn","Предупреждение"},
            {"RestructErr","Ошибка при реструктуризации"},
            {"RestructCritErr","Критическая ошибка при реструктуризации"},
            //События ошибки
            {"GrbgNewPerBuhTot","Бухгалтерские итоги рассчитаны"},
            {"GrbgRclcAllBuhTot","Полный пересчет бухгалтерских итогов"},
            {"GrbgSyntaxErr","Синтаксическая ошибка"},
            {"GrbgRuntimeErr","Ошибка времени выполнения"}

    };

        public static string[][] Find(string FileName, string[] searchfilter, string CurDate)
        {
           // string FileName = @"1cv7.mlg";
            FileStream fs = new FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            fs.Seek(0, SeekOrigin.End);
            byte[] buffer = new byte[5000000];
            byte[] tempbuf = new byte[1];
            Encoding encoder = Encoding.GetEncoding(1251);
            string prefix = string.Empty;
            long pos = 1;
            string[] separator = { "\r\n" };
            int FirstLF = 0;
            long CurrentDate = 0;
            if (CurDate != string.Empty)   
                CurrentDate = long.Parse(CurDate);
            bool found = false;
            string searchfilter_lower = searchfilter[2].ToLower();
            string[][] result = new string[0][];

            while ((pos > 0) && (!found))
            {
                if (fs.Position < buffer.Length)
                    Array.Resize(ref buffer, (int)fs.Position);


                fs.Seek(-buffer.Length, SeekOrigin.Current);
                pos = fs.Position;
                fs.Read(buffer, 0, buffer.Length);

                fs.Seek(-buffer.Length, SeekOrigin.Current);
                pos = fs.Position;
                if ((buffer[FirstLF] != (byte)32) && (pos > 2))
                {
                    fs.Seek(-tempbuf.Length, SeekOrigin.Current);
                    fs.Read(tempbuf, 0, tempbuf.Length);
                    while ((tempbuf[0] != 0x0A) && (fs.Position > 1))
                    {
                        prefix = encoder.GetString(tempbuf) + prefix;
                        fs.Seek(-tempbuf.Length * 2, SeekOrigin.Current);
                        pos = fs.Position;
                        fs.Read(tempbuf, 0, tempbuf.Length);
                    }
                }

                string res = prefix + encoder.GetString(buffer);
                Array.Clear(buffer, 0, buffer.Length);
                prefix = string.Empty;

                string[] loglines = res.Split(separator, StringSplitOptions.RemoveEmptyEntries);
                Array.Reverse(loglines);
                
                

                foreach (string line in loglines)
                {
                    if (!line.Contains(searchfilter[0]))
                        continue;
                    
                    if (!line.Contains(searchfilter[1]))
                        continue;

                    if (!line.ToLower().Contains(searchfilter_lower))
                        continue;

                    string[] logevent = line.Split(';');

                    if (logevent.Length > 9)
                    {
                        string date = logevent[0];

                        long tempdate = long.Parse(date);

                        if (CurrentDate == 0)
                            CurrentDate = tempdate;

                        if (tempdate < CurrentDate)                 //Отсечём предыдущие дни.
                        {
                            fs.Close();
                            return result;
                        }

                        if (logevent[4].Replace("$", "") == searchfilter[0])                       //"$Refs"
                            if ((logevent[5] == searchfilter[1]) || (searchfilter[1] == ""))                   //"RefOpen"
                                if (logevent[9].ToLower().Contains(searchfilter[2].ToLower()) || logevent[8].ToLower().Contains(searchfilter[2].ToLower())) // Номер ДФА
                                {
                                    string[] _1cObject = logevent[8].Split('/');
                                    string _1sUID = string.Empty;
                                    if (_1cObject.Length > 0)
                                        if (_1cObject[0] == "O" || _1cObject[0] == "B")
                                        {
                                            // {"O","0","0","47905","0","0","  20273492RN2"}
                                            // O/47905/(RN2)20273492
                                            string dbID = string.Empty;
                                            if(_1cObject[2].StartsWith("("))
                                            {
                                                dbID = (_1cObject[2].Substring(_1cObject[2].IndexOf(')') + 1) + _1cObject[2].Substring(1, _1cObject[2].IndexOf(')') - 1)).PadLeft(13);
                                            }
                                            else
                                            {
                                                dbID = _1cObject[2].PadRight(13);
                                            }
                                            _1sUID = string.Format("{{\"{0}\",\"0\",\"0\",\"{1}\",\"0\",\"0\",\"{2}\"}}", _1cObject[0], _1cObject[1], dbID);
                                        }

                                    Array.Resize(ref result, result.Length + 1);
                                    result[result.Length - 1] = new string[] { DateTime.ParseExact(logevent[0] + ' ' + logevent[1], "yyyyMMdd HH:mm:ss", null).ToString(), logevent[2], logevent[8], logevent[5], logevent[9], logevent[7],line, _1sUID };
                                }

                    }
                }
            }

            fs.Close();
            return result;

        }
    }
}

namespace Compound
{
    #region Adler32
    public class Adler32
    {
        // parameters
        #region

        public const uint AdlerBase = 0xFFF1;
        public const uint AdlerStart = 0x0001;
        public const uint AdlerBuff = 0x0400;
        /// Adler-32 checksum value
        private uint m_unChecksumValue = 0;
        #endregion
        public uint ChecksumValue
        {
            get
            {
                return m_unChecksumValue;
            }
        }

        public bool MakeForBuff(byte[] bytesBuff, uint unAdlerCheckSum)
        {
            if (Object.Equals(bytesBuff, null))
            {
                m_unChecksumValue = 0;
                return false;
            }
            int nSize = bytesBuff.GetLength(0);
            if (nSize == 0)
            {
                m_unChecksumValue = 0;
                return false;
            }
            uint unSum1 = unAdlerCheckSum & 0xFFFF;
            uint unSum2 = (unAdlerCheckSum >> 16) & 0xFFFF;
            for (int i = 0; i < nSize; i++)
            {
                unSum1 = (unSum1 + bytesBuff[i]) % AdlerBase;
                unSum2 = (unSum1 + unSum2) % AdlerBase;
            }
            m_unChecksumValue = (unSum2 << 16) + unSum1;
            return true;
        }

        public bool Calc(byte[] bytesBuff)
        {
            return MakeForBuff(bytesBuff, AdlerStart);
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            if (this.GetType() != obj.GetType())
                return false;
            Adler32 other = (Adler32)obj;
            return (this.ChecksumValue == other.ChecksumValue);
        }

        public static bool operator ==(Adler32 objA, Adler32 objB)
        {
            if (Object.Equals(objA, null) && Object.Equals(objB, null)) return true;
            if (Object.Equals(objA, null) || Object.Equals(objB, null)) return false;
            return objA.Equals(objB);
        }

        public static bool operator !=(Adler32 objA, Adler32 objB)
        {
            return !(objA == objB);
        }

        public override int GetHashCode()
        {
            return ChecksumValue.GetHashCode();
        }

        public override string ToString()
        {
            if (ChecksumValue != 0)
                return ChecksumValue.ToString();
            return "Unknown";
        }
    }
    #endregion

    #region RC4 encryption
    public class RC4
    {
        public static byte[] MMSkey = { 0x60, 0x46, 0xD2, 0x72, 0x64, 0x25, 0x03, 0x00, 0x09, 0x89, 0x00, 0xC0, 0xDD, 0x3B, 0xE6, 0x36 };

        public static byte[] GMkey = {  0x34, 0x43, 0x33, 0x43, 0x30, 0x42, 0x46, 0x31, 0x31, 0x35, 0x46, 0x38, 0x42, 0x39, 0x35, 0x36, 
                                        0x36, 0x39, 0x46, 0x39, 0x46, 0x43, 0x34, 0x42, 0x36, 0x44, 0x33, 0x41, 0x39, 0x44, 0x36, 0x31, 
                                        0x34, 0x31};
        public byte[] key = null;
        public byte[] data;
        private bool decrypt;
        public bool MakeXOR = true;


        public RC4(byte[] data, byte[] key)
        {
            this.data = data;
            this.key = key;
        }

        public RC4(byte[] data)
        {
            this.data = data;
            this.key = MMSkey;
        }

        public RC4()
        {
        }

        public void Encode()
        {
            bool decrypt = data[0] == 0x25 || data[0] == 0x78;

            byte[] s = new byte[256];
            int i, j, t;

            for (i = 0; i < 256; i++)
                s[i] = (byte)i;

            j = 0;
            for (i = 0; i < 256; i++)
            {
                j = (j + s[i] + key[i % key.Length]) % 256;
                s[i] ^= s[j];
                s[j] ^= s[i];
                s[i] ^= s[j];
            }

            byte tt = s[0];
            i = j = 0;
            for (int x = 0; x < data.Length; x++)
            {

                i = (i + 1) % 256;
                j = (j + s[i]) % 256;

                s[i] ^= s[j];
                s[j] ^= s[i];
                s[i] ^= s[j];

                t = (s[i] + s[j]) % 256;
                data[x] ^= s[t];

                if (MakeXOR)
                {
                    data[x] ^= tt;

                    if (decrypt)
                        tt ^= (byte)(data[x] ^ s[t]);
                    else
                        tt = data[x];
                }
            }
        }
    }
    #endregion

    #region Расширение перечислений
    public static class EnumExtension
    {
        public static string GetDescription(this Enum value)
        {
            var type = value.GetType();
            var fieldInfo = type.GetField(value.ToString(CultureInfo.InvariantCulture));
            var attribs = fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false) as DescriptionAttribute[];
            return attribs != null && attribs.Length > 0 ? attribs[0].Description : value.ToString();
        }
    }
    #endregion

    public class OleStorage
    {

        public delegate void Complete(bool ifComplete);
        public delegate void Progress(string message, int procent);

        #region Enums
        public enum StorageType
        {  //Контейнеры = каталоги
            MetaDataContainer,
            SubcontoContainer,
            SublistContainer,
            SubcontoGroupFolder,
            DocumentContainer,
            JournalContainer,
            ReportContainer,
            TypedTextContainer,
            UserDefContainer,
            PictureContainer,
            CalcJournalContainer,
            CalcVarContainer,
            AccountChartListContainer,
            AccountChartContainer,
            OperationListContainer,
            OperationContainer,
            GlobalDataContainer,
            ProvListContainer,
            TypedObjectContainer,
            WorkBookContainer,
            ModuleContainer,
            WorkPlaceType,
            RigthType,
            //Элементы
            MetaDataStream, //Описание метаданных
            MetaDataHolderContainer,
            GuidHistoryContainer,
            TagStream,
            MetaDataDescription, //Глобальник
            DialogEditor,
            TextDocument,
            [Description("Moxcel.Worksheet")]
            MoxcelWorksheet,
            UsersInterfaceType,
            SubUsersInterfaceType,
            MenuEditorType,
            ToolbarEditorType,
            PictureGalleryContainer
        }

        public static List<StorageType> ListCatalogTypes = new List<StorageType>
        { 
           StorageType.MetaDataContainer,
           //StorageType.MetaDataHolderContainer,
           StorageType.SubcontoContainer,
           StorageType.SublistContainer,
           StorageType.SubcontoGroupFolder,
           StorageType.DocumentContainer,
           StorageType.JournalContainer,
           StorageType.ReportContainer,
           StorageType.TypedTextContainer,
           StorageType.UserDefContainer,
           StorageType.PictureContainer,
           StorageType.CalcJournalContainer,
           StorageType.CalcVarContainer,
           StorageType.AccountChartListContainer,
           StorageType.AccountChartContainer,
           StorageType.OperationListContainer,
           StorageType.OperationContainer,
           StorageType.GlobalDataContainer,
           StorageType.ProvListContainer,
           StorageType.TypedObjectContainer,
           StorageType.WorkBookContainer,
           StorageType.ModuleContainer,
           StorageType.WorkPlaceType,
           StorageType.RigthType,
           StorageType.UsersInterfaceType,
           StorageType.SubUsersInterfaceType
        };

        public enum ValueType
        {
            U, // Неопределенный
            N, // Число
            S, // Строка
            D, // Дата
            E, // Перечисление
            B, // Справочник
            O, // Документ
            C, // Календарь
            A, // ВидРасчета
            T, // Счет
            K, // ВидСубконто
            P // ПланСчетов
        }
        #endregion

        public static Adler32 CRC = new Adler32();

        #region Физическая структура Метаданных

        public class MetaDescriptor
        {
            public int orderid;
            public StorageType Type;
            public MetaDescriptor Parent;
            public string Path;
            public string Name;
            public string Description;
            public string Prop_4;
            public bool isContainer;

            public MetaDescriptor()
            {
            }

            public MetaDescriptor(StorageType Type, string Name, string Description)
            {
                this.Type = Type;
                this.Name = Name;
                this.Description = Description;
                if (Type == StorageType.MoxcelWorksheet && this.Description == "Moxel WorkPlace")
                    this.Description = "Таблица";
            }

            public MetaDescriptor(StorageType Type, string Name, string Description, MetaDescriptor Parent)
                : this(Type, Name, Description)
            {
                this.Parent = Parent;
                if (Type == StorageType.MoxcelWorksheet && this.Description == "Moxel WorkPlace")
                    this.Description = "Таблица";
            }


            public static implicit operator string(MetaDescriptor Item)
            {
                return Item.ToString();
            }

            public string ToString()
            {
                return string.Format("{{\"{0}\",\"{1}\",\"{2}\",\"{3}\"}}", Type.ToString(), Name, Description, Prop_4);
            }
        }


        public enum SubType
        {
            Procedure,
            Function
        }

        public class Sub
        {
            public string PreComment = null;
            public string Name = null;
            public string Body = null;
            public SubType Type = SubType.Function;
            public List<string> Parameters = new List<string>();
            public bool Public = false;
            public bool PreDeclared = false;
            public string Tail = String.Empty;

            public List<string> Developers = new List<string>();
            public List<string> Incidents = new List<string>();
            public Dictionary<string, int> Modifacations = new Dictionary<string, int>();

            public static Dictionary<string, string> DeveloperID = new Dictionary<string, string>
            {
                //{"ДД ","Дрейцер Д."},
                //{"Frog","Козлов В."},
                //{"Berko","Берко А."},
                //{"Iceberg","Берко А."},
                //{"Берко","Берко А."},
                //{"КГ ","Давыдова К."},
                //{"Нина","Подвязникова Н."},
                //{"Максим","Чернушко М."},
                //{"Олег","Шейко О."},
                //{"Соловьёв","Соловьёв И."}
            };

            private int SubstringCount(string sourcestring, string substring)
            {
                return (sourcestring.Length - sourcestring.Replace(substring, "").Length) / substring.Length;
            }

            public void ParceText(string Body)
            {
                string[] splitter = { "\r\n" };
                string[] BodyStrings = Body.Split(splitter, StringSplitOptions.RemoveEmptyEntries);

                string EndOfsub;
                string TypeOfSub;

                string tempBodyString = string.Empty;

                if (!PreDeclared)
                    foreach (string BodyString in BodyStrings)
                    {
                        if (BodyString.Length < 2)
                            continue;

                        if (BodyString.Substring(0, 2) == "//")
                            continue;

                        if (BodyString.Contains("(") || tempBodyString.Contains("("))
                        {
                            if (!BodyString.Contains(")"))
                            {
                                tempBodyString += BodyString.Replace("\t", "");
                                continue;
                            }

                            tempBodyString += BodyString.Replace("\t", "");

                            string[] parameters = tempBodyString.Substring(tempBodyString.IndexOf('(') + 1, tempBodyString.IndexOf(')') - tempBodyString.IndexOf('(') - 1).Split(',');

                            foreach (string parameter in parameters)
                            {
                                if (parameter != "")
                                    Parameters.Add(parameter.Replace("Знач ", "").Split('=')[0].TrimStart(' ').TrimEnd(' '));
                            }

                            Public = tempBodyString.ToLower().Contains("экспорт");
                            PreDeclared = tempBodyString.ToLower().Contains("далее");

                            if (PreDeclared)
                            {
                                EndOfsub = "Далее";
                                string bodyToParce = string.Empty;
                                foreach (string substr in Body.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
                                {
                                    if (substr.Length > 2)
                                        if (substr.Substring(0, 2) != "//")
                                            bodyToParce += substr + "\r\n";
                                }

                                Tail = bodyToParce.Substring(bodyToParce.IndexOf(EndOfsub) + EndOfsub.Length);
                                return;
                            }

                            tempBodyString = string.Empty;
                            break;
                        }
                    }

                if (Type == SubType.Function)
                {
                    EndOfsub = "КонецФункции";
                    TypeOfSub = "Функция";
                }
                else
                {
                    EndOfsub = "КонецПроцедуры";
                    TypeOfSub = "Процедура";
                }

                if (Body.ToLower().Contains("\r\n" + EndOfsub.ToLower()))
                {

                    int SubStart = Body.IndexOf("\r\n" + TypeOfSub, StringComparison.OrdinalIgnoreCase);
                    if(SubStart <= 0)
                        SubStart = Body.IndexOf(TypeOfSub, StringComparison.OrdinalIgnoreCase);

                    this.PreComment = Body.Substring(0, 2 + SubStart);
                    int len = SubStart + EndOfsub.Length + 2 - PreComment.Length;
                    if (len < 0)
                    {
                        throw new Exception("Ошибка при разборе процедуры " + this.Name + " в модуле ");
                    }
                    this.Body = Body.Substring(PreComment.Length, len);
                    Tail = Body.Substring(Body.IndexOf(EndOfsub, StringComparison.OrdinalIgnoreCase) + EndOfsub.Length);

                    foreach (string ID in DeveloperID.Keys)
                    {
                        if (Body.Contains(ID))
                        {
                            Developers.Add(DeveloperID[ID]);

                            if (Modifacations.ContainsKey(DeveloperID[ID]))
                                Modifacations[DeveloperID[ID]] += SubstringCount(this.PreComment + "\r\n" + this.Body, ID);
                            else
                                Modifacations.Add(DeveloperID[ID], SubstringCount(this.PreComment + "\r\n" + this.Body, ID));
                        }
                    }

                    if (Tail.IndexOf("\r\n//+") >= 0)
                    {
                        Tail = Tail.Substring(Tail.IndexOf("\r\n//+") + 2);
                    }
                    else
                    {
                        if (Tail.IndexOf("\r\n//*") >= 0)
                        {
                            Tail = Tail.Substring(Tail.IndexOf("\r\n//*") + 2);
                        }
                        else
                        {
                            if (Tail.IndexOf("\r\n///") >= 0)
                            {
                                Tail = Tail.Substring(Tail.IndexOf("\r\n///") + 2);
                                Tail = Tail.Substring(Tail.IndexOf("\r\n") + 2);
                            }
                        }
                    }


                }
            }

            public Sub(SubType Type, string Body, string Name)
            {
                this.Type = Type;
                this.Name = Name;
                ParceText(Body);
            }

        }

        public class ProgramModule
        {
            public string GlobalVars = null;
            public string GlobalContext = null;
            public Dictionary<string, Sub> Procedures = new Dictionary<string, Sub>();
            public string Text = null;

            private void ParceProcedures(string ModuleText)
            {
                string[] splitter = { "\r\nПроцед", "\r\nФунк" };
                string[] Procedures_txt = ModuleText.Split(splitter, StringSplitOptions.RemoveEmptyEntries);

                string previous = string.Empty;
                foreach (string procedure in Procedures_txt)
                {
                    if (procedure.Length < 4)
                        continue;

                    SubType TypeOfSub = SubType.Function;
                    string EndOfsub = null;
                    string SubKeyWord = null;

                    switch (procedure.Substring(0, 4))
                    {
                        case "ция ":
                            {
                                TypeOfSub = SubType.Function;
                                EndOfsub = "КонецФункции";
                                SubKeyWord = "Функция";
                                break;
                            }
                        case "ура ":
                            {
                                TypeOfSub = SubType.Procedure;
                                EndOfsub = "КонецПроцедуры";
                                SubKeyWord = "Процедура";
                                break;
                            }
                        default:
                            {
                                if (procedure.ToLower().Contains("перем "))
                                {
                                    //Вырежем комментарии
                                    foreach (string str in procedure.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
                                    {
                                        if (str.Substring(0, 1) != "/")
                                            GlobalVars += str + "\r\n";
                                    }
                                    previous = null;
                                }
                                continue;
                            }
                    }

                    Sub Procedure;
                    string proceduretext = string.Empty;

                    proceduretext = previous + "\r\n" + SubKeyWord + " " + procedure.Substring(4);

                    string Name = SubKeyWord + " " + procedure.Substring(4, procedure.IndexOf('(') - 4).TrimStart();

                    if (Procedures.ContainsKey(Name))
                    {
                        Procedure = Procedures[Name];
                        Procedure.ParceText(proceduretext);
                    }
                    else
                    {
                        Procedure = new Sub(TypeOfSub, proceduretext, Name);
                        Procedures.Add(Procedure.Name, Procedure);
                    }

                    previous = Procedure.Tail;

                    if (previous.ToLower().Contains("перем "))
                    {
                        GlobalVars += previous;
                        previous = null;
                    }
                }

                if (previous != null)
                    foreach (string str in previous.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
                    {   //Вырежем комментарии
                        if (str.Substring(0, 1) != "/")
                            GlobalContext += str + "\r\n";
                    }
            }

            public ProgramModule(string text)
            {
                this.Text = text;
                ParceProcedures(text);
            }
        }

        public class MetaItem : MetaDescriptor
        {
            public string MetaPath = string.Empty;

            private byte[] Decompress(byte[] compressed)
            {
                MemoryStream ms = new MemoryStream(compressed.Length);
                ms.Write(compressed, 0, compressed.Length);
                ms.Position = 0;
                DeflateStream compressedzipStream = new DeflateStream(ms, System.IO.Compression.CompressionMode.Decompress, false);

                byte[] data = new byte[compressed.Length * 10];
                int totalCount = 0;
                int bytesRead = compressedzipStream.Read(data, 0, data.Length);
                totalCount += bytesRead;
                compressedzipStream.Close();
                //compressedzipStream.Dispose();
                Array.Resize<byte>(ref data, totalCount);
                return data;
            }

            static List<StorageType> Modules = new List<StorageType> 
            { 
                StorageType.TextDocument, 
                StorageType.MetaDataDescription
            };

            static List<StorageType> Dialogs = new List<StorageType> 
            { 
                StorageType.DialogEditor
            };

            static List<StorageType> Moxels = new List<StorageType> 
            { 
                StorageType.MoxcelWorksheet
            };

            public uint CheckSumm = 0;
            bool IsCompressed = false;
            bool IsEncrypted = false;
            public string Moduletext = string.Empty;
            public ProgramModule Module;
            private byte[] _data;
            public byte[] data
            {
                get { return _data; }

                set
                {

                    CheckSumm = 0;
                    if (IsCompressed)
                    {
                        byte[] decompressed = Decompress(value);

                        if (decompressed.Length > 0)
                        {
                            if (Type == StorageType.MetaDataDescription)
                            {
                                IsEncrypted = decompressed[0] == 0x9e; // определяем зашифрован или нет

                                if (IsEncrypted)
                                {
                                    byte[] Crypted = new byte[decompressed.Length];

                                    //У шифрованного глобальника первые 510 байт - хз что, после обрезки расшифровывается успешно.
                                    Array.ConstrainedCopy(decompressed, 510, Crypted, 0, decompressed.Length - 510);
                                    decompressed = Crypted;

                                    RC4 Codec = new RC4(decompressed, RC4.GMkey);
                                    Codec.MakeXOR = false;      //Тут используется не модифицированый RC4
                                    Codec.Encode();
                                }

                            }
                        }

                        Moduletext = Encoding.GetEncoding(1251).GetString(decompressed);
                        Module = new ProgramModule(Moduletext);
                    }

                    _data = value;

                    if (Type == StorageType.MetaDataStream)
                    {
                        IsEncrypted = _data[0] != 0xFF;
                        if (IsEncrypted)
                        {
                            RC4 Codec = new RC4(_data, RC4.MMSkey);
                            Codec.Encode();
                        }

                        Moduletext = Encoding.GetEncoding(1251).GetString(_data);
                    }

                    if (Type == StorageType.TagStream)
                    {
                        IsEncrypted = true; //Это всегда зашифровано.
                        RC4 Codec = new RC4(_data, RC4.MMSkey);
                        Codec.Encode();

                        Moduletext = Encoding.GetEncoding(1251).GetString(_data);
                    }

                    if (Type == StorageType.DialogEditor)
                    {
                        Moduletext = Encoding.GetEncoding(1251).GetString(_data); //Надо распарсить в XML
                    }


                    if (Moduletext != null && Moduletext != string.Empty)
                    {
                        CheckSumm = (uint)Moduletext.GetHashCode();
                    }

                    if (CheckSumm == 0)
                    {
                        CRC.Calc(value);
                        CheckSumm = CRC.ChecksumValue;
                    }

                    _data = value;

                }
            }


            public MetaItem(StorageType Type, string Name, string Description)
                : base(Type, Name, Description)
            {
                IsCompressed = Modules.Contains(Type);
                isContainer = false;
            }

            public MetaItem(StorageType Type, string Name, string Description, MetaDescriptor Parent)
                : base(Type, Name, Description, Parent)
            {
                IsCompressed = Modules.Contains(Type);
                isContainer = false;
            }

            public MetaItem()
            {
            }
        }

        public class MetaContainer : MetaDescriptor
        {
            public List<MetaDescriptor> Items = null;
            public string Contents
            {
                get
                {
                    string res = "{Container.Contents";
                    foreach (MetaDescriptor Item in Items)
                    {
                        res += "," + Item.ToString();
                    }
                    res += "}";
                    return res;
                }
            }

            public MetaItem GetSubItem(string path)
            {
                string[] PathDirectories = path.Split('\\');

                MetaContainer Root = this;

                foreach (string subdir in PathDirectories)
                {
                    if (subdir != "")
                    {
                        if (Name == subdir)
                            continue;
                        MetaDescriptor item = Root.Items.Find(x => x.Name == subdir);
                        if (item == null)
                            return null;
                        if (item.isContainer)
                            Root = (MetaContainer)item;
                        else
                        {
                            return (MetaItem)item;
                        }


                    }
                }
                return null;
            }

            public MetaContainer GetSubCatalog(string path)
            {
                string[] PathDirectories = path.Split('\\');

                MetaContainer Root = this;

                foreach (string subdir in PathDirectories)
                {
                    if (subdir != "")
                    {
                        if (Name == subdir)
                            continue;
                        MetaDescriptor item = Root.Items.Find(x => x.Name == subdir);
                        if (item == null)
                            return null;

                        if (item.isContainer)
                            Root = (MetaContainer)item;
                        else
                        {
                            return (MetaContainer)item;
                        }


                    }
                }
                return Root;
            }

            public MetaContainer(StorageType Type, string Name, string Description)
                : base(Type, Name, Description)
            {
                Items = new List<MetaDescriptor>();
                isContainer = true;
            }

            public MetaContainer(StorageType Type, string Name, string Description, MetaDescriptor Parent)
                : base(Type, Name, Description, Parent)
            {
                Items = new List<MetaDescriptor>();
                isContainer = true;
            }


            public List<MetaDescriptor> ParceContainer_Contents(string Contents, MetaDescriptor Prent)
            {
                List<MetaDescriptor> Items = new List<MetaDescriptor>();
                Contents = Contents.Replace("{\"Container.Contents\",", "");
                Contents = Contents.Substring(0, Contents.Length - 2);
                StorageType Type;
                foreach (string subitem in Contents.Replace("},{", "#").Split('#'))
                {
                    string[] subelements = subitem.Replace("{", "").Replace("}", "").Replace("\"", "").Split(',');
                    if (StorageType.TryParse(subelements[0].Replace(".", ""), out Type))
                    {
                        if (ListCatalogTypes.Contains(Type))
                            Items.Add(new MetaContainer(Type, subelements[1], subelements[2], Parent));
                        else
                            Items.Add(new MetaItem(Type, subelements[1], subelements[2], Parent));
                    }

                }


                return Items;
            }

            bool ReadStorage(NativeMethods.IStorage RootStorage)
            {
                NativeMethods.IStorage storage = null;
                IStream pIStream = null;
                NativeMethods.IEnumSTATSTG pIEnumStatStg = null;
                byte[] data = { 0 };
                uint fetched = 0;

                System.Runtime.InteropServices.ComTypes.STATSTG[] regelt =
                {
                    new System.Runtime.InteropServices.ComTypes.STATSTG()
                };

                RootStorage.OpenStream("Container.Contents", IntPtr.Zero, (NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE), 0, out pIStream);

                data = NativeMethods.ReadIStream(pIStream);
                Marshal.ReleaseComObject(pIStream);
                Marshal.FinalReleaseComObject(pIStream);
                pIStream = null;

                Items = ParceContainer_Contents(Encoding.GetEncoding(1251).GetString(data), this);

                RootStorage.EnumElements(0, IntPtr.Zero, 0, out pIEnumStatStg);
                while (pIEnumStatStg.Next(1, regelt, out fetched) == 0)
                {
                    string filePage = regelt[0].pwcsName;
                    if (filePage != "Container.Contents")
                    {

                        if ((NativeMethods.STGTY)regelt[0].type == NativeMethods.STGTY.STGTY_STREAM)
                        {
                            MetaItem item = (MetaItem)Items.Find(x => x.Name == filePage);
                            if (item != null)
                            {
                                RootStorage.OpenStream(filePage, IntPtr.Zero, (NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE),
                                    0,
                                    out pIStream);
                                if (pIStream != null)
                                {
                                    item.Path += Path + "\\" + Name;
                                    item.data = NativeMethods.ReadIStream(pIStream);
                                    Marshal.ReleaseComObject(pIStream);
                                    Marshal.FinalReleaseComObject(pIStream);
                                    pIStream = null;
                                }
                            }
                            else
                            {
                                //if (filePage != "MetaContainer.Profile")
                                //    MessageBox.Show("В структуре не найден объект " + filePage);
                            }

                        }

                        if ((NativeMethods.STGTY)regelt[0].type == NativeMethods.STGTY.STGTY_STORAGE)
                        {
                            MetaContainer item = (MetaContainer)Items.Find(x => (x.Name == filePage) && (x.isContainer));
                            if (item != null)
                            {
                                RootStorage.OpenStorage(filePage, null, (NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE),
                                    IntPtr.Zero, 0, out storage);
                                item.Path += Path + "\\" + Name;
                                item.Parent = this;
                                item.ReadStorage(storage);
                                Marshal.ReleaseComObject(storage);
                                Marshal.FinalReleaseComObject(storage);
                                storage = null;
                            }
                            else
                            {
                                // MessageBox.Show("В структуре не найден объект " + filePage);
                            }
                        }
                    }
                }
                return true;
            }

            //это ссылка на массив байт в памяти
            private IntPtr hGlobal;                     //при уничтожениии экземпляра нужно освободить память
            private NativeMethods.ILockBytes LockBytes;
            //Это ссылка на открытый компаунд
            private NativeMethods.IStorage RootStorage; //при уничтожениии экземпляра нужно освободить COM-объект

            bool OpenStorage(string FileName)
            {
                //Читаем MD в массив байт
                byte[] buffer = File.ReadAllBytes(FileName);
                //Выделяем область в памяти для буфера
                hGlobal = Marshal.AllocHGlobal(buffer.Length);
                //Помещаем считанный файл в буфер.
                Marshal.Copy(buffer, 0, hGlobal, buffer.Length);
                buffer = null;
                NativeMethods.CreateILockBytesOnHGlobal(hGlobal, false, out LockBytes);
                //Открываем считанный файл из памяти
                GC.Collect(GC.MaxGeneration);
                return NativeMethods.StgOpenStorageOnILockBytes(LockBytes, null, NativeMethods.STGM.READWRITE | NativeMethods.STGM.SHARE_EXCLUSIVE, IntPtr.Zero, 0, out RootStorage) == 0;
            }

            public string FileName;

            public MetaContainer(string mdFileNAme, Progress progressproc = null)
            {
                this.FileName = mdFileNAme;

                if (!OpenStorage(FileName))
                    throw new Exception("Не удалось открыть конфигурацию.");

                NativeMethods.IStorage Storage;
                System.Runtime.InteropServices.ComTypes.STATSTG MetaDataInfo;

                NativeMethods.IStorage storage = null;
               // NativeMethods.IStorage RootStorage = null;
                IStream pIStream = null;
                NativeMethods.IEnumSTATSTG pIEnumStatStg = null;
                byte[] data = { 0 };
                uint fetched = 0;

                //if (NativeMethods.StgOpenStorage(mdFileNAme, null, STGM.READ | STGM.SHARE_DENY_WRITE, IntPtr.Zero, 0, out RootStorage) == 0)
                {

                    System.Runtime.InteropServices.ComTypes.STATSTG[] regelt =
                    {
                        new System.Runtime.InteropServices.ComTypes.STATSTG()
                    };

                    Name = "Root";

                    RootStorage.OpenStream("Container.Contents", IntPtr.Zero, (NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE), 0, out pIStream);
                    data = NativeMethods.ReadIStream(pIStream);
                    Marshal.ReleaseComObject(pIStream);
                    Marshal.FinalReleaseComObject(pIStream);
                    pIStream = null;

                    Items = ParceContainer_Contents(Encoding.GetEncoding(1251).GetString(data), this);
                    RootStorage.EnumElements(0, IntPtr.Zero, 0, out pIEnumStatStg);
                    while (pIEnumStatStg.Next(1, regelt, out fetched) == 0)
                    {
                        string filePage = regelt[0].pwcsName;
                        if (filePage != "Container.Contents")
                        {
                            if ((NativeMethods.STGTY)regelt[0].type == NativeMethods.STGTY.STGTY_STORAGE)
                            {
                                MetaContainer item = (MetaContainer)Items.Find(x => x.Name == filePage);

                                if (item != null)
                                {
                                    item.Path += Path + "\\" + Name;
                                    RootStorage.OpenStorage(filePage, null, (NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE),
                                        IntPtr.Zero, 0, out storage);
                                    item.Parent = this;
                                    item.ReadStorage(storage);
                                    Marshal.ReleaseComObject(storage);
                                    Marshal.FinalReleaseComObject(storage);
                                    storage = null;
                                }
                                else
                                {
                                    //throw new Exception("В структуре не найден объект " + filePage);
                                }

                            }

                            if ((NativeMethods.STGTY)regelt[0].type == NativeMethods.STGTY.STGTY_STREAM)
                            {
                                MetaItem item = (MetaItem)Items.Find(x => x.Name == filePage);
                                if (item != null)
                                {
                                    RootStorage.OpenStream(filePage, IntPtr.Zero, (NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE),
                                        0,
                                        out pIStream);
                                    if (pIStream != null)
                                    {
                                        item.Path += Path + "\\" + Name;
                                        item.data = NativeMethods.ReadIStream(pIStream);
                                        Marshal.ReleaseComObject(pIStream);
                                        Marshal.FinalReleaseComObject(pIStream);
                                        pIStream = null;
                                    }
                                }
                                else
                                {
                                    //if (filePage != "MetaContainer.Profile")
                                    //    MessageBox.Show("В структуре не найден объект " + filePage);
                                }

                            }

                        }
                    }

                    if (pIStream != null)
                    {
                        Marshal.ReleaseComObject(pIStream);
                        Marshal.FinalReleaseComObject(pIStream);
                        pIStream = null;
                    }

                    if (storage != null)
                    {
                        Marshal.ReleaseComObject(storage);
                        Marshal.FinalReleaseComObject(storage);
                        storage = null;
                    }

                    if (RootStorage != null)
                    {
                        Marshal.ReleaseComObject(RootStorage);
                        Marshal.FinalReleaseComObject(RootStorage);
                        RootStorage = null;
                    }
                }

                //else
                //{
                //    throw new Exception("Не удалось открыть конфигурацию");
                //}
            }

            ~MetaContainer()
            {
                //Освободим паямть
                if (RootStorage != null)
                {
                    Marshal.ReleaseComObject(RootStorage);
                    Marshal.FinalReleaseComObject(RootStorage);
                    RootStorage = null;
                }

                if (LockBytes != null)
                {
                    Marshal.ReleaseComObject(LockBytes);
                    Marshal.FinalReleaseComObject(LockBytes);
                    LockBytes = null;
                }
                Marshal.FreeHGlobal(hGlobal);
                hGlobal = IntPtr.Zero;
            }
        }

        #endregion

        #region Логическая структура Метаданных
        public class MetaObject
        {
            public int ID = 0;
            public string Identity = string.Empty;
            public string Alias = string.Empty;
            public string Description = string.Empty;
            public string MetaPath = string.Empty;

            public override string ToString()
            {
                return string.Format("{{\"{0}\",\"{1}\",\"{2}\",\"{3}\"}}", ID, Identity, Alias, Description);
            }

            public static implicit operator string(MetaObject obj)
            {
                return obj.ToString();
            }

        }

        public class Form : MetaObject
        {
            MetaContainer Catalog;
            public MetaItem Dialog;
            public MetaItem DialogModule;
            public List<Moxel> Moxels = new List<Moxel>();

            public Form()
                : base()
            {
            }

            public Form(MetaContainer Catalog)
            {
                this.Catalog = Catalog;
                this.DialogModule = (MetaItem)Catalog.Items.Find(x => x.Type == StorageType.TextDocument); //.GetSubItem("\\MD Programm text");
                this.Dialog = (MetaItem)Catalog.Items.Find(x => x.Type == StorageType.DialogEditor);// Catalog.GetSubItem("\\Dialog Stream");
            }

            public Form(MetaContainer Catalog, string Path)
                : this(Catalog)
            {
                MetaPath = Path;
                foreach (MetaItem item in Catalog.Items.FindAll(x => x.Type == StorageType.MoxcelWorksheet))
                    Moxels.Add(new Moxel(item, this));
            }

        }

        public class Parameter : MetaObject
        {
            public ValueType Type;
            public int Length = 0;
            public int Precision = 0;
            public MetaObject TypedObject;
            public bool NoNegative = false;
            public bool Splittriades = false;

            public override string ToString()
            {
                return string.Format("{{\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{7}\",\"{4}\",\"{5}\",\"{6}\"}}", ID, Identity, Alias, Description, Length, Precision, TypedObject.ID, Type.ToString());
            }

        }

        public enum NumeratorType
        {
            Text = 1,
            Number = 2
        }

        public enum AutoNumeration
        {
            No = 1,
            Yes = 2
        }

        public class SubcontoParameter : Parameter
        {
            public bool Perodical = false;
            public bool UseForItem = false;
            public bool UseForGroup = false;
            public bool Sort = false;
            public bool HistoryManual = false;
            public bool ChangeByDocument = false;
            public bool Selection = false;

        }

        public class Subconto : MetaObject
        {
            public int Parent = 0;
            public int CodeLength = 0;
            public int CodeSeries = 1;
            public NumeratorType CodeType = NumeratorType.Text;
            public AutoNumeration AutoNum = AutoNumeration.Yes;
            public int NameLength = 0;
            public int MainRepresent = 0;
            public int EditMode = 0;
            public int LevelCount = 0;
            public int SelectFormID = 0;
            public int MainFormID = 0;
            public int OneForm = 0;
            public int UniqCode = 0;
            public int GroupsInTop = 0;

            public Form SelectForm = null;
            public Form MainForm = null;
            public Form Form;
            public Form FolderForm;
            public List<Form> ListForms = new List<Form>();
            public List<SubcontoParameter> Params = new List<SubcontoParameter>();

            public Subconto(string[] objparams) //Справочники
            {
                ID = int.Parse(objparams[0]);
                Identity = objparams[1];
                Description = objparams[2];
                Alias = objparams[3];
                Parent = int.Parse(objparams[4]);
                CodeLength = int.Parse(objparams[5]);
                CodeSeries = int.Parse(objparams[6]);
                CodeType = (NumeratorType)int.Parse(objparams[7]);
                AutoNum = (AutoNumeration)int.Parse(objparams[8]);
                NameLength = int.Parse(objparams[9]);
                MainRepresent = int.Parse(objparams[10]);
                EditMode = int.Parse(objparams[11]);
                LevelCount = int.Parse(objparams[12]);
                SelectFormID = int.Parse(objparams[13]);
                MainFormID = int.Parse(objparams[14]);
                OneForm = int.Parse(objparams[15]);
                UniqCode = int.Parse(objparams[16]);
                GroupsInTop = int.Parse(objparams[17].Replace(",", ""));
                MetaPath = "Справочник." + Identity; 
            }
        }

        public class DocumentParameter : Parameter
        {
        }

        public class DocumentTableParameter : Parameter
        {
            public bool TotalColumn = false;
        }

        public class Document : MetaObject
        {
            public enum DocumentPeriodicity
            {
                All = 0,
                InYead = 1,
                InQuartal,
                InMonth,
                InDay
            }

            public int NumberLength = 0;
            public DocumentPeriodicity Periodicity = 0;    //Периодичность: 0 - по всем данного вида, 1 - в пределах года, 2 - в пределах квартала, 3 - в пределах месяца, 4 - в пределах дня.
            public NumeratorType NumberType = NumeratorType.Text;
            public AutoNumeration AutoNum = AutoNumeration.Yes;
            public int JournalID = 0;
            public int Unkown = -1;
            public bool UniqNumber = true;
            public int NumeratorID = 0;

            public bool Operative = false;
            public bool Calculation = false;
            public bool Accounting = false;
            public List<Document> InputOnBasis;
            public bool BaseForAnyDocument = false;
            public int CreateOperation = 2;
            public bool AutoLineNumber = true;
            public bool AutoRemovActions = true;
            public bool EditOperation = false;
            public bool CanDoActions = true;


            public Form Form;
            public MetaItem TransactionModule;
            public List<DocumentParameter> HeadFileds = new List<DocumentParameter>();
            public List<DocumentTableParameter> TableFields = new List<DocumentTableParameter>();
        }

        public class Journal : MetaObject
        {
            public int Unknown = 0;
            public int JornalType = 0;//       Тип журнала: 0 - обычный, 1 - общий
            public Form SelectForm;//       Числовой идентификатор формы для выбора
            public Form MainForm;//         Числовой идентификатор основной формы  
            public bool NoAdditional;//     Не дополнительный: 0 - журнал дополнительный, 1 - нет

            public List<Form> ListForms = new List<Form>();
        }

        public class EnumVal : MetaObject
        {
            int OrderID = 1;
        }

        public class EnumItem : MetaObject
        {
            public List<EnumVal> Values = new List<EnumVal>();
        }

        public class CalculationAlgorithm : MetaObject
        {
            public MetaItem CalculationModule;
        }

        public class CalcJournal : MetaObject
        {
            public int Unknown = 0;
            public int JornalType = 0;//       Тип журнала: 0 - обычный, 1 - общий
            public Form SelectForm;//       Числовой идентификатор формы для выбора
            public Form MainForm;//         Числовой идентификатор основной формы  
            public List<Form> ListForms = new List<Form>();
        }

        public class BuhParameters : MetaObject         //Параметры бухгалтерии
        {
            public List<Form> AccountChart = new List<Form>();
            public List<Form> AccountChartList = new List<Form>();
            public List<Form> OperationList = new List<Form>();
            public List<Form> ProvListList = new List<Form>();
        }

        public class ReportItem : MetaObject
        {
            public Form Form;
        }

        public class CalcVarItem : MetaObject
        {
            public Form Form;
        }

        public enum DescriptorType
        {
            Null,
            MainDataContDef,
            TaskItem,
            GenJrnlFldDef,
            DocSelRefObj,
            DocNumDef,
            Registers,
            Documents,
            Journalisters,
            EnumList,
            ReportList,
            CJ,
            Calendars,
            Algorithms,
            RecalcRules,
            CalcVars,
            Groups,
            [Description("Document Streams")]
            DocumentStreams,
            Buh,
            CRC,
            //***************************************
            Refers, //Ссылки. Для разных типов разные
            Consts, //Константы
            SbCnts, //Справочники
            Params, //Атрибуты
            Form    //Описатель списка форм
        }

        public class ExtReport : CalcVarItem
        {
            public string FilePath = string.Empty;
            MetaContainer MetadataFileTree = null;
            string MMS = string.Empty;

            public ExtReport()
                : base()
            {

            }

            public ExtReport(string FileName)
                : base()
            {
                try
                {
                    this.FilePath = Path.GetDirectoryName(FileName);

                    Identity = Path.GetFileName(FileName);
                    Alias = Path.GetFileNameWithoutExtension(FileName);

                    MetadataFileTree = new MetaContainer(FileName);
                    if (MetadataFileTree.Items == null)
                        return;

                    MMS = MetadataFileTree.GetSubItem("Root\\Main MetaData Stream").Moduletext;
                    Form = new Form(MetadataFileTree.GetSubCatalog("Root\\"), Identity + ".Форма");
                }
                catch (Exception ex)
                {
                    throw new Exception("Ошибка открытия внешней обработки: " + ex.ToString());
                }
            }
        }

        public class Moxel : MetaItem
        {
            public string Header = string.Empty;
            public List<Sub> UsedIn = null;
            public MetaObject Owner;
            public bool IsEmpty = false;
            
            public Moxel(MetaItem item)
                : base(StorageType.MoxcelWorksheet, item.Name, item.Description)
            {
                data = item.data;
                Parent = item.Parent;
                Path = item.Path;
            }

            public Moxel(MetaItem item, MetaObject Owner)
                : this(item)
            {
                this.Owner = Owner;
                this.MetaPath = Owner.MetaPath + "." + this.Description;
                IsEmpty = this.data.Length <= 147;
                if (!(Owner is TaskItem) && !IsEmpty)
                {
                    UsedIn = ((Form)Owner).DialogModule.Module.Procedures.Values.ToList<Sub>().FindAll(x => (x.Body != null) && (x.Body.Contains("ИсходнаяТаблица(\"" + this.Description + "\")")));
                    if (UsedIn.Count == 0 && ((Form)Owner).Moxels.Count < 2)
                    {
                        UsedIn = ((Form)Owner).DialogModule.Module.Procedures.Values.ToList<Sub>().FindAll(x => (x.Body != null) && (x.Body.Contains("СоздатьОбъект(\"Таблица\")")));
                        if (UsedIn.Count == 0)
                            IsEmpty = true;

                        foreach (Sub sub in UsedIn)
                        {
                            string[] body = sub.Body.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                            foreach (string codestring in body.ToList<string>().FindAll(x => (x.Contains(".Показать(\"")) && (!x.Trim().StartsWith("//"))))
                            {
                                Header = codestring.Substring(codestring.IndexOf(".Показать(\"") + 11);
                                Header = Header.Substring(0, Header.IndexOf(");")).TrimEnd(new char[] { ',', '"'});
                            }

                        }
                    }
                    else
                    {

                        if (UsedIn.Count == 0)
                            IsEmpty = true;

                        foreach (Sub sub in UsedIn)
                        {
                            string[] body = sub.Body.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                            string varname = string.Empty;
                            foreach (string codestring in body.ToList<string>().FindAll(x => x.Contains("ИсходнаяТаблица(\"" + this.Description + "\")")))
                            {
                                varname = codestring.Substring(0, codestring.IndexOf(".ИсходнаяТаблица")).Trim();
                            }

                            if (varname != string.Empty)
                            {
                                string searchstring = varname + ".Показать(\"";
                                foreach (string codestring in body.ToList<string>().FindAll(x => (x.Contains(searchstring)) && (!x.Trim().StartsWith("//"))))
                                {
                                    Header = codestring.Substring(codestring.IndexOf(searchstring) + searchstring.Length);
                                    if (Header.IndexOf(");") > 0)
                                        Header = Header.Substring(0, Header.IndexOf(");")).TrimEnd(new char[] { ',', '"' });
                                }
                            }
                        }
                    }
                    
                }

            }

            public Moxel(MetaItem item, MetaObject Owner, string MetaPath)
                : this(item)
            {
                this.Owner = Owner;
                this.MetaPath = MetaPath + "." + this.Description;
            }
        }

        public class TaskItem : MetaObject
        {
            public List<Document> Documents = new List<Document>();
            public List<Subconto> Subcontos = new List<Subconto>();
            public List<Journal> Journals = new List<Journal>();
            public List<ReportItem> Reports = new List<ReportItem>();
            public List<CalcVarItem> CalcVars = new List<CalcVarItem>();
            public List<CalculationAlgorithm> Algorithms = new List<CalculationAlgorithm>();
            public List<CalcJournal> CJ = new List<CalcJournal>();
            public List<Moxel> GlobalWorkSheets = new List<Moxel>();

            public BuhParameters Buh = new BuhParameters();
            public MetaItem GlobalModule = null;
            public MetaItem TagStream = null;
            public List<Guid> GUIDData = new List<Guid>();
            MetaContainer MetadataFileTree = null;
            public bool Isencrypted = false;
            public string MDFileName;
            public string MMS = null;

            private string CompareModules(MetaItem First, MetaItem Second)
            {
                return CompareModules(First.Module, Second.Module);
            }

            private string CompareModules(ProgramModule First, ProgramModule Second, bool GetEditors = false)
            {
                string report = string.Empty;

                Dictionary<string, Sub> Procedures = First.Procedures;
                Dictionary<string, Sub> Procedures_second = Second.Procedures;

                if (First.GlobalVars != null)
                {
                    if (Second.GlobalVars != null)
                    {
                        if (First.GlobalVars.GetHashCode() != Second.GlobalVars.GetHashCode())
                            report += "   Блок глобальных переменных модуля\n";
                    }
                    else
                    {
                        report += "   Добавлен блок глобальных переменных модуля\n";
                    }
                }
                else
                {
                    if (Second.GlobalVars != null)
                    {
                        report += "   Удален блок глобальных переменных модуля\n";
                    }
                }

                if (First.GlobalContext != null)
                {
                    if (Second.GlobalContext != null)
                    {
                        if (First.GlobalContext.GetHashCode() != Second.GlobalContext.GetHashCode())
                            report += "   Блок кода вне процедур\n";
                    }
                    else
                    {
                        report += "   Добавлен блок кода вне процедур\n";
                    }
                }
                else
                {
                    if (Second.GlobalContext != null)
                    {
                        report += "   Удален блок кода вне процедур\n";
                    }
                }



                foreach (KeyValuePair<string, Sub> Procedure in Procedures)
                {
                    if (Procedures_second.ContainsKey(Procedure.Key))
                    {
                        if (Procedure.Value.Body.GetHashCode() != Procedures_second[Procedure.Key].Body.GetHashCode())
                        {
                            report += "    " + Procedure.Key + "()\n";
                            foreach (string Parameter in Procedure.Value.Parameters)
                            {
                                if (!Procedures_second[Procedure.Key].Parameters.Contains(Parameter))
                                {
                                    report += "       Добавлен параметр \"" + Parameter + "\"\n";
                                }
                            }

                            foreach (string Parameter in Procedures_second[Procedure.Key].Parameters)
                            {
                                if (!Procedure.Value.Parameters.Contains(Parameter))
                                {
                                    report += "       Удален параметр \"" + Parameter + "\"\n";
                                }
                            }

                            string editors = null;

                            if (GetEditors)
                                foreach (KeyValuePair<string, int> mod in Procedure.Value.Modifacations)
                                {

                                    if (Procedures_second[Procedure.Key].Modifacations.ContainsKey(mod.Key))
                                    {
                                        //Если количество изменений этого разработчика больше чем в предыдущей версии - значит он автор изменений
                                        if (mod.Value > Procedures_second[Procedure.Key].Modifacations[mod.Key])
                                            editors += mod.Key + ',';
                                    }
                                    else
                                    {
                                        editors += mod.Key + ',';
                                    }
                                }

                            if (GetEditors)
                                foreach (KeyValuePair<string, int> mod in Procedures_second[Procedure.Key].Modifacations)
                                {
                                    if (!Procedure.Value.Modifacations.ContainsKey(mod.Key))
                                    {
                                        editors += mod.Key + ',';
                                    }
                                }

                            if (editors != null)
                                editors = "    " + "    " + "Авторы изменений: " + editors.Substring(0, editors.Length - 1) + "\n";

                            report += editors;
                        }
                    }
                    else
                    {
                        report += "    Добавлена " + Procedure.Key + "()\n";
                        string editors = null;

                        if (GetEditors)
                            foreach (KeyValuePair<string, int> mod in Procedure.Value.Modifacations)
                            {
                                editors += mod.Key + ',';
                            }

                        if (editors != null)
                            editors = "    " + "    " + "Авторы изменений: " + editors.Substring(0, editors.Length - 1) + "\n";

                        report += editors;
                    }
                }

                foreach (KeyValuePair<string, Sub> Procedure in Procedures_second)
                {
                    if (!Procedures.ContainsKey(Procedure.Key))
                    {
                        report += "    Удалена " + Procedure.Key + "()\n";
                    }
                }
                return report;
            }

            public Dictionary<string, Moxel> Moxels
            {
                get
                {
                    Dictionary<string, Moxel> Moxels = new Dictionary<string, Moxel>();

                    foreach (Moxel item in GlobalWorkSheets)
                        Moxels.Add("ОбщиеТаблицы." + item.Description, item);


                    foreach (var item in Subcontos)
                    {
                        if (item.Form != null)
                        if (item.Form.Moxels != null)
                            foreach (Moxel PrintForm in item.Form.Moxels)
                            {
                                Moxels.Add(item.Form.MetaPath + "." + PrintForm.Description, PrintForm);
                            }

                        foreach (var frm in item.ListForms)
                        {
                            if (frm != null)
                            if (frm.Moxels != null)
                                foreach (Moxel PrintForm in frm.Moxels)
                                {
                                    Moxels.Add(PrintForm.MetaPath, PrintForm);
                                }
                        }
                    }

                    foreach (var item in Documents)
                    {
                        if (item.Form.Moxels != null)
                            foreach (Moxel PrintForm in item.Form.Moxels)
                            {
                                Moxels.Add(PrintForm.MetaPath, PrintForm);
                            }
                    }

                    foreach (var item in Journals)
                    {
                        foreach (var frm in item.ListForms)
                        {
                            if (frm.Moxels != null)
                                foreach (Moxel PrintForm in frm.Moxels)
                                {
                                    Moxels.Add(PrintForm.MetaPath, PrintForm);
                                }
                        }
                    }

                    foreach (var item in CalcVars)
                    {
                        if (item.Form.Moxels != null)
                            foreach (Moxel PrintForm in item.Form.Moxels)
                            {
                                Moxels.Add(PrintForm.MetaPath, PrintForm);
                            }
                    }

                    foreach (var item in Reports)
                    {
                        if (item.Form.Moxels != null)
                            foreach (Moxel PrintForm in item.Form.Moxels)
                            {
                                Moxels.Add(PrintForm.MetaPath, PrintForm);
                            }
                    }

                    return Moxels;

                }
            }


            public string CompareWith(TaskItem Second, bool comparemodeules, bool GetEditors = false)
            {
                string report = string.Empty;

                //Сравним структуру метаданных
                if (MMS.GetHashCode() != Second.MMS.GetHashCode())
                    report += "Изменена структура метаданных\n";

                //Сравним глобальники
                #region Глобальный модуль
                if (GlobalModule.CheckSumm != Second.GlobalModule.CheckSumm)
                {
                    report += "Глобальный модуль\n";

                    if (comparemodeules)
                        report += CompareModules(GlobalModule.Module, Second.GlobalModule.Module, GetEditors);

                }
                #endregion

                //Сравним справочники
                #region  справочники
                foreach (var item in Subcontos)
                {
                    var item2 = Second.Subcontos.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        if (item.Form != null)
                        {
                            if (item2.Form.DialogModule.CheckSumm != item.Form.DialogModule.CheckSumm)
                            {
                                report += string.Format("Справочник.{0}.ФормаЭлемента.Модуль\n", item.Identity);

                                if (comparemodeules)
                                    report += CompareModules(item.Form.DialogModule.Module, item2.Form.DialogModule.Module, GetEditors);

                                if (item2.Form.Dialog.CheckSumm != item.Form.Dialog.CheckSumm)
                                    report += string.Format("Справочник.{0}.ФормаЭлемента.Диалог\n", item.Identity);
                            }

                            if (item.Form.Moxels != null)
                                foreach (MetaItem Moxel in item.Form.Moxels)
                                {
                                    MetaItem Moxel2 = item2.Form.Moxels.Find(x => x.Name == Moxel.Name);
                                    if (Moxel2 != null)
                                    {
                                        if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                            report += string.Format("Таблица Справочник.{0}.ФормаЭлемента.{1}\n", item.Identity, Moxel.Description);
                                    }
                                    else
                                    {
                                        report += string.Format("Добавлена таблица Справочник.{0}.ФормаЭлемента{1}\n", item.Identity, Moxel.Description);
                                    }

                                }
                        }

                        if (item.FolderForm != null)
                        {
                            if (item2.FolderForm.DialogModule.CheckSumm != item.FolderForm.DialogModule.CheckSumm)
                            {
                                report += string.Format("Справочник.{0}.ФормаГруппы.Модуль\n", item.Identity);

                                if (comparemodeules)
                                    report += CompareModules(item.FolderForm.DialogModule.Module, item2.FolderForm.DialogModule.Module, GetEditors);

                                if (item2.FolderForm.Dialog.CheckSumm != item.FolderForm.Dialog.CheckSumm)
                                    report += string.Format("Справочник.{0}.ФормаГруппы.Диалог\n", item.Identity);
                            }

                            if (item.FolderForm.Moxels != null)
                                foreach (MetaItem Moxel in item.FolderForm.Moxels)
                                {
                                    MetaItem Moxel2 = item2.FolderForm.Moxels.Find(x => x.Name == Moxel.Name);
                                    if (Moxel2 != null)
                                    {
                                        if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                            report += string.Format("Таблица Справочник.{0}.ФормаГруппы.{1}\n", item.Identity, Moxel.Description);
                                    }
                                    else
                                    {
                                        report += string.Format("Добавлена таблица Справочник.{0}.ФормаГруппы{1}\n", item.Identity, Moxel.Description);
                                    }

                                }
                        }

                        foreach (var frm in item.ListForms)
                        {
                            var frm2 = item2.ListForms.Find(x => x.Identity == frm.Identity);
                            if (frm2 != null)
                            {
                                if (frm.DialogModule != null)
                                {
                                    if (frm.DialogModule.CheckSumm != frm2.DialogModule.CheckSumm)
                                    {
                                        report += string.Format("Справочник.{1}.ФормаСписка.{0}.Модуль\n", frm.Identity, item.Identity);
                                        if (comparemodeules)
                                            report += CompareModules(frm.DialogModule.Module, frm2.DialogModule.Module, GetEditors);

                                        if (frm.Dialog.CheckSumm != frm2.Dialog.CheckSumm)
                                            report += string.Format("Справочник.{1}.ФормаСписка.{0}.Диалог\n", frm.Identity, item.Identity);
                                    }
                                }

                                if (frm.Moxels != null)
                                    foreach (MetaItem Moxel in frm.Moxels)
                                    {
                                        MetaItem Moxel2 = frm2.Moxels.Find(x => x.Name == Moxel.Name);

                                        if (Moxel2 != null)
                                        {
                                            if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                                report += string.Format("Таблица Справочник.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }
                                        else
                                        {
                                            report += string.Format("Добавлена таблица Справочник.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }

                                    }
                            }
                            else
                                report += string.Format("Добавлена форма: Справочник.{1}.ФормаСписка.{0}\n", frm.Identity, item.Identity);
                        }
                    }
                    else
                        report += string.Format("Добавлен: Справочник.{0}\n", item.Identity);
                }
                #endregion

                //Сравним документы
                #region Документы
                foreach (var item in Documents)
                {
                    var item2 = Second.Documents.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        if (item.Form != null)
                        {
                            if (item2.Form.DialogModule.CheckSumm != item.Form.DialogModule.CheckSumm)
                            {
                                report += string.Format("Документ.{0}.Форма.Модуль\n", item.Identity);

                                if (comparemodeules)
                                    report += CompareModules(item.Form.DialogModule.Module, item2.Form.DialogModule.Module, GetEditors);

                                if (item2.Form.Dialog.CheckSumm != item.Form.Dialog.CheckSumm)
                                    report += string.Format("Документ.{0}.Форма.Диалог\n", item.Identity);
                            }

                            if (item.Form.Moxels != null)
                                foreach (MetaItem Moxel in item.Form.Moxels)
                                {
                                    MetaItem Moxel2 = item2.Form.Moxels.Find(x => x.Name == Moxel.Name);
                                    if (Moxel2 != null)
                                    {
                                        if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                            report += string.Format("Таблица Документ.{0}.Форма.{1}\n", item.Identity, Moxel.Description);
                                    }
                                    else
                                    {
                                        report += string.Format("Добавлена таблица Документ.{0}.Форма{1}\n", item.Identity, Moxel.Description);
                                    }

                                }
                        }

                        if (item.TransactionModule != null)
                        {
                            if (item2.TransactionModule.CheckSumm != item.TransactionModule.CheckSumm)
                            {
                                report += string.Format("Документ.{0}.МодульПроведения\n", item.Identity);
                                if (comparemodeules)
                                    report += CompareModules(item.TransactionModule.Module, item2.TransactionModule.Module, GetEditors);
                            }
                        }
                    }
                    else
                        report += string.Format("Добавлен: Документ.{0}\n", item.Identity);
                }
                #endregion

                //Сравним Журналы
                #region Журналы
                foreach (var item in Journals)
                {
                    var item2 = Second.Journals.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        foreach (var frm in item.ListForms)
                        {
                            var frm2 = item2.ListForms.Find(x => x.Identity == frm.Identity);
                            if (frm2 != null)
                            {
                                if (frm.DialogModule.CheckSumm != frm2.DialogModule.CheckSumm)
                                {
                                    report += string.Format("Журнал.{0}.ФормаСписка.{1}.Модуль\n", frm.Identity, item.Identity);

                                    if (comparemodeules)
                                        report += CompareModules(frm.DialogModule.Module, frm2.DialogModule.Module, GetEditors);

                                    if (frm.Dialog.CheckSumm != frm2.Dialog.CheckSumm)
                                        report += string.Format("Журнал.{0}.ФормаСписка.{1}.Диалог\n", frm.Identity, item.Identity);
                                }

                                if (frm.Moxels != null)
                                    foreach (MetaItem Moxel in frm.Moxels)
                                    {
                                        MetaItem Moxel2 = frm2.Moxels.Find(x => x.Name == Moxel.Name);

                                        if (Moxel2 != null)
                                        {
                                            if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                                report += string.Format("Таблица Журнал.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }
                                        else
                                        {
                                            report += string.Format("Добавлена таблица Журнал.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }

                                    }

                            }
                            else
                                report += string.Format("Добавлена форма: Журнал.{0}.ФормаСписка.{1}\n", frm.Identity, item.Identity);
                        }
                    }
                    else
                        report += string.Format("Добавлен: Журнал.{0}\n", item.Identity);
                }
                #endregion

                //Сравним виды расчетов
                #region Отчеты
                foreach (var item in Algorithms)
                {
                    var item2 = Second.Algorithms.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        if (item2.CalculationModule.CheckSumm != item.CalculationModule.CheckSumm)
                        {
                            report += string.Format("ВидРасчета.{0}.МодульРасчета\n", item.Identity);
                            if (comparemodeules)
                                report += CompareModules(item.CalculationModule.Module, item2.CalculationModule.Module, GetEditors);
                        }
                    }
                    else
                        report += string.Format("Добавлен: ВидРасчета.{0}\n", item.Identity);
                }
                #endregion

                //Сравним Журналы Расчетов
                #region Журналы
                foreach (var item in CJ)
                {
                    var item2 = Second.Journals.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        foreach (var frm in item.ListForms)
                        {
                            var frm2 = item2.ListForms.Find(x => x.Identity == frm.Identity);
                            if (frm2 != null)
                            {
                                if (frm.DialogModule.CheckSumm != frm2.DialogModule.CheckSumm)
                                {
                                    report += string.Format("ЖурналРасчетов.{0}.ФормаСписка.{1}.Модуль\n", frm.Identity, item.Identity);

                                    if (comparemodeules)
                                        report += CompareModules(frm.DialogModule.Module, frm2.DialogModule.Module, GetEditors);

                                    if (frm.Dialog.CheckSumm != frm2.Dialog.CheckSumm)
                                        report += string.Format("ЖурналРасчетов.{0}.ФормаСписка.{1}.Диалог\n", frm.Identity, item.Identity);
                                }

                                if (frm.Moxels != null)
                                    foreach (MetaItem Moxel in frm.Moxels)
                                    {
                                        MetaItem Moxel2 = frm2.Moxels.Find(x => x.Name == Moxel.Name);

                                        if (Moxel2 != null)
                                        {
                                            if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                                report += string.Format("Таблица ЖурналРасчетов.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }
                                        else
                                        {
                                            report += string.Format("Добавлена таблица ЖурналРасчетов.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }

                                    }

                            }
                            else
                                report += string.Format("Добавлена форма: ЖурналРасчетов.{0}.ФормаСписка.{1}\n", frm.Identity, item.Identity);
                        }
                    }
                    else
                        report += string.Format("Добавлен: ЖурналРасчетов.{0}\n", item.Identity);
                }
                #endregion

                //Сравним обработки
                #region Обработки
                foreach (var item in CalcVars)
                {
                    var item2 = Second.CalcVars.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        if (item2.Form.DialogModule.CheckSumm != item.Form.DialogModule.CheckSumm)
                        {
                            if (item.Identity != "DefCls")
                            {
                                report += string.Format("Обработка.{0}.Форма.Модуль\n", item.Identity);
                                if (comparemodeules)
                                    report += CompareModules(item.Form.DialogModule.Module, item2.Form.DialogModule.Module, GetEditors);

                                if (item2.Form.Dialog.CheckSumm != item.Form.Dialog.CheckSumm)
                                    report += string.Format("Обработка.{0}.Форма.Диалог\n", item.Identity);

                                if (item.Form.Moxels != null)
                                    foreach (MetaItem Moxel in item.Form.Moxels)
                                    {
                                        MetaItem Moxel2 = item2.Form.Moxels.Find(x => x.Name == Moxel.Name);
                                        if (Moxel2 != null)
                                        {
                                            if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                                report += string.Format("Таблица Обработка.{0}.Форма.{1}\n", item.Identity, Moxel.Description);
                                        }
                                        else
                                        {
                                            report += string.Format("Добавлена таблица Обработка.{0}.Форма{1}\n", item.Identity, Moxel.Description);
                                        }

                                    }
                            }
                            else
                            {
                                string classesChange = string.Empty;

                                string[] splitter = { "\r\n" };
                                string[] classes1 = item.Form.DialogModule.Moduletext.Split(splitter, StringSplitOptions.RemoveEmptyEntries);
                                string[] classes2 = item2.Form.DialogModule.Moduletext.Split(splitter, StringSplitOptions.RemoveEmptyEntries);

                                foreach (string classesitem in classes1)
                                {
                                    if (classesitem.Substring(0, 5) == "//# {")
                                        continue;
                                    string classname = classesitem.Substring(classesitem.LastIndexOf(' ', classesitem.IndexOf('=') - 2), classesitem.IndexOf("=") - classesitem.LastIndexOf(' ', classesitem.IndexOf('=') - 2)).TrimEnd(' ').TrimStart(' ');
                                    if (!item2.Form.DialogModule.Moduletext.Contains(classname))
                                    {
                                        string calcvarname = classesitem.Substring(classesitem.IndexOf("=") + 1, classesitem.IndexOf("@") - classesitem.IndexOf("=") - 1).TrimEnd(' ').TrimStart(' ');
                                        classesChange += "    Добавлен класс " + classname + " (обработка \"" + calcvarname + "\")\n";
                                    }
                                }


                                foreach (string classesitem in classes2)
                                {
                                    if (classesitem.Substring(0, 5) == "//# {")
                                        continue;
                                    string classname = classesitem.Substring(classesitem.LastIndexOf(' ', classesitem.IndexOf('=') - 2), classesitem.IndexOf("=") - classesitem.LastIndexOf(' ', classesitem.IndexOf('=') - 2)).TrimEnd(' ').TrimStart(' ');
                                    if (!item.Form.DialogModule.Moduletext.Contains(classname))
                                        classesChange += "    Удален класс " + classname + "\n";
                                }

                                if (classesChange != string.Empty)
                                {
                                    report += item.Identity + ":\n" + classesChange;
                                }
                            }
                        }
                    }
                    else
                        report += string.Format("Добавлен: Обработка.{0}\n", item.Identity);
                }
                #endregion

                //Сравним отчеты
                #region Отчеты
                foreach (var item in Reports)
                {
                    var item2 = Second.Reports.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        if (item2.Form.DialogModule.CheckSumm != item.Form.DialogModule.CheckSumm)
                        {
                            report += string.Format("Отчет.{0}.Форма.Модуль\n", item.Identity);
                            if (comparemodeules)
                                report += CompareModules(item.Form.DialogModule.Module, item2.Form.DialogModule.Module, GetEditors);
                            if (item2.Form.Dialog.CheckSumm != item.Form.Dialog.CheckSumm)
                                report += string.Format("Отчет.{0}.Форма.Диалог\n", item.Identity);
                        }

                        if (item.Form.Moxels != null)
                            foreach (MetaItem Moxel in item.Form.Moxels)
                            {
                                MetaItem Moxel2 = item2.Form.Moxels.Find(x => x.Name == Moxel.Name);
                                if (Moxel2 != null)
                                {
                                    if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                        report += string.Format("Таблица Отчет.{0}.Форма.{1}\n", item.Identity, Moxel.Description);
                                }
                                else
                                {
                                    report += string.Format("Добавлена таблица Отчет.{0}.Форма{1}\n", item.Identity, Moxel.Description);
                                }

                            }
                    }
                    else
                        report += string.Format("Добавлен: Отчет.{0}\n", item.Identity);
                }
                #endregion

                return report;
            }

            private void LoadFile(string FileName)
            {
                MDFileName = FileName;
                MetadataFileTree = new MetaContainer(FileName);

                if (MetadataFileTree.Items == null)
                    return;

                MMS = MetadataFileTree.GetSubItem("Root\\Metadata\\Main MetaData Stream").Moduletext;
                ParseMMS(MMS);

                GlobalModule = MetadataFileTree.GetSubItem("Root\\TypedText\\ModuleText_Number1\\MD Programm text");
                foreach (MetaItem item in MetadataFileTree.GetSubCatalog("Root\\GlobalData\\GlobalData_Number1\\WorkBook").Items.FindAll(x => x.Type == StorageType.MoxcelWorksheet))
                    GlobalWorkSheets.Add(new Moxel(item, this, "ОбщиеТаблицы"));
               
                TagStream = MetadataFileTree.GetSubItem("Root\\Metadata\\TagStream");
                byte[] GUIDData_raw = MetadataFileTree.GetSubItem("Root\\Metadata\\GUIDData").data;
                byte[] buffer = new byte[16];

                while (GUIDData_raw.Length - (20 + 16 * GUIDData.Count) >= 16)
                {
                    Array.ConstrainedCopy(GUIDData_raw, 20 + 16 * GUIDData.Count, buffer, 0, 16);
                    GUIDData.Add(new Guid(buffer));
                }
            }

            public TaskItem(string FileName)
            {
                LoadFile(FileName);
            }

            private void ParseMMS(string MMS)
            {
                int First = MMS.IndexOf("{\r\n");
                MMS = MMS.Substring(First);
                string errorlog = string.Empty;
                string[] elements = MMS.Replace("\r\n", "").Split('{');
                string[][][] objects = new string[elements.Length][][];
                int level = 0;
                int index = -1;
                int BuhFormsCount = 0;
                int i = 0;
                DescriptorType Type = DescriptorType.Null;
                MetaObject Current = null;
                string Current5 = null;
                foreach (string sub in elements)
                {
                    level++;
                    string[] ss = sub.Split("}".ToCharArray());
                    foreach (string subss in ss)
                    {
                        string[] properties = subss.Split(',');
                        string trytype = properties[0].Replace("\"", "").Replace(" ", "");

                        bool toplevel = level == 3;

                        if (toplevel)
                        {
                            if (trytype != "")
                            {
                                if (Enum.TryParse<DescriptorType>(trytype, out Type))
                                {
                                    index++;
                                    objects[index] = new string[3][];
                                    objects[index][0] = new string[1];
                                    objects[index][0][0] = Type.GetDescription();
                                    objects[index][1] = properties;
                                    objects[index][2] = new string[1];
                                    i = 0;
                                }
                                else
                                {
                                    errorlog += "DescriptorType." + trytype + ",\r\n";
                                }

                            }
                        }

                        string[] ObjParams = subss.Replace("\",\"", "|").Replace("\"", "").Split('|');

                        if (level == 4)
                        {
                            if (objects[index][2].Length <= i + 1)
                                Array.Resize<string>(ref objects[index][2], i + 1);
                            objects[index][2][i] = subss;
                            i++;
                            if (ObjParams.Length >= 4)
                            {
                                if (Type == DescriptorType.SbCnts)
                                {
                                    Subconto SB = new Subconto(ObjParams);
                                    MetaContainer FormContainer;

                                    if (SB.EditMode != 0) //Способ редактирования не "В списке"
                                    {
                                        FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\Subconto\\Subconto_Number{0}\\WorkBook", SB.ID));
                                        if (FormContainer != null)
                                            SB.Form = new Form(FormContainer, SB.MetaPath + ".ФормаЭлемента");
                                    }
                                    if (SB.OneForm != 1)
                                        if (SB.LevelCount > 1)
                                        {
                                            FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\SubFolder\\SubFolder_Number{0}\\WorkBook", SB.ID));
                                            if (FormContainer != null)
                                                SB.FolderForm = new Form(FormContainer, SB.MetaPath + ".ФормаГруппы");
                                        }

                                    Subcontos.Add(SB);
                                    Current = SB;
                                }

                                if (Type == DescriptorType.Documents)
                                {
                                    Document Doc = new Document();
                                    Doc.ID = int.Parse(ObjParams[0]);
                                    Doc.Identity = ObjParams[1];
                                    Doc.Alias = ObjParams[2];
                                    Doc.MetaPath = "Документ." + Doc.Identity;
                                    MetaContainer FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\Document\\Document_Number{0}\\WorkBook", Doc.ID));
                                    if (FormContainer != null)
                                        Doc.Form = new Form(FormContainer, Doc.MetaPath + ".Форма");
                                    Doc.TransactionModule = MetadataFileTree.GetSubItem(String.Format("Root\\TypedText\\Transact_Number{0}\\MD Programm text", Doc.ID));
                                    Documents.Add(Doc);
                                    Current = Doc;
                                }

                                if (Type == DescriptorType.Buh)
                                {
                                    Buh = new BuhParameters();
                                    Buh.ID = int.Parse(ObjParams[0]);
                                    Buh.Identity = ObjParams[1];
                                    Buh.Alias = ObjParams[2];

                                    Document Doc = Documents.Find(x => x.Identity == "Операция");
                                    MetaContainer FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\Operation\\Operation_Number{0}\\WorkBook", ObjParams[0]));
                                    if (FormContainer != null)
                                        Doc.Form = new Form(FormContainer, Doc.MetaPath + ".Форма");

                                    FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\AccountChart\\AccountChart_Number{0}\\WorkBook", Buh.ID));

                                    if (FormContainer != null)
                                        Buh.AccountChart.Add(new Form(FormContainer));

                                    Current = Doc;
                                }

                                if (Type == DescriptorType.Journalisters)
                                {
                                    Journal obj = new Journal();
                                    obj.ID = int.Parse(ObjParams[0]);
                                    obj.Identity = ObjParams[1];
                                    obj.Alias = ObjParams[2];
                                    obj.MetaPath = "ЖурналДокументов." + obj.Identity;
                                    Journals.Add(obj);
                                    Current = obj;
                                }

                                if (Type == DescriptorType.CJ)
                                {
                                    CalcJournal obj = new CalcJournal();
                                    obj.ID = int.Parse(ObjParams[0]);
                                    obj.Identity = ObjParams[1];
                                    obj.Alias = ObjParams[2];
                                    obj.MetaPath = "ЖурналРасчетов." + obj.Identity;
                                    CJ.Add(obj);
                                    Current = obj;
                                }

                                if (Type == DescriptorType.ReportList)
                                {
                                    ReportItem obj = new ReportItem();

                                    obj.ID = int.Parse(ObjParams[0]);
                                    obj.Identity = ObjParams[1];
                                    obj.Alias = ObjParams[2];
                                    obj.MetaPath = "Отчет." + obj.Identity;
                                    obj.Form = new Form(MetadataFileTree.GetSubCatalog(String.Format("Root\\Report\\Report_Number{0}\\WorkBook", obj.ID)), obj.MetaPath+".Форма");
                                    Reports.Add(obj);
                                    Current = obj;
                                }

                                if (Type == DescriptorType.CalcVars)
                                {
                                    CalcVarItem obj = new CalcVarItem();

                                    obj.ID = int.Parse(ObjParams[0]);
                                    obj.Identity = ObjParams[1];
                                    obj.Alias = ObjParams[2];
                                    obj.MetaPath = "Обработка." + obj.Identity;
                                    obj.Form = new Form(MetadataFileTree.GetSubCatalog(String.Format("Root\\CalcVar\\CalcVar_Number{0}\\WorkBook", obj.ID)), obj.MetaPath + ".Форма");
                                    CalcVars.Add(obj);
                                    Current = obj;
                                }

                                if (Type == DescriptorType.Algorithms)
                                {
                                    CalculationAlgorithm obj = new CalculationAlgorithm();

                                    obj.ID = int.Parse(ObjParams[0]);
                                    obj.Identity = ObjParams[1];
                                    obj.Alias = ObjParams[2];
                                    obj.MetaPath = "ВидРасчета." + obj.Identity;
                                    obj.CalculationModule = MetadataFileTree.GetSubItem(String.Format("Root\\TypedText\\CalcAlg_Number{0}\\MD Programm text", obj.ID));
                                    Algorithms.Add(obj);
                                    Current = obj;
                                }
                            }
                        }

                        if (level == 5)
                        {
                            Current5 = trytype;

                            if (Type == DescriptorType.Buh && Current5 == "Form")
                                BuhFormsCount++;
                        }

                        if (level == 7)
                        {
                            Current5 = trytype;
                            if (Type == DescriptorType.Buh && trytype == "Form")
                                BuhFormsCount++;
                        }

                        if ((level == 6 || level == 8) && Current5 == "Form" && ObjParams.Length == 4)
                        {
                            MetaContainer FormContainer = null;

                            if (Type == DescriptorType.SbCnts)
                                FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\SubList\\SubList_Number{0}\\WorkBook", ObjParams[0]));

                            if (Type == DescriptorType.Journalisters)
                                FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\Journal\\Journal_Number{0}\\WorkBook", ObjParams[0]));

                            if (Type == DescriptorType.CJ)
                                FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\CalcJournal\\CalcJournal_Number{0}\\WorkBook", ObjParams[0]));

                            if (Type == DescriptorType.Buh)
                            {
                                switch (BuhFormsCount)
                                {
                                    case 1:            //Форма списка плана счетов
                                        FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\AccountChartList\\AccountChartList_Number{0}\\WorkBook", ObjParams[0]));
                                        break;
                                    case 2:            //ХЗ
                                        FormContainer = null;//MetadataFileTree.GetSubCatalog(String.Format("Root\\AccountChart\\AccountChart_Number{0}\\WorkBook", Buh.ID));
                                        break;
                                    case 3:            //Форма списка проводок
                                        FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\ProvList\\ProvList_Number{0}\\WorkBook", ObjParams[0]));
                                        break;
                                    case 4:            //Форма списка операции
                                        FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\OperationList\\OperationList_Number{0}\\WorkBook", ObjParams[0]));
                                        break;
                                }
                            }

                            Form frm = new Form();

                            if (FormContainer != null)
                                frm = new Form(FormContainer, Current.MetaPath + ".ФормаСписка." + ObjParams[1]);

                            frm.ID = int.Parse(ObjParams[0]);
                            frm.Identity = ObjParams[1];
                            frm.Alias = ObjParams[2];
                            frm.Description = ObjParams[3];

                            if (Type == DescriptorType.SbCnts)
                                ((Subconto)Current).ListForms.Add(frm);

                            if (Type == DescriptorType.Journalisters)
                                ((Journal)Current).ListForms.Add(frm);

                            if (Type == DescriptorType.CJ)
                                ((CalcJournal)Current).ListForms.Add(frm);

                            if (Type == DescriptorType.Buh)
                            {
                                switch (BuhFormsCount)
                                {
                                    case 1:            //Форма списка плана счетов
                                        Buh.AccountChartList.Add(frm);
                                        break;
                                    case 2:            //ХЗ
                                        Buh.AccountChart.Add(frm);
                                        break;
                                    case 3:            //Форма списка проводок
                                        Buh.ProvListList.Add(frm);
                                        break;
                                    case 4:            //Форма списка операции
                                        Buh.OperationList.Add(frm);
                                        break;
                                }
                            }

                        }
                    }
                    level -= ss.Length - 1;
                }

                Array.Resize<string[][]>(ref objects, index + 1);
                if (errorlog.Length > 1)
                    File.WriteAllText(MDFileName + ".errorlog", errorlog);
            }
        }

        public class MetaData
        {
            MetaObject MainDataContDef;

        }
        #endregion

        #region Интерфейсы OLE
        //[ComImport]
        //[Guid("0000000d-0000-0000-C000-000000000046")]
        //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        //interface IEnumSTATSTG
        //{
        //    // The user needs to allocate an STATSTG array whose size is celt.
        //    [PreserveSig]
        //    uint Next(
        //        uint celt,
        //        [MarshalAs(UnmanagedType.LPArray), Out]
        //            System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt,
        //        out uint pceltFetched
        //    );
        //    void Skip(uint celt);
        //    void Reset();
        //    [return: MarshalAs(UnmanagedType.Interface)]
        //    IEnumSTATSTG Clone();
        //}

        //[ComImport]
        //[Guid("0000000b-0000-0000-C000-000000000046")]
        //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        //interface IStorage
        //{
        //    void CreateStream(
        //        /* [string][in] */ string pwcsName,
        //        /* [in] */ uint grfMode,
        //        /* [in] */ uint reserved1,
        //        /* [in] */ uint reserved2,
        //        /* [out] */ out IStream ppstm);

        //    void OpenStream(
        //        /* [string][in] */ string pwcsName,
        //        /* [unique][in] */ IntPtr reserved1,
        //        /* [in] */ uint grfMode,
        //        /* [in] */ uint reserved2,
        //        /* [out] */ out IStream ppstm);

        //    void CreateStorage(
        //        /* [string][in] */ string pwcsName,
        //        /* [in] */ uint grfMode,
        //        /* [in] */ uint reserved1,
        //        /* [in] */ uint reserved2,
        //        /* [out] */ out IStorage ppstg);

        //    void OpenStorage(
        //        /* [string][unique][in] */ string pwcsName,
        //        /* [unique][in] */ IStorage pstgPriority,
        //        /* [in] */ uint grfMode,
        //        /* [unique][in] */ IntPtr snbExclude,
        //        /* [in] */ uint reserved,
        //        /* [out] */ out IStorage ppstg);

        //    void CopyTo(
        //        /* [in] */ uint ciidExclude,
        //        /* [size_is][unique][in] */ Guid rgiidExclude, // should this be an array?
        //        /* [unique][in] */ IntPtr snbExclude,
        //        /* [unique][in] */ IStorage pstgDest);

        //    void MoveElementTo(
        //        /* [string][in] */ string pwcsName,
        //        /* [unique][in] */ IStorage pstgDest,
        //        /* [string][in] */ string pwcsNewName,
        //        /* [in] */ uint grfFlags);

        //    void Commit(
        //        /* [in] */ STGC grfCommitFlags);

        //    void Revert();

        //    void EnumElements(
        //        /* [in] */ uint reserved1,
        //        /* [size_is][unique][in] */ IntPtr reserved2,
        //        /* [in] */ uint reserved3,
        //        /* [out] */ out IEnumSTATSTG ppenum);

        //    void DestroyElement(
        //        /* [string][in] */ string pwcsName);

        //    void RenameElement(
        //        /* [string][in] */ string pwcsOldName,
        //        /* [string][in] */ string pwcsNewName);

        //    void SetElementTimes(
        //        /* [string][unique][in] */ string pwcsName,
        //        /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pctime,
        //        /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME patime,
        //        /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

        //    void SetClass(
        //        /* [in] */ Guid clsid);

        //    void SetStateBits(
        //        /* [in] */ uint grfStateBits,
        //        /* [in] */ uint grfMask);

        //    void Stat(
        //        /* [out] */ out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg,
        //        /* [in] */ uint grfStatFlag);

        //}

        //[Flags]
        //private enum STGC : int
        //{
        //    DEFAULT = 0,
        //    OVERWRITE = 1,
        //    ONLYIFCURRENT = 2,
        //    DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
        //    CONSOLIDATE = 8
        //}

        //[Flags]
        //private enum STGM : int
        //{
        //    DIRECT = 0x00000000,
        //    TRANSACTED = 0x00010000,
        //    SIMPLE = 0x08000000,
        //    READ = 0x00000000,
        //    WRITE = 0x00000001,
        //    READWRITE = 0x00000002,
        //    SHARE_DENY_NONE = 0x00000040,
        //    SHARE_DENY_READ = 0x00000030,
        //    SHARE_DENY_WRITE = 0x00000020,
        //    SHARE_EXCLUSIVE = 0x00000010,
        //    PRIORITY = 0x00040000,
        //    DELETEONRELEASE = 0x04000000,
        //    NOSCRATCH = 0x00100000,
        //    CREATE = 0x00001000,
        //    CONVERT = 0x00020000,
        //    FAILIFTHERE = 0x00000000,
        //    NOSNAPSHOT = 0x00200000,
        //    DIRECT_SWMR = 0x00400000,
        //}

        //[Flags]
        //private enum STATFLAG : uint
        //{
        //    STATFLAG_DEFAULT = 0,
        //    STATFLAG_NONAME = 1,
        //    STATFLAG_NOOPEN = 2
        //}

        //[Flags]
        //private enum STGTY : int
        //{
        //    STGTY_STORAGE = 1,
        //    STGTY_STREAM = 2,
        //    STGTY_LOCKBYTES = 3,
        //    STGTY_PROPERTY = 4
        //}

        ////Читает IStream в массив байт
        //private static byte[] ReadIStream(IStream pIStream)
        //{
        //    System.Runtime.InteropServices.ComTypes.STATSTG StreamInfo;
        //    pIStream.Stat(out StreamInfo, 0);
        //    byte[] data = new byte[StreamInfo.cbSize];
        //    pIStream.Read(data, (int)StreamInfo.cbSize, IntPtr.Zero);
        //    return data;
        //}


        //[ComVisible(false)]
        //[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000A-0000-0000-C000-000000000046")]
        //public interface ILockBytes
        //{
        //    void ReadAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbRead);
        //    void WriteAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbWritten);
        //    void Flush();
        //    void SetSize(long cb);
        //    void LockRegion(long libOffset, long cb, int dwLockType);
        //    void UnlockRegion(long libOffset, long cb, int dwLockType);
        //    void Stat(out System.Runtime.InteropServices.STATSTG pstatstg, int grfStatFlag);

        //}

        static class NativeMethods
        {
            [ComImport]
            [Guid("0000000d-0000-0000-C000-000000000046")]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IEnumSTATSTG
            {
                // The user needs to allocate an STATSTG array whose size is celt.
                [PreserveSig]
                uint Next(
                    uint celt,
                    [MarshalAs(UnmanagedType.LPArray), Out]
                    System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt,
                    out uint pceltFetched
                );
                void Skip(uint celt);
                void Reset();
                [return: MarshalAs(UnmanagedType.Interface)]
                IEnumSTATSTG Clone();
            }

            [ComImport]
            [Guid("0000000b-0000-0000-C000-000000000046")]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IStorage
            {
                void CreateStream(
                    /* [string][in]*/  string pwcsName,
                    /* [in]*/  STGM grfMode,
                    /* [in]*/  uint reserved1,
                    /* [in] */ uint reserved2,
                    /* [out] */ out IStream ppstm);

                void OpenStream(
                    /* [string][in] */ string pwcsName,
                    /* [unique][in] */ IntPtr reserved1,
                    /* [in] */ STGM grfMode,
                    /* [in] */ uint reserved2,
                    /* [out] */ out IStream ppstm);

                void CreateStorage(
                    /* [string][in] */ string pwcsName,
                    /* [in] */ STGM grfMode,
                    /* [in] */ uint reserved1,
                    /* [in] */ uint reserved2,
                    /* [out] */ out IStorage ppstg);

                void OpenStorage(
                    /* [string][unique][in] */ string pwcsName,
                    /* [unique][in] */ IStorage pstgPriority,
                    /* [in] */ STGM grfMode,
                    /* [unique][in] */ IntPtr snbExclude,
                    /* [in] */ uint reserved,
                    /* [out] */ out IStorage ppstg);

                void CopyTo(
                    /* [in] */ uint ciidExclude,
                    /* [size_is][unique][in] */ ref Guid rgiidExclude, // should this be an array?
                    /* [unique][in] */ IntPtr snbExclude,
                    /* [unique][in] */ IStorage pstgDest);

                void MoveElementTo(
                    /* [string][in] */ string pwcsName,
                    /* [unique][in] */ IStorage pstgDest,
                    /* [string][in] */ string pwcsNewName,
                    /* [in] */ uint grfFlags);

                void Commit(
                    /* [in] */ STGC grfCommitFlags);

                void Revert();

                void EnumElements(
                    /* [in] */ uint reserved1,
                    /* [size_is][unique][in] */ IntPtr reserved2,
                    /* [in] */ uint reserved3,
                    /* [out] */out IEnumSTATSTG ppenum);

                void DestroyElement(
                    /* [string][in] */ string pwcsName);

                void RenameElement(
                    /* [string][in] */ string pwcsOldName,
                    /* [string][in] */ string pwcsNewName);

                void SetElementTimes(
                    /* [string][unique][in] */ string pwcsName,
                    /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pctime,
                    /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME patime,
                    /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

                void SetClass(
                    /* [in] */ ref Guid clsid);

                void SetStateBits(
                    /* [in] */ uint grfStateBits,
                    /* [in] */ uint grfMask);

                void Stat(
                    /* [out] */ out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg,
                    /* [in] */ uint grfStatFlag);

            }

            [Flags]
            public enum STGC : int
            {
                DEFAULT = 0,
                OVERWRITE = 1,
                ONLYIFCURRENT = 2,
                DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
                CONSOLIDATE = 8
            }

            [Flags]
            public enum STGM : int
            {
                DIRECT = 0x00000000,
                TRANSACTED = 0x00010000,
                SIMPLE = 0x08000000,
                READ = 0x00000000,
                WRITE = 0x00000001,
                READWRITE = 0x00000002,
                SHARE_DENY_NONE = 0x00000040,
                SHARE_DENY_READ = 0x00000030,
                SHARE_DENY_WRITE = 0x00000020,
                SHARE_EXCLUSIVE = 0x00000010,
                PRIORITY = 0x00040000,
                DELETEONRELEASE = 0x04000000,
                NOSCRATCH = 0x00100000,
                CREATE = 0x00001000,
                CONVERT = 0x00020000,
                FAILIFTHERE = 0x00000000,
                NOSNAPSHOT = 0x00200000,
                DIRECT_SWMR = 0x00400000,
            }

            [Flags]
            public enum STATFLAG : uint
            {
                STATFLAG_DEFAULT = 0,
                STATFLAG_NONAME = 1,
                STATFLAG_NOOPEN = 2
            }

            [Flags]
            public enum STGTY : int
            {
                STGTY_STORAGE = 1,
                STGTY_STREAM = 2,
                STGTY_LOCKBYTES = 3,
                STGTY_PROPERTY = 4
            }

            //Читает IStream в массив байт
            public static byte[] ReadIStream(IStream pIStream)
            {
                System.Runtime.InteropServices.ComTypes.STATSTG StreamInfo;
                pIStream.Stat(out StreamInfo, 0);
                byte[] data = new byte[StreamInfo.cbSize];
                pIStream.Read(data, (int)StreamInfo.cbSize, IntPtr.Zero);
                Marshal.ReleaseComObject(pIStream);
                Marshal.FinalReleaseComObject(pIStream);
                return data;
            }

            //Читает IStream в массив байт и разжимает
            public static byte[] ReadCompressedIStream(IStream pIStream)
            {
                System.Runtime.InteropServices.ComTypes.STATSTG StreamInfo;
                pIStream.Stat(out StreamInfo, 0);
                byte[] data = new byte[StreamInfo.cbSize];
                pIStream.Read(data, (int)StreamInfo.cbSize, IntPtr.Zero);
                Marshal.ReleaseComObject(pIStream);
                Marshal.FinalReleaseComObject(pIStream);
                DeflateStream ZLibCompressed = new DeflateStream(new MemoryStream(data), CompressionMode.Decompress, false);
                MemoryStream Decompressed = new MemoryStream();
                ZLibCompressed.CopyTo(Decompressed);
                ZLibCompressed.Dispose();
                data = Decompressed.ToArray();
                Array.Resize(ref data, (int)Decompressed.Length);
                return data;
            }

            //Читает IStream в строку
            public static string ReadIStreamToString(IStream pIStream)
            {
                System.Runtime.InteropServices.ComTypes.STATSTG StreamInfo;
                pIStream.Stat(out StreamInfo, 0);
                byte[] data = new byte[StreamInfo.cbSize];
                pIStream.Read(data, (int)StreamInfo.cbSize, IntPtr.Zero);
                Marshal.ReleaseComObject(pIStream);
                Marshal.FinalReleaseComObject(pIStream);
                return Encoding.GetEncoding(1251).GetString(data);
            }

            [ComVisible(false)]
            [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000A-0000-0000-C000-000000000046")]
            public interface ILockBytes
            {
                void ReadAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbRead);
                void WriteAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbWritten);
                void Flush();
                void SetSize(long cb);
                void LockRegion(long libOffset, long cb, int dwLockType);
                void UnlockRegion(long libOffset, long cb, int dwLockType);
                void Stat(out System.Runtime.InteropServices.STATSTG pstatstg, int grfStatFlag);
            }


            [DllImport("ole32.dll")]
            public static extern int StgIsStorageFile(
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsName);

            [DllImport("ole32.dll")]
            public static extern int StgOpenStorage(
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
                IStorage pstgPriority,
                STGM grfMode,
                IntPtr snbExclude,
                uint reserved,
                out IStorage ppstgOpen);

            [DllImport("ole32.dll")]
            public static extern int StgCreateDocfile(
                [MarshalAs(UnmanagedType.LPWStr)]string pwcsName,
                STGM grfMode,
                uint reserved,
                out IStorage ppstgOpen);

            [DllImport("ole32.dll")]
            public static extern int StgOpenStorageOnILockBytes(ILockBytes plkbyt,
               IStorage pStgPriority, STGM grfMode, IntPtr snbEnclude, uint reserved,
               out IStorage ppstgOpen);

            [DllImport("ole32.dll")]
            public static extern int CreateILockBytesOnHGlobal(IntPtr hGlobal, [MarshalAs(UnmanagedType.Bool)] bool fDeleteOnRelease, out ILockBytes ppLkbyt);
        }

        #endregion
    }
}

namespace UserDef
{
    public class UserDefworks
    {

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool LockFile(IntPtr hFile, int dwFileOffsetLow, int dwFileOffsetHigh, int nNumberOfBytesToLockLow, int nNumberOfBytesToLockHigh);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool UnlockFile(IntPtr hFile, int dwFileOffsetLow, int dwFileOffsetHigh, int nNumberOfBytesToLockLow, int nNumberOfBytesToLockHigh);

        public enum UserParameters
        {
            Header = 0,
            DontCheckRights,
            PasswordHash,
            FullName,
            UserCatalog,
            RightsEnabled,
            UserInterface,
            UserRights
        }

        public static string GetDBCatalog(string filename)
        {
            string DBcatalog = filename;

            if (Path.GetFileName(DBcatalog) != "")
                DBcatalog = Path.GetDirectoryName(DBcatalog);
            
            if (DBcatalog == "")
                return string.Empty;

            while (!File.Exists(DBcatalog + "\\1cv7.md") && DBcatalog != null)
            {
                DBcatalog = Path.GetDirectoryName(DBcatalog);
            }

            if (DBcatalog == null)
                return string.Empty;

            if (!DBcatalog.EndsWith("\\", StringComparison.OrdinalIgnoreCase))
                DBcatalog += "\\";

            if (!File.Exists(DBcatalog + "1cv7.md"))
                DBcatalog = string.Empty;

            return DBcatalog;
        }

        public static bool IsSQLDatabase(string DBcatalog)
        {
            bool IsSQL = File.Exists(DBcatalog + "1cv7.dba");
            return IsSQL;
        }

        public static bool IsConfigRunning(string DBcatalog)
        {
            bool IsConfigRunning = false;

            if (IsSQLDatabase(DBcatalog))
            {
                // проверим, можно ли перезаписать users.usr
                string fLockFile = DBcatalog + "1cv7.lck";
                if (File.Exists(fLockFile))
                {
                    //Откроем 1cv7.lck в разделенном режиме
                    FileStream f = File.Open(fLockFile, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);

                    //Конфигуратор при открытии блокирует первые 100000 байт файла 1cv7.lck в корне базы.
                    //Если не удастся их заблокировать, значит конфигуратор запущен
                    IsConfigRunning = !LockFile(f.Handle, 0, 0, 100000, 0);

                    //не забываем разблокировать и закрыть файл, иначе получим блокировку на запуск конфигуратора.
                    UnlockFile(f.Handle, 0, 0, 10000000, 0);
                    f.Close();
                    f.Dispose();

                }
            }
            return IsConfigRunning;
        }

        public static string GetStringHash(string instr)
        {
            if (instr.Length == 0)
                return "233"; //1С воспринимает это как хэш пустой строки

            string strHash = string.Empty;

            //1C принимате только пароли не больше 10 симвлов длиной и в верхнем регистре.
            instr = instr.Substring(0, Math.Min(10, instr.Length)).ToUpper();

            foreach (byte b in new MD5CryptoServiceProvider().ComputeHash(Encoding.Default.GetBytes(instr)))
            {
                strHash += b.ToString("X2");
            }
            return strHash;
        }

        public static Dictionary<UserParameters, string> UserParamNames = new Dictionary<UserParameters, string>()
        {
            {UserParameters.Header,"Заголовок"},
            {UserParameters.DontCheckRights,"Отключить контроль прав"},
            {UserParameters.PasswordHash,"Хэш пароля"},
            {UserParameters.FullName,"Полное имя"},
            {UserParameters.UserCatalog,"Каталог пользователя"},
            {UserParameters.RightsEnabled,"Заданы права"},
            {UserParameters.UserInterface,"Интерфейс"},
            {UserParameters.UserRights,"Набор прав"}
        };

        // алгоритм подсчета CheckSum - представить файл в виде массива DWORD и сложить все элементы.
        public static string CheckSum(string USRfileName)
        {
            if (!File.Exists(USRfileName))
                return "00000000";
            byte[] Users_usr = { 0 };
            try
            {
                Users_usr = File.ReadAllBytes(USRfileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Файл {0} заблокирован. Невозможно вычислить \"Checksum\" {1}", USRfileName, ex.Message);
                return "00000000";
            }
            UInt32 crc = 0;
            UInt32 x = 0;
            int k = Users_usr.Length % 4;

            if (k > 0)
            {
                Array.Resize(ref Users_usr, Users_usr.Length + k);
            }

            for (int i = 0; i < Users_usr.Length; i += 4)
            {
                x = BitConverter.ToUInt32(Users_usr, i);
                if (Math.Abs(crc + x) > Int32.MaxValue)
                    crc += (UInt32)(x - Math.Sign(crc + x) * 4294967296);
                else
                    crc += x;
            }
            return string.Format("\"{0}\"", Convert.ToString(crc, 16));
        }
        //Расксорить 1cv7.DBA и вернуть строку с параметрами подключения к БД
        public static string ReadDBA(string DBAfileName)
        {
            byte[] ByteBuffer = File.ReadAllBytes(DBAfileName);
            byte[] SQLKey = Encoding.ASCII.GetBytes("19465912879oiuxc ensdfaiuo3i73798kjl");
            for (int i = 0; i < ByteBuffer.Length; i++)
            {
                ByteBuffer[i] = (byte)(ByteBuffer[i] ^ SQLKey[i % 36]);
            }
            return Encoding.ASCII.GetString(ByteBuffer);
        }

        //Заксорить и записать параметры подключения к БД в 1cv7.DBA
        public static bool WriteDBA(string DBAfileName, string Connect)
        {
            byte[] ByteBuffer = Encoding.ASCII.GetBytes(Connect);
            byte[] SQLKey = Encoding.ASCII.GetBytes("19465912879oiuxc ensdfaiuo3i73798kjl");

            for (int i = 0; i < ByteBuffer.Length; i++)
            {
                ByteBuffer[i] = (byte)(ByteBuffer[i] ^ SQLKey[i % 36]);
            }
            try
            {
                File.WriteAllBytes(DBAfileName, ByteBuffer);
                return true;
            }
            catch
            {
                return false;
            }
        }

        //Класс описвает объект потока пользователя UserItem
        [Serializable]
        public class UserItem
        {
            //Класс описывает строку в формате Pascal - массив байт в первом элементе длина, остальные - значение
            [Serializable]
            public class PascalString
            {
                public byte Length;
                public byte[] Value;

                //Создание из строки
                public PascalString(string InStr)
                {
                    Value = Encoding.Default.GetBytes(InStr);
                    Length = (byte)Value.Length;
                }

                //Создание из строки
                public PascalString(byte[] InStr, int SourceIndex = 0)
                {
                    if (SourceIndex >= InStr.Length)
                        return;
                    Length = InStr[SourceIndex];
                    if (SourceIndex + 1 + Length >= InStr.Length)
                        Length = (byte)(InStr.Length - 1 - SourceIndex);
                    Value = new byte[Length];
                    Array.Copy(InStr, SourceIndex + 1, Value, 0, Length);
                }

                // Для удобства зададим неявное преобразование из строки (используется при присваиваниии)
                public static implicit operator PascalString(string InStr)
                {
                    return new PascalString(InStr);
                }

                // Для удобства зададим неявное преобразование из массива байт (используется при присваиваниии)
                public static implicit operator PascalString(byte[] InStr)
                {
                    return new PascalString(InStr);
                }

                public static implicit operator byte[](PascalString InStr)
                {
                    return InStr.Serialize();
                }

                // для удобства - преобразование в обычную строку
                override public string ToString()
                {
                    return Encoding.Default.GetString(Value);
                }

                public static implicit operator string(PascalString InStr)
                {
                    return InStr.ToString();
                }

                // Заполнение из строки
                public void FromString(string InStr)
                {
                    Value = Encoding.Default.GetBytes(InStr);
                    Length = (byte)(Value.Length - 1);
                }

                // возвращает массив байт в нужном формате
                public byte[] Serialize()
                {
                    byte[] ByteBuffer = new byte[Length + 1];
                    ByteBuffer[0] = Length;
                    for (int i = 1; i < ByteBuffer.Length; i++)
                        ByteBuffer[i] = Value[i - 1];

                    return ByteBuffer;
                }

                public PascalString Deserialize(byte[] data)
                {
                    return data;
                }

                public byte[] GetObjectData()
                {
                    return Serialize();
                }
            }

            public string Name;
            public string PageName;
            public int CheckRights = 1;
            public PascalString PasswordHash;
            public PascalString FullName;
            public PascalString UserCatalog;
            public int RightsEnabled = 1;
            public PascalString UserInterface;
            public PascalString UserRights;
            public bool modified = false;

            //Возвращаяет позицию указаного параметра в массиве байт потока пользователя
            // 0 - пустой параметр, заголовк записи. всегда = 1
            // 1 - контролировать права число 1/0
            // 2 - хэш пароля, длина всегда либо 32 либо 3 если пароль не задан
            // 3 - Полное имя пользователя
            // 4 - каталог
            // 5 - флаг наличия прав, число 1/0
            // 6 - интерфейс
            // 7 - набор прав
            static int GetPos(byte[] data, UserParameters Param)
            {
                int StartPosition = 0;
                int Count = 0;
                while (Count != (int)Param)
                {
                    if (StartPosition >= data.Length)
                        return 0;

                    if (data[StartPosition] < 2 && (Count == 0 || Count == 1 || Count == 5)) //булевы параметры длина = 4
                    {
                        StartPosition += 4;
                    }
                    else //строковые параметры длина в первом байте
                    {
                        StartPosition += data[StartPosition] + 1;
                    }
                    Count++;
                }
                return StartPosition;
            }

            //Возвращает значение переданного параметра из массива байт
            static dynamic ParseByteArray(byte[] data, UserParameters paramNuber)
            {
                dynamic param = null;

                int paramstart = GetPos(data, paramNuber);
                switch (paramNuber)
                {
                    case UserParameters.RightsEnabled:
                    case UserParameters.Header:
                        {
                            param = (int)(data[paramstart]);
                            break;
                        }
                    case UserParameters.DontCheckRights:
                        {
                            param = (int)(1 - data[paramstart]);
                            break;
                        }
                    default:
                        {
                            param = new PascalString(data, paramstart);
                            break;
                        }
                }
                return param;
            }

            public dynamic this[UserParameters param]
            {
                get
                {
                    switch (param)
                    {
                        case UserParameters.Header:
                            return 1;
                        case UserParameters.DontCheckRights:
                            return CheckRights;
                        case UserParameters.RightsEnabled:
                            return RightsEnabled;
                        case UserParameters.PasswordHash:
                            return PasswordHash;
                        case UserParameters.FullName:
                            return FullName;
                        case UserParameters.UserCatalog:
                            return UserCatalog;
                        case UserParameters.UserInterface:
                            return UserInterface;
                        case UserParameters.UserRights:
                            return UserRights;
                        default:
                            return null;
                    }
                }
                set
                {
                    modified = true;
                    switch (param)
                    {
                        case UserParameters.DontCheckRights:
                            { CheckRights = value; break; }
                        case UserParameters.RightsEnabled:
                            { RightsEnabled = value; break; }
                        case UserParameters.PasswordHash:
                            { PasswordHash = value; break; }
                        case UserParameters.FullName:
                            { FullName = value; break; }
                        case UserParameters.UserCatalog:
                            { UserCatalog = value; break; }
                        case UserParameters.UserInterface:
                            { UserInterface = value; break; }
                        case UserParameters.UserRights:
                            { UserRights = value; break; }
                        default:
                            break;
                    }

                }
            }

            public void SetParam(UserParameters param, dynamic value)
            {
                this[param] = value;
            }

            public dynamic GetParam(UserParameters param)
            {
                return this[param];
            }

            //Созадает структуру из массива байт
            public UserItem(byte[] data, string Name = "", string PageName = "")
            {
                this.Name = Name;
                this.PageName = PageName;
                for (UserParameters param = UserParameters.Header; param <= UserParameters.UserRights; param++)
                {
                    SetParam(param, ParseByteArray(data, param));
                }
                this.CheckRights = 1;
                modified = false;
            }

            //создает структуру из массива байт при присваивании
            public static implicit operator UserItem(byte[] data)
            {
                return new UserItem(data);
            }

            //Созадает структуру из набора параметров
            public UserItem(string PageName,
                            int _CheckRights = 1,
                            string HashCode = "233",
                            string FullName = "",
                            string UserCatalog = "",
                            string Interface = "",
                            string Rights = "",
                            string Name = "")
            {
                CheckRights = _CheckRights;
                RightsEnabled = 1;

                this.Name = Name;
                this.PageName = PageName;
                if (Name == "")
                    this.Name = PageName;

                this[UserParameters.PasswordHash] = HashCode;
                this[UserParameters.FullName] = FullName;
                this[UserParameters.UserCatalog] = UserCatalog;
                this[UserParameters.UserInterface] = Interface;
                this[UserParameters.UserRights] = Rights;
                modified = false;
            }

            //Возвращает массив байт для записи в поток файла
            public byte[] Serialyze()
            {
                //посчитаем размер массива. 17 = числовые поля плюс по 1 байту на каждое строковое поле для хранения длины. В конце должны быть Int(0)
                int rawsize = 17 + PasswordHash.Length + FullName.Length + UserCatalog.Length + UserInterface.Length + UserRights.Length + 4;

                byte[] rawdata = new byte[rawsize];
                byte[] buffer;
                int lastCount = 0;

                //преобразуем каждое поле в массив байт и сложим в общий массив в нужном порядке
                buffer = BitConverter.GetBytes((int)1);
                Array.Copy(buffer, 0, rawdata, lastCount, buffer.Length);
                lastCount += buffer.Length;

                for (UserParameters i = UserParameters.DontCheckRights; i <= UserParameters.UserRights; i++)
                {
                    dynamic param = GetParam(i);

                    if (param.GetType().Name == "Int32")
                        buffer = BitConverter.GetBytes(param);
                    else
                        buffer = param;

                    Array.Copy(buffer, 0, rawdata, lastCount, buffer.Length);
                    lastCount += buffer.Length;
                }
               return rawdata;
            }

            public byte[] GetObjectData()
            {
                return Serialyze();
            }
        }

        //Класс описывает список элементов пользователей.
        [Serializable]
        public class UsersList : Dictionary<string, UserItem>
        {
            private Dictionary<string, string> Container;

            public string DBCAtalog = string.Empty;
            public string DBAPath = string.Empty;
            public string USRPath = string.Empty;

            public List<string> InterfaceList = new List<string>();
            public List<string> RightsList = new List<string>();

            public bool modified = false;

            [ComImport]
            [Guid("0000000d-0000-0000-C000-000000000046")]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            interface IEnumSTATSTG
            {
                // The user needs to allocate an STATSTG array whose size is celt.
                [PreserveSig]
                uint Next(
                    uint celt,
                    [MarshalAs(UnmanagedType.LPArray), Out]
                    System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt,
                    out uint pceltFetched
                );
                void Skip(uint celt);
                void Reset();
                [return: MarshalAs(UnmanagedType.Interface)]
                IEnumSTATSTG Clone();
            }

            [ComImport]
            [Guid("0000000b-0000-0000-C000-000000000046")]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            interface IStorage
            {
                void CreateStream(
                    /* [string][in] */ string pwcsName,
                    /* [in] */ uint grfMode,
                    /* [in] */ uint reserved1,
                    /* [in] */ uint reserved2,
                    /* [out] */ out IStream ppstm);

                void OpenStream(
                    /* [string][in] */ string pwcsName,
                    /* [unique][in] */ IntPtr reserved1,
                    /* [in] */ uint grfMode,
                    /* [in] */ uint reserved2,
                    /* [out] */ out IStream ppstm);

                void CreateStorage(
                    /* [string][in] */ string pwcsName,
                    /* [in] */ uint grfMode,
                    /* [in] */ uint reserved1,
                    /* [in] */ uint reserved2,
                    /* [out] */ out IStorage ppstg);

                void OpenStorage(
                    /* [string][unique][in] */ string pwcsName,
                    /* [unique][in] */ IStorage pstgPriority,
                    /* [in] */ uint grfMode,
                    /* [unique][in] */ IntPtr snbExclude,
                    /* [in] */ uint reserved,
                    /* [out] */ out IStorage ppstg);

                void CopyTo(
                    /* [in] */ uint ciidExclude,
                    /* [size_is][unique][in] */ Guid rgiidExclude, // should this be an array?
                    /* [unique][in] */ IntPtr snbExclude,
                    /* [unique][in] */ IStorage pstgDest);

                void MoveElementTo(
                    /* [string][in] */ string pwcsName,
                    /* [unique][in] */ IStorage pstgDest,
                    /* [string][in] */ string pwcsNewName,
                    /* [in] */ uint grfFlags);

                void Commit(
                    /* [in] */ STGC grfCommitFlags);

                void Revert();

                void EnumElements(
                    /* [in] */ uint reserved1,
                    /* [size_is][unique][in] */ IntPtr reserved2,
                    /* [in] */ uint reserved3,
                    /* [out] */ out IEnumSTATSTG ppenum);

                void DestroyElement(
                    /* [string][in] */ string pwcsName);

                void RenameElement(
                    /* [string][in] */ string pwcsOldName,
                    /* [string][in] */ string pwcsNewName);

                void SetElementTimes(
                    /* [string][unique][in] */ string pwcsName,
                    /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pctime,
                    /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME patime,
                    /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

                void SetClass(
                    /* [in] */ Guid clsid);

                void SetStateBits(
                    /* [in] */ uint grfStateBits,
                    /* [in] */ uint grfMask);

                void Stat(
                    /* [out] */ out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg,
                    /* [in] */ uint grfStatFlag);

            }

            [Flags]
            private enum STGC : int
            {
                DEFAULT = 0,
                OVERWRITE = 1,
                ONLYIFCURRENT = 2,
                DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
                CONSOLIDATE = 8
            }

            [Flags]
            private enum STGM : int
            {
                DIRECT = 0x00000000,
                TRANSACTED = 0x00010000,
                SIMPLE = 0x08000000,
                READ = 0x00000000,
                WRITE = 0x00000001,
                READWRITE = 0x00000002,
                SHARE_DENY_NONE = 0x00000040,
                SHARE_DENY_READ = 0x00000030,
                SHARE_DENY_WRITE = 0x00000020,
                SHARE_EXCLUSIVE = 0x00000010,
                PRIORITY = 0x00040000,
                DELETEONRELEASE = 0x04000000,
                NOSCRATCH = 0x00100000,
                CREATE = 0x00001000,
                CONVERT = 0x00020000,
                FAILIFTHERE = 0x00000000,
                NOSNAPSHOT = 0x00200000,
                DIRECT_SWMR = 0x00400000,
            }

            [Flags]
            private enum STATFLAG : uint
            {
                STATFLAG_DEFAULT = 0,
                STATFLAG_NONAME = 1,
                STATFLAG_NOOPEN = 2
            }

            [Flags]
            private enum STGTY : int
            {
                STGTY_STORAGE = 1,
                STGTY_STREAM = 2,
                STGTY_LOCKBYTES = 3,
                STGTY_PROPERTY = 4
            }

            //Читает IStream в массив байт
            private static byte[] ReadIStream(IStream pIStream)
            {
                System.Runtime.InteropServices.ComTypes.STATSTG StreamInfo;
                pIStream.Stat(out StreamInfo, 0);
                byte[] data = new byte[StreamInfo.cbSize];
                pIStream.Read(data, (int)StreamInfo.cbSize, IntPtr.Zero);
                return data;
            }

            class NativeMethods
            {
                [DllImport("ole32.dll")]
                public static extern int StgIsStorageFile(
                    [MarshalAs(UnmanagedType.LPWStr)] string pwcsName);

                [DllImport("ole32.dll")]
                public static extern int StgOpenStorage(
                    [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
                    IStorage pstgPriority,
                    STGM grfMode,
                    IntPtr snbExclude,
                    uint reserved,
                    out IStorage ppstgOpen);

                [DllImport("ole32.dll")]
                public static extern int StgCreateDocfile(
                    [MarshalAs(UnmanagedType.LPWStr)]string pwcsName,
                    STGM grfMode,
                    uint reserved,
                    out IStorage ppstgOpen);
            }

            public UsersList()
                : base()
            {
                Container = new Dictionary<string, string>();
                modified = true;
            }

            public string ToXml()
            {
                string xml = string.Empty;
                xml += "<?xml version=\"1.0\" encoding=\"cp866\"?>\r\n";
                xml += string.Format("<UserdefList Filename = \"{0}\" databaseformat=\"{1}\">\r\n", this.USRPath, "undefined");
                foreach (UserDefworks.UserItem User in this.Values)
                {
                    xml += string.Format("     <UserItemType UserName = \"{0}\" StreamName=\"{1}\">\r\n", User.Name, User.PageName);
                    for (UserDefworks.UserParameters i = UserDefworks.UserParameters.DontCheckRights; i <= UserDefworks.UserParameters.UserRights; i++)
                    {
                        xml += string.Format("         <{0} description=\"{2}\">{1}</{0}>\r\n", i, User.GetParam(i), UserDefworks.UserParamNames[i]);
                    }
                    xml += string.Format("     </UserItemType>\r\n");
                }
                xml += string.Format("</UserdefList>\r\n");

                return xml;
            }

            //разбирает строку Container.Contents в словарь
            private Dictionary<string, string> ParseContainerD(string Container)
            {
                Dictionary<string, string> res = new Dictionary<string, string>();
                string[] UserItemType;

                if (!Container.Contains("Container.Contents"))
                    return res;

                Container = Container.Replace("{\"Container.Contents\",{", "").Replace("}}", "").Replace("},{", ";");

                foreach (string UserItem in Container.Split(';'))
                {
                    UserItemType = UserItem.Replace("\"", "").Split(',');
                    if (UserItemType.Length > 3)
                        res.Add(UserItemType[1], UserItemType[2]);
                }
                return res;
            }

            //разбирает строку Container.Contents в словарь
            private List<string> ParseContainerL(string Container)
            {
                List<string> res = new List<string>();
                string[] UserItemType;

                if (!Container.Contains("Container.Contents"))
                    return res;

                Container = Container.Replace("{\"Container.Contents\",{", "").Replace("}}", "").Replace("},{", ";");

                foreach (string UserItem in Container.Split(';'))
                {
                    UserItemType = UserItem.Replace("\"", "").Split(',');
                    res.Add(UserItemType[2]);
                }
                return res;
            }
            //Собирает строку Container.Contents из словаря
            private string ConstructContainer()
            {

                string strContainer = "{\"Container.Contents\"";
                Container = new Dictionary<string, string>();
                foreach (UserItem item in this.Values)
                {
                    strContainer += string.Format(",{{\"UserItemType\",\"{0}\",\"{1}\",\"\"}}", item.PageName, item.Name);
                    Container.Add(item.PageName, item.Name);
                }
                strContainer += "}\r";
                return strContainer;
            }

            public void Add(UserItem Item)
            {
                if (this.Keys.Contains(Item.Name))
                {
                    this.Remove(Item.Name);
                }

                if (Container.ContainsKey(Item.PageName))
                {
                    Container.Remove(Item.PageName);
                }

                base.Add(Item.Name, Item);
                Container.Add(Item.PageName, Item.Name);
                modified = true;
            }

            new public void Remove(string Key)
            {
                if (this.Keys.Contains(Key))
                {
                    Container.Remove(this[Key].PageName);
                    base.Remove(Key);
                }

                if (Container.Values.Contains(Key))
                {
                    Container.Remove(Container.Keys.ElementAt(Container.Values.ToList().IndexOf(Key)));
                }
                modified = true;

            }

            public UserItem GetValue(int Index)
            {
                return this[Index];
            }

            public UserItem this[int Index]
            {
                get
                {
                    return this.ElementAt(Index).Value;
                }
                set
                {
                    string key = this.ElementAt(Index).Key;
                    this.Remove(key);
                    this.Add(key, value);
                    modified = true;
                }
            }
   
            new public UserItem this[string key]
            {
                get
                {
                    return base[key];
                }
                set
                {
                    this.Remove(key);
                    this.Add(key, value);
                }
            }

            public void FillLists()
            {
                IStorage storage = null;
                IStream pIStream = null;
                byte[] data = { 0 };

                string MDFileName = DBCAtalog + "1cv7.md";
                IStorage RootStorage = null;
                if (NativeMethods.StgOpenStorage(MDFileName, null, STGM.READ | STGM.TRANSACTED, IntPtr.Zero, 0, out RootStorage) == 0)
                {

                    RootStorage.OpenStorage("UserDef", null, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), IntPtr.Zero, 0, out RootStorage);

                    RootStorage.OpenStorage("Page.1", null, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), IntPtr.Zero, 0, out storage);
                    try
                    {
                        storage.OpenStream("Container.Contents", IntPtr.Zero, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), 0, out pIStream);
                        data = ReadIStream(pIStream);
                        // ОБЯЗАТЕЛЬНО освобождть поток, иначе доступа к нему позже не будет
                        Marshal.FinalReleaseComObject(pIStream);
                    }
                    catch (Exception ex)
                    {
                    }
                    //Разберем список интерфейсов
                    InterfaceList = ParseContainerL(Encoding.Default.GetString(data, 0, data.Length));
                    Marshal.ReleaseComObject(storage);
                    Marshal.FinalReleaseComObject(storage);
                    storage = null;

                    RootStorage.OpenStorage("Page.2", null, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), IntPtr.Zero, 0, out storage);
                    try
                    {
                        storage.OpenStream("Container.Contents", IntPtr.Zero, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), 0, out pIStream);
                        data = ReadIStream(pIStream);
                        // ОБЯЗАТЕЛЬНО освобождть поток, иначе доступа к нему позже не будет
                        Marshal.FinalReleaseComObject(pIStream);
                    }
                    catch (Exception ex)
                    {
                    }
                    //Разберем список наборов прав
                    RightsList = ParseContainerL(Encoding.Default.GetString(data, 0, data.Length));
                    Marshal.ReleaseComObject(storage);
                    Marshal.FinalReleaseComObject(storage);
                    storage = null;

                    Marshal.ReleaseComObject(RootStorage);
                    Marshal.FinalReleaseComObject(RootStorage);
                    RootStorage = null;

                    GC.Collect();
                    GC.Collect();
                }
            }

            public UsersList(string USRfileName)
                : this()
            {
                USRPath = USRfileName;
                DBCAtalog = GetDBCatalog(USRfileName);
                DBAPath = DBCAtalog + "1cv7.dba";

                try
                {
                    FillLists();
                }
                catch
                {

                }

                if (!File.Exists(USRfileName))
                    return;

                IStorage storage = null;
                uint fetched = 0;
                IStream pIStream = null;
                byte[] data = { 0 };
                IEnumSTATSTG pIEnumStatStg = null;
                

                if (NativeMethods.StgIsStorageFile(USRfileName) == 0)
                {
                    if (NativeMethods.StgOpenStorage(USRfileName, null, STGM.READ | STGM.TRANSACTED, IntPtr.Zero, 0, out storage) == 0)
                    {
                        //System.Runtime.InteropServices.ComTypes.STATSTG statstg = new System.Runtime.InteropServices.ComTypes.STATSTG();
                        System.Runtime.InteropServices.ComTypes.STATSTG[] regelt = { new System.Runtime.InteropServices.ComTypes.STATSTG() };

                        //Читаем содержимое "Container.Contents"
                        try
                        {
                            storage.OpenStream("Container.Contents", IntPtr.Zero, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), 0, out pIStream);
                            data = ReadIStream(pIStream);
                            // ОБЯЗАТЕЛЬНО освобождть поток, иначе доступа к нему позже не будет
                            Marshal.FinalReleaseComObject(pIStream);
                        }
                        catch 
                        {
                        }



                        //Разберем список соответствия потоков и имен пользователей
                        Container = ParseContainerD(Encoding.Default.GetString(data, 0, data.Length));
                        //Обойдем все элементы хранилища
                        storage.EnumElements(0, IntPtr.Zero, 0, out pIEnumStatStg);
                        while (pIEnumStatStg.Next(1, regelt, out fetched) == 0)
                        {
                            string filePage = regelt[0].pwcsName;
                            if (filePage != "Container.Contents")
                            {
                                string UserName = string.Empty;

                                if (Container.Keys.Contains(filePage))
                                    UserName = Container[filePage];
                                else
                                    UserName = filePage;

                                if ((STGTY)regelt[0].type == STGTY.STGTY_STREAM)
                                {
                                    storage.OpenStream(filePage, IntPtr.Zero, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), 0, out pIStream);
                                    if (pIStream != null)
                                    {
                                        data = ReadIStream(pIStream);
                                        Marshal.ReleaseComObject(pIStream);
                                        UserItem User = new UserItem(data, UserName, filePage);
                                        this.Add(User);
                                    }
                                }
                            }
                        }

                        Marshal.ReleaseComObject(storage);
                        Marshal.FinalReleaseComObject(storage);
                        storage = null;
                        Marshal.ReleaseComObject(pIEnumStatStg);
                        Marshal.FinalReleaseComObject(pIEnumStatStg);
                        pIEnumStatStg = null;

                        GC.Collect();
                        GC.Collect();
                       // GC.WaitForPendingFinalizers();
                    }
                    else
                    {
                        throw new Exception(string.Format("Файл {0} занят.", USRfileName));
                    }
                }
                else
                {
                    throw new Exception(string.Format("Файл {0} не является хранилищем списка пользователей.", USRfileName));
                }
                modified = false;
            }

            public bool Save(string USRfileName, bool MakeBackUp = true)
            {
                if (DBCAtalog == string.Empty)
                    DBCAtalog = GetDBCatalog(USRfileName);

                DBAPath = Path.Combine(DBCAtalog,"1cv7.dba");

                // Для базы SQL нужно обновить 1cv7.dba. 
                // И если октрыт конфигуратор, он очистит параметры подключения к БД
                // проверим, не открыт ли конфигуратор
                if (IsConfigRunning(DBCAtalog))
                            throw new Exception(string.Format("В базе {0} открыт конфигуратор. Нельзя сохранять список пользователей, иначе будут сброшены параметры подключения к SQL.", Path.GetDirectoryName(Path.GetDirectoryName(USRfileName))));
                

                IStorage ppstgOpen = null;
                IStream pIStream = null;
                IEnumSTATSTG pIEnumStatStg = null;
                System.Runtime.InteropServices.ComTypes.STATSTG[] regelt = { new System.Runtime.InteropServices.ComTypes.STATSTG() };
                uint fetched = 0;
                byte[] data;
                string TempFileName = USRfileName + ".tmp";
                string oldfilename = USRfileName + ".old";

                if (File.Exists(TempFileName))
                {
                    File.Delete(TempFileName);
                }


                if (File.Exists(oldfilename))
                {
                    File.Delete(oldfilename);
                }

                if (File.Exists(USRfileName))
                {
                    if (MakeBackUp)
                    {
                        string BkpMAsk = string.Format("1cv7_{0}_bkp", DateTime.Now.ToString().Replace(".", "").Replace(":", "").Replace(" ", ""));
                        string BAckUpPath = Path.Combine(DBCAtalog, "usrdef", "BackUp");

                        //BkpMAsk = ;
                        if (!Directory.Exists(BAckUpPath))
                            Directory.CreateDirectory(BAckUpPath);

                        File.Copy(USRfileName, Path.Combine(BAckUpPath, BkpMAsk + ".usr"));
                        Console.WriteLine("{0}; {1}; Сохранён бэкап списка пользователей: {2}", DateTime.Now.ToString(), Environment.UserName, Path.Combine(BAckUpPath, BkpMAsk + ".usr"));

                        if (File.Exists(DBAPath))
                        {
                            File.Copy(DBAPath, Path.Combine(BAckUpPath, BkpMAsk + ".dba"));
                            Console.WriteLine("{0}; {1}; Сохранён бэкап настроек подключения к БД: {2}", DateTime.Now.ToString(), Environment.UserName, Path.Combine(BAckUpPath, BkpMAsk + ".dba"));
                        }
                    }

                    File.Copy(USRfileName, TempFileName);
                    if (!File.Exists(TempFileName))
                    {
                        throw new Exception("Не удалось создать временный файл.");
                    }


                    if (NativeMethods.StgIsStorageFile(TempFileName) == 0)
                    {
                        if (NativeMethods.StgOpenStorage(TempFileName, null, STGM.READWRITE | STGM.TRANSACTED, IntPtr.Zero, 0, out ppstgOpen) == 0)
                        {
                            ppstgOpen.EnumElements(0, IntPtr.Zero, 0, out pIEnumStatStg);
                            while (pIEnumStatStg.Next(1, regelt, out fetched) == 0)
                            {
                                ppstgOpen.DestroyElement(regelt[0].pwcsName);
                            }
                            Marshal.ReleaseComObject(pIEnumStatStg);
                            pIEnumStatStg = null;
                        }
                        else
                        {
                            throw new Exception(string.Format("Ошибка открытия файла {0} для записи.", USRfileName));
                        }
                    }
                    else
                    {
                        throw new Exception(string.Format("Файл {0} не является хранилищем списка пользователей.", USRfileName));
                    }
                }
                else
                {
                    if (!Directory.Exists(Path.GetDirectoryName(USRfileName)))
                        Directory.CreateDirectory(Path.GetDirectoryName(USRfileName));

                    NativeMethods.StgCreateDocfile(USRfileName, STGM.CREATE | STGM.WRITE | STGM.SHARE_EXCLUSIVE, 0, out ppstgOpen);
                }

                if (ppstgOpen != null)
                {
                    ppstgOpen.Commit(STGC.OVERWRITE);

                    ppstgOpen.CreateStream("Container.Contents", (uint)(STGM.CREATE | STGM.WRITE | STGM.SHARE_EXCLUSIVE), 0, 0, out pIStream);
                    data = Encoding.Default.GetBytes(ConstructContainer());
                    pIStream.Write(data, data.Length, new IntPtr(fetched));
                    pIStream.Commit((int)STGC.OVERWRITE);
                    Marshal.ReleaseComObject(pIStream);
                    ppstgOpen.Commit(0);
                    foreach (UserItem item in this.Values)
                    {
                        ppstgOpen.CreateStream(item.PageName, (uint)(STGM.CREATE | STGM.WRITE | STGM.SHARE_EXCLUSIVE), 0, 0, out pIStream);
                        data = item.Serialyze();
                        pIStream.Write(data, data.Length, new IntPtr(fetched));
                        pIStream.Commit((int)STGC.OVERWRITE);
                        Marshal.ReleaseComObject(pIStream);
                        ppstgOpen.Commit(STGC.OVERWRITE);
                    }
                }
                else
                {
                    throw new Exception(string.Format("Ошибка записи файла {0}", TempFileName));
                }
                try
                {
                    ppstgOpen.Commit(STGC.OVERWRITE);
                    Marshal.ReleaseComObject(ppstgOpen);
                    Marshal.FinalReleaseComObject(ppstgOpen);
                    ppstgOpen = null;

                    GC.Collect();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Ошибка записи файла {0} : {1}", TempFileName, ex.Message), ex);
                }
                
                GC.Collect();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (File.Exists(TempFileName))
                {
                    File.Move(USRfileName, oldfilename);
                    File.Move(TempFileName, USRfileName);
                }

               // try
                {
                    DBAPath = DBCAtalog + "1cv7.dba";
                    if (File.Exists(DBAPath))
                    {
                        string Connect = ReadDBA(DBAPath);
                        Thread.Sleep(300); 
                        string checksum = CheckSum(USRfileName);
                        checksum = CheckSum(USRfileName);
                        // Console.WriteLine("New Checksum: {0}", CheckSum(USRfileName));
                        Connect = string.Format("{0}{1}}}}}", Connect.Substring(0, Connect.IndexOf("sum\",") + 5), checksum);
                        WriteDBA(DBAPath, Connect);
                    }
                }
                //catch (Exception ex)
                //{
                //    throw new Exception(string.Format("Ошибка записи файла {0} : {1}", DBAPath, ex.Message), ex);
                //}
                return true;
            }
        }
    }
}

namespace Utilities.Exchange
{
    /// <summary>
    /// Utility class for creating and sending an appointment
    /// </summary>
    public class Appointment
    {
        #region Constructor
        /// <summary>
        /// Constructor
        /// </summary>
        public Appointment()
        {
            AttendeeList = new MailAddressCollection();
            Attachments = new List<Attachment>();
        }
        #endregion

        #region Public Functions

        public string ToString()
        {
            return GetHTML(false);
        }

        /// <summary>
        /// Adds an appointment to a user's calendar
        /// </summary>
        public virtual void AddAppointment()
        {
            string XMLNSInfo = "xmlns:g=\"DAV:\" "
                + "xmlns:e=\"http://schemas.microsoft.com/exchange/\" "
                + "xmlns:mapi=\"http://schemas.microsoft.com/mapi/\" "
                + "xmlns:mapit=\"http://schemas.microsoft.com/mapi/proptag/\" "
                + "xmlns:x=\"xml:\" xmlns:cal=\"urn:schemas:calendar:\" "
                + "xmlns:dt=\"urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/\" "
                + "xmlns:header=\"urn:schemas:mailheader:\" "
                + "xmlns:mail=\"urn:schemas:httpmail:\"";

            string CalendarInfo = "<cal:location>" + Location + "</cal:location>"// + Location + "</cal:location>"
                + "<cal:dtstart dt:dt=\"dateTime.tz\">" + StartDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.000Z") + "</cal:dtstart>"// + StartDate.ToUniversalTime().ToString("yyyyMMddTHHmmssZ") + "</cal:dtstart>"
                + "<cal:dtend dt:dt=\"dateTime.tz\">" + EndDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.000Z") + "</cal:dtend>"// + EndDate.ToUniversalTime().ToString("yyyyMMddTHHmmssZ") + "</cal:dtend>"
                + "<cal:instancetype dt:dt=\"int\">0</cal:instancetype>"
                + "<cal:busystatus>BUSY</cal:busystatus>"
                + "<cal:meetingstatus>CONFIRMED</cal:meetingstatus>"
                + "<cal:alldayevent dt:dt=\"boolean\">0</cal:alldayevent>"
                + "<cal:responserequested dt:dt=\"boolean\">1</cal:responserequested>"
                + "<cal:reminderoffset dt:dt=\"int\">900</cal:reminderoffset>"
                + "<cal:uid>" + MeetingGUID.ToString("B") + "</cal:uid>";

            string HeaderInfo = "<header:to>" + AttendeeList.ToString() + "</header:to>";

            string MailInfo = "<mail:subject>" + Subject + "</mail:subject>"
                + "<mail:htmldescription>" + Summary + "</mail:htmldescription>";

            string AppointmentRequest = "<?xml version=\"1.0\"?>"
                + "<g:propertyupdate " + XMLNSInfo + ">"
                + "<g:set><g:prop>"
                + "<g:contentclass>urn:content-classes:appointment</g:contentclass>"
                + "<e:outlookmessageclass>IPM.Appointment</e:outlookmessageclass>"
                + MailInfo
                + CalendarInfo
                + HeaderInfo
                + "<mapi:finvited dt:dt=\"boolean\">1</mapi:finvited>"
                + "</g:prop></g:set>"
                + "</g:propertyupdate>";

            System.Net.HttpWebRequest PROPPATCHRequest = (System.Net.HttpWebRequest)HttpWebRequest.Create(ServerName + "/exchange/" + Directory + "/Calendar/" + MeetingGUID.ToString() + ".eml");

            System.Net.CredentialCache MyCredentialCache = new System.Net.CredentialCache();

            if (!string.IsNullOrEmpty(UserName) && !string.IsNullOrEmpty(Password))
            {
                MyCredentialCache.Add(new System.Uri(ServerName + "/exchange/" + Directory + "/Calendar/" + MeetingGUID.ToString() + ".eml"),
                                       "NTLM",
                                       new System.Net.NetworkCredential(UserName, Password));
            }
            else
            {
                MyCredentialCache.Add(new System.Uri(ServerName + "/exchange/" + Directory + "/Calendar/" + MeetingGUID.ToString() + ".eml"),
                                       "Negotiate",
                                       (System.Net.NetworkCredential)CredentialCache.DefaultCredentials);
            }

            PROPPATCHRequest.Credentials = MyCredentialCache;
            PROPPATCHRequest.Method = "PROPPATCH";
            byte[] bytes = Encoding.UTF8.GetBytes((string)AppointmentRequest);
            PROPPATCHRequest.ContentLength = bytes.Length;
            using (System.IO.Stream PROPPATCHRequestStream = PROPPATCHRequest.GetRequestStream())
            {
                PROPPATCHRequestStream.Write(bytes, 0, bytes.Length);
                PROPPATCHRequestStream.Close();
                PROPPATCHRequest.ContentType = "text/xml";
                System.Net.WebResponse PROPPATCHResponse = (System.Net.HttpWebResponse)PROPPATCHRequest.GetResponse();
                PROPPATCHResponse.Close();
            }
        }

        /// <summary>
        /// Emails an appointment to the specified users
        /// </summary>
        public virtual void EmailAppointment()
        {
            using (MailMessage Mail = new MailMessage())
            {
                System.Net.Mime.ContentType TextType = new System.Net.Mime.ContentType("text/plain");
                using (AlternateView TextView = AlternateView.CreateAlternateViewFromString(GetText(), TextType))
                {
                    System.Net.Mime.ContentType HTMLType = new System.Net.Mime.ContentType("text/html");
                    using (AlternateView HTMLView = AlternateView.CreateAlternateViewFromString(GetHTML(false), HTMLType))
                    {
                        System.Net.Mime.ContentType CalendarType = new System.Net.Mime.ContentType("text/calendar");
                        CalendarType.Parameters.Add("method", "REQUEST");
                        CalendarType.Parameters.Add("name", "meeting.ics");
                        using (AlternateView CalendarView = AlternateView.CreateAlternateViewFromString(GetCalendar(false), CalendarType))
                        {
                            CalendarView.TransferEncoding = System.Net.Mime.TransferEncoding.SevenBit;

                            Mail.AlternateViews.Add(TextView);
                            Mail.AlternateViews.Add(HTMLView);
                            Mail.AlternateViews.Add(CalendarView);

                            Mail.From = new MailAddress(OrganizerEmail);

                            foreach (MailAddress attendee in AttendeeList)
                            {
                                Mail.To.Add(attendee);
                            }


                            Mail.Subject = Subject;

                            foreach (Attachment Attachment in Attachments)
                            {
                                Mail.Attachments.Add(Attachment);
                            }


                            SmtpClient Server = new SmtpClient(ServerName, Port);
#if DEBUG
                            Server.EnableSsl = false;
#else
                            Server.EnableSsl = true;
#endif

                            if (!string.IsNullOrEmpty(UserName) && !string.IsNullOrEmpty(Password))
                            {
                                
                                Server.Credentials = new System.Net.NetworkCredential(UserName, Password);
                            }
                            if (AttendeeList.Count > 0)
                            {
                                Server.Send(Mail);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Sends a cancellation to the people specified
        /// </summary>
        public virtual void SendCancelEmails()
        {
            using (MailMessage Mail = new MailMessage())
            {
                System.Net.Mime.ContentType TextType = new System.Net.Mime.ContentType("text/plain");
                using (AlternateView TextView = AlternateView.CreateAlternateViewFromString(GetText(), TextType))
                {
                    System.Net.Mime.ContentType HTMLType = new System.Net.Mime.ContentType("text/html");
                    using (AlternateView HTMLView = AlternateView.CreateAlternateViewFromString(GetHTML(true), HTMLType))
                    {
                        System.Net.Mime.ContentType CalendarType = new System.Net.Mime.ContentType("text/calendar");
                        CalendarType.Parameters.Add("method", "CANCEL");
                        CalendarType.Parameters.Add("name", "meeting.ics");
                        using (AlternateView CalendarView = AlternateView.CreateAlternateViewFromString(GetCalendar(true), CalendarType))
                        {
                            CalendarView.TransferEncoding = System.Net.Mime.TransferEncoding.SevenBit;

                            Mail.AlternateViews.Add(TextView);
                            Mail.AlternateViews.Add(HTMLView);
                            Mail.AlternateViews.Add(CalendarView);

                            Mail.From = new MailAddress(OrganizerEmail);
                            foreach (MailAddress attendee in AttendeeList)
                            {
                                Mail.To.Add(attendee);
                            }
                            Mail.Subject = Subject + " - Cancelled";
                            foreach (Attachment Attachment in Attachments)
                            {
                                Mail.Attachments.Add(Attachment);
                            }
                            SmtpClient Server = new SmtpClient(ServerName, Port);
                            if (!string.IsNullOrEmpty(UserName) && !string.IsNullOrEmpty(Password))
                            {
                                Server.Credentials = new System.Net.NetworkCredential(UserName, Password);
                            }
                            if (AttendeeList.Count > 0)
                            {
                                Server.Send(Mail);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Cancels an appointment on someone's calendar
        /// </summary>
        public virtual void CancelAppointment()
        {
            System.Net.HttpWebRequest PROPPATCHRequest = (System.Net.HttpWebRequest)HttpWebRequest.Create(ServerName + "/exchange/" + Directory + "/Calendar/" + MeetingGUID.ToString() + ".eml");

            System.Net.CredentialCache MyCredentialCache = new System.Net.CredentialCache();
            if (!string.IsNullOrEmpty(UserName) && !string.IsNullOrEmpty(Password))
            {
                MyCredentialCache.Add(new System.Uri(ServerName + "/exchange/" + Directory + "/Calendar/" + MeetingGUID.ToString() + ".eml"),
                                   "NTLM",
                                   new System.Net.NetworkCredential(UserName, Password));
            }
            else
            {
                MyCredentialCache.Add(new System.Uri(ServerName + "/exchange/" + Directory + "/Calendar/" + MeetingGUID.ToString() + ".eml"),
                                       "Negotiate",
                                       (System.Net.NetworkCredential)CredentialCache.DefaultCredentials);
            }

            PROPPATCHRequest.Credentials = MyCredentialCache;
            PROPPATCHRequest.Method = "DELETE";
            System.Net.WebResponse PROPPATCHResponse = (System.Net.HttpWebResponse)PROPPATCHRequest.GetResponse();
            PROPPATCHResponse.Close();
        }

        #endregion

        #region Private Functions

        /// <summary>
        /// Returns the text version of the appointment
        /// </summary>
        /// <returns>A text version of the appointment</returns>
        private string GetText()
        {
            string Body = "Type:Single Meeting\n" +
                "Organizer:" + OrganizerName + "\n" +
                "Start Time:" + StartDate.ToLongDateString() + " " + StartDate.ToLongTimeString() + "\n" +
                "End Time:" + EndDate.ToLongDateString() + " " + EndDate.ToLongTimeString() + "\n" +
                "Time Zone:" + System.TimeZone.CurrentTimeZone.StandardName + "\n" +
                "Location: " + Location + "\n\n" +
                "*~*~*~*~*~*~*~*~*~*\n\n" +
                Summary;
            return Body;
        }

        /// <summary>
        /// Gets an HTML version of the appointment
        /// </summary>
        /// <param name="Canceled">If true it returns a cancellation, false it returns a request</param>
        /// <returns>An HTML version of the appointment</returns>
        private string GetHTML(bool Canceled)
        {
            string bodyHTML = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 3.2//EN\">\r\n<HTML>\r\n<HEAD>\r\n<META HTTP-EQUIV=\"Content-Type\" CONTENT=\"text/html; charset=utf-8\">\r\n<META NAME=\"Generator\" CONTENT=\"MS Exchange Server version 6.5.7652.24\">\r\n<TITLE>{0}</TITLE>\r\n</HEAD>\r\n<BODY>\r\n<!-- Converted from text/plain format -->\r\n<P><FONT SIZE=2>Type:Single Meeting<BR>\r\nOrganizer:{1}<BR>\r\nStart Time:{2}<BR>\r\nEnd Time:{3}<BR>\r\nTime Zone:{4}<BR>\r\nLocation:{5}<BR>\r\n<BR>\r\n*~*~*~*~*~*~*~*~*~*<BR>\r\n<BR>\r\n{6}<BR>\r\n</FONT>\r\n</P>\r\n\r\n</BODY>\r\n</HTML>";
            string TempSummary = Summary;
            if (Canceled)
            {
                TempSummary += " - Canceled";
            }
            return string.Format(bodyHTML,
                TempSummary,
                OrganizerName,
                StartDate.ToLongDateString() + " " + StartDate.ToLongTimeString(),
                EndDate.ToLongDateString() + " " + EndDate.ToLongTimeString(),
                System.TimeZone.CurrentTimeZone.StandardName,
                Location,
                TempSummary);
        }

        /// <summary>
        /// Gets an iCalendar version of the appointment
        /// </summary>
        /// <param name="Canceled">If true, it returns a cancellation version.
        /// If false, it returns a request version.</param>
        /// <returns>An iCalendar version of the appointment</returns>
        private string GetCalendar(bool Canceled)
        {
            string DateFormatUsing = "yyyyMMddTHHmmssZ";
            string Method;

            string attendees = string.Empty;

            foreach (MailAddress Attendee in AttendeeList)
            {
                attendees += "ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=\"" + Attendee.DisplayName + "\":MAILTO:" + Attendee.Address + "\r\n";
            }

            

            if (!Canceled)
                Method = "REQUEST";
            else
                Method = "CANCEL";
            string bodyCalendar = "BEGIN:VCALENDAR\r\nPRODID:-//Microsoft Corporation//Outlook 12.0 MIMEDIR//EN\r\nVERSION:2.0\r\nMETHOD:{10}\r\nX-MS-OLK-FORCEINSPECTOROPEN:TRUE\r\nBEGIN:VTIMEZONE\r\nTZID:(GMT+10) Vladivostok\r\nBEGIN:STANDARD\r\nDTSTART:16010101T000000\r\nTZOFFSETFROM:+1100\r\nTZOFFSETTO:+1000\r\nEND:STANDARD\r\nEND:VTIMEZONE\r\nBEGIN:VEVENT\r\n{9}\r\nCLASS:PUBLIC\r\nCREATED:{8}\r\nDESCRIPTION:{7}\r\nDTEND;TZID=\"(GMT+10) Vladivostok\":{1}\r\nDTSTAMP:{8}\r\nDTSTART;TZID=\"(GMT+10) Vladivostok\":{0}\r\nLAST-MODIFIED:{8}\r\nLOCATION:{2}\r\nORGANIZER;CN=\"{3}\":mailto:{4}\r\nPRIORITY:5\r\nRECURRENCE-ID;TZID=\"(GMT+10) Vladivostok\":{0}\r\nSEQUENCE:19\r\nSUMMARY;LANGUAGE=ru:{6}\r\nTRANSP:OPAQUE\r\nUID:{5}\r\nX-ALT-DESC;FMTTYPE=text/html:{7}\r\nX-MICROSOFT-CDO-BUSYSTATUS:TENTATIVE\r\nX-MICROSOFT-CDO-IMPORTANCE:1\r\nX-MICROSOFT-CDO-INTENDEDSTATUS:BUSY\r\nX-MICROSOFT-DISALLOW-COUNTER:FALSE\r\nX-MS-OLK-APPTSEQTIME:{8}\r\nX-MS-OLK-CONFTYPE:0\r\nBEGIN:VALARM\r\nACTION:DISPLAY\r\nDESCRIPTION:REMINDER\r\nTRIGGER;RELATED=START:-PT00H05M00S\r\nEND:VALARM\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n";
            //string bodyCalendar = "BEGIN:VCALENDAR\r\nMETHOD:{10}\r\nPRODID:-//Microsoft Corporation//Outlook 12.0 MIMEDIR//EN\r\nVERSION:2.0\r\nX-MS-OLK-FORCEINSPECTOROPEN:TRUE\r\nBEGIN:VTIMEZONE\r\nTZID:(GMT+10) Vladivostok\r\nBEGIN:STANDARD\r\nDTSTART:16010101T000000\r\nTZOFFSETFROM:+1100\r\nTZOFFSETTO:+1000\r\nEND:STANDARD\r\nEND:VTIMEZONE\r\nBEGIN:VEVENT\r\nCLASS:PUBLIC\r\nCREATED:{8}\r\nDTSTART;TZID=\"(GMT+10) Vladivostok\":{0}\r\nSUMMARY:{7}\r\nUID:{5}\r\n{9}\r\nACTION;RSVP=TRUE;CN=\"{4}\":MAILTO:{4}\r\nORGANIZER;CN=\"{3}\":mailto:{4}\r\nLOCATION:{2}\r\nDTEND;TZID=\"(GMT+10) Vladivostok\":{1}\r\nDESCRIPTION:{7}\r\nSEQUENCE:1\r\nPRIORITY:5\r\nCLASS:\r\nCREATED:{8}\r\nLAST-MODIFIED:{8}\r\nTRANSP:OPAQUE\r\nX-MICROSOFT-CDO-BUSYSTATUS:BUSY\r\nX-MICROSOFT-CDO-INSTTYPE:0\r\nX-MICROSOFT-CDO-INTENDEDSTATUS:BUSY\r\nX-MICROSOFT-CDO-ALLDAYEVENT:FALSE\r\nX-MICROSOFT-CDO-IMPORTANCE:1\r\nX-MICROSOFT-CDO-OWNERAPPTID:-1\r\nX-MICROSOFT-CDO-ATTENDEE-CRITICAL-CHANGE:{8}\r\nX-MICROSOFT-CDO-OWNER-CRITICAL-CHANGE:{8}\r\nBEGIN:VALARM\r\nACTION:DISPLAY\r\nDESCRIPTION:REMINDER\r\nTRIGGER;RELATED=START:-PT00H05M00S\r\nEND:VALARM\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n";
            bodyCalendar = string.Format(bodyCalendar,
                StartDate.ToUniversalTime().ToString(DateFormatUsing),
                EndDate.ToUniversalTime().ToString(DateFormatUsing),
                Location,
                OrganizerName,
                OrganizerEmail,
                MeetingGUID.ToString("B"),
                Summary,
                Subject,
                DateTime.Now.ToUniversalTime().ToString(DateFormatUsing),
                attendees,
                Method);
            return bodyCalendar;
        }

        #endregion

        #region Variables
        public DateTime StartDate;
        public DateTime EndDate;
        public string Subject;
        public string Summary;
        public string Location;
        public MailAddressCollection AttendeeList;
        public string OrganizerName;
        public string OrganizerEmail;

        public List<Attachment> Attachments;

        public Guid MeetingGUID;

        public string Directory;
        public string ServerName;
        public string UserName;
        public string Password;
        public int Port;

        #endregion
    }
}

//<Guid("7492A893-8CD4-458C-839E-B17AFDDAB5D3")>
namespace HTTPServer
{

    public static class Options
    {
        public static string FileName = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath),  Path.GetFileNameWithoutExtension(Application.ExecutablePath)) + ".ini";

        public static string Get(string OptionName)
        {
            string result = string.Empty;
            if(File.Exists(FileName))
            {
                string optionLine = File.ReadAllLines(FileName).FirstOrDefault(x => x.Split(new string[] { " = " }, StringSplitOptions.RemoveEmptyEntries)[0].Equals(OptionName, StringComparison.OrdinalIgnoreCase));
                if (optionLine != string.Empty && optionLine != null)
                    result = optionLine.Replace(string.Format("{0} = ", OptionName),"");
                else
                    Set(OptionName, string.Empty);
            }
            else
            {
                Set(OptionName, string.Empty);
            }
            return result;
        }

        public static void Set(string OptionName, string Value)
        {
           
            if (File.Exists(FileName))
            {
                StringBuilder OptionsList = new StringBuilder();

                foreach(string OptionLine in File.ReadAllLines(FileName))
                {
                    string[] OptionValue = OptionLine.Split(new string[] { " = " }, StringSplitOptions.RemoveEmptyEntries);
                    if (OptionValue[0].Equals(OptionName, StringComparison.OrdinalIgnoreCase))
                    {
                        OptionsList.AppendLine(string.Format("{0} = {1}", OptionName, Value));
                    }
                    else
                    {
                        OptionsList.AppendLine(OptionLine);
                    }
                    
                }
                File.WriteAllText(FileName, OptionsList.ToString());
            }
            else
            {
                File.WriteAllText(FileName, string.Format("{0} = {1}", OptionName, Value));
            }
        }
    }

    public class Utils
    {

        class Entry
        {
           public SqlConnection Connection;
           public bool isBisy = false;
           public int UID;

            public Entry()
            {

                string ConnectionString = Options.Get("ConnectionString");
                if (ConnectionString == string.Empty)
                {
#if DEBUG
                    ConnectionString = @"Data Source=PB3000070115PC;Integrated Security=SSPI;Initial Catalog=Portal";
#else
                    ConnectionString = @"Data Source=RUVVOS20005SRV;Integrated Security=SSPI;Initial Catalog=Portal";
#endif
                    Options.Set("ConnectionString", ConnectionString);
                }
                Connection = new SqlConnection(@ConnectionString);
                isBisy = true;
            try
            {
                Connection.Open();
                Utils.SQLPresent = true;
            }
            catch
            {
                Utils.SQLPresent = false;
            }

            UID = Connection.GetHashCode();
           }


        }


        public class ConnectionPool 
        {
            string DefaultConnectionProperties = string.Empty;
            static int Limit = 5; //максимальное количество открытых подключений.
            static int WaitingQueryLenght = 0;
            static List<Entry> pool = new List<Entry>();
            static Dictionary<int,SqlDataReader> readerbindings = new Dictionary<int,SqlDataReader>();


            public static SqlDataReader ExecuteReader(string QueryText)
            {
                SqlConnection Conn = Current;
                SqlCommand command = new SqlCommand(QueryText, Conn);
                SqlDataReader reader = command.ExecuteReader();
                command.StatementCompleted += command_StatementCompleted;
                Utils.ConnectionPool.BindReader(Conn, reader);
                return reader;                
            }

            static void command_StatementCompleted(object sender, System.Data.StatementCompletedEventArgs e)
            {
                SqlConnection Conn = ((SqlCommand)sender).Connection;
                readerbindings.Remove(Conn.GetHashCode());
            }

            public static void BindReader(SqlConnection Conn, SqlDataReader Reader)
            {
                if (!readerbindings.ContainsKey(Conn.GetHashCode()))
                    readerbindings.Add(Conn.GetHashCode(), Reader);
            }

            private static Entry GetFreeConnection()
            {
                try
                {
                    foreach (Entry Conn in pool)
                    {
                        if (!readerbindings.ContainsKey(Conn.UID))
                            return Conn;
                        else
                        {
                            if (readerbindings[Conn.UID] == null)
                            {
                                readerbindings.Remove(Conn.UID);
                                continue;
                            }

                            if (readerbindings[Conn.UID].IsClosed)
                                readerbindings.Remove(Conn.UID);
                        }
                    }
                }
                catch
                {

                }

                return null;
            }

            public static SqlConnection Current
            {
                get
                {
                    Entry Conn = GetFreeConnection();

                    if (Conn != null)
                    {
                        Conn.isBisy = true;
                        return Conn.Connection;
                    }

                    if (pool.Count < Limit)
                    {
                        Conn = new Entry();
                        pool.Add(Conn);
                        return Conn.Connection;
                    }
                    else
                    {
                        WaitingQueryLenght++;
                        while (Conn == null) // Ждем освобождения 
                        {
                            Conn = GetFreeConnection();
                            Thread.Sleep(100);
                        }
                        WaitingQueryLenght--;
                        return Conn.Connection;
                    }


                }

            }



        }

      //  public static SqlConnection thisConnection = null;
        public static bool SQLPresent = false;
        public static SortedList SessID = new SortedList();
        public static byte[] Logo = (byte[])new System.Drawing.ImageConverter().ConvertTo(Properties.Resources.top_rlogo, typeof(byte[]));
        public static IntPtr ServerHandle = IntPtr.Zero;

        [Serializable()]
        public enum SchedulerTasks
        {
            Null = 0,
            ResetSession = 1
        };

        [Serializable()]
        public class KickOffSchedulerEntry
        {
            public string Username = string.Empty;
            public SchedulerTasks TaskType;
            public string Time = string.Format("{0:D2}:{0:D2}",0);
            public bool Enabled = false;
            public string id;

            public KickOffSchedulerEntry(string username, SchedulerTasks TaskType, string Time, bool Enabled)
            {
                this.Username = username;
                this.TaskType = TaskType;
                this.Time = Time;
                this.Enabled = Enabled;
                this.id = Guid.NewGuid().ToString();
            }

            public bool Run()
            {
                bool success = false;
                switch (TaskType)
                {
                    case SchedulerTasks.ResetSession:
                        {
                            int SessionID = GetSessionIDByUserName(Username);

                            if (SessionID == -1)
                            {
                                Console.WriteLine("{0}; {1}; Ошибка: Не удалось запустить задачу по расписанию ({2}). Не удалсь определить сессию.", DateTime.Now.ToString(), Environment.UserName, Username);
                            }
                            else
                            {

                                if (SessionID != Process.GetCurrentProcess().SessionId)
                                {
                                    success = WTSLogoffSession(ServerHandle, SessionID, true);
                                    if (!success)
                                        Console.WriteLine("{0}; {1}; Ошибка: Не удалось запустить задачу по расписанию ({2}). Ошибка сброса сессии {3}.", DateTime.Now.ToString(), Environment.UserName, Username, SessionID);
                                    else
                                        Console.WriteLine("{0}; {1}; Задача по расписанию ({2}). Сессия {3} сброшена.", DateTime.Now.ToString(), Environment.UserName, Username, SessionID);
                                }
                                else
                                    Console.WriteLine("{0}; {1}; Ошибка: Задача по расписанию ({2}) попыталась завершить сессию системного пользователя - {3}.", DateTime.Now.ToString(), Environment.UserName, Username, SessionID);
                            }
                            break;
                        }
                    default:
                        break;
                }
                return success;
            }
        }

        [Serializable()]
        public class KickOffScheduler : List<KickOffSchedulerEntry> 
        {
            void Scheduler_Elapsed(object sender, ElapsedEventArgs e)
            {
                string time = string.Format("{0:D2}:{1:D2}", DateTime.Now.Hour, DateTime.Now.Minute);
                ThreadPool.QueueUserWorkItem(new WaitCallback(CheckTasks), time);
            }

            System.Timers.Timer SchedTimer;
            public void Start()
            {
                SchedTimer = new System.Timers.Timer();
                SchedTimer.Interval = 60000;
                SchedTimer.Elapsed += Scheduler_Elapsed;
                SchedTimer.Start();
            }

            void CheckTasks(object strTime)
            {
               string Time = (string)strTime;
               foreach(KickOffSchedulerEntry Entry in this)
                {
                    if (Entry.Enabled)
                        if (Entry.Time == Time)
                            Entry.Run();
                }
            }

            public KickOffScheduler()
                : base()
            {
                try
                {
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(List<KickOffSchedulerEntry>));
                    FileStream Scheduler = File.Open(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath.ToLower()) + "\\.scheduler", FileMode.Open);
                    List<KickOffSchedulerEntry> tmp = (List<KickOffSchedulerEntry>)ser.ReadObject(Scheduler);
                    Scheduler.Close();
                    foreach (KickOffSchedulerEntry Entry in tmp)
                    {
                        Add(Entry);
                    }
                    tmp = null;
                    ser = null;
                    
                }
                catch
                {
                }
                Start();
            }


            ~KickOffScheduler()
            {
                SchedTimer.Stop();
                try
                {
                    Save();
                }
                catch
                { 
                }

            }

            public void Save()
            {
                try
                {
                    FileStream Scheduler = File.Open(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath.ToLower()) + "\\.scheduler", FileMode.Create);
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(List<KickOffSchedulerEntry>));
                    ser.WriteObject(Scheduler, this);
                    Scheduler.Flush();
                    Scheduler.Close();
                    ser = null;
                }
                catch (Exception ex)
                {
                    throw new Exception("Ошибка сохранения расписания заданий: " + ex.Message);
                }
            }

            public void Remove(KickOffSchedulerEntry value)
            {
                base.Remove(value);
                Save();
            }

            public void Add( KickOffSchedulerEntry value)
            {
                base.Add(value);
                Save();
            }
                                    
        }


        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern IntPtr WTSOpenServer(string pServerName);

        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern IntPtr WTSOpenServerEx(string pServerName);

        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern void WTSCloseServer(IntPtr pServerHandle);

        [DllImport("wtsapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool WTSEnumerateSessionsEx(
        // Always use IntPtr for handles, and then something derived from SafeHandle in .NET 2.0
        IntPtr hServer,
        // Not required, but a good practice.
        [MarshalAs(UnmanagedType.U4)]
        ref int pLevel,
        [MarshalAs(UnmanagedType.U4)]
        int Filter,
        // You are going to create the memory block yourself.
        ref IntPtr ppSessionInfo,
        [MarshalAs(UnmanagedType.U4)]
        ref int pCount);

        [DllImport("wtsapi32.dll", CharSet = CharSet.Auto)]
        public static extern bool WTSLogoffSession(
        IntPtr hServer,
        [MarshalAs(UnmanagedType.U4)]
        int SessionID,
        [MarshalAs(UnmanagedType.Bool)]
        bool bWait);

        public enum WtsInfoClass
        {
            WTSInitialProgram,
            WTSApplicationName,
            WTSWorkingDirectory,
            WTSOEMId,
            WTSSessionId,
            WTSUserName,
            WTSWinStationName,
            WTSDomainName,
            WTSConnectState,
            WTSClientBuildNumber,
            WTSClientName,
            WTSClientDirectory,
            WTSClientProductId,
            WTSClientHardwareId,
            WTSClientAddress,
            WTSClientDisplay,
            WTSClientProtocolType,
            WTSIdleTime,
            WTSLogonTime,
            WTSIncomingBytes,
            WTSOutgoingBytes,
            WTSIncomingFrames,
            WTSOutgoingFrames,
            WTSClientInfo,
            WTSSessionInfo
        }

        [DllImport("kernel32.dll")]
        public static extern uint GetLastError();

        [DllImport("wtsapi32.dll")]
        private static extern void WTSFreeMemory(IntPtr pMemory);

        [DllImport("Wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSQuerySessionInformation(
            System.IntPtr hServer, 
            int sessionId,
            WtsInfoClass wtsInfoClass, 
            out System.IntPtr ppBuffer, 
            out int pBytesReturned);

        public enum WTS_CONNECTSTATE_CLASS
        {
            WTSActive,
            WTSConnected,
            WTSConnectQuery, // This was misspelled.
            WTSShadow,
            WTSDisconnected,
            WTSIdle,
            WTSListen,
            WTSReset,
            WTSDown,
            WTSInit
        };

        static PerformanceCounter cpuCounterCurrentProcess; 
        static PerformanceCounter cpuCounter;
        public static KickOffScheduler Sched;
        public static double ProcessorLoad = 0;
        public static double ProcessorLoadTotal = 0;
        public static DateTime Uptime = DateTime.Now;
        public static string strUptime = string.Empty;
        public static Dictionary<string, string> Employees = new Dictionary<string,string>();
        public static Dictionary<string, string> Clients = new Dictionary<string, string>();
        static bool isRefilling = false;
        //static SortedList SessID = new SortedList();

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct WTS_SESSION_INFO_1
        {
           // [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
            public int ExecEnvId;
            public WTS_CONNECTSTATE_CLASS State;
            public int SessionId;
            public string pSessionName;
            public string pHostName;
            public string pUserName;
            public string pDomainName;
            public string pFarmName;

        }

        public static string GetSessionClientName(IntPtr ServerHandle, int sessionid)
        {
            IntPtr buffer = IntPtr.Zero;
            int BytesReturned;
            string clientName = null;

            bool success = WTSQuerySessionInformation(ServerHandle, sessionid, WtsInfoClass.WTSClientName, out buffer, out BytesReturned);
            if (success)
            {

                clientName = Marshal.PtrToStringAnsi(buffer);
                WTSFreeMemory(buffer);
                if (ServerHandle!=IntPtr.Zero)
                    Console.WriteLine("Определили имя машины:{0}, Хэндл сервера: {1}", clientName, ServerHandle.ToInt32());
            }
            else
            {
                Console.WriteLine("Не смогли определить имя машины");
            }
            return clientName.ToLower();
        }


        //*********
        private enum WINSTATIONINFOCLASS
        {
            WinStationCreateData,
            WinStationConfiguration,
            WinStationPdParams,
            WinStationWd,
            WinStationPd,
            WinStationPrinter,
            WinStationClient,
            WinStationModules,
            WinStationInformation,
            WinStationTrace,
            WinStationBeep,
            WinStationEncryptionOff,
            WinStationEncryptionPerm,
            WinStationNtSecurity,
            WinStationUserToken,
            WinStationUnused1,
            WinStationVideoData,
            WinStationInitialProgram,
            WinStationCd,
            WinStationSystemTrace,
            WinStationVirtualData,
            WinStationClientData,
            WinStationSecureDesktopEnter,
            WinStationSecureDesktopExit,
            WinStationLoadBalanceSessionTarget,
            WinStationLoadIndicator,
            WinStationShadowInfo,
            WinStationDigProductId,
            WinStationLockedState,
            WinStationRemoteAddress,
            WinStationIdleTime,
            WinStationLastReconnectType,
            WinStationDisallowAutoReconnect,
            WinStationUnused2,
            WinStationUnused3,
            WinStationUnused4,
            WinStationUnused5,
            WinStationReconnectedFromId,
            WinStationEffectsPolicy,
            WinStationType,
            WinStationInformationEx           
        } //Works only on Windows 2000/2003

        [DllImport("winsta.dll", SetLastError = true)]
        private static extern bool WinStationQueryInformationW(
        IntPtr hServer,
        int SessionId,
        WINSTATIONINFOCLASS WinStationInformation,
        [Out] IntPtr Buf,
        uint BufLen,
        ref uint RetLen);


        public struct WinstaInfo
        {
            public int SessionId;
            public DateTime ConnectTime;
            public DateTime DisconnectTime;
            public DateTime LastInputTime;
            public DateTime LoginTime;
            public DateTime CurrentTime;
        }

        private static DateTime FileTimeToDateTime(System.Runtime.InteropServices.ComTypes.FILETIME ft)
        {
            long hFT = (((long)ft.dwHighDateTime) << 32) + ft.dwLowDateTime;
            return DateTime.FromFileTime(hFT);
        }


        [StructLayout(LayoutKind.Sequential)]
        struct WINSTATIONINFORMATIONW
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 70)]
            public Byte[] ConnectState;
            public UInt32 WinStationName;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public Byte[] LogonId;
            public System.Runtime.InteropServices.ComTypes.FILETIME ConnectTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME DisconnectTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastInputTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME LoginTime;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1096)]
            public Byte[] UserName;
            public System.Runtime.InteropServices.ComTypes.FILETIME CurrentTime;
        }

        public static TimeSpan GetSessionIdleTime(IntPtr ServerHandle, int SessionId)
        {
            uint RetLen = 0;
            bool ret;
            WINSTATIONINFORMATIONW wsInfo = new WINSTATIONINFORMATIONW();
            System.IntPtr pwsInfo = Marshal.AllocHGlobal(Marshal.SizeOf(wsInfo));
            WinstaInfo RetWsInfo = new WinstaInfo();
            try
            {
                ret = WinStationQueryInformationW(ServerHandle, SessionId, WINSTATIONINFOCLASS.WinStationInformation, pwsInfo, (uint)Marshal.SizeOf(typeof(WINSTATIONINFORMATIONW)), ref RetLen);
                if (ret)
                {
                    wsInfo = (WINSTATIONINFORMATIONW)Marshal.PtrToStructure(pwsInfo, typeof(WINSTATIONINFORMATIONW));
                    
                    RetWsInfo.ConnectTime = FileTimeToDateTime(wsInfo.ConnectTime);
                    RetWsInfo.CurrentTime = FileTimeToDateTime(wsInfo.CurrentTime);
                    RetWsInfo.DisconnectTime = FileTimeToDateTime(wsInfo.DisconnectTime);
                    RetWsInfo.LastInputTime = FileTimeToDateTime(wsInfo.LastInputTime);
                    RetWsInfo.LoginTime = FileTimeToDateTime(wsInfo.LoginTime);
                    RetWsInfo.SessionId = (int)SessionId;

                    //Console.WriteLine("Logon time: {0}", RetWsInfo.LoginTime);
                }
                else
                    Console.WriteLine("Не удалось выполнить WinStationQueryInformationW: {0}, {1}", Marshal.GetLastWin32Error(), ret);
            }
            catch (Win32Exception w32ex) {
                Console.WriteLine("Исключение при выполнении WinStationQueryInformationW");
            }
            finally
            {
                Marshal.FreeHGlobal(pwsInfo);
            }
            return RetWsInfo.CurrentTime.Subtract(RetWsInfo.LastInputTime);
        }
        //*********

        public static Dictionary<string,int> GetSessionList(IntPtr ServerHandle)
        {
            Dictionary<string,int> result = new Dictionary<string,int>();
            IntPtr buffer = IntPtr.Zero;
            int count = 0;
            int pLevel = 1;

            WTSEnumerateSessionsEx(ServerHandle, ref pLevel, 0, ref buffer, ref count);
            if (count != 0)
            {
                // Marshal to a structure array here. Create the array first.
                WTS_SESSION_INFO_1 sessionInfo = new WTS_SESSION_INFO_1();
                for (int index = 0; index < count; index++)
                {
                    try
                    {
                        // Marshal the value over.
                        sessionInfo = (WTS_SESSION_INFO_1)Marshal.PtrToStructure(buffer + (Marshal.SizeOf(sessionInfo) * index), typeof(WTS_SESSION_INFO_1));
                        // Work with the array here.
                        // Console.WriteLine("{0} {1} {2}", sessionInfo[index].pUserName, sessionInfo[index].SessionId, sessionInfo[index].State);
                        if (sessionInfo.pUserName != null)
                        {
                            string username = sessionInfo.pUserName.ToLower();

                            if (!result.ContainsKey(username))
                                result.Add(username, sessionInfo.SessionId);

                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("{0}; {1}; Ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                    }
                }
            }
            else
                Console.WriteLine("Не смогли получить список сессий с удаленного сервера. Ошибка:{0}", Marshal.GetLastWin32Error());

            return result;
        }

        public static char[] NumericChars = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.', '-', ',' };

        public static bool IsNumeric(string Value)
        {
            if (Value != null)
                return Value.ToCharArray().All(item => item.ToString().IndexOfAny(NumericChars) >= 0);
            return false;
        }

        public static SortedList GetSessionList(ref string loginlist, bool dropdisconnected = false)
        {
                // Create the pointer that will get the buffer.
                SortedList SessID = new SortedList();
                IntPtr buffer = IntPtr.Zero;
                int count = 0;
                int pLevel = 1;
                string ClientName = null;

                if (WTSEnumerateSessionsEx(ServerHandle, ref pLevel, 0, ref buffer, ref count))
                {
                    // Marshal to a structure array here. Create the array first.
                    WTS_SESSION_INFO_1 sessionInfo = new WTS_SESSION_INFO_1();
                    // Cycle through and copy the array over.

                    if (dropdisconnected)
                    {
                        Console.WriteLine("{0}; {1}; Запущен процесс сброса отключенных сессий ", DateTime.Now.ToString(), Environment.UserName);
                    }
                    for (int index = 0; index < count; index++)
                    {
                        try
                        {
                            // Marshal the value over.
                            sessionInfo = (WTS_SESSION_INFO_1)Marshal.PtrToStructure(buffer + (Marshal.SizeOf(sessionInfo) * index), typeof(WTS_SESSION_INFO_1));
                            // Work with the array here.
                            // Console.WriteLine("{0} {1} {2}", sessionInfo[index].pUserName, sessionInfo[index].SessionId, sessionInfo[index].State);
                            if (sessionInfo.pUserName != null)
                            {
                                string username = sessionInfo.pUserName.ToLower();


                                TimeSpan idle = GetSessionIdleTime(ServerHandle, sessionInfo.SessionId);

                                

                                if (idle.TotalMinutes >= 15 && sessionInfo.State == WTS_CONNECTSTATE_CLASS.WTSActive)
                                    sessionInfo.State = WTS_CONNECTSTATE_CLASS.WTSIdle;

                                //if (idle.TotalMinutes >= 1)
                                  //  Console.WriteLine("{0}; {1}; Пользователь {2}. Время простоя {3}", DateTime.Now.ToString(), Environment.UserName, username, idle);

                                if (!SessID.ContainsKey(username))
                                    SessID.Add(username, sessionInfo);
                                
                                loginlist += "'" + Server.DomainName + "\\" + username + "',";

                                if ((!username.Contains("w99sdkf0")))
                                {

                                    if (sessionInfo.State != WTS_CONNECTSTATE_CLASS.WTSActive)
                                    {
                                        if (dropdisconnected)
                                        {
                                            if (username != Environment.UserName.ToLower())
                                            {
                                                if (WTSLogoffSession(ServerHandle, sessionInfo.SessionId, false))
                                                    Console.WriteLine("{0}; {1}; Сессия {2} сброшена. ", DateTime.Now.ToString(), Environment.UserName, sessionInfo.SessionId);
                                                else
                                                    Console.WriteLine("{0}; {1}; Сессия {2} НЕ сброшена!", DateTime.Now.ToString(), Environment.UserName, sessionInfo.SessionId);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        ClientName = GetSessionClientName(ServerHandle, sessionInfo.SessionId);

                                        if (ClientName.Length > 1)
                                        {
                                            if (Clients.ContainsKey(ClientName))
                                                Clients.Remove(ClientName);
                                            Clients.Add(ClientName, username);
                                            //Console.WriteLine("Определено имя машины {0} для пользователя {1}", ClientName, username);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("{0}; {1}; Ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                        }
                    }
                    if (dropdisconnected)
                    {
                        Console.WriteLine("{0}; {1}; Завершен процесс сброса отключенных сессий ", DateTime.Now.ToString(), Environment.UserName);
                    }
                    // Close the buffer.

                   // XmlSerializer ser = new XmlSerializer(Clients.GetType());
                    try
                    {
                        FileStream ClientList = File.Open(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath.ToLower()) + "\\.clients", FileMode.OpenOrCreate);
                        DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Dictionary<string, string>));
                        ser.WriteObject(ClientList, Clients);

                        ClientList.Flush();
                        ClientList.Close();
                        ser = null;
                    }
                    catch
                    {
                    }

                    WTSFreeMemory(buffer);
                    //Console.WriteLine(loginlist);
                }
                
                if (dropdisconnected)
                    return GetSessionList(ref loginlist);
                else
                    return SessID;
        }

        public static string GetEmployeeEmail(string Username)
        {
            string useremail = string.Empty;

            if (SQLPresent)
            {
                using (SqlDataReader reader = Utils.ConnectionPool.ExecuteReader(@"select 
                                                            useremail
                                                         from dlfe.employees 
                                                         Where 
                                                            UserName like '%\\" + Username + "' and status = 0"))
                {
                    while (reader.Read())
                    {
                        useremail = reader[0].ToString();
                    }
                    reader.Close();
                    
                }
            }

            return useremail;
        }

        public static void RefillLists(bool dropdisconnected = false)
        {
            if (isRefilling) return;

            isRefilling = true;
            string loginlist = string.Empty;
            SessID = GetSessionList(ref loginlist, dropdisconnected);

            Server.sessionListHTML = "<select name=\"sessionid\">";
            Server.secIDListHTML = "<select name=\"login\" class=\"loginlist\">";
            WTS_SESSION_INFO_1 Session;
            string name;
            string fullname = string.Empty;
            /////////////////////
            //SQLPresent = false;
            /////////////////////
            if (SQLPresent)
            {                                                                                       
                try
                {
                    {
                        using (SqlDataReader reader = Utils.ConnectionPool.ExecuteReader(@"
select 
SUBSTRING(UserName,PATINDEX('%\%', UserName) + 1,100) as Userlogin
,   LastName + ' ' + FirstName As UserFIO
from dlfe.employees 
Order by LastName"))
                        {
                            while (reader.Read())
                            {

                                string userlogin = reader[0].ToString().ToLower();

                                if (!Employees.ContainsKey(userlogin))
                                    Employees.Add(userlogin, reader[1].ToString());

                                if (!SessID.ContainsKey(userlogin)) 
                                        continue;

                                Session = ((WTS_SESSION_INFO_1)SessID[userlogin]);
                                if (Session.State != WTS_CONNECTSTATE_CLASS.WTSActive)
                                {
                                    string Colour = "#CCCCCC";
                                    if(Session.State == WTS_CONNECTSTATE_CLASS.WTSIdle)
                                        Colour = "#FFCCCC";


                                    Server.sessionListHTML += "<option style=\"background-color:" + Colour + ";\" value='" + Session.SessionId + "'>" + reader[1] + "</option>\r\n";
                                    Server.secIDListHTML += "<option style=\"background-color:" + Colour +";\" value='" + userlogin + "'>" + reader[1] + "</option>\r\n";
                                }
                                else
                                {
                                    Server.secIDListHTML += "<option value='" + userlogin + "'>" + reader[1] + "</option>\r\n";
                                    Server.sessionListHTML += "<option value='" + Session.SessionId + "'>" + reader[1] + "</option>\r\n";
                                };

                            }
                            reader.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    isRefilling = false;
                    Console.WriteLine("{0}; {1}; Ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                    return;
                };
            }
            else
                if(Server.InDomain)
            {

                DirectoryEntry searchRoot;
                searchRoot = new DirectoryEntry("LDAP://" + Server.DomainController, null, null, AuthenticationTypes.Secure);
                DirectorySearcher search = new DirectorySearcher(searchRoot);
                loginlist = loginlist.Replace(Server.DomainName + "\\", "(samaccountname=").Replace(",", ")").Replace("'","");
                search.Filter = "(&(objectClass=user)(objectCategory=person)(|"+loginlist+"))";
               // search.PropertiesToLoad.Add("cn,samaccountname");
                SearchResult result;
                SearchResultCollection resultCol = search.FindAll();

                if (resultCol != null)
                    for (int counter = 0; counter < resultCol.Count; counter++)
                    {
                        result = resultCol[counter];
                        if (result.Properties.Contains("samaccountname"))
                        {
                            name = result.Properties["samaccountname"][0].ToString().ToLower();
                            fullname = name;
                            if (result.Properties.Contains("cn"))
                                fullname = result.Properties["cn"][0].ToString().Replace(name, "").TrimEnd();

                            Session = (WTS_SESSION_INFO_1)SessID[name];

                            if (Session.State == WTS_CONNECTSTATE_CLASS.WTSDisconnected)
                            {
                                Server.sessionListHTML += "<option style=\"background-color:#CCCCCC;\" value='" + Session.SessionId + "'>" + fullname + "</option>\r\n";
                                Server.secIDListHTML += "<option style=\"background-color:#CCCCCC;\" value='" + Session.pUserName + "'>" + fullname + "</option>\r\n";
                            }
                            else
                            {
                                Server.sessionListHTML += "<option value='" + Session.SessionId + "'>" + fullname + "</option>\r\n";
                                Server.secIDListHTML += "<option value='" + Session.pUserName + "'>" + fullname + "</option>\r\n";
                            };

                            if (!Employees.ContainsKey(Session.pUserName.ToLower()))
                                Employees.Add(Session.pUserName.ToLower(), fullname);
                        }
                    }


            }
                else
                {
                    loginlist = loginlist.Replace(Server.DomainName + "\\", "(name=").Replace(",", ")").Replace("'", "");
                    using (var accountSearcher = new System.Management.ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_UserAccount WHERE Disabled=False"))
                    using (var accountCollection = accountSearcher.Get())
                        foreach (var account in accountCollection)
                        {
                            name = account["Name"].ToString().ToLower();
                            fullname = account["FullName"].ToString().ToLower(); ;
                            if (fullname == string.Empty)
                                fullname = name;
                            if (SessID.ContainsKey(name))
                            {

                                Session = (WTS_SESSION_INFO_1)SessID[name];

                                if (Session.State == WTS_CONNECTSTATE_CLASS.WTSDisconnected)
                                {
                                    Server.sessionListHTML +=
                                        "<option style=\"background-color:#CCCCCC;\" value='" + Session.SessionId +
                                        "'>" + fullname + "</option>\r\n";
                                    Server.secIDListHTML += "<option style=\"background-color:#CCCCCC;\" value='" +
                                                            Session.pUserName + "'>" + fullname + "</option>\r\n";
                                }
                                else
                                {
                                    Server.sessionListHTML += "<option value='" + Session.SessionId + "'>" +
                                                              fullname + "</option>\r\n";
                                    Server.secIDListHTML += "<option value='" + Session.pUserName + "'>" + fullname +
                                                            "</option>\r\n";
                                }

                                if (!Employees.ContainsKey(Session.pUserName.ToLower()))
                                    Employees.Add(Session.pUserName.ToLower(), fullname);
                            }
                        }                    
                }
            Server.sessionListHTML += "</select>";
            Server.secIDListHTML += "</select>";

            Server.DBlistHTML = "<select name=\"filepath\" class=\"dblist\" size=3 onchange=\"whoisinconfig(this)\">";

            try
            {
                using (RegistryKey reg = Registry.CurrentUser.OpenSubKey(@"Software\1C\1Cv7\7.7\Titles"))
                    if (reg != null)
                    {
                        string[] dblist = reg.GetValueNames();
                        foreach (string dbpath in dblist)
                        {

                            if (UserDef.UserDefworks.IsConfigRunning(dbpath))
                                Server.DBlistHTML += "<option style=\"background-color:#CCCCCC;\" value='" + dbpath + @"usrdef\users.usr'>" + reg.GetValue(dbpath) + "[" + dbpath + "]" + "</option>\r\n";
                            else
                                Server.DBlistHTML += "<option value='" + dbpath + @"usrdef\users.usr'>" + reg.GetValue(dbpath) + "[" + dbpath + "]" + "</option>\r\n";
                        }
                    }
                Server.DBlistHTML += @"<option value='\\vv1000070155pc\Deploy\Application\usrdef\users.usr'>ЦЕНТР</option>\r\n";
                Server.DBlistHTML += "</select>";
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0}; {1}; Ошибка заполнения списка баз: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                Server.DBlistHTML = "<b>Ошибка зполнения</b>";
            }

            isRefilling = false;
        }

        public static int GetSessionIDByUserName(string UserName)
        {
            if (SessID.Contains(UserName))
                return ((WTS_SESSION_INFO_1)SessID[UserName]).SessionId;
            else
            {
                RefillLists();
                if (SessID.Contains(UserName))
                    return ((WTS_SESSION_INFO_1)SessID[UserName]).SessionId;
            };
            return -1;
        }
        static bool iscontrol = false; 

        public static void CounterTick(dynamic source=null, ElapsedEventArgs e=null)
        {
            ProcessorLoad = Math.Round((ProcessorLoad + cpuCounterCurrentProcess.NextValue() / Environment.ProcessorCount) / 2, 2);
            ProcessorLoadTotal = Math.Round((ProcessorLoadTotal + cpuCounter.NextValue()) / 2, 2);
            TimeSpan span = DateTime.Now - Uptime;
            strUptime = string.Format("В строю: <b>{0}д. {1}ч. {2}м. {3}с.</b>", span.Days, span.Hours, span.Minutes, span.Seconds);
            if (Process.GetCurrentProcess().WorkingSet64  > 70 * 1024 * 1024)
            {
                GC.Collect();
                GC.Collect();
                GC.WaitForFullGCComplete(500);

                if (Process.GetCurrentProcess().WorkingSet64 > 950 * 1024 * 1024)
                {
                    Console.WriteLine("{0}; {1}; Превышение порога используемой памяти, сервис будет перезапущен", DateTime.Now.ToString(), Process.GetCurrentProcess().WorkingSet64 / 1024 / 1024);
                    Server.Restart();
                }
            }   
        }

        public Utils()
        {
            Sched = new KickOffScheduler();
            cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_total");
            cpuCounterCurrentProcess = new PerformanceCounter("Process", "% Processor Time", Process.GetCurrentProcess().ProcessName);
            string CurrentUser = Environment.UserName;

            //DEBUG
           // ServerHandle = WTSOpenServer("ruvvos20003srv");
            ///

            //Console.WriteLine("{0}; {1}; Начато первичное заполнение списков процессов пользователей : ", DateTime.Now.ToString(), CurrentUser);
//            thisConnection = new SqlConnection(@"Data Source=PB3000070115PC\SQL2012;Integrated Security=SSPI;Initial Catalog=Portal");
//#if DEBUG
//            thisConnection = new SqlConnection(@"Data Source=PB3000070115PC\SQL2012;Integrated Security=SSPI;Initial Catalog=Portal");
//#else
//            thisConnection = new SqlConnection(@"Data Source=RUVVOS20005SRV;Integrated Security=SSPI;Initial Catalog=Portal");
//#endif


            if (ConnectionPool.Current.State != System.Data.ConnectionState.Open)
            {
                try
                {
                    ConnectionPool.Current.Open();
                    SQLPresent = true;
                }
                catch
                {
                    SQLPresent = false;
                }
            }
            else
            {
                SQLPresent = true;
            }


            if (File.Exists(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath.ToLower()) + "\\.clients"))
            {
                try
                {
                    FileStream ClientList = File.Open(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath.ToLower()) + "\\.clients", FileMode.OpenOrCreate);
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Dictionary<string, string>));
                    Clients = (Dictionary<string, string>)ser.ReadObject(ClientList);
                    ClientList.Close();
                    ser = null;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("{0}; {1}; Ошибка чтения кэша рабочих станций: {2}", DateTime.Now.ToString(), CurrentUser, ex);
                }
            }

            RefillLists();
            System.Timers.Timer tt = new System.Timers.Timer();
            tt.Interval = 1000;
            tt.Elapsed += tt_Elapsed;
            tt.Start();
            //Console.WriteLine("{0}; {1}; Завершено первичное заполнение списков процессов пользователей : ", DateTime.Now.ToString(), CurrentUser);
        }

        void tt_Elapsed(object sender, ElapsedEventArgs e)
        {
            CounterTick(sender);
        }

    }

    // Класс-обработчик клиента
    class Client
    {
        private HttpListenerContext thisClient;
        public SqlConnection thisConnection;
        string username = string.Empty;
        private bool SendMessage(string strMessage, int HeaderType = 200, string[] header = null, bool admin = false, bool closeConnection = true, string autorefresh = "100000;URL='/../")
        {
            byte[] byteBuffer;
            //bool closeConnection = true;
            bool result = true;
            try
            {
                string HTML_styles = @"
<style type=""text/css"">
INPUT       {border:1px solid #94ADCC;margin:1px;width:240px;background-color:transparent;FONT:bold 10px Arial;COLOR:#442222;}
INPUT:hover {background-color:E0E0EF}
.radio      {border: none; margin:1px;width:20px; background-color:transparent;FONT:bold 10px Arial;COLOR:#442222;}
.b_w        {border: 1px solid #94ADCC;margin: 1px; background-color:#FFFFFF; font: bold 10px Arial; color: #4FAA21;}
A           {FONT: 12px Tahoma;COLOR:#004990; TEXT-DECORATION:none;}
A:hover	    {FONT: 12px Tahoma;COLOR:#45a1cf; TEXT-DECORATION:underline;}
SELECT      {border:1px solid #94ADCC; margin:1px; width:300px; background-color:TRANSPARENT;FONT: 10px Arial;COLOR:#442222;}
BODY        {background-color:#F2F6Fb; FONT: 11px Arial;COLOR:#442222; PADDING: 0px; MARGIN: 0px;}
.text       {FONT: 10px Arial; COLOR:#442222; MARGIN-LEFT: 0px}
A.locker    {FONT: bold 10px Arial; COLOR:#442222; MARGIN-LEFT: 0px}
A.locker:hover     {FONT: bold 10px Arial; COLOR:#442222; MARGIN-LEFT: 0px}
.head1      {FONT: 20px ""Times New Roman"";COLOR: #004990;LINE-HEIGHT: 1.1;MARGIN-LEFT: 8px;}
.head2      {FONT: 16px ""Times New Roman"";COLOR: #45a1cf;LINE-HEIGHT: 1.2;MARGIN-LEFT: 8px;}
.white      {background-color: white;}
.base       {background-color:#F2F6Fb; FONT: 11px Arial;COLOR:#442222; MARGIN-LEFT: 20px}
span.top    {FONT: 20px ""Arial"";COLOR: #ff5500;LINE-HEIGHT: 1.2;MARGIN-LEFT: 8px;}
table       {background-color:#FFFFFF; border: 8px solid white; border-top:0px; }

.panel   {background-color: #F2F6FB; border: 8px solid white; vertical-align: top; MARGIN-LEFT: 12px; font: 11px Arial; font-color:#442222;}
.pheader {border-top:0px; background-color: #E8E8F8; border: 0px solid white; vertical-align: middle; text-align:left; MARGIN-LEFT: 12px; font:bold 11px Arial; font-color:#442222;}

.topmenu {background-color: #E8E8F8; border: 0px; vertical-align: middle; font: 11px arial;  text-align:center; border-right: #C7E3F2 1px solid; border-spacing: 0px;}
.topmenu:hover {background: #C7E3F2;}
A.menu {FONT: 11px arial; COLOR:#004990; TEXT-DECORATION:none; text-align:center;}

</style>
";

                string Message = @"
<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Strict//EN"">
<html xmlns=""http://www.w3.org/1999/xhtml"">
<html><head>" + HTML_styles + @"
<meta charset=windows-1251>
<meta http-equiv=""X-UA-Compatible"" content=""IE=8"" />
<meta http-equiv=""refresh"" content=""" + autorefresh + @""">
<script src=""/resources/jquery.js""></script>
<link id=""DATEBOXCSS"" href=""/portal/shared/_jscB/jscalendar/skins/dl/theme.css"" type=""text/css"" rel=""stylesheet"" />
<style type=""text/css"">
	.CodeMirror {border-top: 1px solid black; border-bottom: 1px solid black;}
</style>

</head>
<body>
<title>«Сименс Финанс» | Сервис технической помощи</title>
<script> 
function act(obj)
{
    if(obj.formAction == undefined)
    {
        $(obj.form).attr(""action"",obj.id);
    }
        else
    {
        $(obj.form).attr(""action"",obj.formAction);
    }
    obj.form.submit();
}
";
                if (admin) Message += @"

$(document).ready(function()
{  
    show(); 
    setInterval('show()',1000); 
   
}); 

function show()  
{  
    $.ajaxSetup({cache: true});
    $.ajax({  
        url: ""cpu_counter.html"",  
        cache: false,  
        success: function(html){  
            $(""#CPU"").html(html);  
        },
        error: function(){  
            $(""#CPU"").html(""<span class=top>Сервис не доступен!</span>"");  
        }  
  
    });  
}  

";

                Message += @"
</script>
<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">

<tr width=""100%"">
<td width=""62%""> 
                <div>
                    <span class=""head1"">Сервис технической помощи на сервере <b>" + Environment.MachineName + @"</b></span>
                    <br>";
                if (admin) Message += @"<span class=""head2""><a class=""head2"" href=""/?asuser=True"" onmouseover=""this.innerHTML='Переключиться в интерфейс пользователя'"" onmouseout=""this.innerHTML='Интерфейс Администратора'"">Интерфейс Администратора.</a></span>";
                else Message += @"<span class=""head2""> Ваш логин Siemens: " + username + "</span>";
                Message += @"
                </div>
</td>
<td width=""20%"">";
                if (admin) Message += @"<div id=CPU></div>";
                Message += @"
<td width=""18%""> 
<a href='/../' ><img border=""0"" src=""/resources/top_rlogo.gif""></a>
</td>   
</tr>
</table>

<div style=""background-color: #E8E8F8; height: 20px;"">
";
                if (admin) Message += @"
<table cellpadding=0 cellspasing=0 border=0 style=""background-color: #E8E8F8; width: 100%; height: 100%; border: 0px; border-spacing: 0px;"">
<tr style=""background-color: #E8E8F8; height: 15px;"" >
<td width=1%><td>
<td class=topmenu><a class=menu href='../'>Главная</a><td>
<td class=topmenu><a class=menu href='../colorer'>Раскраска</a><td>
<td class=topmenu><a class=menu href='../sqlquery'>Запрос к порталу</a><td>
<td class=topmenu><a class=menu href='../v7configreports'>Конфигурации и отчеты</a><td>

<td width=60%><td>
</tr>
</table>
                ";
                Message += @"
</div>

<div class=""white""><br></div>

<div class=""base"">
";
                //closeConnection = true;
                if (!strMessage.Contains("<script"))
                    strMessage = strMessage.Replace("\n", "<br>");
                Message += strMessage + @"<br><a href='/../' >В начало</a></div></body></html>";
                byteBuffer = Encoding.GetEncoding(1251).GetBytes(Message);
                thisClient.Response.StatusCode = HeaderType;
                thisClient.Response.AddHeader("Cache-Control", "max-age=86400");//, must-revalidate

                if (thisClient.Request.Headers["Accept-Encoding"].Contains("deflate"))
                {

                    MemoryStream MS = new MemoryStream(byteBuffer);
                    thisClient.Response.ContentEncoding = Encoding.GetEncoding(1251);
                    thisClient.Response.ContentType = "text/html";
                    thisClient.Response.AppendHeader("Content-Encoding", "deflate");
                    DeflateStream gzhtml = new System.IO.Compression.DeflateStream(thisClient.Response.OutputStream, System.IO.Compression.CompressionMode.Compress);
                    MS.CopyTo(gzhtml);
                    gzhtml.Close();
                }
                else
                {
                    thisClient.Response.ContentType = "text/html";
                    thisClient.Response.ContentEncoding = Encoding.GetEncoding(1251);
                }
                thisClient.Response.OutputStream.Write(byteBuffer, 0, byteBuffer.Length);
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0}; {1}; Ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                closeConnection = true;
                result = false;
            }

            if (closeConnection)
            {
                try
                {
                    thisClient.Response.Close();
                }
                catch //(Exception ex)
                {
                    // Console.WriteLine("{0}; {1}; Ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                }
            }

            return result;
        }



        private string GetUserSecurityID(string UserName, string DomainName = null)

        {
            if (DomainName == null)
                DomainName = Server.DomainName;

            string UserSID = string.Empty;
            ManagementObjectSearcher objAccount =
            new ManagementObjectSearcher("root\\CIMV2",
                                            "SELECT * FROM Win32_UserAccount Where Domain = '" + DomainName + "' and Name = '" + UserName + "'");
            foreach (ManagementObject UserEntry in objAccount.Get())
            {
                UserSID = UserEntry["SID"].ToString();
            }
            return UserSID;
        }

        private bool UnlockFile(string fileName, bool sendmessage = true)
        {
            string mess = "Запрос разблокировки файла <b>" + fileName + "</b>\r\n";
            try
            {
                string utilname = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\handle.exe";
                if (!File.Exists(utilname))
                {
                    //System.IO.File.WriteAllBytes(@"handle.exe", Properties.Resources.handle);
                    //if (!File.Exists("handle.exe"))
                    {
                        Console.WriteLine("{0}; : Ошибка разблокировки файла {1}: утилита handle.exe не найдена.", DateTime.Now.ToString(), fileName);
                        if (sendmessage)
                            SendMessage("Ошибка разблокировки файла " + fileName + ": утилита handle.exe не найдена.", 200, null, true, true);
                        return false;
                    }
                }

                Process tool = new Process();
                tool.StartInfo.FileName = utilname;
                tool.StartInfo.WorkingDirectory = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
                tool.StartInfo.Arguments = fileName + " /accepteula";
                tool.StartInfo.UseShellExecute = false;
                tool.StartInfo.RedirectStandardOutput = true;
                tool.Start();
                tool.WaitForExit();
                string outputTool = tool.StandardOutput.ReadToEnd();
                int count = 0;

                string matchPattern = @"(?<=\s+pid:\s+)\b(\d+)\b(?=\s+)";
                foreach (Match match in Regex.Matches(outputTool, matchPattern))
                {
                    count++;
                    mess += "Закрыт процесс " + int.Parse(match.Value) + "\r\n";
                    Process.GetProcessById(int.Parse(match.Value)).Kill();
                    Console.WriteLine("{0}; : Закрыт процесс ; {1}", DateTime.Now.ToString(), match.Value);
                }
                if (count == 0)
                {
                    if (outputTool.Contains("error"))
                    {
                        mess += "Ошибка разблокировки файла\r\n.";
                        mess += outputTool;
                        Console.WriteLine("{0}; : Ошибка разблокировки файла ; {1}", DateTime.Now.ToString(), outputTool);
                    }
                    else
                    {
                        mess += "Не найдено блокирующих процессов.";
                        Console.WriteLine("{0}; : Не найдено блокирующих процессов; {1}", DateTime.Now.ToString(), fileName);
                    }

                }
                if (sendmessage)
                    SendMessage(mess, 200, null, true, true, "5;URL='/../");
                return true;
            }
            catch (Exception ex)
            {
                if (sendmessage)
                    SendMessage("Ошибка разблокировки файла " + fileName + ": " + ex.Message, 200, null, true);
                Console.WriteLine("{0}; : Ошибка разблокировки файла - \"{1}\" ; {2}", DateTime.Now.ToString(), ex.Message, fileName);
                return false;
            }

        }

        struct ConfigSession
        {
            public string Name;
            public string RunMode;
            public DateTime TimeStamp;
        }

        // Конструктор класса. Ему нужно передавать принятого клиента от TcpListener
        public Client(HttpListenerContext Context)
        {
            this.thisClient = Context;
            this.thisConnection = Utils.ConnectionPool.Current;

            // Объявим строку, в которой будет хранится запрос клиента
            string Request = string.Empty;
            string domain = string.Empty;
            string host = string.Empty;
            string Option = string.Empty;
            string Command = string.Empty;
            bool admin = false;
            bool asuser = false;
            int timeshift = 0;
            int UserCenterCode = -1;
            string mess;

            #region in connection
            Stream OutStream = thisClient.Response.OutputStream;
            StreamReader InStream = new StreamReader(thisClient.Request.InputStream, Encoding.GetEncoding(1251));
            thisClient.Response.KeepAlive = true;
            thisClient.Response.AddHeader("Keep-Alive", "timeout = 20000 max = 1000");
            bool gzip = false;

            if (thisClient.Request.Headers.AllKeys.Contains("Accept-Encoding"))
                gzip = thisClient.Request.Headers["Accept-Encoding"].Contains("deflate");


            Request = thisClient.Request.RawUrl;
            try
            {
                //Request = thisClient.Request.RawUrl;
                thisClient.Response.KeepAlive = true;
                if (Request.Length == 0)
                {
                    thisClient.Response.Close();
                    return;
                }


                thisClient.Response.Headers.Set(HttpResponseHeader.CacheControl, "no-cache");

                if (Request.ToLower().Contains("cpu_counter.html"))
                {   //отдадим счетчики
                    mess = @"<span class=""text"">Нагрузка CPU (сервис):<b>" + Utils.ProcessorLoad +
                                       @"%</b><br>Нагрузка CPU (общая):<b>" + Utils.ProcessorLoadTotal + @"%</b>
                                              <br>Использовано памяти: <b>" + (Process.GetCurrentProcess().WorkingSet64 / 1024).ToString() + @" K</b>
                                              <br>" + Utils.strUptime + "</span>";
                    byte[] HeadersBuffer = Encoding.UTF8.GetBytes(mess);
                    thisClient.Response.StatusCode = 200;
                    thisClient.Response.Headers.Clear();
                    thisClient.Response.Headers.Add("Cache-Control: no-store, no-cache, must-revalidate");
                    thisClient.Response.Headers.Add("Expires: Mon, 26 Jul 1997 05:00:00 GMT");
                    thisClient.Response.Headers.Add("Pragma: no-cache");

                    thisClient.Response.ContentType = "text/html";
                    OutStream.Write(HeadersBuffer, 0, HeadersBuffer.Length);
                    thisClient.Response.Close();
                    return;
                }

                if (Request.ToLower().Contains("/resources/") || Request.ToLower().Contains("/portal/"))
                { //отдадим jquery

                    bool getfromportal = Request.ToLower().Contains("/portal/");
                    byte[] buf = null;

                    string filename = Path.GetFileNameWithoutExtension(Request);
                    string restype = Path.GetExtension(Request).ToLower();

                    if (!getfromportal)
                    {
                        System.Resources.ResourceManager ResMan = Properties.Resources.ResourceManager;
                        object res = ResMan.GetObject(filename.Replace('-', '_'));
                        if (res == null)
                        {
                            thisClient.Response.StatusCode = 404;
                            thisClient.Response.Close();
                            return;
                        }
                        thisClient.Response.Headers.Clear();

                        switch (restype)
                        {

                            case ".js":
                                buf = Encoding.GetEncoding(1251).GetBytes(((string)res));
                                thisClient.Response.ContentType = "application/x-javascript";
                                thisClient.Response.ContentEncoding = Encoding.GetEncoding(1251);
                                gzip = true;
                                break;
                            case ".css":
                                buf = Encoding.ASCII.GetBytes((string)res);
                                thisClient.Response.ContentEncoding = Encoding.ASCII;
                                break;
                            case ".jpg":
                            case ".gif":
                                thisClient.Response.ContentType = "image/gif";
                                buf = (byte[])new System.Drawing.ImageConverter().ConvertTo(res, typeof(byte[]));
                                gzip = true;
                                break;
                            case ".rdp":
                                thisClient.Response.ContentType = "application/octet-stream";
                                thisClient.Response.AddHeader("Content-disposition", "attachment; filename=\"" + Path.GetFileName(Request) + "\"");

                                buf = (byte[])res;
                                gzip = false;
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        if (!Server.FileCache.ContainsKey(filename))
                        {
                            string servername = "https://sfs.siemens.ru/";
                            WebRequest req = WebRequest.Create(Request.Replace("/portal/", servername));
                            req.UseDefaultCredentials = true;
                            Stream resp = req.GetResponse().GetResponseStream();
                            int portion = (int)req.GetResponse().ContentLength;
                            buf = new byte[portion];
                            int total = 0;
                            int readbytes = resp.Read(buf, total, portion);

                            while (readbytes > 0)
                            {
                                total += readbytes;
                                Array.Resize<byte>(ref buf, buf.Length + portion);
                                readbytes = resp.Read(buf, total, portion);
                            }
                            total += readbytes;
                            Array.Resize<byte>(ref buf, total);

                            resp.Close();
                            try
                            {
                                Server.FileCache.Add(filename, buf);
                            }
                            catch
                            {

                            }
                        }
                        else
                        {
                            buf = Server.FileCache[filename];
                        }

                        switch (restype)
                        {

                            case ".js":
                                thisClient.Response.ContentType = "application/x-javascript";
                                thisClient.Response.ContentEncoding = Encoding.GetEncoding(1251);
                                gzip = true;
                                break;
                            case ".css":
                                thisClient.Response.ContentEncoding = Encoding.ASCII;
                                gzip = true;
                                break;
                            case ".jpg":
                            case ".gif":
                                thisClient.Response.ContentType = "image/gif";
                                gzip = true;
                                break;
                            case ".rdp":
                                thisClient.Response.ContentType = "application/octet-stream";
                                thisClient.Response.AddHeader("Content-disposition", "attachment; filename=\"" + Path.GetFileName(Request) + "\"");
                                gzip = false;
                                break;
                            default:
                                break;
                        }


                    }


                    thisClient.Response.AppendHeader("Connection", "keep-alive");
                    thisClient.Response.AppendHeader("Cache-Control", "max-age=186400000");

                    if (gzip)
                    {

                        MemoryStream MS = new MemoryStream(buf);
                        thisClient.Response.ContentEncoding = Encoding.GetEncoding(1251);
                        thisClient.Response.AppendHeader("Content-Encoding", "deflate");
                        DeflateStream gzhtml = new System.IO.Compression.DeflateStream(thisClient.Response.OutputStream, System.IO.Compression.CompressionMode.Compress);
                        MS.CopyTo(gzhtml);
                        gzhtml.Close();
                    }
                    else
                    {
                        thisClient.Response.ContentLength64 = buf.Length;
                        OutStream.Write(buf, 0, buf.Length);
                        OutStream.Flush();
                    }
                    thisClient.Response.Close();
                    return;
                }


            }
            catch (Exception ex)
            {
                SendMessage("Вы не авторизованы", 200, null, false);
                Console.WriteLine("{0}; {1}; Ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                return;
            }

            //    if (Utils.ProcessorLoad > 10)
            //    {
            //        SendMessage("В настоящий момент сервер перегружен. Попробуйте через минуту.", 200, null, false, true, "30");
            //        
            //        return;
            //    }
            #endregion


            Request = InStream.ReadToEnd();
            if (thisClient.Request.HttpMethod != "POST")
                Request = thisClient.Request.RawUrl;

            Match ReqMatch;
            {
                domain = thisClient.User.Identity.Name.Split('\\')[0];
                username = thisClient.User.Identity.Name.Split('\\')[1].ToLower();
                /* Получаем строку запроса */
                if (username == string.Empty)
                {
                    SendMessage("Не удалось определить Вашу учетную запись", 200, null, false);
                    return;
                }
                try
                {
                    if (!Server.admins.Contains(username))
                    {
                        if (Utils.SQLPresent)
                        {
                            if (thisConnection.State != System.Data.ConnectionState.Open) thisConnection.Open();
                            using (SqlDataReader reader = Utils.ConnectionPool.ExecuteReader(@"  Use portal
                                                                    select 
                                                                    CenterCode,
                                                                    [ServerTimeOffset] as TimeShiftFromVladivostok
                                                                    from dlfe.employees as emp
                                                                    left join dlfe.OfficeList as ol on emp.officeid = ol.officeid
                                                                    left join [dlfe].[WorkingHours] as wh on ol.WorkingHoursId = wh.RecID
                                                                    Where 
                                                                    UserName like '%" + username + @"'
                                                                    and status = 0 "))

                                // using (SqlDataReader reader = command.ExecuteReader())
                                try
                                {
                                    while (reader.Read())
                                    {
                                        UserCenterCode = Convert.ToInt32(reader[0]);
                                        int.TryParse(reader[1].ToString(), out timeshift);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    SendMessage("Не удалось получить информацию об учетной записи <b>" + username + @"</b><br>Ошибка:" + ex.Message, 200, null, false);
                                    Console.WriteLine("{0}; {1}; Ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                                    return;
                                };

                            admin = (UserCenterCode == 8);
                            if (admin)
                                Console.WriteLine("{0}; {1}; Сотрудник {2} относится к ЦФО 8, внесен в чписок администраторов", DateTime.Now.ToString(), Environment.UserName, username);

                            if (UserCenterCode == -1)
                            {
                                SendMessage("Сотрудник с учетной записью <b>" + username + @"</b> не найден", 200, null, false);
                                return;
                            }
                        }
                        else
                        {
                            if (username == Environment.UserName)
                            {
                                admin = true;
                            }
                        }

                        if (admin)
                            Server.admins.Add(username);

                    }
                    else
                    {
                        admin = true;
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("{0}; {1}; Ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                    Console.WriteLine("{0}; {1}; Не удалось проверить учетную запись:", DateTime.Now.ToString(), username);
                    SendMessage("Не удалось проверить учетную запись <b>" + username + @"</b><br>Ошибка:" + ex.Message, 200, null, false);
                    return;
                }
                // admin = false;

                if (admin)
                {
                    if (Request.ToLower().Contains("loginlist.scr"))
                    {
                        Utils.RefillLists();
                        byte[] buf = Encoding.UTF8.GetBytes(Server.secIDListHTML.Replace("<select name=\"login\" class=\"loginlist\">", "").Replace("</select>", ""));
                        thisClient.Response.StatusCode = 200;
                        thisClient.Response.ContentType = "text/html";

                        if (gzip)
                        {

                            MemoryStream MS = new MemoryStream(buf);
                            thisClient.Response.ContentEncoding = Encoding.GetEncoding(1251);
                            thisClient.Response.AppendHeader("Content-Encoding", "deflate");
                            DeflateStream gzhtml = new System.IO.Compression.DeflateStream(OutStream, System.IO.Compression.CompressionMode.Compress);
                            MS.CopyTo(gzhtml);
                            gzhtml.Close();
                        }
                        else
                            OutStream.Write(buf, 0, buf.Length);

                        thisClient.Response.Close();
                        return;
                    }

                    if (Request.ToLower().Contains("dblist.scr"))
                    {
                        Utils.RefillLists();
                        byte[] buf = Encoding.UTF8.GetBytes(Server.DBlistHTML.Replace("<select name=\"filepath\" class=\"dblist\" size=3 onchange=\"whoisinconfig(this)\">", "").Replace("</select>", ""));
                        thisClient.Response.StatusCode = 200;
                        thisClient.Response.ContentType = "text/html";
                        if (gzip)
                        {

                            MemoryStream MS = new MemoryStream(buf);
                            thisClient.Response.ContentEncoding = Encoding.GetEncoding(1251);
                            thisClient.Response.AppendHeader("Content-Encoding", "deflate");
                            DeflateStream gzhtml = new System.IO.Compression.DeflateStream(OutStream, System.IO.Compression.CompressionMode.Compress);
                            MS.CopyTo(gzhtml);
                            gzhtml.Close();
                        }
                        else
                            OutStream.Write(buf, 0, buf.Length);
                        OutStream.Write(buf, 0, buf.Length);
                        thisClient.Response.Close();
                        return;
                    }

                    if (Request.ToLower().Contains("showlog.scr"))
                    {
                        byte[] buf = new byte[Server.logfileR.Length];

                        try
                        {
                            Server.logfileR.Seek(0, SeekOrigin.Begin);
                            buf = Encoding.UTF8.GetBytes((new StreamReader(Server.logfileR)).ReadToEnd().Replace("\r", "<br>"));
                            thisClient.Response.StatusCode = 200;
                        }
                        catch (Exception ex)
                        {
                            buf = Encoding.UTF8.GetBytes("Ошибка получения содержимого лога: " + ex.Message);
                        }

                        if (gzip)
                        {

                            MemoryStream MS = new MemoryStream(buf);
                            thisClient.Response.ContentEncoding = Encoding.GetEncoding(1251);
                            thisClient.Response.AppendHeader("Content-Encoding", "deflate");
                            DeflateStream gzhtml = new System.IO.Compression.DeflateStream(OutStream, System.IO.Compression.CompressionMode.Compress);
                            MS.CopyTo(gzhtml);
                            gzhtml.Close();
                        }
                        else
                            OutStream.Write(buf, 0, buf.Length);

                        thisClient.Response.Close();
                        return;
                    }

                    if (Request.Contains("/temp/"))
                    {
                        string filename = System.Web.HttpUtility.UrlDecode(Request);
                        filename = filename.Replace("/temp/", Path.GetTempPath()).Replace("/", "\\");
                        byte[] buf = null;

                        if (Server.FileCache.ContainsKey(filename))
                        {
                            buf = Server.FileCache[filename];
                            if (File.Exists(filename))
                                File.Delete(filename);
                        }
                        else
                            if (File.Exists(filename))
                        {
                            buf = File.ReadAllBytes(filename);
                            Server.FileCache.Add(filename, buf);
                        }
                        if (buf != null)
                        {
                            thisClient.Response.ContentEncoding = Encoding.UTF8;
                            thisClient.Response.ContentType = thisClient.Request.AcceptTypes[0];
                            if (gzip)
                            {
                                MemoryStream MS = new MemoryStream(buf);
                                thisClient.Response.ContentEncoding = Encoding.GetEncoding(1251);
                                thisClient.Response.AppendHeader("Content-Encoding", "deflate");
                                DeflateStream gzhtml = new System.IO.Compression.DeflateStream(OutStream, System.IO.Compression.CompressionMode.Compress);
                                MS.CopyTo(gzhtml);
                                gzhtml.Close();
                            }
                            else
                                OutStream.Write(buf, 0, buf.Length);
                            OutStream.Flush();
                        }
                        thisClient.Response.Close();
                        return;
                    }
                }

                Command = thisClient.Request.Url.AbsolutePath.Replace("/", "");

                if (thisClient.Request.HttpMethod == "GET")
                {
                    ReqMatch = Regex.Match(Request, @"^/(.*)\?(.*)", RegexOptions.IgnoreCase);
                    Request = ReqMatch.Groups[2].Value;
                }

                Dictionary<string, string> Options = new Dictionary<string, string>();
                try
                {
                    foreach (string opt in Request.Split('&'))
                    {
                        if (opt != "")
                        {
                            string[] item = opt.Split('=');
                            if (item.Length > 1)
                            {
                                Options.Add(item[0].ToLower(), Encoding.GetEncoding(1251).GetString(Encoding.Unicode.GetBytes(Regex.Unescape(item[1].Replace("%", @"\x").Replace('+', ' ')))).Replace("\0", ""));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    SendMessage("Ошибка чтения параметров: " + ex.Message, 200, null, admin, true, null);
                }

                Option = Options.ElementAtOrDefault(0).Key;
                Request = Options.ElementAtOrDefault(0).Value;

                if (Request == null)
                    Request = string.Empty;

                #region Default page

                if (Options.ContainsKey("asuser"))
                    asuser = bool.Parse(Options["asuser"]);



                if ((Command == string.Empty)) //Command == string.Empty
                {
                    Utils.RefillLists();
                    string body = string.Empty;


                    if ((!admin) || (asuser))
                    {
                        string databaseListHTML = @"<select name=""filename"">";

                        try
                        {
                            string securityID = GetUserSecurityID(username, domain);
                            using (RegistryKey reg = Registry.Users.OpenSubKey(securityID + "\\Software\\1C\\1Cv7\\7.7\\Titles"))
                                if (reg != null)
                                {
                                    string[] dblist = reg.GetValueNames();
                                    foreach (string dbpath in dblist)
                                    {

                                        if (UserDef.UserDefworks.IsConfigRunning(dbpath))
                                            databaseListHTML += "<option style=\"background-color:#CCCCCC;\" value='" + dbpath + @"usrdef\users.usr'>" + reg.GetValue(dbpath) + "[" + dbpath + "]" + "</option>\r\n";
                                        else
                                            databaseListHTML += "<option value='" + dbpath + @"usrdef\users.usr'>" + reg.GetValue(dbpath) + "[" + dbpath + "]" + "</option>\r\n";
                                    }
                                }
                            databaseListHTML += "</select>\r\n";
                        }
                        catch (Exception ex)
                        {
                            //SendMessage("Не удалось получить список баз 1С из реестра<br>Ошибка:" + ex.Message, 200, null, admin);
                            Console.WriteLine("{0}; {1}; Ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                            databaseListHTML += "</select>\r\n";
                            //return;
                        }

                        string SchedulerTasks = @"<table width=30%><tr><td class=""pheader"" colspan=""2"">Текущие задания</td></tr>";
                        if (Utils.Sched != null)
                            foreach (Utils.KickOffSchedulerEntry Entry in Utils.Sched)
                            {
                                if (Entry.Username == username)
                                {
                                    SchedulerTasks += string.Format(@"
<form>
<tr class=""panel"">
    <td width=*>Сброс сессии в <b>{0} по Владивостоку</b><input type=hidden name=taskid value={1}></td>
    <td><input style=""width: 12px;"" align=""center"" type=submit formAction=sched_remove id=sched_remove class=b_w value='Х' onclick=""act(this)""></td>
</tr>
</form>", Entry.Time, Entry.id);
                                }
                            }

                        SchedulerTasks += "</table>";
                        string hours = "<select name=hour  style=\"width: 50px;\">";
                        for (int i = 0; i < 24; i++)
                            hours += string.Format("<option value = {0:d2}>{0:d2}</option>", i);
                        hours += "</select>";

                        string minutes = "<select name=minute  style=\"width: 50px;\">";
                        for (int i = 0; i < 60; i++)
                            minutes += string.Format("<option value = {0:d2}>{0:d2}</option>", i);
                        minutes += "</select>";

                        body = @"
    <script></script>
    <table width=100% >
        <tr>
            <td class=""pheader"">Утилиты пользователя</td>
        </tr>
        <tbody class=""panel"">
                <tr>
                    <td width=*>
                        <form action="""" method=""post"">
                            <input type=hidden name=login value='" + username + @"'>
                            <br><input type=submit formAction=""resetsession"" id=""resetsession"" class=""b_w"" value='Завершить сессию' onclick=""act(this)"">
                            <br><input type=submit formAction=""resetprinter"" id=""resetprinter"" class=""b_w"" value='Сбросить настройки печати 1С' onclick=""act(this)"">
                            <br><input type=submit formAction=""resetwindow"" id=""resetwindow"" class=""b_w"" value='Сбросить настройки окна 1С' onclick=""act(this)"">
                        </form>
                        <form  method=""post"">
                            <input type=hidden name=accountname value=" + username + @">
                            
                            <br><input type=submit formAction=""db_findinlog"" id=""db_findinlog"" class=""b_w"" value='Узнать кто блокирует ДФА:' onclick=""act(this)"">
                            <input type=hidden name=objecttype value='Refs'>
                            <input type=hidden name=eventtype  value='RefOpen'>
                            <input type=hidden name=eventdate  value='" + DateTime.Today.ToString() + @"'> <input name=searchfilter value=''> 
в базе " + databaseListHTML + @"
                            <br><a href='/resources/_1C_General.rdp'>Получить ярлык для входа в 1С</a>                            

                        </form> 
                        <form  method=""post"">
                             <input type=hidden name=accountname value=" + username + @">
                             <input type=hidden name=tasktypeid value=1>
                             <input type=submit formAction=""sched_addtask"" id=""sched_addtask"" class=""b_w"" value='Добавить сброс сессии по расписанию' onclick=""act(this)""> 
                             Час*: " + hours + @" минута: " + minutes + @"
                             <span style=""font:8px Arial;COLOR:#664444;"">  *Время указывается по Владивостоку</span>
                        </form>
                        " + SchedulerTasks + @"
                    </td>
                </tr>
                <tr>
                    <td class=""pheader"">Обратиться за помощью:</td>
                </tr>
                <tr> 
                    <td>
                    <table width=100% >
                    <tr><td class=""pheader"">Владивосток (" + DateTime.ParseExact("09:00", "HH:mm", null).AddHours(timeshift).ToShortTimeString() + "-" + DateTime.ParseExact("18:00", "HH:mm", null).AddHours(timeshift).ToShortTimeString() + ((DateTime.Now.Hour < 18) && (DateTime.Now.Hour > 7) ? "" : " - рабочий день окончен") + @"):</td></tr><td>
                        <a href=""sip:alexander.berko@siemens.com"">Александр Берко</a>
                        <br>    <a href=""sip:kristina.davydova@siemens.com"">Кристина Давыдова</a>
                        <br>    <a href=""sip:nina.podvyaznikova@siemens.com"">Нина Подвязникова</a>
                        <br>    <a href=""sip:maxim.chernushko@siemens.com"">Максим Чернушко</a>
                        <br>    <a href=""sip:igor.soloviev@siemens.com"">Игорь Соловьев</a>
                        <br>    <a href=""sip:oleg.sheyko@siemens.com"">Олег Шейко</a> </td>
                    <tr><td class=""pheader"">Санкт-Петербург (" + DateTime.ParseExact("15:00", "HH:mm", null).AddHours(timeshift).ToShortTimeString() + "-" + DateTime.ParseExact("00:00", "HH:mm", null).AddHours(timeshift).ToShortTimeString() + ((DateTime.Now.Hour > 14) && (DateTime.Now.Hour <= 23) ? "" : " - рабочий день окончен") + @"):</td></tr><td>
                            <a href=""sip:dmitry.dreytser@siemens.com"">Дмитрий Дрейцер</a>
                        </td>
                    </table>
                    </td>
                </tr>
        </tbody>
    </table>";
                    }
                    else
                    {
                        //string sessionListHTML = Server.sessionListHTML;
                        string secIDListHTML = Server.secIDListHTML;
                        string databaseListHTML = Server.DBlistHTML;
                        body = @"
<script>
function updateloginList()
{
    $("".loginlist"").each(function ()
          {$(this).html(""<option>Список обновляется...</option>"")});
    
    $.ajaxSetup({cache: false});
    $.ajax({
            url: ""loginlist.scr"",
            cache: false,
            success: function(html)
                {
                    $("".loginlist"").each(function ()
                        {$(this).html(html)})
                } 
            });
};

function updatedbList()
{
    $("".dblist"").each(function ()
          {$(this).html(""<option>Список обновляется...</option>"")});
    $.ajaxSetup({cache: false});
    $.ajax({
            url: ""dblist.scr"",
            cache: false,
            success: function(html)
                {
                    $("".dblist"").each(function ()
                        {$(this).html(html)})
                } 
            });
};

function whoisinconfig(list)
{
    $(""#dbinfo"").html('Статус обновляется...');
    var resp = $.get('dbinfo.scr',{filepath:$(list).val()})
        .success(
                function(html)
                {
                    $(""#dbinfo"").html(html);
               })
        .error(function() {$(""#dbinfo"").html('Ошибка получения имени сотрудника, занявшего конфигуратор');});
}
</script>

    <table width=""*"" >
         <tr>
           <td colspan=""2"" rowspan=""1"" class=""pheader"">Работа с пользователями.</td>
         </tr>

        <tbody class=""panel"">
            <tr>
                <td width=50%>
                    <form id=usersaction  method=""post"">
                    Сотрудник*: " + secIDListHTML + @" <input style=""width: 100px;"" value='Обновить список' onclick=""updateloginList()"" type=button><br>
                    <span style=""font:8px Arial;COLOR:#664444;"">  *Серым выделены отключенные сессии, розовым - простаивающие дольше 15 минут</span>
                    <br><input  type=submit formAction=""resetsession"" id=""resetsession"" class=""b_w"" value='Завершить сессию' onclick=""act(this)"">
                    <br><input  type=submit formAction=""resetprinter"" id=""resetprinter"" class=""b_w"" value='Сбросить настройки печати 1С' onclick=""act(this)"">
                    <br><input  type=submit formAction=""resetwindow"" id=""resetwindow""  class=""b_w"" value='Сбросить настройки окна 1С' onclick=""act(this)"">
                    </form>
                    <form action=""resetsession"" method=""post"">
                            <input type=submit class=""b_w"" value='Завершить все отключенные сессии'>
                            <input type=hidden name=dropdisconnected value='true'>
                    </form>
                    <form action=""viewsessions"" method=""post"">
                        <br><input  type=submit formAction=""viewsessions"" id=""viewsessions""  class=""b_w"" value='Просмотреть список сессий' onclick=""act(this)"">
                    </form>
                </td>
                <td>
<script src=""/resources/mlgparse.js"" type=""text/javascript""></script>
<form id=dbaction  method=""post"" name=dbaction>
База*:<span id=""dbinfo"" class=""text""> </span>
<br>" + databaseListHTML + @" <input style=""width: 100px;"" value='Обновить список' onclick=""updatedbList()"" type=button>
<br><span style=""font:8px Arial;COLOR:#664444;"">  *Серым выделены SQL-базы в которых запущен конфигуратор</span>
<br><input  formAction=""unlockfile"" id=""unlockfile"" type=submit class=""b_w"" value='Разблокировать список пользователей базы' onclick=""act(this)"">
<br><input  formAction=""db_edituserlist"" id=""db_edituserlist"" type=submit class=""b_w"" value='Редактировать список пользователей базы' onclick=""act(this)"">
<br><input  formAction=""db_shownightlog"" id=""db_shownightlog"" type=submit class=""b_w"" value='Показать лог ночных операций' onclick=""act(this)"">

<div>
    <br><input type=submit formAction=""db_findinlog"" id=""db_findinlog"" class=""b_w"" value='Просмотреть историю лога 1C' onclick=""act(this)"">
                       
    <br>Объект: <select name=objecttype id=objecttype style=""width: 250px;"" onchange=""logoptionsselect()"">
			    <option value ='Refs' selected >Справочник</option>
			    <option value ='Docs'>Документ</option> 
			    <option value ='Distr'>Обмен</option> 
                <option value ='Restruct'>Изменение конфигурации</option> 
                <option value ='Grbgs'>Ошибки</option> 
    </select>
	

	    <br>Событие: <select name=eventtype id=eventtype style=""width: 250px;"">
            <option value=''>Все</option>
            <option selected value='RefOpen'>Открыт</option>
            <option value='RefWrite'>Записан</option>
            <option value='RefNew'>Создан</option>
            <option value='RefMarkDel'>Помечен на удаление</option>
            <option value='RefUnmarkDel'>Снята пометка удаления</option>
            <option value='RefDel'>Удален</option>
            <option value='RefGrpMove'>Перенесен в другую группу</option>
            <option value='RefAttrWrite'>Значение реквизита записано</option>
            <option value='RefAttrDel'>Значение реквизита удалено</option>
    </select>
    <script src=""/portal/shared/_jscB/jscalendar/calendar.js"" type=""text/javascript""></script>
    <script src=""/portal/shared/_jscB/jscalendar/lang/calendar-ru_win_.js"" type=""text/javascript""></script>
    <script src=""/portal/shared/_jscB/jscalendar/calendar-setup.js"" type=""text/javascript""></script>
    <script src=""/portal/shared/_jscB/inputmask.js"" type=""text/javascript""></script>

    <br>Глубина выборки: <input type=text id='eventdate' name='eventdate' value='" + DateTime.Today.AddDays(-1).ToShortDateString() + @"' style=""width: 70px;"">

    <script type=""text/javascript"">
        //<![CDATA[

{
        var eventdate = InputMask(document.getElementById('eventdate'), '##.##.####', {'space':'_'});
        var __cultureInfo = {""name"":""ru-RU"",""numberFormat"":{""CurrencyDecimalDigits"":2,""CurrencyDecimalSeparator"":"","",""IsReadOnly"":true,""CurrencyGroupSizes"":[3],""NumberGroupSizes"":[3],""PercentGroupSizes"":[3],""CurrencyGroupSeparator"":"" "",""CurrencySymbol"":""р."",""NaNSymbol"":""NaN"",""CurrencyNegativePattern"":5,""NumberNegativePattern"":1,""PercentPositivePattern"":1,""PercentNegativePattern"":1,""NegativeInfinitySymbol"":""-бесконечность"",""NegativeSign"":""-"",""NumberDecimalDigits"":2,""NumberDecimalSeparator"":"","",""NumberGroupSeparator"":"" "",""CurrencyPositivePattern"":1,""PositiveInfinitySymbol"":""бесконечность"",""PositiveSign"":""+"",""PercentDecimalDigits"":2,""PercentDecimalSeparator"":"","",""PercentGroupSeparator"":"" "",""PercentSymbol"":""%"",""PerMilleSymbol"":""‰"",""NativeDigits"":[""0"",""1"",""2"",""3"",""4"",""5"",""6"",""7"",""8"",""9""],""DigitSubstitution"":1},""dateTimeFormat"":{""AMDesignator"":"""",""Calendar"":{""MinSupportedDateTime"":""\/Date(-62135596800000)\/"",""MaxSupportedDateTime"":""\/Date(253402264799999)\/"",""AlgorithmType"":1,""CalendarType"":1,""Eras"":[1],""TwoDigitYearMax"":2029,""IsReadOnly"":true},""DateSeparator"":""."",""FirstDayOfWeek"":1,""CalendarWeekRule"":0,""FullDateTimePattern"":""d MMMM yyyy \u0027г.\u0027 H:mm:ss"",""LongDatePattern"":""d MMMM yyyy \u0027г.\u0027"",""LongTimePattern"":""H:mm:ss"",""MonthDayPattern"":""MMMM dd"",""PMDesignator"":"""",""RFC1123Pattern"":""ddd, dd MMM yyyy HH\u0027:\u0027mm\u0027:\u0027ss \u0027GMT\u0027"",""ShortDatePattern"":""dd.MM.yyyy"",""ShortTimePattern"":""H:mm"",""SortableDateTimePattern"":""yyyy\u0027-\u0027MM\u0027-\u0027dd\u0027T\u0027HH\u0027:\u0027mm\u0027:\u0027ss"",""TimeSeparator"":"":"",""UniversalSortableDateTimePattern"":""yyyy\u0027-\u0027MM\u0027-\u0027dd HH\u0027:\u0027mm\u0027:\u0027ss\u0027Z\u0027"",""YearMonthPattern"":""MMMM yyyy"",""AbbreviatedDayNames"":[""Вс"",""Пн"",""Вт"",""Ср"",""Чт"",""Пт"",""Сб""],""ShortestDayNames"":[""Вс"",""Пн"",""Вт"",""Ср"",""Чт"",""Пт"",""Сб""],""DayNames"":[""воскресенье"",""понедельник"",""вторник"",""среда"",""четверг"",""пятница"",""суббота""],""AbbreviatedMonthNames"":[""янв"",""фев"",""мар"",""апр"",""май"",""июн"",""июл"",""авг"",""сен"",""окт"",""ноя"",""дек"",""""],""MonthNames"":[""Январь"",""Февраль"",""Март"",""Апрель"",""Май"",""Июнь"",""Июль"",""Август"",""Сентябрь"",""Октябрь"",""Ноябрь"",""Декабрь"",""""],""IsReadOnly"":true,""NativeCalendarName"":""григорианский календарь"",""AbbreviatedMonthGenitiveNames"":[""янв"",""фев"",""мар"",""апр"",""май"",""июн"",""июл"",""авг"",""сен"",""окт"",""ноя"",""дек"",""""],""MonthGenitiveNames"":[""января"",""февраля"",""марта"",""апреля"",""мая"",""июня"",""июля"",""августа"",""сентября"",""октября"",""ноября"",""декабря"",""""]},""eras"":[1,""A.D."",null,0]};

        //]]>

        Calendar.setup({""ifFormat"":""%d.%m.%Y"",""daFormat"":""%d.%m.%Y"", ""range"" : [" + DateTime.Today.AddYears(-4).Year + @"," + DateTime.Today.Year + @"],  ""firstDay"":1,""showsTime"":false,""showOthers"":true,""inputField"":""eventdate"",""button"":""eventdate""});
}
    </script>

    <br>Представление объекта: <input name=searchfilter value=''>
</div>
</form>
                </td>
            </tr>
        </tbody>
    
        <tr>
           <td colspan=""2"" rowspan=""1"" class=""pheader"">Утилиты для обмена</td>
        </tr>

        <tbody class=""panel"">
            <tr>
                    <td width=50%>
                        <form action=""runlocal"" method=""post"">
                            <input type=submit class=""b_w"" value='Выгрузить данные из RN2'>
                            <input type=hidden name=ruvvos20003srv value='""D:\1C\Additions\1C\Утилиты для обменов\Начать обмен.cmd""'> Выгрузит данные из RN2 для начала дневного обмена
                        </form><form action=""runlocal"" method=""post"">
                            <input type=submit class=""b_w"" value='Провести обмен в базе RN2'>
                            <input type=hidden name=ruvvos20003srv value='""D:\1C\Additions\1C\Утилиты для обменов\RN2.cmd""'> Начнет обмен в базе RN2, предварительно выгнав всех пользователей и отобрав права на запуск 1С. <b>ВНИМАНИЕ!</b> Всех выбросит из базы! </b>
                        </form>
                    </td>
                    <td>
                        <form action=""runremote"" method=""post"">
                                <input type=submit class=""b_w"" value='Выгрузить данные из центра'>
                                <input type=hidden name=VV1000070142NB value='""C:\Deploy\Additions\1C\Утилиты для обменов\Выгрузить центр.cmd""'> Выгрузит данные из центра во все периферийные базы
                        </form>
                        <form action=""runremote"" method=""post"">
                                <input type=submit class=""b_w"" value='Загрузить данные в центр'>
                                <input type=hidden name=VV1000070142NB value='""C:\Deploy\Additions\1C\Утилиты для обменов\Exchange.cmd"" RN2'> Загрузит данные в Центр из RN2 для окончания обмена
                        </form>
                    </td>
            </tr>
        </tbody>
        <tr>
           <td class=""pheader"">Лог операций</td>
           <td class=""pheader"">Служебные</td>
        </tr>

        <tbody class=""panel"">
            <tr>
                <td>
                    <form method=""post"">
                        <input type=hidden name=dummy value='dummy'>
                        <input formAction=""showlog"" id=""showlog"" type=submit class=""b_w""  value='Показать лог' onclick=""act(this)"">
                        <br><input type=submit  formAction=""savelog"" id=""savelog"" class=""b_w"" value='Cохранить лог' onclick=""act(this)"">
                    </form>
                </td>
                <td>
                    <form method=""post"">
                        <input type=hidden name=do value='true'>    
                        <input type=submit formAction=""update"" id=""update"" class=""b_w"" value='Обновить версию' onclick=""act(this)"">
                        <br><input type=submit formAction=""restart"" id=""restart"" class=""b_w"" value='Перезапустить сервис' onclick=""act(this)"">
                    </form>
                </td>
            </tr>
        </tbody>    
    </table>";

                    }
                    SendMessage(body, 200, null, admin, true, "-1");
                    return;
                }
                #endregion

                #region Command executing
                else
                {
                    Request = Request.ToLower();
                    username = username.ToLower();
                    switch (Command)
                    {
                        case "dbinfo.scr":
                            {
                                string filepath = null;
                                if (Options.ContainsKey("filepath"))
                                    filepath = Options["filepath"];

                                if (filepath == null)
                                {
                                    thisClient.Response.Close();
                                    return;
                                }

                                string DBCatalog = UserDefworks.GetDBCatalog(filepath);

                                if (DBCatalog == string.Empty)
                                {
                                    thisClient.Response.Close();
                                    return;
                                }

                                mess = "";
                                string filename = DBCatalog + "SYSLOG\\links.tmp";
                                string locker = null;
                                string ServerName = Path.GetPathRoot(DBCatalog).Replace(@"\\", "").Split('\\')[0];

                                bool RemoteServer = (ServerName.Length > 2) && (ServerName != Environment.MachineName);

                                DateTime TimeStamp = DateTime.Now;

                                try
                                {

                                    if (File.Exists(filename))
                                    {
                                        FileStream fs = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                                        StreamReader links = new StreamReader(fs);
                                        string strLinks = links.ReadToEnd();
                                        links.Close();
                                        links.Dispose();
                                        fs.Close();
                                        fs.Dispose();

                                        ConfigSession LastSession = new ConfigSession();
                                        LastSession.TimeStamp = DateTime.MinValue;
                                        LastSession.Name = null;

                                        foreach (string entry in strLinks.Split(new string[] { "}}" }, StringSplitOptions.RemoveEmptyEntries))
                                        {
                                            ConfigSession Session = new ConfigSession();
                                            Session.TimeStamp = DateTime.MinValue;
                                            Session.RunMode = null;
                                            Session.Name = null;

                                            foreach (string item in entry.Replace("\r\n", "").Replace(" ", "").Replace("{{", "").Replace("},{", ";").Replace("{", "").Split(';'))
                                            {
                                                string[] parameter = item.Replace("\"", "").Split(',');
                                                if (parameter.Length > 1)
                                                {
                                                    string option = parameter[0];
                                                    string value = parameter[1];
                                                    switch (option)
                                                    {
                                                        case "Name":
                                                            Session.Name = value.ToLower();
                                                            break;
                                                        case "Runmode":
                                                            Session.RunMode = value;
                                                            break;
                                                        case "Date&Time":
                                                            Session.TimeStamp = DateTime.Parse(value.Replace(".", "/").Replace(",", " "));
                                                            break;
                                                    }
                                                }
                                            }

                                            if (Session.RunMode == "C")
                                            {
                                                if (Session.TimeStamp > LastSession.TimeStamp)
                                                    LastSession = Session;
                                            }

                                        }

                                        locker = LastSession.Name;
                                        bool isSystemUser = false;

                                        if (locker != null)
                                        {
                                            IntPtr ServerHandle = Utils.ServerHandle;

                                            if (RemoteServer)
                                            {
                                                ServerHandle = Utils.WTSOpenServerEx(ServerName);
                                            }

                                            Dictionary<string, int> SessionList = Utils.GetSessionList(ServerHandle);
                                            string ClientName = string.Empty;

                                            //Utils.RefillLists();
                                            // 
                                            {
                                                if (locker.Contains("w99sdkf0"))
                                                {
                                                    isSystemUser = true;
                                                    int SessionID = -1;

                                                    if (SessionList.ContainsKey(locker))
                                                        SessionID = SessionList[locker];

                                                    if (SessionID != -1)
                                                    {
                                                        ClientName = Utils.GetSessionClientName(ServerHandle, SessionID);

                                                        if (ClientName.Length > 1)
                                                        {
                                                            if (Utils.Clients.ContainsKey(ClientName))
                                                                locker = Utils.Clients[ClientName];
                                                        }
                                                        else
                                                        {
                                                            Console.WriteLine("Сессия системного пользователя {0} отключена.", locker);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (!Utils.Employees.ContainsKey(locker))
                                                        locker = locker.Substring(0, locker.Length - 1);
                                                }

                                            }

                                            if (RemoteServer)
                                                Utils.WTSCloseServer(ServerHandle);

                                            if (Utils.Employees.ContainsKey(locker))
                                            {
                                                string useremail = Utils.GetEmployeeEmail(locker);
                                                if (useremail != string.Empty)
                                                    useremail = string.Format("<a class='locker' href='sip:{0}'>", useremail);

                                                if (UserDefworks.IsConfigRunning(DBCatalog))
                                                    if (isSystemUser)
                                                        mess = "Сейчас в конфигураторе: <b>" + useremail + Utils.Employees[locker] + "</a> от имени системного пользователя</b>";
                                                    else
                                                        mess = "Сейчас в конфигураторе: <b>" + useremail + Utils.Employees[locker] + "</a></b>";
                                                else
                                                    if (isSystemUser)
                                                    mess = "Конфигуратор последний раз был открыт от имени системного пользователя. <br> Сейчас под системным пользователем на сервере:<b>" + useremail + Utils.Employees[locker] + "</a></b>";
                                                else
                                                    mess = "Последним был в конфигураторе: <b>" + useremail + Utils.Employees[locker] + "</a></b>";
                                            }
                                            else
                                            {
                                                if (isSystemUser)
                                                    if (UserDefworks.IsConfigRunning(DBCatalog))
                                                        mess = "Конфигуратор открыт под системным пользователем - " + locker + ", с клиента - " + ClientName + ". Сотрудника определить не удалось.";
                                                    else
                                                        mess = "Конфигуратор последний раз был открыт под системным пользователем - " + locker + ", с клиента - " + ClientName + ". Сотрудника определить не удалось.";
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    mess = "Ошибка получения имени сотрудника, занявшего конфигуратор: " + ex.Message;
                                }
                                byte[] buf = Encoding.UTF8.GetBytes(mess);
                                thisClient.Response.StatusCode = 200;
                                thisClient.Response.ContentType = "text/html";
                                OutStream.Write(buf, 0, buf.Length);
                                thisClient.Response.Close();
                                return;
                            }
                        case "resetsession":
                            {
                                if (!admin)
                                    if (username != Request)
                                    {
                                        SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                        Console.WriteLine("{0}; {1}; Несанкционированный запрос; {2}", DateTime.Now.ToString(), username, Request);
                                        return;
                                    }

                                int SessionID = -1;

                                string userlogin = Request;
                                switch (Option)
                                {
                                    case "login":
                                        {
                                            userlogin = Request;
                                            SessionID = Utils.GetSessionIDByUserName(Request);
                                            break;
                                        }
                                    case "sessionid":
                                        {
                                            SessionID = Convert.ToInt32(Request);
                                            break;
                                        }
                                    case "dropdisconnected":
                                        {
                                            SessionID = -2;
                                            break;
                                        }
                                    default:
                                        {
                                            SendMessage("Неизвестный параметр: " + Option, 200, null, admin, true);
                                            return;
                                        }
                                }
                                mess = string.Empty;
                                bool dropdisconnected = false;
                                if (SessionID > 0)
                                {
                                    mess = "Запрос на завершение сессии пользователя " + userlogin + " от " + thisClient.Request.RemoteEndPoint.ToString() + " (Siemens GID: " + username + ")";
                                    Console.WriteLine("{0}; {1}; Запрос на завершение сессии; {2}", DateTime.Now.ToString(), username, userlogin);

                                    if (SessionID == Process.GetCurrentProcess().SessionId || userlogin.Contains("w99sdkf0"))
                                    {
                                        mess = "Отклонен запрос на завершение сессии системного пользователя.";
                                        Console.WriteLine("{0}; {1}; Отклонен на завершение сессии системного пользователя: {2}", DateTime.Now.ToString(), Environment.UserName, userlogin);
                                    }
                                    else
                                    {
                                        try
                                        {
                                            if (Utils.WTSLogoffSession(IntPtr.Zero, SessionID, true))
                                            {
                                                Console.WriteLine("{0}; {1};Сессия {2} закрыта; {3}", DateTime.Now.ToString(), username, SessionID, userlogin);
                                                mess += "\r\nСессия " + SessionID + " закрыта.";
                                            }
                                            else
                                            {
                                                Console.WriteLine("{0}; {1}; Ошибка закрытия сессии; {2}", DateTime.Now.ToString(), username, userlogin);
                                                mess += "\r\nОшибка закрытия сессии " + SessionID + ". ";
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            mess += "\r\nОшибка закрытия сессии " + SessionID + ": " + ex.Message;
                                            Console.WriteLine("{0}; {1}; Ошибка закрытия сессии: {2} в методе {3}; {4}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType, userlogin);
                                        }
                                    }
                                }
                                else
                                {
                                    switch (SessionID)
                                    {
                                        case -1:
                                            {
                                                mess = "Запрос на завершение сессии пользователя " + userlogin + " от " + thisClient.Request.RemoteEndPoint.ToString() + " (Siemens GID: " + username + ")";
                                                Console.WriteLine("{0}; {1}; Запрос на завершение сессии; {2}", DateTime.Now.ToString(), username, userlogin);
                                                mess += "\r\nНе удалось определить SessionID";
                                                Console.WriteLine("{0}; {1}; Не удалось определить SessionID; {2}", DateTime.Now.ToString(), username, userlogin);
                                                break;
                                            }
                                        case -2:
                                            {
                                                mess = "Запрос на завершение отключенных сессий от " + thisClient.Request.RemoteEndPoint.ToString() + " (Siemens GID: " + username + ")";
                                                Console.WriteLine("{0}; {1}; Запрос на завершение отключенных сессий", DateTime.Now.ToString(), username);
                                                mess += "\r\nОтключенные сессии сбрасываются.";
                                                dropdisconnected = true;
                                                break;
                                            }
                                        default:
                                            {
                                                mess += "\r\nНе удалось определить SessionID";
                                                Console.WriteLine("{0}; {1}; Не удалось определить SessionID; {2}", DateTime.Now.ToString(), username, userlogin);
                                                break;
                                            }

                                    }
                                }
                                SendMessage(mess, 200, null, admin, true, "2;URL='/../");
                                Utils.RefillLists(dropdisconnected);
                                break;
                            }
                        case "resetprinter":
                            {
                                ///
                                //SendMessage("Временно не доступно. Приносим извинения за не удобства.", 200, null, admin, true, "1;URL='/../");
                                ///

                                if (!admin)
                                {
                                    if (username != Request)
                                    {
                                        SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                        Console.WriteLine("{0}; {1}; Несанкционированный запрос сброса настроек печати; {2}", DateTime.Now.ToString(), username, Request);
                                        return;
                                    }
                                    ///
                                    //SendMessage("Временно не доступно. Приносим извинения за не удобства.", 200, null, admin, true, "1;URL='/../");
                                    //return;
                                    ///
                                }
                                string securityID = string.Empty;
                                string userlogin = string.Empty;
                                switch (Option)
                                {
                                    case "login":
                                        {
                                            securityID = GetUserSecurityID(Request);
                                            userlogin = Request;

                                            if (securityID == string.Empty)
                                            {

                                                SendMessage("Не удалось определить SecurityID пользователя " + Request, 200, null, admin, true);
                                                Console.WriteLine("{0}; {1}; Не удалось определить SecurityID пользователя; {2}", DateTime.Now.ToString(), username, userlogin);
                                                return;
                                            }
                                            break;
                                        }
                                    default:
                                        {
                                            thisClient.Response.Close();
                                            return;
                                        }
                                }
                                Console.WriteLine("{0}; {1}; Запрос на сброс настроек печати 1C; {2}", DateTime.Now.ToString(), username, userlogin);
                                mess = "Запрос на сброс настроек печати пользователя " + userlogin + "\r\n";

                                //using (
                                RegistryKey reg = Registry.Users.OpenSubKey(securityID + "\\Software\\1C\\1Cv7\\7.7\\Titles");
                                  //  )
                                {

                                    if (reg == null)
                                    {
                                        securityID = GetUserSecurityID(Request, "ww600");
                                        reg = Registry.Users.OpenSubKey(securityID + "\\Software\\1C\\1Cv7\\7.7\\Titles");
                                    }
                                    if (reg != null)
                                    {
                                        mess += "Ключ реестра пользователя <b>найден</b>\r\n";
                                        Console.WriteLine("{0}; {1}; Найден ключ реестра \"{2}\"; {3}", DateTime.Now.ToString(), username, reg.ToString(), userlogin);
                                        string[] dblist = reg.GetValueNames();
                                        foreach (string dbname in dblist)
                                        {
                                            mess += "База  <b>'" + reg.GetValue(dbname) + "'</b> - ";
                                            try
                                            {
                                                Registry.Users.DeleteSubKeyTree(securityID + @"\Software\1C\1Cv7\7.7\" + reg.GetValue(dbname) + @"\V7\" + userlogin + @"\Moxel\Default");
                                                mess += "<b>ОК!</b>\r\n";
                                                Console.WriteLine("{0}; {1}; Удалены настройки печати для базы \"{2}\"; {3}", DateTime.Now.ToString(), username, dbname, userlogin);
                                            }
                                            catch (Exception ex)
                                            {
                                                mess += "\r\nОшибка удаления ключа настроек:</b> " + ex.Message + "\r\n";
                                                Console.WriteLine("{0}; {1}; Ошибка удаления настроек печати : {2} в методе {3}; {4}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType, userlogin);
                                            }

                                        }
                                        SendMessage(mess, 200, null, admin, true, "1;URL='/../");
                                    }
                                    else
                                    {
                                        mess += "Ключ реестра <b>не найден</b> : " + securityID + "\r\n";
                                        Console.WriteLine("{0}; {1}; НЕ найден ключ реестра сотрудника \"{2}\"; {3}", DateTime.Now.ToString(), username, securityID, userlogin);
                                        SendMessage(mess, 200, null, admin, true, "5;URL='/../");
                                    }

                                }
                                break;
                            }
                        case "resetwindow":
                            {

                                if (!admin)
                                {
                                    if (username != Request)
                                    {
                                        SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                        Console.WriteLine("{0}; {1}; Несанкционированный запрос сброса настроек окна; {2}", DateTime.Now.ToString(), username, Request);
                                        return;
                                    }
                                    ///
                                    //SendMessage("Временно не доступно. Приносим извинения за не удобства.", 200, null, admin, true, "1;URL='/../");
                                    //return;
                                    ///
                                }
                                string securityID = string.Empty;
                                string userlogin = string.Empty;
                                switch (Option)
                                {
                                    case "login":
                                        {
                                            userlogin = Request;
                                            securityID = GetUserSecurityID(Request);
                                            if (securityID == string.Empty)
                                            {
                                                SendMessage("Не удалось определить SecurityID пользователя " + userlogin, 200, null, admin, true);
                                                Console.WriteLine("{0}; {1}; Не удалось определить SecurityID пользователя; {2}", DateTime.Now.ToString(), username, userlogin);
                                                return;
                                            }
                                            break;
                                        }
                                    default:
                                        {
                                            thisClient.Response.Close();
                                            return;
                                        }
                                }
                                Console.WriteLine("{0}; {1}; Запрос на сброс настроек окна 1C; {2}", DateTime.Now.ToString(), username, userlogin);
                                mess = "Запрос на сброс настроек окна пользователя " + userlogin + "\r\n";
                                //using (
                                    RegistryKey reg = Registry.Users.OpenSubKey(securityID + "\\Software\\1C\\1Cv7\\7.7\\Titles");
                                  //  )
                                {
                                    if (reg == null)
                                    {
                                        securityID = GetUserSecurityID(Request, "ww600");
                                        reg = Registry.Users.OpenSubKey(securityID + "\\Software\\1C\\1Cv7\\7.7\\Titles");
                                    }

                                    if (reg != null)
                                    {
                                        mess += "Ключ реестра пользователя <b>найден</b>\r\n";
                                        Console.WriteLine("{0}; {1}; Найден ключ реестра \"{2}\"; {3}", DateTime.Now.ToString(), username, reg.ToString(), userlogin);
                                        string[] dblist = reg.GetValueNames();
                                        foreach (string dbpath in dblist)
                                        {
                                            mess += "База  <b>'" + reg.GetValue(dbpath) + "'</b> - ";
                                            try
                                            {
                                                Registry.Users.DeleteSubKeyTree(securityID + @"\Software\1C\1Cv7\7.7\" + reg.GetValue(dbpath) + @"\V7\" + userlogin + @"\Windows");
                                                Registry.Users.DeleteSubKeyTree(securityID + @"\Software\1C\1Cv7\7.7\" + reg.GetValue(dbpath) + @"\V7\" + userlogin + @"\ToolbarSystem");
                                                mess += "<b>ОК!</b>\r\n";
                                                Console.WriteLine("{0}; {1}; Удалены настройки окна для базы \"{2}\"; {3}", DateTime.Now.ToString(), username, dbpath, userlogin);
                                            }
                                            catch (Exception ex)
                                            {
                                                mess += "\r\nОшибка удаления ключа настроек:</b> " + ex.Message + "\r\n";
                                                Console.WriteLine("{0}; {1}; Ошибка удаления настроек окна для базы: {2} в методе {3}; {4}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType, userlogin);
                                            }

                                        }
                                        SendMessage(mess, 200, null, admin, true, "1;URL='/../resetsession?login=" + userlogin);
                                    }
                                    else
                                    {
                                        mess += "Ключ реестра <b>не найден </b>: " + securityID + "\r\n";
                                        SendMessage(mess, 200, null, admin, true, "5;URL='/../");
                                        Console.WriteLine("{0}; {1}; НЕ найден ключ реестра сотрудника \"{2}\"; {3}", DateTime.Now.ToString(), username, securityID, userlogin);

                                    }


                                }
                                break;
                            }
                        case "unlockfile":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }
                                string filename = string.Empty;
                                switch (Option)
                                {
                                    case "filepath":
                                        {
                                            filename = Uri.UnescapeDataString(Request);
                                            break;
                                        }
                                    default:
                                        {
                                            thisClient.Response.Close();
                                            return;
                                        }
                                }
                                try
                                {
                                    Console.WriteLine("{0}; {1}; Запрос разблокировки файла; {2}", DateTime.Now.ToString(), username, filename);
                                    UnlockFile(filename);
                                }
                                catch (Exception ex)
                                {
                                    SendMessage("Ошибка разблокировки файла " + filename + "\r\nОписание ошибки: " + ex.Message, 200, null, admin, true);
                                    Console.WriteLine("{0}; {1}; Ошибка разблокировки файла: {2} в методе {3}; {4}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType, filename);
                                    return;
                                }
                                break;
                            }
                        case "db_shownightlog":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, admin, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }


                                string BasePath = string.Empty;
                                if (Options.ContainsKey("filepath"))
                                    BasePath = Path.GetDirectoryName(Path.GetDirectoryName(Uri.UnescapeDataString(Options["filepath"])));
                                string FileName = BasePath + "\\SYSLOG\\logfile.txt";
                                string line;
                                string[] logstring;
                                bool found = false;
                                string log = null;

                                string today = string.Format("{0:dd.MM.yy}", DateTime.Now);
                                Encoding Win1251 = Encoding.GetEncoding(866);

                                try
                                {
                                    System.IO.StreamReader file = new System.IO.StreamReader(FileName, Encoding.Default);
                                    while ((line = file.ReadLine()) != null)
                                    {
                                        if (!found)
                                        {
                                            logstring = line.Split(' ');
                                            if (logstring.Length > 3)
                                            {
                                                if (logstring[1] == today)
                                                    found = true;
                                            }
                                        }

                                        if (found)
                                        {
                                            log += Encoding.GetEncoding(1251).GetString(Encoding.Convert(Encoding.Default, Encoding.GetEncoding(1251), Encoding.Default.GetBytes(line))) + "\r\n";
                                        }
                                    }
                                    file.Close();
                                    SendMessage("<p>" + log + "</p>", 200, null, admin, true, "-1");
                                }
                                catch (Exception ex)
                                {
                                    SendMessage("<p>" + ex.Message + "</p>", 200, null, admin, true, "-1");
                                };



                                break;
                            }
                        case "showlog":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, admin, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }

                                Server.logfileR.Seek(0, SeekOrigin.Begin);
                                TextReader SR = new StreamReader(Server.logfileR);
                                string ttt = string.Empty;
                                try
                                {
                                    ttt = SR.ReadToEnd();
                                    ttt += @"
                                        <script>
                function showlog()  
                { 
                $.ajaxSetup({cache: false}); 
                $.ajax({  
                url: ""showlog.scr"",  
                cache: false,  
                success: function(html){  
                $(""#log"").html(html);  
                }  
                });  
                }  
      
                $(document).ready(function(){  
                showlog(); 
                setInterval('showlog()',2500);  
                }); 
                </script>
                <div id=""log"" class=""white""></div>
                                    ";
                                }
                                catch (Exception ex)
                                {
                                    ttt = "Ошибка: " + ex.Message;
                                }

                                SendMessage("<p>" + ttt + "</p>", 200, null, admin, true, "-1");
                                break;
                            }
                        case "savelog":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }

                                Stream outStream = thisClient.Response.OutputStream;
                                Server.logfileR.Seek(0, SeekOrigin.Begin);
                                System.IO.StreamReader file = new System.IO.StreamReader(Server.logfileR, Encoding.Default);
                                byte[] byteBuffer = Encoding.Default.GetBytes(file.ReadToEnd());

                                thisClient.Response.ContentType = "plaintext";
                                thisClient.Response.AddHeader("Content-disposition", "attachment; filename=\"kickOffServer.log\"");
                                outStream.Write(byteBuffer, 0, byteBuffer.Length);
                                thisClient.Response.Close();
                                break;
                            }
                        case "runremote":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }
                                string ttt = string.Empty;
                                if (File.Exists(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\psexec.exe"))
                                {
                                    string RemoteCommand = Uri.UnescapeDataString(Request).Replace("+", " ");
                                    Encoding Win1251 = Encoding.GetEncoding("Windows-1251");
                                    Encoding DOS866 = Encoding.GetEncoding(866);
                                    Console.WriteLine(@"{0}; {1}; psexec.exe -accepteula -h \\{2} {3}", DateTime.Now.ToString(), username, Option, RemoteCommand);
                                    Process usrLogoff = new Process();
                                    usrLogoff.StartInfo.FileName = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\psexec.exe";
                                    usrLogoff.StartInfo.Arguments = @"-accepteula -i -h \\" + Option + " " + RemoteCommand;
                                    usrLogoff.StartInfo.UseShellExecute = false;
                                    usrLogoff.StartInfo.CreateNoWindow = true;
                                    usrLogoff.StartInfo.RedirectStandardOutput = true;
                                    usrLogoff.StartInfo.StandardOutputEncoding = DOS866;
                                    string DOS_text = null;
                                    usrLogoff.Start();
                                    StreamReader stdOut = usrLogoff.StandardOutput;
                                    usrLogoff.WaitForExit();
                                    stdOut.BaseStream.Flush();
                                    Console.WriteLine("{0}; {1}; psexec.exe \\\\:{2} {3}: Код выполнения - {4} ", DateTime.Now.ToString(), username, Option, RemoteCommand, usrLogoff.ExitCode);
                                    DOS_text = stdOut.ReadToEnd();
                                    usrLogoff.Dispose();
                                    ttt += Encoding.GetEncoding(1251).GetString(Encoding.Convert(DOS866, UTF8Encoding.GetEncoding(1251), DOS866.GetBytes(DOS_text)));
                                }
                                else
                                {
                                    ttt = "Утилита psexec.exe не найдена. Выполнение команд на удаленных системах невозможно.";
                                }
                                SendMessage("<p>" + ttt + "</p>", 200, null, admin, true, "-1");
                                break;
                            }
                        case "runlocal":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }
                                string RemoteCommand = Uri.UnescapeDataString(Request).Replace("+", " ");
                                Option = Uri.UnescapeDataString(Option).Replace("+", " ");
                                Encoding Win1251 = Encoding.GetEncoding("Windows-1251");
                                Encoding DOS866 = Encoding.GetEncoding(866);
                                Console.WriteLine("{0}; {1}; cmd.exe /C {2}", DateTime.Now.ToString(), username, RemoteCommand);
                                Process usrLogoff = new Process();
                                usrLogoff.StartInfo.FileName = "cmd.exe";
                                usrLogoff.StartInfo.Arguments = "/A /C " + RemoteCommand;
                                usrLogoff.StartInfo.UseShellExecute = false;
                                usrLogoff.StartInfo.RedirectStandardOutput = true;
                                usrLogoff.StartInfo.RedirectStandardError = true;
                                usrLogoff.StartInfo.StandardOutputEncoding = DOS866;
                                usrLogoff.Start();

                                StreamReader stdOut = usrLogoff.StandardOutput;
                                usrLogoff.WaitForExit();

                                stdOut.BaseStream.Flush();
                                Console.WriteLine("{0}; {1}; cmd.exe /C {2}; Код выполнения - {3}", DateTime.Now.ToString(), username, RemoteCommand, usrLogoff.ExitCode);
                                string DOS_text = stdOut.ReadToEnd();
                                usrLogoff.Dispose();
                                //Перекодируем для разборчивости 
                                string ttt = Encoding.GetEncoding(1251).GetString(Encoding.Convert(DOS866, UTF8Encoding.GetEncoding(1251), DOS866.GetBytes(DOS_text)));
                                SendMessage("<p>" + ttt + "</p>", 200, null, admin, true, null);
                                ttt = null;
                                Win1251 = null;
                                break;
                            }
                        case "colorer":
                            {

                                string ttt = @"
<link rel=""stylesheet"" href=""/resources/default_css.css"">
<script src=""/resources/json2.js""></script>
<script src=""/resources/underscore.js""></script>
<script src=""/resources/highlight.js""></script>


<script>
    window.onready = function() {

	$('#source').on('change keyup paste mosemove blur', function() {
                        text = this.value;

                        $('pre code').each(function(i, block) {
                            $(block).html(text);
                            hljs.highlightBlock(block);
                        ;
                        });
});
}
</script>

<table width=""98%"">
<tr> <td class=""pheader"">Код 1С</td> </tr>
<tbody>
<tr> <td class=""panel""><textarea id=source class='1c base' style=""width:98%;"" rows=40'></textarea></td></tr>
</tbody>
<tr> <td class=""pheader"">Раскрашенный код для копирования и вставки</td> </tr>
<tbody class=""panel"">
<tr><td><div><pre><code class='1c'><code></pre></div></td> </tr>
</tbody>
</table>
";

                                SendMessage("<p>" + ttt + "</p>", 200, null, admin, true, null);
                                break;
                            }
                        case "sqlquery":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }

                                string Query = string.Empty;
                                bool doquery = false;

                                string ttt = string.Empty;
                                string mode = "csv";
                                bool noheaders = false;


                                if (Options.ContainsKey("query"))
                                    Query = Options["query"];

                                if (Query == string.Empty)
                                {
                                    string tCoo = thisClient.Request.Cookies["query"] != null ? thisClient.Request.Cookies["query"].Value.ToString() : "";
                                    Query = Uri.UnescapeDataString(tCoo.Replace('+', ' '));
                                }

                                if (Options.ContainsKey("doquery"))
                                    doquery = bool.Parse(Options["doquery"]);

                                if (Options.ContainsKey("mode"))
                                    mode = Options["mode"];

                                try
                                {
                                    if (Options.ContainsKey("noheaders"))
                                        if (int.Parse(Options["noheaders"]) == 1)
                                            noheaders = true;
                                }
                                catch
                                {

                                }

                                if (doquery)
                                {
                                    string CrLf = "\r\n";

                                    if ((mode != "xls") && (!noheaders))
                                        ttt = "<A HREF=\"javascript:history.go(-1)\">Вернуться в консоль</A><br><br>";

                                    if ((mode == "table") || (mode == "xls"))
                                    {
                                        ttt += "<table name='result'> <tbody> <tr class=\"pheader\">";
                                        CrLf = "</tr><tr>";
                                    }


                                    if (Utils.SQLPresent)
                                    {
                                        if (thisConnection.State != System.Data.ConnectionState.Open) thisConnection.Open();
                                        {
                                            try
                                            {
                                                if (mode.ToLower() == "json")
                                                {
                                                    SqlDataAdapter DA = new SqlDataAdapter(Query, thisConnection);
                                                    DataTable Data = new DataTable("Result");
                                                    DA.Fill(Data);

                                                    thisClient.Response.Headers.Clear();
                                                    System.Runtime.Serialization.Formatters.Binary.BinaryFormatter bf1 = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
                                                    bf1.Serialize(thisClient.Response.OutputStream, Data);
                                                    thisClient.Response.OutputStream.Flush();
                                                    thisClient.Response.Close();
                                                    return;
                                                }

                                                if (mode.ToLower() == "xls")
                                                {
                                                    SqlDataAdapter DA = new SqlDataAdapter(Query, thisConnection);
                                                    DataTable Data = new DataTable("Result");
                                                    DA.Fill(Data);
                                                    thisClient.Response.Headers.Clear();
                                                    thisClient.Response.AddHeader("content-disposition", "attachment; filename=\"Query_" + DateTime.Now.ToString().Replace(".", "").Replace(':', '-').Replace(' ', '_') + ".xls\"");
                                                    thisClient.Response.ContentType = "application/vnd.ms-excel";

                                                    System.Web.UI.WebControls.GridView excel = new System.Web.UI.WebControls.GridView();
                                                    excel.DataSource = Data;
                                                    excel.DataBind();
                                                    TextWriter SW = new StreamWriter(thisClient.Response.OutputStream, Encoding.GetEncoding(1251));
                                                    excel.RenderControl(new System.Web.UI.HtmlTextWriter(SW));
                                                    SW.Flush();
                                                    thisClient.Response.Close();
                                                    return;
                                                }

                                                using (SqlDataReader reader = Utils.ConnectionPool.ExecuteReader(Query))
                                                {
                                                    for (int i = 0; i < reader.FieldCount; i++)
                                                    {
                                                        if (mode == "csv")
                                                            ttt += reader.GetName(i) + (i == reader.FieldCount ? "" : ";");
                                                        else
                                                            ttt += "<td>" + reader.GetName(i) + "</td>";
                                                    };

                                                    if ((mode == "table") || (mode == "xls"))
                                                        ttt += "</tr></tbody> <tbody class=\"panel\"><tr>";
                                                    else
                                                        ttt += CrLf;

                                                    while (reader.Read())
                                                    {
                                                        for (int i = 0; i < reader.FieldCount; i++)
                                                        {

                                                            if (mode == "csv")
                                                                ttt += reader[i].ToString() + (i == reader.FieldCount ? "" : ";");
                                                            else
                                                                if (Utils.IsNumeric(reader[i].ToString()))
                                                                if (!reader.GetName(i).Contains("Id"))
                                                                    ttt += "<td class=number>" + reader[i].ToString() + "</td>";
                                                                else
                                                                    ttt += "<td class=recid>" + reader[i].ToString() + "</td>";
                                                            else
                                                                    if (reader[i].GetType() == typeof(DateTime))
                                                                if (((DateTime)reader[i]).Hour == 0)
                                                                    ttt += "<td class=date>" + ((DateTime)reader[i]).ToShortDateString() + "</td>";
                                                                else
                                                                    ttt += "<td class=date>" + reader[i].ToString() + "</td>";
                                                            else
                                                                ttt += "<td>" + reader[i].ToString() + "</td>";


                                                        };
                                                        ttt += CrLf;
                                                    }
                                                    if ((mode == "table") || (mode == "xls"))
                                                        ttt += "</tr></tbody></table>";
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                ttt = string.Empty;
                                                if (!noheaders)
                                                    ttt = "<A HREF=\"javascript:history.go(-1)\">Вернуться в консоль</A><br><br>";
                                                else
                                                    ttt = "ERROR: ";
                                                ttt += ex.Message;
                                            };
                                        }
                                    }
                                    else
                                    {
                                        mode = "";
                                        ttt = "Не поддерживается в данной версии.";
                                    }
                                }
                                else
                                {
                                    ttt = @"
                                                    <table width=""98%"" >
                                                         <tr>
                                                           <td colspan=""2"" rowspan=""1"" class=""pheader"">Выполнить запрос к базе портала:</b></td>
                                                         </tr>

                                                        <tbody class=""panel"">
                                                            <tr>
                                                                <td width=*>
                                                                    <form id=sqlquery  method=""post"">
                                                                        <div>
                                                                            <TEXTAREA class=panel name=query id=query style=""Width:200px"">" + Query + @"</TEXTAREA> 
                                                                        </div> 
                                                                        <input type=hidden name=doquery value=true> 
                                                                        <br><input class=radio type=radio name=mode value=table checked style=""Width:20"">Результат в виде таблицы</input>
                                                                        <br><input class=radio type=radio name=mode value=csv style=""Width:20"">Результат в виде текста с разделителями</input>
                                                                        <br><input class=radio type=radio name=mode value=xls style=""Width:20"">Результат в Excel</input>
                                                                        <br><input type=submit formAction=""sqlquery"" id=""sqlquery"" class=""b_w"" value='Выполнить запрос' onclick=""act(this)"">
                                                                    </form>
<link href=""/resources/codemirror_css.css"" rel=""stylesheet""> 
<link href=""/resources/show-hint_css.css"" rel=""stylesheet""> 
<SCRIPT src=""/resources/codemirror.js""></SCRIPT>
<script src=""/resources/sql-hint.js""></script>
<script src=""/resources/show-hint.js""></script>
<SCRIPT src=""/resources/matchbrackets.js""></SCRIPT>

<SCRIPT src=""/resources/closebrackets.js""></SCRIPT>
<script src=""/resources/sql.js""></script>






	                                                                <script>


                                                                    window.onload = function() {
                                                                      window.editor = CodeMirror.fromTextArea(document.getElementById('query'), {
                                                                        mode: {name:'text/x-mssql', globalVars: true},
                                                                        indentWithTabs: true,
                                                                        smartIndent: true,
                                                                        lineNumbers: true,
                                                                        matchBrackets : true,
                                                                        autofocus: true,
                                                                        dragDrop: true,
                                                                        viewportMargin: 27,
                                                                        lineWrapping: true,
                                                                        autoCloseBrackets: true,
                                                                        extraKeys: {
                                                                         ""Shift-Space"": ""autocomplete""
                                                                            }
                                                                       
                                                                      });

                                                                      //  editor.on(""keyup"",function(cm) { 
                                                                      //  CodeMirror.showHint(cm,CodeMirror.hint.deluge,{completeSingle: false});
                                                                      //  });
                                                                    };
                                                                    
                                                                    </script>
                                                                </td>
                                                            </tr>
                                                        </tbody>     
                                                    </table>";
                                }

                                if (Query != string.Empty)
                                {
                                    string qCoo = Uri.EscapeDataString(Query);
                                    thisClient.Response.Cookies.Add(new Cookie("query", qCoo, "/sqlquery"));
                                }
                                if (!noheaders)
                                    SendMessage("<p>" + ttt + "</p>", 200, null, admin, true, "-1");
                                else
                                {
                                    byte[] byteBuffer = Encoding.Default.GetBytes(ttt);
                                    thisClient.Response.Headers.Clear();
                                    thisClient.Response.ContentType = "plaintext";

                                    thisClient.Response.AppendHeader("Connection", "keep-alive");
                                    thisClient.Response.OutputStream.Write(byteBuffer, 0, byteBuffer.Length);
                                    thisClient.Response.OutputStream.Flush();
                                    thisClient.Response.Close();
                                }
                                break;
                            }
                        case "update":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }
                                string updateserver = @"\\ruvvos20008v01\SFS.Works\Obmen\";
                                string fileName = System.Windows.Forms.Application.ExecutablePath.ToLower();
                                string updateFileName = updateserver + Path.GetFileName(fileName);
                                string myFileVersionInfo = FileVersionInfo.GetVersionInfo(fileName).FileVersion;
                                DateTime myFileDateTime = File.GetLastWriteTime(fileName);
                                string UpdFileVersionInfo = null;
                                DateTime UpdFileDateTime = new DateTime(0);
                                if (File.Exists(updateFileName))
                                {
                                    UpdFileVersionInfo = FileVersionInfo.GetVersionInfo(updateFileName).FileVersion;
                                    UpdFileDateTime = File.GetLastWriteTime(updateFileName);
                                }
                                else
                                {
                                    SendMessage("Текущая версия: " + myFileVersionInfo + "(" + myFileDateTime + ")\r\n Файл обновления не найден\"" + updateserver + "\" не найдено", 200, null, admin, true, "-1");
                                };

                                if (UpdFileVersionInfo != null)
                                {
                                    if ((Convert.ToInt64(UpdFileVersionInfo.Replace(".", "")) > Convert.ToInt64(myFileVersionInfo.Replace(".", ""))) || (UpdFileDateTime > myFileDateTime))
                                    {
                                        if (File.Exists(fileName.Replace(".exe", ".exe.old")))
                                        {
                                            File.Delete(fileName.Replace(".exe", ".exe.old"));
                                        };

                                        if (File.Exists(fileName.Replace(".exe", ".exe.update")))
                                        {
                                            File.Delete(fileName.Replace(".exe", ".exe.update"));
                                        };

                                        SendMessage("Текущая версия: " + myFileVersionInfo + " (" + myFileDateTime + ")\r\nВерсия обновления: " + UpdFileVersionInfo + " (" + UpdFileDateTime + ")\r\nУстанавливается обновление.\r\n<b> Дождитесь перезагрузки страницы.<b>", 200, null, admin, true, "15");
                                        File.Copy(updateserver + Path.GetFileName(fileName), fileName.Replace(".exe", ".exe.update"));
                                        File.Move(fileName, fileName.Replace(".exe", ".exe.old"));
                                        File.Move(fileName.Replace(".exe", ".exe.update"), fileName);
                                        Console.WriteLine("{0}; {1}; Установлено обновление. Версия {2}", DateTime.Now.ToString(), Environment.UserName, UpdFileVersionInfo);
                                        Server.Restart();

                                    }
                                    else
                                    {
                                        SendMessage("Текущая версия: " + myFileVersionInfo + "(" + myFileDateTime + ")\r\nВерсия обновления: " + UpdFileVersionInfo + " (" + UpdFileDateTime + ")", 200, null, admin, true, "-1");
                                    }
                                }
                                else
                                {
                                    SendMessage("Текущая версия: " + myFileVersionInfo + "(" + myFileDateTime + ")\r\nОбновлений на сервере обновлений \"" + updateserver + "\" не найдено", 200, null, admin, true, "-1");
                                }
                                break;
                            }
                        case "restart":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }
                                SendMessage("<b>Сервис перезагружается. Дождитесь перезагрузки страницы.<b>", 200, null, admin, true, "10;URL='/../");
                                Server.Restart();
                                break;
                            }
                        case "currenttask":
                            {
                                string message = string.Empty;
                                try
                                {
                                    if (Utils.SQLPresent)
                                        using (SqlCommand command = new SqlCommand(@"
                                                                        SELECT 
                                                                            Concat([IncidentId],'.',[BranchNo])
                                                                        FROM [dlfe].[DTSWorkFact] AS wf
                                                                        inner join [dlfe].[Employees] as emp on wf.EmployeeId = emp.EmployeeID 
                                                                        AND emp.UserName = '" + Server.DomainName + "\\" + username + "' AND wf.EndDateTime IS NULL"
                                                                         , thisConnection))
                                        {
                                            using (SqlDataReader reader = command.ExecuteReader())
                                            {
                                                Utils.ConnectionPool.BindReader(command.Connection, reader);
                                                while (reader.Read())
                                                {
                                                    message = reader[0].ToString();
                                                }
                                                reader.Close();

                                            }
                                            command.Dispose();
                                        }


                                    //message = "TEST + TEST";
                                }
                                catch (Exception ex)
                                {
                                    message = ex.Message;
                                }

                                byte[] byteBuffer = Encoding.Default.GetBytes(message);
                                //thisClient.Response.ContentType = "plaintext";
                                thisClient.Response.OutputStream.Write(byteBuffer, 0, byteBuffer.Length);
                                thisClient.Response.Close();

                                break;
                            }
                        case "db_edituserlist":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос редактирования списка пользователей базы; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }

                                string htmlDBUserList = string.Empty;
                                string filename = string.Empty;
                                //string htmlDBUserList = "<table class=\"panel\" width=\"100%\">";
                                string body = string.Empty;
                                try
                                {
                                    filename = Options["filepath"];
                                    UserDefworks.UsersList DBUserList = new UserDefworks.UsersList(filename);

                                    bool IsConfigRunning = UserDefworks.IsConfigRunning(DBUserList.DBCAtalog);

                                    string buttons = @"
                                                        <form id=usersaction   method=""post"">
                                                        <input type=hidden name=filename value=" + filename + @">
                                                        <input type=submit formAction=""db_adduser"" id=""db_adduser"" class=""b_w"" value='Добавить нового пользователя' onclick=""act(this)""> 
                                                        </form>";
                                    string tbuttons = @"<input type=submit formAction=""db_edituser"" id=""db_edituser"" class=""b_w"" style=""width:70"" value='Изменить' onclick=""act(this)""> 
                                                        <input type=submit formAction=""db_deleteuser"" id=""db_deleteuser"" class=""b_w"" style=""width:70"" value='Удалить' onclick=""act(this)"">";
                                    string formheader = @"<form id=usersaction   method=""post""><input type=hidden name=filename value=" + filename + @">";

                                    if (IsConfigRunning)
                                    {
                                        buttons = "В базе [" + DBUserList.DBCAtalog + "] запущен конфигуратор. Редактировать список пользователей запрещено, чтобы не сбросить настройки подключения к базе.";
                                        tbuttons = "";
                                    }

                                    var items = from pair in DBUserList orderby pair.Value.FullName.ToString() ascending select pair.Value;

                                    foreach (UserDefworks.UserItem User in items)
                                    {
                                        htmlDBUserList += string.Format("<tr>{5}<td><input type=\"hidden\" name=\"accountname\" value='{0}'> {1} </td><td>{0}</td><td>{2}</td><td>{3}</td><td>  {4}</td></form><tr>", User.Name, User.FullName, User.UserInterface, User.UserRights, tbuttons, formheader);
                                        // htmlDBUserList += "<option value='" + User.Name + "'>" + User.FullName + " [" + User.Name + "] </option>\r\n";
                                    }

                                    body = @"<script></script>
                                                    <table width=98%>
                                                         <tr>
                                                           <td  width=* colspan=""5"" rowspan=""1"" class=""pheader"">Работа с пользователями в базе [" + DBUserList.DBCAtalog + @"].</td>
                                                         </tr>

                                                        <tr class=""panel""><td width=*  colspan=""5"">
                                                        " + buttons + @"
                                                        </td></tr>

                                                         <tr class=""pheader"">
                                                           <td>Полное имя</td>
                                                           <td>Имя пользователя</td> 
                                                           <td>Интерфейс</td> 
                                                           <td colspan=2>Набор прав</td> 
                                                         </tr>

                                                        <tbody class=""panel"">
                                                                    " + htmlDBUserList + @"
                                                        </tbody>   
                                                    </table>";

                                }
                                catch (Exception ex)
                                {
                                    body = "Ошибка открытия списка пользователей: " + ex.Message;
                                }

                                SendMessage(body, 200, null, admin, true);
                                break;
                            }
                        case "db_edituser":
                            {

                                string body = string.Empty;
                                string filename = string.Empty; ;
                                string userlogin = string.Empty; ;

                                if (Options.ContainsKey("accountname"))
                                    userlogin = Options["accountname"];


                                if ((!admin) && (userlogin != username))
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос редактирования учетной записи 1С; {2}", DateTime.Now.ToString(), username, userlogin);
                                    return;
                                }


                                bool DoSave = false;

                                try
                                {
                                    if (Options.ContainsKey("filename"))
                                        filename = Options["filename"];

                                    if (filename.Length < 6)
                                    {
                                        SendMessage("Не выбрана база данных.", 200, null, false, true);
                                        return;
                                    }

                                    UserDefworks.UsersList DBUserList = new UserDefworks.UsersList(filename);
                                    if (!DBUserList.ContainsKey(userlogin))
                                    {
                                        SendMessage("Учетная запись <b>" + userlogin + "<b> не найдена в базе " + filename + ".", 200, null, false, true);
                                        return;
                                    }

                                    UserDefworks.UserItem User = DBUserList[userlogin];

                                    if (Options.ContainsKey("dosave"))
                                        DoSave = bool.Parse(Options["dosave"]);

                                    if (!DoSave)
                                    {
                                        string InterfaceListHTML = "<select name=\"userinterface\">";
                                        string RightsListHTML = "<select name=\"userrights\">";

                                        if (admin)
                                        {

                                            foreach (string interfacename in DBUserList.InterfaceList)
                                            {
                                                if (User.UserInterface == interfacename)
                                                    InterfaceListHTML += "<option value='" + interfacename + "' selected>" + interfacename + "</option>\r\n";
                                                else
                                                    InterfaceListHTML += "<option value='" + interfacename + "'>" + interfacename + "</option>\r\n";
                                            }

                                            foreach (string ringhtname in DBUserList.RightsList)
                                            {
                                                if (User.UserRights == ringhtname)
                                                    RightsListHTML += "<option value='" + ringhtname + "' selected>" + ringhtname + "</option>\r\n";
                                                else
                                                    RightsListHTML += "<option value='" + ringhtname + "'>" + ringhtname + "</option>\r\n";
                                            }

                                            InterfaceListHTML += "</select>";
                                            RightsListHTML += "</select>";

                                        }
                                        else
                                        {
                                            InterfaceListHTML = User.UserInterface;
                                            RightsListHTML = User.UserRights;
                                        }

                                        string editblockHTML = null;
                                        if (!UserDefworks.IsConfigRunning(DBUserList.DBCAtalog))
                                        {
                                            if (admin)
                                            {
                                                editblockHTML += "<br> Новое имя пользователя: <input type=text name=newname value=''>";
                                                foreach (UserDefworks.UserParameters UserOption in Enum.GetValues(typeof(UserDefworks.UserParameters)))
                                                {
                                                    if (UserOption == UserDefworks.UserParameters.PasswordHash)
                                                    {
                                                        editblockHTML += "<br> Пароль: <input type=text name=userpassword value=''>";
                                                        continue;
                                                    }

                                                    if (UserOption == UserDefworks.UserParameters.UserInterface)
                                                    {
                                                        editblockHTML += "<br> " + UserDefworks.UserParamNames[UserOption] + ": " + InterfaceListHTML;
                                                        continue;
                                                    }

                                                    if (UserOption == UserDefworks.UserParameters.UserRights)
                                                    {
                                                        editblockHTML += "<br> " + UserDefworks.UserParamNames[UserOption] + ": " + RightsListHTML;
                                                        continue;
                                                    }

                                                    if (User[UserOption].GetType().Name == "Int32")
                                                        continue;//editblockHTML += "<br> " + UserDefworks.UserParamNames[UserOption] + ": <input type=checkbox name=" + UserOption.ToString() + " value=" + User[UserOption] + ">";
                                                    else
                                                        editblockHTML += "<br> " + UserDefworks.UserParamNames[UserOption] + ": <input type=text name=" + UserOption.ToString() + " value=\"" + User[UserOption] + "\" >";
                                                }
                                            }
                                            else
                                            {
                                                editblockHTML += @"<br>Пароль: <input type=text name=userpassword value=''>
                                                    <br>Полное имя: " + User.FullName + @"
                                                    <br>Интерфейс: " + User.UserInterface + @" 
                                                    <br>Набор прав: " + User.UserRights;
                                            }
                                        }
                                        else
                                            editblockHTML = "В базе [" + UserDefworks.GetDBCatalog(filename) + "] запущен конфигуратор. Редактировать список пользователей запрещено, чтобы не сбросить настройки подключения к базе.";

                                        body = @"<script></script>
                                                    <table width=""100%"" >
                                                         <tr>
                                                           <td colspan=""2"" rowspan=""1"" class=""pheader"">Учетная запись: <b>" + userlogin + @" в базе [" + UserDefworks.GetDBCatalog(filename) + @"]</b></td>
                                                         </tr>

                                                        <tbody class=""panel"">
                                                            <tr>
                                                                <td width=*>
                                                                    <form id=usersaction  method=""post"">
                                                                    <input type=hidden name=filename value=" + filename + @">
                                                                    <input type=hidden name=dosave value=True>
                                                                    <input type=hidden name=accountname value=" + userlogin + @">
                                                                    " + editblockHTML + @"
                                                                    <br><input type=submit formAction=""db_edituser"" id=""db_edituser"" class=""b_w"" value='Сохранить изменения' onclick=""act(this)"">
                                                                    </form>
                                                                </td>
                                                            </tr>
                                                        </tbody>     
                                                    </table>";
                                    }
                                    else
                                    {
                                        string usercatalog = null;

                                        body += "\nЗапрос изменения учетной записи " + userlogin;
                                        Console.WriteLine("{0}; {1}; Запрос редактирования учетной записи 1С; {2}", DateTime.Now.ToString(), username, userlogin);
                                        if (admin)
                                        {
                                            if (Options.ContainsKey("newname"))
                                            {
                                                string newname = Options["newname"];
                                                if (newname != string.Empty)
                                                    if (User.Name != newname)
                                                    {
                                                        User.Name = newname;
                                                        User.modified = true;
                                                    }
                                            }
                                            foreach (UserDefworks.UserParameters UserOption in Enum.GetValues(typeof(UserDefworks.UserParameters)))
                                            {
                                                string optionkey = UserOption.ToString().ToLower();

                                                if (UserOption == UserDefworks.UserParameters.PasswordHash)
                                                {
                                                    continue;
                                                }

                                                if (Options.ContainsKey(optionkey))
                                                    if (User[UserOption] != Options[optionkey])
                                                    {
                                                        User[UserOption] = Options[optionkey];
                                                        Console.WriteLine("{0}; {1}; Изменен параметр [{2}] учетной записи 1С; {3}", DateTime.Now.ToString(), username, UserDefworks.UserParamNames[UserOption], userlogin);
                                                    }
                                            }
                                        }

                                        if (Options.ContainsKey("userpassword"))
                                        {
                                            string userpassword = Options["userpassword"];
                                            if (userpassword != "")
                                            {
                                                string userpasswordhash = UserDefworks.GetStringHash(userpassword);
                                                if (userpassword == "#empty")
                                                {
                                                    userpasswordhash = UserDefworks.GetStringHash(string.Empty);
                                                }

                                                if (User.PasswordHash != userpasswordhash)
                                                {
                                                    User.PasswordHash = userpasswordhash;
                                                    User.modified = true;
                                                    Console.WriteLine("{0}; {1}; Изменен пароль учетной записи 1С; {2}", DateTime.Now.ToString(), username, userlogin);
                                                }
                                            }
                                        }

                                        if (!User.modified)
                                            body += "\nНи один из параметров не был изменен.";
                                        else
                                        {

                                            usercatalog = User[UserDefworks.UserParameters.UserCatalog];
                                            if (usercatalog != "")
                                            {
                                                string dbcatalog = UserDefworks.GetDBCatalog(filename);
                                                if (!usercatalog.StartsWith(".\\"))
                                                    usercatalog.Replace(dbcatalog, ".\\");

                                                string userFullcatalog = usercatalog.Replace(".\\", dbcatalog);

                                                if (!Directory.Exists(userFullcatalog))
                                                {
                                                    Directory.CreateDirectory(userFullcatalog);
                                                    body = "\nСоздана директория пользователя [" + userFullcatalog + "] для учетной записи " + userlogin;
                                                    Console.WriteLine("{0}; {1}; Создан каталог пользователя [{2}] 1С; {3}", DateTime.Now.ToString(), username, userFullcatalog, userlogin);
                                                }
                                            }
                                            DBUserList[userlogin] = User;
                                            UnlockFile(filename, false);

                                            if (DBUserList.Save(filename))
                                            {
                                                body += "\nУчетная запись " + userlogin + " в базе " + UserDefworks.GetDBCatalog(filename) + " сохранена успешно.";
                                                Console.WriteLine("{0}; {1}; Учетная запись 1С сохранена; {2}", DateTime.Now.ToString(), username, userlogin);
                                            }
                                            else
                                            {
                                                body += "\nНеизвестная ошибка! Учетная запись " + userlogin + " в базе " + UserDefworks.GetDBCatalog(filename) + " НЕ ИЗМЕНЕНА.";
                                                Console.WriteLine("{0}; {1};Ошибка сохранения списка пользователей базы 1C {3} ; {2}", DateTime.Now.ToString(), username, userlogin, DBUserList.DBCAtalog);
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    body = "\nОшибка редактирования учетной записи " + userlogin + ": " + ex.Message;
                                    Console.WriteLine("{0}; {1};Ошибка редактирования списка пользователей базы 1C {2}; {3}", DateTime.Now.ToString(), username, UserDefworks.GetDBCatalog(filename), ex.Message);
                                }
                                SendMessage(body, 200, null, admin, true);
                                break;
                            }
                        case "db_deleteuser":
                            {

                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос удаления учетной записи 1С; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }

                                string body = string.Empty;
                                string filename = string.Empty; ;
                                string userlogin = string.Empty; ;

                                try
                                {
                                    filename = Options["filename"];
                                    userlogin = Options["accountname"];
                                    body += "\nЗапрос удаления учетной записи " + userlogin;

                                    UserDefworks.UsersList DBUserList = new UserDefworks.UsersList(filename);
                                    UserDefworks.UserItem User = DBUserList[userlogin];

                                    string usercatalog = User.UserCatalog;
                                    if (usercatalog != "")
                                    {
                                        string dbcatalog = UserDefworks.GetDBCatalog(filename);
                                        if (!usercatalog.StartsWith(".\\"))
                                            usercatalog.Replace(dbcatalog, ".\\");

                                        string userFullcatalog = usercatalog.Replace(".\\", dbcatalog);

                                        if (Directory.Exists(userFullcatalog))
                                        {
                                            Directory.Delete(userFullcatalog, true);
                                            body = "\nУдален каталог пользователя [" + userFullcatalog + "] для учетной записи " + userlogin;
                                        }
                                    }

                                    DBUserList.Remove(userlogin);
                                    UnlockFile(filename, false);

                                    if (DBUserList.Save(filename))
                                        body += "\nУчетная запись " + userlogin + " в базе " + UserDefworks.GetDBCatalog(filename) + " удалена успешно.";
                                    else
                                        body += "\nНеизвестная ошибка! Учетная запись " + userlogin + " в базе " + UserDefworks.GetDBCatalog(filename) + " НЕ УДАЛЕНА.";
                                }
                                catch (Exception ex)
                                {
                                    body = "\nОшибка удаления учетной записи " + userlogin + ": " + ex.Message;
                                }

                                SendMessage(body, 200, null, admin, true);
                                break;
                            }
                        case "db_adduser":
                            {
                                if (!admin)
                                {
                                    SendMessage("Несанкционированный запрос. <b>Лог операций сохранен</b>", 200, null, false, true);
                                    Console.WriteLine("{0}; {1}; Несанкционированный запрос добавления учетной записи 1С; {2}", DateTime.Now.ToString(), username, Request);
                                    return;
                                }

                                string body = string.Empty;
                                string filename = string.Empty; ;
                                string userlogin = string.Empty; ;
                                bool DoSave = false;

                                try
                                {
                                    filename = Options["filename"];



                                    UserDefworks.UsersList DBUserList = new UserDefworks.UsersList(filename);
                                    if (Options.ContainsKey("dosave"))
                                        DoSave = bool.Parse(Options["dosave"]);

                                    if (!DoSave)
                                    {
                                        string editblockHTML = null;
                                        string InterfaceListHTML = "<select name=\"userinterface\">";
                                        string RightsListHTML = "<select name=\"userrights\">";
                                        foreach (string iterfacename in DBUserList.InterfaceList)
                                            InterfaceListHTML += "<option value='" + iterfacename + "'>" + iterfacename + "</option>\r\n";
                                        foreach (string ringhtname in DBUserList.RightsList)
                                            RightsListHTML += "<option value='" + ringhtname + "'>" + ringhtname + "</option>\r\n";
                                        InterfaceListHTML += "</select>";
                                        RightsListHTML += "</select>";

                                        UserDefworks.UserItem User = new UserDefworks.UserItem("#temp", Name: "#temp");
                                        if (!UserDefworks.IsConfigRunning(UserDefworks.GetDBCatalog(filename)))
                                        {
                                            foreach (UserDefworks.UserParameters UserOption in Enum.GetValues(typeof(UserDefworks.UserParameters)))
                                                switch (UserOption)
                                                {
                                                    case UserDefworks.UserParameters.PasswordHash:
                                                        editblockHTML += "<br> Пароль: <input type=text name=UserPassword value=''>";
                                                        break;
                                                    case UserDefworks.UserParameters.UserInterface:
                                                        editblockHTML += "<br> " + UserDefworks.UserParamNames[UserOption] + ": " + InterfaceListHTML;
                                                        continue;
                                                    case UserDefworks.UserParameters.UserRights:
                                                        editblockHTML += "<br> " + UserDefworks.UserParamNames[UserOption] + ": " + RightsListHTML;
                                                        continue;
                                                    default:
                                                        if (User[UserOption].GetType().Name == "Int32")
                                                            continue;
                                                        else
                                                            editblockHTML += "<br> " + UserDefworks.UserParamNames[UserOption] + ": <input type=text name=" + UserOption.ToString().ToLower() + " value=\"" + User[UserOption] + "\" id=" + UserOption.ToString().ToLower() + " >";
                                                        break;
                                                }
                                        }
                                        else
                                            editblockHTML = "В базе [" + UserDefworks.GetDBCatalog(filename) + "] запущен конфигуратор. Редактировать список пользователей запрещено, чтобы не сбросить настройки подключения к базе.";

                                        body = @"
<script>
window.onready = function(){
    $('#accountname').on('change keypress keyup blur', function(){

    $('#usercatalog').val('.\\Users\\' + this.value)
});
};

</script>
                                                    <table width=""100%"" >
                                                         <tr>
                                                           <td colspan=""2"" rowspan=""1"" class=""pheader"">Добавление новой учетной записи</td>
                                                         </tr>

                                                        <tbody class=""panel"">
                                                            <tr>
                                                                <td width=*>
                                                                    <form id=usersaction  method=""post"">
                                                                    <input type=hidden name=filename value=" + filename + @">
                                                                    <input type=hidden name=DoSave value=True>
                                                                    Имя пользователя: <input type=text name=accountname id=accountname>
                                                                    " + editblockHTML + @"
                                                                    <br><input type=submit formAction=""db_adduser"" id=""db_adduser"" class=""b_w"" value='Сохранить изменения' onclick=""act(this)"">
                                                                    </form>
                                                                </td>
                                                            </tr>
                                                        </tbody>     
                                                    </table>";
                                    }
                                    else
                                    {
                                        string usercatalog = null;
                                        userlogin = Options["accountname"];
                                        if (DBUserList.ContainsKey(userlogin))
                                        {
                                            body = "\nТакая учетная запись уже есть в базе [" + UserDefworks.GetDBCatalog(filename) + "]";
                                        }
                                        else
                                        {
                                            UserDefworks.UserItem User = new UserDefworks.UserItem(userlogin, Name: userlogin);
                                            body += "\nЗапрос добавления учетной записи " + userlogin;
                                            foreach (UserDefworks.UserParameters UserOption in Enum.GetValues(typeof(UserDefworks.UserParameters)))
                                            {
                                                string optionkey = UserOption.ToString().ToLower();

                                                if (UserOption == UserDefworks.UserParameters.PasswordHash)
                                                {
                                                    if (Options.ContainsKey("userpassword"))
                                                    {
                                                        string userpassword = Options["userpassword"];
                                                        if (userpassword != "")
                                                        {
                                                            string userpasswordhash = UserDefworks.GetStringHash(userpassword);
                                                            if (userpassword == "#empty")
                                                            {
                                                                userpasswordhash = UserDefworks.GetStringHash(string.Empty);
                                                            }
                                                            User[UserOption] = userpasswordhash;
                                                        }
                                                    }
                                                    continue;
                                                }

                                                if (Options.ContainsKey(optionkey))
                                                {
                                                    User[UserOption] = Options[optionkey];
                                                }
                                            }

                                            usercatalog = User[UserDefworks.UserParameters.UserCatalog];
                                            if (usercatalog != "")
                                            {
                                                string dbcatalog = UserDefworks.GetDBCatalog(filename);
                                                if (!usercatalog.StartsWith(".\\"))
                                                    usercatalog.Replace(dbcatalog, ".\\");

                                                string userFullcatalog = usercatalog.Replace(".\\", dbcatalog);

                                                if (!Directory.Exists(userFullcatalog))
                                                {
                                                    Directory.CreateDirectory(userFullcatalog);
                                                    body = "\nСоздана директория пользователя [" + userFullcatalog + "] для учетной записи " + userlogin;
                                                }
                                            }

                                            DBUserList.Add(User);
                                            UnlockFile(filename, false);
                                            if (DBUserList.Save(filename))
                                                body += "\nУчетная запись " + userlogin + " в базе " + UserDefworks.GetDBCatalog(filename) + " добавлена успешно.";
                                            else
                                                body += "\nНеизвестная ошибка! Учетная запись " + userlogin + " в базе " + UserDefworks.GetDBCatalog(filename) + " НЕ ДОБАВЛЕНА.";
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    body = "\nОшибка добавления учетной записи " + userlogin + ": " + ex.Message;
                                }
                                SendMessage(body, 200, null, admin, true);
                                break;
                            }
                        case "sched_addtask":
                            {
                                try
                                {
                                    string accountname = string.Empty;
                                    string hour = string.Empty;
                                    string minute = string.Empty;
                                    Utils.SchedulerTasks TaskId = Utils.SchedulerTasks.Null;

                                    if (Options.ContainsKey("accountname"))
                                        accountname = Options["accountname"];
                                    if (Options.ContainsKey("hour"))
                                        hour = Options["hour"];
                                    if (Options.ContainsKey("minute"))
                                        minute = Options["minute"];
                                    if (Options.ContainsKey("tasktypeid"))
                                        TaskId = (Utils.SchedulerTasks)int.Parse(Options["tasktypeid"]);

                                    string time = hour + ":" + minute;
                                    Utils.Sched.Add(new Utils.KickOffSchedulerEntry(accountname, TaskId, time, true));

                                    SendMessage("Задача добавлена в расписание.", 200, null, admin, true);
                                }
                                catch (Exception ex)
                                {
                                    SendMessage("Ошибка добавления задачи." + ex.Message, 200, null, admin, true);
                                }
                                break;
                            }
                        case "sched_remove":
                            {
                                try
                                {
                                    string taskid = string.Empty;
                                    if (Options.ContainsKey("taskid"))
                                        taskid = Options["taskid"];

                                    foreach (Utils.KickOffSchedulerEntry Entry in Utils.Sched)
                                    {
                                        if (Entry.id == taskid)
                                        {
                                            Utils.Sched.Remove(Entry);
                                            SendMessage("Задача удалена", 200, null, admin, true);
                                            break;
                                        }

                                    }
                                }
                                catch
                                {
                                    SendMessage("Ошибка удаления задачи", 200, null, admin, true);
                                }
                                SendMessage("Ошибка удаления задачи: задача не найдена.", 200, null, admin, true);
                                break;
                            }
                        case "db_findinlog":
                            {
                                string filename = string.Empty;
                                string searchfilter = string.Empty;
                                string ObjectType = "$Refs";
                                string EventType = "RefOpen";
                                DateTime EventDate = DateTime.Today;

                                string message = string.Empty;

                                if (Options.ContainsKey("filename"))
                                    filename = Path.GetDirectoryName(Path.GetDirectoryName(Options["filename"])) + "\\SYSLOG\\1cv7.mlg";

                                if (Options.ContainsKey("filepath"))
                                    filename = Path.GetDirectoryName(Path.GetDirectoryName(Options["filepath"])) + "\\SYSLOG\\1cv7.mlg";

                                if (Options.ContainsKey("searchfilter"))
                                    searchfilter = Options["searchfilter"];

                                if (Options.ContainsKey("objecttype"))
                                    ObjectType = Options["objecttype"];

                                if (Options.ContainsKey("eventtype"))
                                    EventType = Options["eventtype"];

                                if (Options.ContainsKey("eventdate"))
                                    DateTime.TryParse(Options["eventdate"], out EventDate);


                                if (filename == string.Empty)
                                {
                                    filename = @"F:\1C\RN2_NEW\SYSLOG\1cv7.mlg";
                                    //message = "(лучше все-таки выбрать базу из списка)<br>";
                                }


                                if (File.Exists(filename))
                                {
                                    string[] filter = { ObjectType, EventType, searchfilter };
                                    bool LC = (EventType == "RefOpen") && (searchfilter != "");

                                    string datefilter = string.Format("{0:yyyyMMdd}", EventDate);

                                    if (DateTime.Now.Hour < 08)
                                        datefilter = string.Format("{0:yyyyMMdd}", EventDate.AddDays(-1));

                                    string[][] res = LogParser.Find(filename, filter, datefilter);
                                    Utils.RefillLists();
                                    message = "История лога <b>" + searchfilter + "</b>: \n";

                                    if (EventType == "RefOpen")
                                        message = "История блокирования ДФА <b>" + searchfilter + "</b>: \n";

                                    bool first = true;
                                    if (res.Length > 0)
                                    {

                                        foreach (string[] result in res)
                                        {
                                            if (result != null)
                                            {

                                                string logoff = string.Empty;

                                                string employee = result[1];

                                                if (employee.EndsWith("1") & !employee.EndsWith("-e01"))
                                                    employee = employee.Substring(0, employee.Length - 1);

                                                string sessioninfo = string.Empty;

                                                DateTime EventMoment = DateTime.Parse(result[0]).AddHours(timeshift);

                                                if (LC)
                                                {
                                                    if (Utils.SessID.Contains(employee))
                                                    {

                                                        if (admin && first)
                                                        {
                                                            logoff = string.Format(@"<form id=usersaction method=""post""><input type=hidden name=login value={0}><input type=submit formAction=""resetsession"" id=""resetsession"" class=""b_w"" value='Завершить сессию сотрудника' onclick=""act(this)""></form>", employee);
                                                            first = false;
                                                        }

                                                        Utils.WTS_SESSION_INFO_1 Session = (Utils.WTS_SESSION_INFO_1)Utils.SessID[employee];
                                                        switch (Session.State)
                                                        {
                                                            case Utils.WTS_CONNECTSTATE_CLASS.WTSActive:
                                                                sessioninfo = "(Сейчас сотрудник <b>подключен</b> к серверу и работает)";
                                                                break;
                                                            case Utils.WTS_CONNECTSTATE_CLASS.WTSConnected:
                                                                sessioninfo = "(Сейчас сотрудник <b>не подключен</b> к серверу, но из 1С не вышел)";
                                                                break;
                                                            case Utils.WTS_CONNECTSTATE_CLASS.WTSDisconnected:
                                                                sessioninfo = "(Сейчас сотрудник <b>не подключен</b> к серверу, но из 1С не вышел)";
                                                                break;
                                                            case Utils.WTS_CONNECTSTATE_CLASS.WTSIdle:
                                                                TimeSpan IdleTime = Utils.GetSessionIdleTime(IntPtr.Zero, Session.SessionId);
                                                                sessioninfo = "(Сейчас сотрудник <b>подключен</b> к серверу, время простоя  " + IdleTime + ")";
                                                                break;
                                                            default:
                                                                sessioninfo = "(Сейчас сотрудник в статусе <b>" + Session.State.ToString() + "</b> к серверу, но из 1С не вышел)";
                                                                break;
                                                        }

                                                    }
                                                    else
                                                        sessioninfo = "(Сейчас сотрудник <b>не подключен</b> к серверу и к 1С. Это значит, что скорее всего ДФА не заблокирован)";


                                                    if (Utils.Employees.ContainsKey(employee))
                                                        message += string.Format("ДФА {0} был заблокирован сотрудником <a href='sip:{1}'>{2}</a> {3} {4}{5}\n", searchfilter, Utils.GetEmployeeEmail(employee), Utils.Employees[employee], EventMoment, sessioninfo, logoff);
                                                    else
                                                        message += string.Format("ДФА {0} был заблокирован сотрудником <a href='sip:{1}'>{2}</a> {3} {4}{5}\n", searchfilter, Utils.GetEmployeeEmail(employee), Utils.GetEmployeeEmail(employee), EventMoment, sessioninfo, logoff);
                                                }
                                                else
                                                {         //datetime, user, eventobject, eventname, eventdescription,comment, line
                                                    if (Utils.Employees.ContainsKey(employee))
                                                        message += string.Format("<a href='sip:{0}'>{1}</a>; {2}; {3}; <b>{4}</b>; {5}; {6}; <b>{7}</b>\n", Utils.GetEmployeeEmail(employee), Utils.Employees[employee], EventMoment, result[2], LogParser.logevents[result[3]], result[5], result[4], result[7].Replace(" ","&nbsp"));
                                                    else
                                                        message += string.Format("{0}; {1}; {2}; <b>{3}</b>; {4}; {5}; {6}\n", result[1], EventMoment, result[2], LogParser.logevents[result[3]], result[5], result[4], result[7]);

                                                }
                                            }
                                        }
                                    }
                                    else
                                        if (LC)
                                        message += string.Format("Сотрудник, заблокировавший ДФА {0} не найден", searchfilter);
                                    else
                                        message += string.Format("Записи лога {0} не найдены", searchfilter);
                                }
                                else
                                    message += string.Format("Файл лога не найден: {0}", filename);

                                SendMessage(message, 200, null, admin, true);

                                break;
                            }
                        case "makeappointment":
                            {
                                Appointment app = new Appointment();

                                app.AttendeeList.Add(new MailAddress(Utils.GetEmployeeEmail(username), Utils.Employees[username]));

                                app.AttendeeList.Add("conf.5fl.vvo.sflr.ru@internal.siemens.com");

                                app.Summary = @"Описание тестового приглашения.";
                                app.Subject = @"Тема тестового приглашения. Тестовое письмо, пожалуйста игнорируйте.";
                                //app.OrganizerEmail = @"ptchk.ru@siemens.com";//Utils.GetEmployeeEmail(username);

                                app.OrganizerEmail = "Oleg.Sheyko@siemens.com";

                                app.Location = "RG RU SFLR VLADIVOSTOK 5 FL MEETING ROOM";
                                app.StartDate = DateTime.Now.AddDays(1);
                                app.EndDate = app.StartDate.AddHours(1);
                                //#if DEBUG
                                app.ServerName = "rumowmc2010msx.ww600.siemens.net";
                                //#else
                                //app.ServerName = "ruvvos20004srv.ww600.siemens.net";
                                //#endif
                                app.Port = 25;

                                app.UserName = @"";
                                app.Password = @"";

                                app.MeetingGUID = new Guid();
                                try
                                {
                                    app.EmailAppointment();
                                    SendMessage(app.ToString());
                                }
                                catch (Exception ex)
                                {
                                    SendMessage(ex.Message);
                                }

                                break;
                            }
                        case "v7configreports":
                            {
                                mess = @"
<script>
function updatedbList()
{
    $("".dblist"").each(function ()
          {$(this).html(""<option>Список обновляется...</option>"")});
    $.ajaxSetup({cache: false});
    $.ajax({
            url: ""dblist.scr"",
            cache: false,
            success: function(html)
                {
                    $("".dblist"").each(function ()
                        {$(this).html(html)})
                } 
            });
};

function whoisinconfig(list)
{
}
</script>
<form id=dbaction  method=""post"" name=dbaction>
<br>" + Server.DBlistHTML + @" <input style=""width: 100px;"" value='Обновить список' onclick=""updatedbList()"" type=button>
<br><input  formAction=""countcodestrings"" id=""countcodestrings"" type=submit class=""b_w"" value='Отчет о количестве строк в конфигурации' onclick=""act(this)"">
<br><input  formAction=""countcodeinreports"" id=""countcodeinreports"" style=""width: 278px;"" type=submit class=""b_w"" value='Отчет о количестве строк во внешней отчетности' onclick=""act(this)"">
<br><input  formAction=""getmoxelslist""  id=""getmoxelslist"" type=submit class=""b_w"" value='Список печатных форм конфигурации' onclick=""act(this)"">
<br><input  formAction=""getmoxelsextlist"" id=""getmoxelsextlist"" type=submit class=""b_w"" value='Список печатных форм внешней отчетности' onclick=""act(this)"">
<br>
<br><input class=radio type=radio name=mode value=html checked style=""Width:20"">Результат в виде таблицы</input>
<br><input class=radio type=radio name=mode value=xls style=""Width:20"">Результат в Excel</input>
<br>
<br>
<br><input  formAction=""searchinextreports"" id=""searchinextreports"" type=submit class=""b_w"" value='Поиск по внешней отчетности' onclick=""act(this)"">

</form>";
                                SendMessage(mess, 200, null, admin, true);
                                break;

                            }
                        case "countcodestrings":
                            {
                                string filename = string.Empty;
                                string mode = "html";
                                mess = string.Empty;
                                try
                                {
                                    if (Options.ContainsKey("filepath"))
                                        filename = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Options["filepath"])), "1cv7.md");

                                    if (Options.ContainsKey("mode"))
                                        mode = Options["mode"];

                                    if (!File.Exists(filename))
                                    {
                                        SendMessage("Файл " + filename + " не найден.", 200, null, admin, true);
                                        break;
                                    }

                                    Compound.OleStorage.TaskItem v7Config = new Compound.OleStorage.TaskItem(filename);

                                    DataTable report = new DataTable("v7ConfigReport");

                                    report.Columns.Add(new DataColumn("Объект", System.Type.GetType("System.String")));
                                    report.Columns.Add(new DataColumn("Синоним", System.Type.GetType("System.String")));
                                    report.Columns.Add(new DataColumn("Количество Строк", System.Type.GetType("System.Int32")));
                                    report.Columns.Add(new DataColumn("Количество Процедур", System.Type.GetType("System.Int32")));
                                    report.Columns.Add(new DataColumn("Количество Функций", System.Type.GetType("System.Int32")));
                                    int result = 0;

                                    DataRow row;
                                    mess = string.Format("Считаем Количество Строк конфигурации :<b>{0}</b><br><script></script>", filename);

                                    if (v7Config.GlobalModule != null)
                                    {
                                        var Obj = v7Config.GlobalModule;
                                        row = report.NewRow();
                                        row["Объект"] = "ГлобальныйМодуль";
                                        row["Количество Строк"] = Obj.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                        row["Количество Процедур"] = Obj.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                        row["Количество Функций"] = Obj.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                        report.Rows.Add(row);
                                    }

                                    foreach (var Obj in v7Config.Subcontos)
                                    {

                                        if (Obj.Form != null)
                                        {
                                            row = report.NewRow();
                                            row["Объект"] = "Справочник." + Obj.Identity + ".ФормаЭлемента.Модуль";
                                            row["Синоним"] = Obj.Alias;
                                            row["Количество Строк"] = Obj.Form.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                            row["Количество Процедур"] = Obj.Form.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                            row["Количество Функций"] = Obj.Form.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                            report.Rows.Add(row);
                                        }


                                        if (Obj.FolderForm != null)
                                        {
                                            row = report.NewRow();
                                            row["Объект"] = "Справочник." + Obj.Identity + ".ФормаГруппы.Модуль";
                                            row["Синоним"] = Obj.Alias;
                                            row["Количество Строк"] = Obj.FolderForm.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                            row["Количество Процедур"] = Obj.FolderForm.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                            row["Количество Функций"] = Obj.FolderForm.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                            report.Rows.Add(row);
                                        }


                                        if (Obj.ListForms != null)
                                            foreach (var ListForm in Obj.ListForms)
                                            {
                                                if (ListForm.DialogModule != null)
                                                {
                                                    row = report.NewRow();
                                                    row["Объект"] = string.Format("Справочник.{0}.ФормаСписка.{1}.Модуль", Obj.Identity, ListForm.Identity);
                                                    row["Синоним"] = ListForm.Alias;
                                                    row["Количество Строк"] = ListForm.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                                    row["Количество Процедур"] = ListForm.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                                    row["Количество Функций"] = ListForm.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                                    report.Rows.Add(row);
                                                }
                                            }
                                    }

                                    foreach (var Obj in v7Config.Documents)
                                    {
                                        if (Obj.Form != null)
                                        {
                                            row = report.NewRow();
                                            row["Объект"] = "Документ." + Obj.Identity + ".Форма.Модуль";
                                            row["Синоним"] = Obj.Alias;
                                            row["Количество Строк"] = Obj.Form.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                            row["Количество Процедур"] = Obj.Form.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                            row["Количество Функций"] = Obj.Form.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                            report.Rows.Add(row);
                                        }


                                        if (Obj.TransactionModule != null)
                                        {
                                            row = report.NewRow();
                                            row["Объект"] = "Документ." + Obj.Identity + ".МодульПроведения";
                                            row["Синоним"] = Obj.Alias;
                                            row["Количество Строк"] = Obj.TransactionModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                            row["Количество Процедур"] = Obj.TransactionModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                            row["Количество Функций"] = Obj.TransactionModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                            report.Rows.Add(row);
                                        }

                                    }

                                    foreach (var Obj in v7Config.Algorithms)
                                    {
                                        if (Obj.CalculationModule != null)
                                        {
                                            row = report.NewRow();
                                            row["Объект"] = "ВидРасчета." + Obj.Identity + ".МодульРасчета";
                                            row["Синоним"] = Obj.Alias;
                                            row["Количество Строк"] = Obj.CalculationModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                            row["Количество Процедур"] = Obj.CalculationModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                            row["Количество Функций"] = Obj.CalculationModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                            report.Rows.Add(row);
                                        }

                                    }


                                    foreach (var Obj in v7Config.CJ)
                                    {
                                        if (Obj.ListForms != null)
                                            foreach (var ListForm in Obj.ListForms)
                                            {
                                                if (ListForm.DialogModule != null)
                                                {
                                                    row = report.NewRow();
                                                    row["Объект"] = string.Format("ЖурналРасчетов.{0}.ФормаСписка.{1}.Модуль", Obj.Identity, ListForm.Identity);
                                                    row["Синоним"] = ListForm.Alias;
                                                    row["Количество Строк"] = ListForm.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                                    row["Количество Процедур"] = ListForm.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                                    row["Количество Функций"] = ListForm.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                                    report.Rows.Add(row);
                                                }
                                            }
                                    }

                                    foreach (var Obj in v7Config.Journals)
                                    {
                                        if (Obj.ListForms != null)
                                            foreach (var ListForm in Obj.ListForms)
                                            {
                                                if (ListForm.DialogModule != null)
                                                {
                                                    row = report.NewRow();
                                                    row["Объект"] = string.Format("Журнал.{0}.ФормаСписка.{1}.Модуль", Obj.Identity, ListForm.Identity);
                                                    row["Синоним"] = ListForm.Alias;
                                                    row["Количество Строк"] = ListForm.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                                    row["Количество Процедур"] = ListForm.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                                    row["Количество Функций"] = ListForm.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                                    report.Rows.Add(row);
                                                }
                                            }
                                    }

                                    if (v7Config.Buh != null)
                                    {
                                        if (v7Config.Buh.AccountChart != null)
                                        {

                                            foreach (var Obj in v7Config.Buh.AccountChart)
                                                if (Obj.DialogModule != null)
                                                {
                                                    row = report.NewRow();
                                                    row["Объект"] = "ПланСчетов." + Obj.Identity + ".Форма.Модуль";
                                                    row["Синоним"] = Obj.Alias;
                                                    row["Количество Строк"] = Obj.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                                    row["Количество Процедур"] = Obj.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                                    row["Количество Функций"] = Obj.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                                    report.Rows.Add(row);
                                                }
                                        }

                                        if (v7Config.Buh.AccountChartList != null)
                                        {
                                            foreach (var Obj in v7Config.Buh.AccountChartList)
                                            {
                                                if (Obj.DialogModule != null)
                                                {
                                                    row = report.NewRow();
                                                    row["Объект"] = "ФормаСпискаПланаСчетов." + Obj.Identity + ".Модуль";
                                                    row["Синоним"] = Obj.Alias;
                                                    row["Количество Строк"] = Obj.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                                    row["Количество Процедур"] = Obj.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                                    row["Количество Функций"] = Obj.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                                    report.Rows.Add(row);
                                                }
                                            }
                                        }

                                        if (v7Config.Buh.OperationList != null)
                                        {
                                            foreach (var Obj in v7Config.Buh.OperationList)
                                            {
                                                if (Obj.DialogModule != null)
                                                {
                                                    row = report.NewRow();
                                                    row["Объект"] = "Операция.ФормаСписка." + Obj.Identity + ".Модуль";
                                                    row["Синоним"] = Obj.Alias;
                                                    row["Количество Строк"] = Obj.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                                    row["Количество Процедур"] = Obj.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                                    row["Количество Функций"] = Obj.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                                    report.Rows.Add(row);
                                                }
                                            }
                                        }

                                        if (v7Config.Buh.ProvListList != null)
                                        {
                                            foreach (var Obj in v7Config.Buh.ProvListList)
                                            {
                                                if (Obj.DialogModule != null)
                                                {
                                                    row = report.NewRow();
                                                    row["Объект"] = "Проводка.ФормаСписка." + Obj.Identity + ".Модуль";
                                                    row["Синоним"] = Obj.Alias;
                                                    row["Количество Строк"] = Obj.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                                    row["Количество Процедур"] = Obj.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                                    row["Количество Функций"] = Obj.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                                    report.Rows.Add(row);
                                                }
                                            }
                                        }

                                    }

                                    foreach (var Obj in v7Config.Reports)
                                    {
                                        row = report.NewRow();
                                        row["Объект"] = "Отчет." + Obj.Identity + ".Форма.Модуль";
                                        row["Синоним"] = Obj.Alias;
                                        row["Количество Строк"] = Obj.Form.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                        row["Количество Процедур"] = Obj.Form.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                        row["Количество Функций"] = Obj.Form.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                        report.Rows.Add(row);
                                    }

                                    foreach (var Obj in v7Config.CalcVars)
                                    {
                                        row = report.NewRow();
                                        row["Объект"] = "Обработка." + Obj.Identity + ".Форма.Модуль";
                                        row["Синоним"] = Obj.Alias;
                                        row["Количество Строк"] = Obj.Form.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                        row["Количество Процедур"] = Obj.Form.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                        row["Количество Функций"] = Obj.Form.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));

                                        report.Rows.Add(row);
                                    }

                                    row = report.NewRow();
                                    row["Синоним"] = "ИТОГО";
                                    row["Количество Строк"] = report.Compute("SUM([Количество Строк])", "");
                                    row["Количество Процедур"] = report.Compute("SUM([Количество Процедур])", "");
                                    row["Количество Функций"] = report.Compute("SUM([Количество Функций])", "");
                                    report.Rows.Add(row);

                                    System.Web.UI.WebControls.GridView excel = new System.Web.UI.WebControls.GridView();
                                    excel.DataSource = report;
                                    excel.DataBind();

                                    MemoryStream MS = new MemoryStream();
                                    TextWriter SW;

                                    if (mode == "html")
                                        SW = new StreamWriter(MS, Encoding.GetEncoding(1251));
                                    else
                                    {
                                        thisClient.Response.Headers.Clear();
                                        thisClient.Response.AddHeader("content-disposition", "attachment; filename=\"Report.xls\"");
                                        thisClient.Response.ContentType = "application/vnd.ms-excel";
                                        SW = new StreamWriter(thisClient.Response.OutputStream, Encoding.GetEncoding(1251));
                                    }
                                    excel.RenderControl(new System.Web.UI.HtmlTextWriter(SW));
                                    SW.Flush();

                                    if (mode != "html")
                                        SW.Close();

                                    MS.Seek(0, SeekOrigin.Begin);
                                    TextReader SR = new StreamReader(MS, Encoding.GetEncoding(1251));
                                    mess += SR.ReadToEnd();
                                    SR.Close();
                                    SW.Close();
                                    MS.Dispose();

                                    SendMessage(mess, 200, null, admin, true);


                                }
                                catch (Exception ex)
                                {
                                    SendMessage("Ошибка <b>" + ex.ToString() + "</b>", 200, null, admin, true);
                                }



                                break;
                            }
                        case "countcodeinreports":
                            {
                                string mode = "html";
                                string filename = "";
                                mess = string.Empty;
                                List<string> FileList = new List<string>();
                                try
                                {
                                    if (Options.ContainsKey("mode"))
                                        mode = Options["mode"];
                                    string Query = @"
SELECT 
     	Reports.FileName  		AS FileName
	, 	Reports.Description     AS Description
	,   FileVersion.Content AS Content
	FROM 
		 dlfe.FILE_1C_Reports AS Reports
	LEFT JOIN
		files.dlfe.FileVersions AS FileVersion ON (FileVersion.FileID = Reports.FileID) AND (FileVersion.VersionID = Reports.VersionID)
	WHERE (Reports.IsActive = 1) AND
	((RTrim(FileVersion.ContentType) = 'application/1C Report'))

	ORDER BY 
	Reports.FileName";
                                    DataTable Data = new DataTable("Result");

                                    if (thisConnection.State != System.Data.ConnectionState.Open) thisConnection.Open();
                                    {
                                        SqlDataAdapter DA = new SqlDataAdapter(Query, thisConnection);

                                        DA.Fill(Data);

                                    }

                                    DataTable report = new DataTable("v7ConfigReport");

                                    report.Columns.Add(new DataColumn("Объект", System.Type.GetType("System.String")));
                                    report.Columns.Add(new DataColumn("Синоним", System.Type.GetType("System.String")));
                                    report.Columns.Add(new DataColumn("Количество Строк", System.Type.GetType("System.Int32")));
                                    report.Columns.Add(new DataColumn("Количество Процедур", System.Type.GetType("System.Int32")));
                                    report.Columns.Add(new DataColumn("Количество Функций", System.Type.GetType("System.Int32")));
                                    mess = string.Format("Считаем количество строк во внешних отчетах <br><script></script>");


                                    Compound.OleStorage.ExtReport v7Report;
                                    DataRow row;

                                    foreach (DataRow Row in Data.Rows)
                                    {
                                        filename = Path.Combine(Path.GetTempPath(), Row["FileName"].ToString());
                                        row = report.NewRow();

                                        try
                                        {
                                            if (File.Exists(filename))
                                            {
                                                GC.Collect();
                                                GC.Collect();
                                                GC.WaitForPendingFinalizers();
                                                File.Delete(filename);
                                            }

                                            File.WriteAllBytes(filename, (byte[])Row["Content"]);
                                            v7Report = new Compound.OleStorage.ExtReport(filename);
                                            FileList.Add(filename);

                                            row["Объект"] = "ВнешняяОбработка." + v7Report.Identity + ".Форма.Модуль";
                                            row["Синоним"] = v7Report.Alias;
                                            row["Количество Строк"] = v7Report.Form.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                                            row["Количество Процедур"] = v7Report.Form.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Procedure));
                                            row["Количество Функций"] = v7Report.Form.DialogModule.Module.Procedures.Count(x => (x.Value.Type == OleStorage.SubType.Function));
                                        }
                                        catch (Exception ex)
                                        {
                                            row["Объект"] = "ВнешняяОбработка." + Path.GetFileName(filename) + ".Форма.Модуль";
                                            row["Синоним"] = "ОШИБКА: " + ex.Message;
                                        }

                                        report.Rows.Add(row);

                                    }

                                    row = report.NewRow();
                                    row["Синоним"] = "ИТОГО";
                                    row["Количество Строк"] = report.Compute("SUM([Количество Строк])", "");
                                    row["Количество Процедур"] = report.Compute("SUM([Количество Процедур])", "");
                                    row["Количество Функций"] = report.Compute("SUM([Количество Функций])", "");
                                    report.Rows.Add(row);

                                    MemoryStream MS = new MemoryStream();
                                    TextWriter SW;

                                    if (mode == "html")
                                        SW = new StreamWriter(MS, Encoding.GetEncoding(1251));
                                    else
                                    {
                                        thisClient.Response.Headers.Clear();
                                        thisClient.Response.AddHeader("content-disposition", "attachment; filename=\"Report.xls\"");
                                        thisClient.Response.ContentType = "application/vnd.ms-excel";
                                        SW = new StreamWriter(thisClient.Response.OutputStream, Encoding.GetEncoding(1251));
                                    }

                                    System.Web.UI.WebControls.GridView excel = new System.Web.UI.WebControls.GridView();
                                    excel.DataSource = report;
                                    excel.DataBind();

                                    excel.RenderControl(new System.Web.UI.HtmlTextWriter(SW));
                                    SW.Flush();

                                    if (mode != "html")
                                        SW.Close();

                                    MS.Seek(0, SeekOrigin.Begin);
                                    TextReader SR = new StreamReader(MS, Encoding.GetEncoding(1251));
                                    mess += SR.ReadToEnd();
                                    SR.Close();
                                    SW.Close();
                                    MS.Dispose();

                                    SendMessage(mess, 200, null, admin, true);

                                }
                                catch (Exception ex)
                                {
                                    mess += "<br>Ошибка: " + ex.ToString();
                                }
                                SendMessage(mess, 200, null, admin, true);

                                GC.Collect();
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                try
                                {
                                    foreach (string FileNAme in FileList)
                                        File.Delete(FileNAme);
                                }
                                catch
                                {
                                }
                                break;

                            }
                        case "getmoxel":
                            {
                                string filename = string.Empty;
                                string moxelid = string.Empty;
                                string mode = "html";

                                if (Options.ContainsKey("filename"))
                                    filename = Options["filename"];

                                if (Options.ContainsKey("moxelid"))
                                    moxelid = Options["moxelid"];

                                if (Options.ContainsKey("mode"))
                                    mode = Options["mode"];



                                if (filename != string.Empty)
                                {
                                    if (moxelid != string.Empty)
                                    {
                                        try
                                        {
                                            OleStorage.Moxel moxel = null;

                                            if (Path.GetExtension(filename) == ".md")
                                            {

                                                Compound.OleStorage.TaskItem v7Config = new Compound.OleStorage.TaskItem(filename);

                                                if (v7Config.Moxels.ContainsKey(moxelid))
                                                    moxel = v7Config.Moxels[moxelid];
                                            }
                                            else
                                            {

                                                string Query = @"
SELECT 
	     FileVersion.Content AS Content
	FROM 
		 dlfe.FILE_1C_Reports AS Reports
	LEFT JOIN
		files.dlfe.FileVersions AS FileVersion ON (FileVersion.FileID = Reports.FileID) AND (FileVersion.VersionID = Reports.VersionID)
	WHERE (Reports.IsActive = 1) 
    AND	((RTrim(FileVersion.ContentType) = 'application/1C Report'))
    AND  Reports.FileName = '" + filename + @"'

	ORDER BY 
	Reports.FileName";
                                                DataTable Data = new DataTable("Result");

                                                if (thisConnection.State != System.Data.ConnectionState.Open) thisConnection.Open();
                                                {
                                                    SqlDataAdapter DA = new SqlDataAdapter(Query, thisConnection);

                                                    DA.Fill(Data);

                                                }

                                                if (Data.Rows.Count > 0)
                                                {
                                                    filename = Path.Combine(Path.GetTempPath(), filename);
                                                    File.WriteAllBytes(filename, ((byte[])Data.Rows[0]["Content"]));
                                                    Compound.OleStorage.ExtReport v7Report = new Compound.OleStorage.ExtReport(filename);
                                                    moxel = v7Report.Form.Moxels.Find(x => x.MetaPath == moxelid);
                                                }
                                            }

                                            if (moxel != null)
                                            {
                                                string attachmentname = System.Web.HttpUtility.UrlEncode(moxel.Description, Encoding.UTF8);

                                                string tmpMxl = Path.Combine(Path.GetTempPath(), moxel.Description + ".mxl");
                                                string tmpExcel = tmpMxl.Replace(".mxl", ".xls");
                                                string tmpHtml = tmpMxl.Replace(".mxl", ".html");


                                                byte[] buffer;
                                                switch (mode)
                                                {
                                                    case "mxl":
                                                        {
                                                            buffer = moxel.data;
                                                            thisClient.Response.Headers.Clear();
                                                            thisClient.Response.KeepAlive = true;
                                                            thisClient.Response.AddHeader("content-disposition", "attachment; filename=\"" + attachmentname + "." + mode);
                                                            thisClient.Response.ContentType = "bin/octet-stream";
                                                            thisClient.Response.OutputStream.Write(buffer, 0, buffer.Length);
                                                            thisClient.Response.OutputStream.Flush();
                                                            thisClient.Response.Close();
                                                            return;
                                                        }
                                                    default:
                                                        {
                                                            File.WriteAllBytes(tmpMxl, moxel.data);
                                                            Yoksel.CoYoksel Cont = new Yoksel.CoYoksel();
                                                            Yoksel.SpreadsheetDocument Document = Cont.CreateSpreadsheetDocument();
                                                            Document.Open(tmpMxl, 1, 0);
                                                            Document.Save(tmpExcel, 1);

                                                            switch (mode)
                                                            {
                                                                case "xls":
                                                                    {
                                                                        buffer = File.ReadAllBytes(tmpExcel);
                                                                        thisClient.Response.Headers.Clear();
                                                                        thisClient.Response.KeepAlive = true;
                                                                        thisClient.Response.AddHeader("content-disposition", "attachment; filename=\"" + attachmentname + "." + mode);
                                                                        thisClient.Response.ContentType = "bin/octet-stream";
                                                                        thisClient.Response.OutputStream.Write(buffer, 0, buffer.Length);
                                                                        thisClient.Response.OutputStream.Flush();
                                                                        thisClient.Response.Close();
                                                                        File.Delete(tmpExcel);
                                                                        File.Delete(tmpMxl);
                                                                        return;
                                                                    }
                                                                case "html":
                                                                    {
                                                                        if (File.Exists(tmpHtml))
                                                                            File.Delete(tmpHtml);

                                                                        if (File.Exists(tmpHtml))
                                                                            SendMessage("Не удалсь удалить временный файл " + tmpHtml, 200, null, admin, true);

                                                                        Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                                                                        Microsoft.Office.Interop.Excel.Workbook Book = Excel.Workbooks.Open(tmpExcel);
                                                                        Microsoft.Office.Interop.Excel.Worksheet Sheet = Book.Worksheets[1];
                                                                        Sheet.SaveAs(tmpHtml, Microsoft.Office.Interop.Excel.XlFileFormat.xlHtml);
                                                                        Book.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlDoNotSaveChanges);
                                                                        Excel.Quit();

                                                                        File.Delete(tmpExcel);
                                                                        File.Delete(tmpMxl);

                                                                        string path = tmpHtml.Replace(".html", ".files\\");

                                                                        string message = "<script></script><div>" + File.ReadAllText(Path.Combine(path, "sheet001.html"), Encoding.GetEncoding(1251)).Replace("window.location.", "//window.location.") + "</div>";
                                                                        message = message.Replace("src=", "src=" + path.Replace(Path.GetTempPath(), "/temp/"));
                                                                        message = message.Replace("stylesheet.css", path.Replace(Path.GetTempPath(), "/temp/") + "stylesheet.css");
                                                                        message = message.Replace("if gte vml 1", "if gte vml 100").Replace("<![if !vml]>", "<![if vml]>");

                                                                        filename = System.Web.HttpUtility.UrlEncode(filename);
                                                                        string link = thisClient.Request.Url.AbsoluteUri.Replace("mode=html", "");
                                                                        message = string.Format("Печатная форма: <b>{2}</b><br>Скачать: <a href={0}&mode=xls>{3}.XLS</a>&nbsp;&nbsp;&nbsp;<a href={0}&mode=mxl>{3}.MXL</a>{1}", link, message, moxel.MetaPath, moxel.Description);
                                                                        SendMessage(message, 200, null, admin, true);
                                                                        File.Delete(tmpHtml);
                                                                        return;
                                                                    }

                                                            }
                                                            break;
                                                        }
                                                }
                                            }
                                            else
                                            {
                                                SendMessage("Печатная форма \"" + moxelid + "\" не найдена в конфигурации \"" + filename + "\"", 200, null, admin, true);
                                            }


                                        }
                                        catch (Exception ex)
                                        {
                                            SendMessage("Ошибка получения печатной формы \"" + moxelid + "\" из конфигурации \"" + filename + "\": " + ex.ToString(), 200, null, admin, true);
                                        }

                                    }
                                    else
                                    {
                                        SendMessage("Не указано имя печатной формы", 200, null, admin, true);
                                    }
                                }
                                else
                                {
                                    SendMessage("Не найден файл " + filename, 200, null, admin, true);
                                }


                                break;
                            }

                        case "getmoxelslist":
                            {
                                string filename = string.Empty;
                                string mode = "html";
                                mess = string.Empty;
                                try
                                {

                                    if (Options.ContainsKey("filepath"))
                                        filename = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Options["filepath"])), "1cv7.md");

                                    if (Options.ContainsKey("mode"))
                                        mode = Options["mode"];

                                    if (!File.Exists(filename))
                                    {
                                        SendMessage("Файл " + filename + " не найден.", 200, null, admin, true);
                                        break;
                                    }

                                    Compound.OleStorage.TaskItem v7Config = new Compound.OleStorage.TaskItem(filename);

                                    DataTable report = new DataTable("v7ConfigReport");
                                    report.Columns.Add(new DataColumn("Название печатной формы для пользователей", System.Type.GetType("System.String")));
                                    report.Columns.Add(new DataColumn("Путь к объекту", System.Type.GetType("System.String")));
                                    report.Columns.Add(new DataColumn("Описание", System.Type.GetType("System.String")));
                                    report.Columns.Add(new DataColumn("Идентификатор объекта", System.Type.GetType("System.String")));


                                    DataRow row;
                                    mess = string.Format("Собираем список форм конфигурации :<b>{0}</b><br><script></script>", filename);
                                    filename = System.Web.HttpUtility.UrlEncode(filename);
                                    if (v7Config.Moxels != null)
                                    {
                                        foreach (OleStorage.Moxel Obj in v7Config.Moxels.Values.ToList<OleStorage.Moxel>().FindAll(x => !x.IsEmpty))
                                        {
                                            row = report.NewRow();

                                            if (mode == "html")
                                                row["Идентификатор объекта"] = "<a href=\\getmoxel?filename=" + filename + "&moxelid=" + System.Web.HttpUtility.UrlEncode(Obj.MetaPath, Encoding.GetEncoding(1251)) + ">" + Obj.MetaPath + "</a>";
                                            else
                                                row["Идентификатор объекта"] = Obj.MetaPath;

                                            if (Obj.Header == string.Empty)
                                                row["Название печатной формы для пользователей"] = Obj.Description;
                                            else
                                                row["Название печатной формы для пользователей"] = Obj.Header;

                                            row["Описание"] = Obj.Description;
                                            row["Путь к объекту"] = "";
                                            report.Rows.Add(row);
                                        }
                                    }

                                    MemoryStream MS = new MemoryStream();
                                    TextWriter SW;

                                    if (mode == "xls")
                                    {
                                        thisClient.Response.Headers.Clear();
                                        thisClient.Response.AddHeader("content-disposition", "attachment; filename=\"Report.xls\"");
                                        thisClient.Response.ContentType = "application/vnd.ms-excel";
                                        SW = new StreamWriter(thisClient.Response.OutputStream, Encoding.GetEncoding(1251));
                                        System.Web.UI.WebControls.GridView excel = new System.Web.UI.WebControls.GridView();
                                        excel.DataSource = report;
                                        excel.DataBind();

                                        excel.RenderControl(new System.Web.UI.HtmlTextWriter(SW));
                                        SW.Flush();
                                        thisClient.Response.OutputStream.Close();
                                        return;
                                    }


                                    if (mode == "html")
                                    {
                                        mess = "<table name='result'> <tbody> <tr class=\"pheader\">";
                                        foreach (DataColumn Col in report.Columns)
                                            mess += "<td>" + Col.ColumnName + "</td>";
                                        mess += "</tr></tbody> <tbody class=\"panel\"><tr>";
                                        foreach (DataRow Row in report.Rows)
                                        {
                                            mess += string.Format("<td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr><tr>", Row["Название печатной формы для пользователей"], Row["Путь к объекту"], Row["Описание"], Row["Идентификатор объекта"]);
                                        }
                                        mess += "</tr></tbody></table>";
                                    }
                                    SendMessage(mess, 200, null, admin, true);

                                }
                                catch (Exception ex)
                                {
                                    SendMessage("Ошибка получения списка печатных форм: " + ex, 200, null, admin, true);
                                }

                                break;
                            }
                        case "getmoxelsextlist":
                            {
                                string mode = "html";
                                string filename = "";
                                mess = string.Empty;
                                List<string> FileList = new List<string>();
                                try
                                {
                                    if (Options.ContainsKey("mode"))
                                        mode = Options["mode"];
                                    string Query = @"

    declare @Top int = null, @exclude int = NULL, @RightID int = 569;

	WITH cte(Level, RecID, FileID, VersionID, Description, ParentID, Name, Path, IsFolder) as
	(
	SELECT 
		0 as Level,
		RecID,
        FileID,
        VersionID, 
        Description,		
        ParentID,
		FileName as Name,
		cast('Внешняя Отчетность' as varchar(500)) as Path,
		IsFolder
	from [dlfe].[FILE_1C_Reports]
	WHERE
			(RecId = @Top OR @Top is NULL)
		AND ParentId IS NULL
	UNION ALL 
	SELECT 
	cte.Level + 1 as Level,
	r.RecID, 
    r.FileID, 
    r.VersionID,
    r.Description,
	r.ParentID,
	r.FileName as Name,
	Cast(cte.Path + '\' +cast(cte.Name as varchar(500)) as varchar(500)) as Path ,
	r.IsFolder
	from [dlfe].[FILE_1C_Reports] as r, cte
	WHERE 		
		cte.RecID = r.ParentId 
		AND r.IsActive = 1
)

SELECT 
    'ВнешнийОтчет.' + Replace([Name],' ','') [ObjectId]
    ,Reports.Name as FileName
	,Reports.Path
    ,Reports.Description  AS Description
    ,FileVersion.Content AS Content
FROM cte as Reports

INNER JOIN
		files.dlfe.FileVersions AS FileVersion ON (FileVersion.FileID = Reports.FileID) AND (FileVersion.VersionID = Reports.VersionID) AND ((RTrim(FileVersion.ContentType) = 'application/1C Report'))

ORDER BY 
	Reports.Name
	";
                                    DataTable Data = new DataTable("Result");

                                    if (thisConnection.State != System.Data.ConnectionState.Open) thisConnection.Open();
                                    {
                                        SqlDataAdapter DA = new SqlDataAdapter(Query, thisConnection);

                                        DA.Fill(Data);

                                    }

                                    DataTable report = new DataTable("v7ConfigReport");
                                    DataRow row;
                                    report.Columns.Add(new DataColumn("Название печатной формы для пользователей", System.Type.GetType("System.String")));
                                    report.Columns.Add(new DataColumn("Путь к объекту", System.Type.GetType("System.String")));
                                    report.Columns.Add(new DataColumn("Описание", System.Type.GetType("System.String")));
                                    report.Columns.Add(new DataColumn("Идентификатор объекта", System.Type.GetType("System.String")));

                                    filename = System.Web.HttpUtility.UrlEncode(filename);

                                    Compound.OleStorage.ExtReport v7Report;
                                    foreach (DataRow Row in Data.Rows)
                                    {
                                        filename = Path.Combine(Path.GetTempPath(), Row["FileName"].ToString());
                                        row = report.NewRow();

                                        try
                                        {

                                            if (File.Exists(filename))
                                            {
                                                GC.Collect();
                                                GC.Collect();
                                                GC.WaitForPendingFinalizers();
                                                File.Delete(filename);
                                            }

                                            File.WriteAllBytes(filename, (byte[])Row["Content"]);
                                            v7Report = new Compound.OleStorage.ExtReport(filename);
                                            FileList.Add(filename);

                                            filename = System.Web.HttpUtility.UrlEncode(v7Report.Identity, Encoding.GetEncoding(1251));
                                            foreach (OleStorage.Moxel Obj in v7Report.Form.Moxels.FindAll(x => !x.IsEmpty))
                                            {

                                                row = report.NewRow();

                                                if (mode == "html")
                                                    row["Идентификатор объекта"] = "<a href=\\getmoxel?filename=" + filename + "&moxelid=" + System.Web.HttpUtility.UrlEncode(Obj.MetaPath, Encoding.GetEncoding(1251)) + ">" + Obj.MetaPath + "</a>";
                                                else
                                                    row["Идентификатор объекта"] = Obj.MetaPath;

                                                if (Obj.Header == string.Empty)
                                                    row["Название печатной формы для пользователей"] = Obj.Description;
                                                else
                                                    row["Название печатной формы для пользователей"] = Obj.Header;

                                                row["Описание"] = Row["Description"].ToString();
                                                row["Путь к объекту"] = Row["Path"].ToString() + "\\" + v7Report.Identity;
                                                report.Rows.Add(row);
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            row["Идентификатор объекта"] = "ВнешняяОбработка." + Path.GetFileName(filename);
                                            row["Описание"] = "ОШИБКА: " + ex.Message;
                                            report.Rows.Add(row);
                                        }



                                    }


                                    MemoryStream MS = new MemoryStream();
                                    TextWriter SW;

                                    switch (mode)
                                    {
                                        case "xls":
                                            {
                                                thisClient.Response.Headers.Clear();
                                                thisClient.Response.AddHeader("content-disposition", "attachment; filename=\"Report.xls\"");
                                                thisClient.Response.ContentType = "application/vnd.ms-excel";
                                                SW = new StreamWriter(thisClient.Response.OutputStream, Encoding.GetEncoding(1251));
                                                System.Web.UI.WebControls.GridView excel = new System.Web.UI.WebControls.GridView();
                                                excel.DataSource = report;
                                                excel.DataBind();
                                                excel.RenderControl(new System.Web.UI.HtmlTextWriter(SW));
                                                SW.Flush();
                                                thisClient.Response.OutputStream.Close();
                                                break;
                                            }
                                        case "html":
                                            {
                                                mess = "<table name='result'> <tbody> <tr class=\"pheader\">";
                                                foreach (DataColumn Col in report.Columns)
                                                    mess += "<td>" + Col.ColumnName + "</td>";
                                                mess += "</tr></tbody> <tbody class=\"panel\"><tr>";
                                                foreach (DataRow Row in report.Rows)
                                                {
                                                    mess += string.Format("<td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr><tr>", Row["Название печатной формы для пользователей"], Row["Путь к объекту"], Row["Описание"], Row["Идентификатор объекта"]);
                                                }
                                                mess += "</tr></tbody></table>";

                                                SendMessage(mess, 200, null, admin, true);
                                                break;
                                            }

                                    }

                                    GC.Collect();
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                    foreach (string FileNAme in FileList)
                                        File.Delete(FileNAme);
                                }
                                catch (Exception ex)
                                {
                                    SendMessage("Ошибка получения списка печатных форм: " + ex, 200, null, admin, true);
                                }

                                break;
                            }
                        case "searchinextreports":
                            {
                                string searchstring = string.Empty;
                                mess = string.Empty;

                                if (Options.ContainsKey("searchstring"))
                                    searchstring = Options["searchstring"].Trim();

                                if (searchstring == string.Empty)
                                {
                                    mess = @"
<form id=repaction  method=""post"" name=repaction>
<br>Строка для поиска: <input  name=searchstring style=""width: 700px;"" value=''>
<br><input  formAction=""searchinextreports"" id=""searchinextreports"" type=submit class=""b_w"" value='Найти во внешней отчетности' onclick=""act(this)"">
</form>";
                                }
                                else
                                {
                                    List<string> FileList = new List<string>();
                                    try
                                    {
                                        string Query = @"
SELECT 
     	Reports.FileName  AS FileName
	,   FileVersion.Content AS Content
    ,   FileVersion.VersionID as Vid
    ,   Reports.FileID as FileID
	FROM 
		 dlfe.FILE_1C_Reports AS Reports
	LEFT JOIN
		files.dlfe.FileVersions AS FileVersion ON (FileVersion.FileID = Reports.FileID) AND (FileVersion.VersionID = Reports.VersionID)
	WHERE (Reports.IsActive = 1) AND
	((RTrim(FileVersion.ContentType) = 'application/1C Report'))
	ORDER BY 
	Reports.FileName";
                                        DataTable Data = new DataTable("Result");

                                        if (thisConnection.State != System.Data.ConnectionState.Open) thisConnection.Open();
                                        {
                                            SqlDataAdapter DA = new SqlDataAdapter(Query, thisConnection);
                                            DA.Fill(Data);
                                        }

                                        DataTable report = new DataTable("v7ConfigReport");

                                        report.Columns.Add(new DataColumn("Объект", System.Type.GetType("System.String")));
                                        report.Columns.Add(new DataColumn("Номер строки", System.Type.GetType("System.Int32")));
                                        report.Columns.Add(new DataColumn("Строка кода", System.Type.GetType("System.String")));

                                        Compound.OleStorage.ExtReport v7Report;
                                        DataRow row;
                                        string filename = "";

                                        foreach (DataRow Row in Data.Rows)
                                        {
                                            filename = Path.Combine(Path.GetTempPath(), Row["FileName"].ToString());
                                            try
                                            {
                                                if (File.Exists(filename))
                                                {
                                                    GC.Collect();
                                                    GC.Collect();
                                                    GC.WaitForPendingFinalizers();
                                                    File.Delete(filename);
                                                }

                                                File.WriteAllBytes(filename, (byte[])Row["Content"]);
                                                v7Report = new Compound.OleStorage.ExtReport(filename);
                                                FileList.Add(filename);
                                                int counter = 0;
                                                if (v7Report.Form.DialogModule.Moduletext.Contains(searchstring))
                                                {
                                                    foreach (string modulestring in v7Report.Form.DialogModule.Moduletext.Split(new string[] { "\r\n" }, StringSplitOptions.None))
                                                    {
                                                        counter++;
                                                        if (modulestring.Contains(searchstring))
                                                        {

                                                            row = report.NewRow();
                                                            row["Объект"] = "<a href=\"https://sfs.siemens.ru/secure/download.ashx?fileid=" + Row["FileID"] + "&vid=" + Row["Vid"] + "\">" + v7Report.Identity + "</a>";
                                                            row["Номер строки"] = counter;
                                                            row["Строка кода"] = modulestring;
                                                            report.Rows.Add(row);
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {

                                            }

                                        }

                                        if (report.Rows.Count > 0)
                                        {

                                            string ttt = "<table name='result'> <tbody> <tr class=\"pheader\">";
                                            foreach (DataColumn Col in report.Columns)
                                                ttt += "<td>" + Col.ColumnName + "</td>";
                                            ttt += "</tr></tbody> <tbody class=\"panel\"><tr>";
                                            foreach (DataRow Row in report.Rows)
                                            {
                                                ttt += string.Format("<td>{0}</td><td class=number>{1}</td><td>{2}</td></tr><tr>", Row["Объект"], Row["Номер строки"], Row["Строка кода"]);
                                            }
                                            ttt += "</tr></tbody></table>";

                                            mess += ttt;

                                        }
                                        else
                                            mess = "Ничего не найдено";

                                        mess += "<script></script>";
                                        SendMessage(mess, 200, null, admin, true);

                                    }
                                    catch (Exception ex)
                                    {
                                        mess += "<br>Ошибка: " + ex.ToString();
                                    }
                                    GC.Collect();
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                    try
                                    {
                                        foreach (string FileNAme in FileList)
                                            File.Delete(FileNAme);
                                    }
                                    catch
                                    {
                                    }

                                }
                                SendMessage(mess, 200, null, admin, true);
                                break;
                            }
                        case "viewsessions":
                            {
                                mess = string.Empty;
                                try
                                {
                                    using (System.Net.NetworkInformation.Ping pinger = new System.Net.NetworkInformation.Ping())
                                    {

                                        Utils.RefillLists();

                                        foreach (Utils.WTS_SESSION_INFO_1 Session in Utils.SessID.Values)
                                        {
                                            if (Utils.Employees.ContainsKey(Session.pUserName))
                                            {
                                                var emp = Utils.Employees.Where(x => x.Key == Session.pUserName).First();

                                                string WorkStation = Utils.GetSessionClientName(IntPtr.Zero, Session.SessionId);

                                                if (WorkStation.Length <= 1)
                                                {
                                                    if (Utils.Clients.ContainsKey(emp.Key))
                                                        WorkStation = Utils.Clients[emp.Key];
                                                }

                                                if (WorkStation.Length > 1)
                                                {
                                                    bool Online = false;
                                                    try
                                                    {
                                                        System.Net.NetworkInformation.PingReply reply = pinger.Send(WorkStation, 1500);
                                                        Online = reply.Status == System.Net.NetworkInformation.IPStatus.Success;
                                                        WorkStation += Online ? " <font color='green'>(Online, пинг:" + reply.RoundtripTime + " мс.)</font>" : " <font color='red'>(Offline)</font>";
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        WorkStation += " (" + ex.Message + ")";
                                                    }

                                                }
                                                else
                                                    WorkStation = "<font color='blue'>Не определена</font>";

                                                TimeSpan IdleTime = Utils.GetSessionIdleTime(IntPtr.Zero, Session.SessionId);
                                                mess += string.Format("<p>Логин: <b>{0}</b>, Имя: <b>{1}</b>, Машина: <b>{2}</b>, Время простоя: {3}</p>", emp.Key, emp.Value, WorkStation, IdleTime.ToString(@"hh\:mm\:ss"));
                                            }
                                        }
                                    }
                                }
                                catch(Exception ex)
                                {
                                    mess += "<br> Ошибка получения списка рабочих станций: " + ex.Message;
                                }

                                SendMessage(mess, 200, null, admin, true);
                                break;
                            }

                        default:
                            {
                                SendMessage("Команда <b>" + Command + "</b> пока не реализована. ", 404, null, admin, true);
                                break;
                            }
                    }

                }
                #endregion

            }

        }

    }


    class Server
    {

        public static string DomainName = Environment.UserDomainName;
        public static string DomainController = DomainName;
        public static bool InDomain = false;
        public static string sessionListHTML = "<b>Список обновляется...</b>\r\n";
        public static string secIDListHTML = "<b>Список обновляется...</b>\r\n";
        public static string DBlistHTML = "<select name=\"filepath\" class=\"dblist\" size=3><option>Список обновляется...</option></select>";
        public static List<string> admins = new List<string>();
        public static Dictionary<string, byte[]> FileCache = new Dictionary<string, byte[]>();

        //public static TcpListener Listener; // Объект, принимающий TCP-клиентов
        public static HttpListener Listener; // Объект, принимающий TCP-клиентов
        //public static MemoryStream ConTmp = new MemoryStream();
        //public static string logfile = ;
        public static FileStream logfileW = File.Open(System.Windows.Forms.Application.ExecutablePath.ToLower().Replace("exe", "log"), FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
        public static FileStream logfileR = File.Open(System.Windows.Forms.Application.ExecutablePath.ToLower().Replace("exe", "log"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        public static StreamWriter sw1 = new StreamWriter(logfileW);
        
        public static bool IsDisposed = false;
        static bool RestartServer = false;

        // Запуск сервера
        public Server(int Port)
        {
                
            string CurrentUser = Environment.UserName;
            try
            {

                using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
                {
                    InDomain = true;
                    
                    DomainController = Options.Get("DomainController");
                    if (DomainController == string.Empty || Server.DomainName != Options.Get("Domain"))
                    {
                        Console.WriteLine("{0}; {1}; Проверяем контролер домена:", DateTime.Now.ToString(), CurrentUser);
                        DomainController = context.ConnectedServer;
                        Console.WriteLine("{0}; {1}; Контроллер домена найден: {2}", DateTime.Now.ToString(), CurrentUser, DomainController);
                        Options.Set("DomainController", DomainController);
                        Options.Set("Domain", Server.DomainName);
                    }
                    else
                    {
                        Console.WriteLine("{0}; {1}; Домен не менялся, имя контроллера прочитано из настроек: {2}", DateTime.Now.ToString(), CurrentUser, DomainController);
                    }
                    
                }
            }
            catch (Exception ex)
            {
                DomainController = string.Empty;
                InDomain = false;
            }

                Listener = new HttpListener();
                Listener.Prefixes.Add("http://*:"+port+"/");
                Listener.Prefixes.Add("https://*/");
                Listener.AuthenticationSchemes = AuthenticationSchemes.Ntlm;
                Listener.Start(); // Запускаем его

                
                Console.WriteLine("{0}; {1}; Запускаем поток чтения списков:", DateTime.Now.ToString(), CurrentUser);
                ThreadPool.QueueUserWorkItem(new WaitCallback(UtilitiesThread), null);
                Console.WriteLine("{0}; {1}; Запускаем поток чтения списков: OK", DateTime.Now.ToString(), CurrentUser);

                // В бесконечном цикле
                Console.WriteLine("{0}; {1}; Сервер запущен, порт: {2}", DateTime.Now.ToString(), CurrentUser, Port);
                while (!IsDisposed)
                {
                    try
                    {
                        HttpListenerContext Context = Listener.GetContext();
                        ThreadPool.QueueUserWorkItem(new WaitCallback(ClientThread), Context);
                        tryNumber = 0;
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine("{0}; {1}; Общая ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
                        if (!Listener.IsListening)
                            Listener.Start();
                    }
                }
           
        }

        static void ClientThread(Object StateInfo)
        {
            new Client((HttpListenerContext)StateInfo);
        }

        static void UtilitiesThread(Object StateInfo)
        {
            new Utils();
        }


        static int port = 80;
        static int tryNumber = 0;

        static void  SystemEvents_EventsThreadShutdown(object sender, SessionEndingEventArgs e)
        {
            IsDisposed = true;
            if (e.Reason == SessionEndReasons.SystemShutdown)
            {
                Console.WriteLine("{0}; {1}; Сервис принудительно остановлен. (Сервер ушел на перезагрузку)", DateTime.Now.ToString(), Environment.UserName);
                   // Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\",true).SetValue("KickOFF Server","\""+System.Windows.Forms.Application.ExecutablePath + "\" /p:80 /r:AfterRestart");
            }
            else
            {
                Console.WriteLine("{0}; {1}; Сервис принудительно остановлен. (Системный пользователь разлогинился)", DateTime.Now.ToString(), Environment.UserName);
            }
            //sw1.Flush();
            //File.AppendAllText(System.Windows.Forms.Application.ExecutablePath.ToLower().Replace("exe","log"), Encoding.GetEncoding(1251).GetString(ConTmp.GetBuffer()).TrimEnd() + "\r\n");
        }

        enum StartReasons
        { 
            NormalStart,
            StartAfterRestart
        }

        private static void HandleError(Exception exception)
        {

            Console.WriteLine("{0}; {1}; Необработанное исключение: {2}", DateTime.Now.ToString(), Environment.UserName, exception.ToString());
            if (((System.Net.HttpListenerException)(exception)).ErrorCode == 183) 
            { 
                Console.WriteLine("{0}; {1}; Порт занят.", DateTime.Now.ToString(), Environment.UserName);
                tryNumber = 30;
            }
                

            if (tryNumber < 30)
            {
                Console.WriteLine("{0}; {1}; Попытка запуска - {2}", DateTime.Now.ToString(), Environment.UserName, tryNumber);
                tryNumber++;

                if (Listener.IsListening)
                {
                    Listener.Stop();
                }
                Server.Restart(tryNumber);
            }
            //sw1.Flush();
            //File.AppendAllText(System.Windows.Forms.Application.ExecutablePath.ToLower().Replace("exe", "log"), Encoding.GetEncoding(1251).GetString(ConTmp.GetBuffer()).TrimEnd() + "\r\n");
            //Environment.Exit(1);

        }

        static void Main(string[] args)
        {

            AppDomain.CurrentDomain.UnhandledException += (sender, e) => HandleError((Exception)e.ExceptionObject);

            Console.SetOut(sw1);
            Server.sw1.AutoFlush = true;

            WindowsPrincipal pricipal = new WindowsPrincipal(WindowsIdentity.GetCurrent());
            bool hasAdministrativeRight = pricipal.IsInRole(WindowsBuiltInRole.Administrator);
            if (!hasAdministrativeRight)
            {
                ProcessStartInfo processInfo = new ProcessStartInfo(); //создаем новый процесс
                processInfo.Verb = "runas"; //в данном случае указываем, что процесс должен быть запущен с правами администратора
                processInfo.FileName = System.Windows.Forms.Application.ExecutablePath; //указываем исполняемый файл (программу) для запуска
                Process.Start(processInfo); //пытаемся запустить процесс
                Environment.Exit(1);
            }
            
            
            int MaxThreadsCount = 36;
            // Установим максимальное количество рабочих потоков
            ThreadPool.SetMaxThreads(MaxThreadsCount, MaxThreadsCount);
            ThreadPool.SetMinThreads(1,1);



            StartReasons StartReason = StartReasons.NormalStart;
           // SetConsoleCtrlHandler(new HandlerRoutine(ConsoleCtrlCheck), true); //Обработаем события консоли
            SystemEvents.SessionEnding += new SessionEndingEventHandler(SystemEvents_EventsThreadShutdown);
           // System.Windows.Forms.Application.ApplicationExit += new EventHandler(Application_ThreadExit);


            
            if (args.Length != 0) 
            {
                string par = string.Empty;
                for (int i=0; i < args.Length; i++)
                {
                    par = args[i].ToLower();
                    switch (par.Substring(0,2)) 
                    { 
                        case "-p":
                        case "/p":
                            port = Convert.ToInt16(par.Split(':')[1]);
                            break;
                        case "-r":
                        case "/r":
                            if (par.Split(':')[1] == "afterrestart") StartReason = StartReasons.StartAfterRestart;
                            break;
                        case "-n":
                        case "/n":
                            tryNumber = Convert.ToInt16(par.Split(':')[1]);
                            break;
                    }
                }

            }

            if (StartReason == StartReasons.StartAfterRestart) 
            {
                Console.WriteLine("{0}; {1}; Сервер был перезагружен", DateTime.Now.ToString(), Environment.UserName);
                //Console.WriteLine("{0}; {1}; Удаляемся из автозагрузки", DateTime.Now.ToString(), Environment.UserName);
                //Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\",true).DeleteValue("KickOFF Server", false);
            }

            if (!IsDisposed) new Server(port);

            if (Listener.IsListening)
            {
                // Остановим его
                Listener.Stop();
            }
            if (RestartServer)
            {
                Process NewVer = new Process();
                NewVer.StartInfo.FileName = System.Windows.Forms.Application.ExecutablePath;
                NewVer.StartInfo.Arguments = "/p:" + port + " /n:" + tryNumber;
                NewVer.Start();
                Console.WriteLine("{0}; {1}; Запущен дочерний процесс.", DateTime.Now.ToString(), Environment.UserName);
            }

            Console.WriteLine("{0}; {1}; Сервер остановлен.", DateTime.Now.ToString(), Environment.UserName);
            return;
        }

        public static void Restart(int tryNumber = 0)
        {
            if (IsDisposed)
                return;
            Console.WriteLine("{0}; {1}; Сервис перезапускается.", DateTime.Now.ToString(), Environment.UserName);
            try
            {
                if (Listener.IsListening)
                {
                    RestartServer = true;
                    IsDisposed = true;
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0}; {1}; Общая ошибка: {2} в методе {3}", DateTime.Now.ToString(), Environment.UserName, ex.Message, ex.TargetSite.DeclaringType);
            }
        }


        private static bool ConsoleCtrlCheck(CtrlTypes ctrlType)
        {
            IsDisposed = true;
            switch (ctrlType)
            {
                case CtrlTypes.CTRL_C_EVENT:
                case CtrlTypes.CTRL_BREAK_EVENT:
                case CtrlTypes.CTRL_CLOSE_EVENT:
                    Console.WriteLine("{0}; {1}; Сервер принудительно остановлен. (Закрыто окно консоли)", DateTime.Now.ToString(), Environment.UserName);
                    //sw1.Flush();
                    //File.AppendAllText(System.Windows.Forms.Application.ExecutablePath.ToLower().Replace("exe", "log"), Encoding.GetEncoding(1251).GetString(ConTmp.GetBuffer()).TrimEnd() + "\r\n");
                    return true;
                case CtrlTypes.CTRL_LOGOFF_EVENT:
                    SystemEvents_EventsThreadShutdown(null, new SessionEndingEventArgs(SessionEndReasons.Logoff));
                    break;
                case CtrlTypes.CTRL_SHUTDOWN_EVENT:
                    SystemEvents_EventsThreadShutdown(null, new SessionEndingEventArgs(SessionEndReasons.SystemShutdown));
                    break;
            }
            return true;
        }

        #region unmanaged
        // Declare the SetConsoleCtrlHandler function
        // as external and receiving a delegate.
        [DllImport("Kernel32")]
        public static extern bool SetConsoleCtrlHandler(HandlerRoutine Handler, bool Add);
        // A delegate type to be used as the handler routine
        // for SetConsoleCtrlHandler.
        public delegate bool HandlerRoutine(CtrlTypes CtrlType);
        // An enumerated type for the control messages
        // sent to the handler routine.

        public enum CtrlTypes
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT,
            CTRL_CLOSE_EVENT,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT
        }
        #endregion
    }
}
