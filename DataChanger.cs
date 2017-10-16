using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.OleDb;
using System.Xml.Serialization;
using System.Xml;
using System.Data;
namespace DataChanger
{
    class Program
    {
        const string _update = "-update";
        const string _configPath = "-config path";
        const string _targetPath = "-target path";
        const string _help = "-help";
        const string _exit = "-exit";

        static PathController pathController;
        static PathRecorder path;
        static bool isFinished = false;
        static bool willExit = false;
        static void Main(string[] args)
        {
            Console.Title = "ExcelChanger";
            //获得配置表和Json路径
            pathController = new PathController();
            path = pathController.GetXMLPath();
            Console.WriteLine("input commond , use {0} to get more", _help);
            while (true)
            {
                string userInput = Console.ReadLine().ToLower();
                Commond(userInput);
                if (isFinished)
                {
                    break;
                }
                if (willExit)
                {
                    return;
                }
            }
            Console.WriteLine("press any key to exit");
            Console.ReadKey();
        }
        private static void Commond(string commond)
        {

            switch (commond)
            {
                case _update:
                    Commond_Update();
                    break;
                case _configPath:
                    Commond_ConfigPathSet();
                    break;
                case _targetPath:
                    Commond_TargetPathSet();
                    break;
                case _help:
                    Commond_Help();
                    break;
                case _exit:
                    willExit = true;
                    break;
                default:
                    Console.WriteLine("useless commond \nyou can input {0} to show more", _help);
                    break;
            }
        }
        private static void Commond_Update()
        {
            if (path == null || path.configPath == null || path.targetPath == null)
            {
                Console.WriteLine("Please set the correct path");
                return;
            }
            //读取Excel并进行数据处理
            MyExcelDataTransformer excelReader = new MyExcelDataTransformer(path.configPath, path.targetPath);
            excelReader.ExcelDataRead();
            isFinished = true;
        }
        private static void Commond_ConfigPathSet()
        {
            PathRecorder tempPath = pathController.SetPath(true);
            if (tempPath != null)
            {
                path = tempPath;
            }
        }
        private static void Commond_TargetPathSet()
        {
            PathRecorder tempPath = pathController.SetPath(false);
            if (tempPath != null)
            {
                path = tempPath;
            }
        }
        private static void Commond_Help()
        {
            Console.WriteLine("{0}\t generate target data", _update);
            Console.WriteLine("{0}\t set config path ", _configPath);
            Console.WriteLine("{0}\t set target path", _targetPath);
            Console.WriteLine("{0}\t exit application", _exit);
        }
    }
    /// <summary>
    /// 本地序列化的路径
    /// </summary>
    public class PathRecorder
    {
        public string configPath;
        public string targetPath;
    }
    public class PathController
    {
        const string pathFileName = "pathRecorder.xml";
        string curPath = System.Environment.CurrentDirectory;
        public PathRecorder SetPath(bool isConfigPath)
        {
            PathRecorder pathRecorder;
            string tempPath;
            while (true)
            {
                tempPath = Console.ReadLine();
                if (new DirectoryInfo(tempPath).Exists)
                {
                    Console.WriteLine("Set {0} path: {1}",
                        isConfigPath ? "config" : "target"
                        , tempPath);
                    break;
                }
                else
                {
                    Console.WriteLine("wrong path,input again");
                    return null;
                }
            }
            DirectoryInfo di = new DirectoryInfo(curPath);
            FileInfo[] files = di.GetFiles(pathFileName);
            string fullPath = string.Format("{0}/{1}", curPath, pathFileName);
            //存在XML文件时
            if (files != null && files.Length > 0)
            {
                FileStream xmlFS;
                XmlSerializer xmlserializer = new XmlSerializer(typeof(PathRecorder));
                xmlFS = new FileStream(fullPath, FileMode.Open, FileAccess.Read);
                xmlFS.Position = 0;
                pathRecorder = xmlserializer.Deserialize(xmlFS) as PathRecorder;
                if (isConfigPath)
                {
                    pathRecorder.configPath = tempPath;
                }
                else
                {
                    pathRecorder.targetPath = tempPath;
                }
                xmlFS.Close();
                xmlFS.Dispose();
                xmlFS = new FileStream(fullPath, FileMode.Create, FileAccess.Write);
                xmlserializer.Serialize(xmlFS, pathRecorder);
                xmlFS.Close();
                xmlFS.Dispose();
            }
            //不存在时
            else
            {
                FileStream fs = new FileStream(curPath + "/" + pathFileName, FileMode.Create, FileAccess.Write);
                XmlSerializer xmlSerialize = new XmlSerializer(typeof(PathRecorder));
                pathRecorder = new PathRecorder();
                if (isConfigPath)
                {
                    pathRecorder.configPath = tempPath;
                }
                else
                {
                    pathRecorder.targetPath = tempPath;
                }
                xmlSerialize.Serialize(fs, pathRecorder);
                fs.Close();
                fs.Dispose();
            }
            return pathRecorder;
        }
        /// <summary>
        /// 获得相应路径文件夹
        /// </summary>
        public PathRecorder GetXMLPath()
        {
            string fullPath = string.Format("{0}/{1}", curPath, pathFileName);
            DirectoryInfo di = new DirectoryInfo(curPath);
            FileInfo[] files = di.GetFiles(pathFileName);
            //存在XML文件时
            if (files != null && files.Length > 0)
            {
                XmlSerializer xmlserializer = new XmlSerializer(typeof(PathRecorder));
                FileStream xmlFS = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite);
                xmlFS.Position = 0;
                PathRecorder pathRecorder;
                //处理非法XML文件
                try
                {
                    pathRecorder = xmlserializer.Deserialize(xmlFS) as PathRecorder;
                }
                catch (Exception)
                {
                    Console.WriteLine("wrong xml path,please set again");
                    xmlFS.Close();
                    xmlFS.Dispose();
                    files[0].Delete();
                    return null;
                }
                xmlFS.Close();
                xmlFS.Dispose();
                Console.WriteLine(String.Format("config path: {0}", pathRecorder.configPath));
                Console.WriteLine(String.Format("target path: {0}", pathRecorder.targetPath));
                return pathRecorder;
            }
            //不存在时
            else
            {
                Console.WriteLine("You should set path first");
                return null;
            }
        }
    }
    public class MyExcelDataTransformer
    {
        public MyExcelDataTransformer(string path, string targetPath)
        {
            this.path = path;
            this.targetPath = targetPath;
        }
        string path;
        string targetPath;
        struct pathStruct
        {
            public string path;
            public bool isXlsx;
            public pathStruct(string path, bool isXlsx)
            {
                this.path = path;
                this.isXlsx = isXlsx;
            }
        }
        /// <summary>
        /// 获取所有EXCEL文件路径
        /// </summary>
        private pathStruct[] GetAllOfExcelFilesPath()
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(path);
            FileInfo[] filesInfo = directoryInfo.GetFiles();
            List<pathStruct> pathStructs = new List<pathStruct>();
            for (int i = 0; i < filesInfo.Length; i++)
            {
                //排除打开EXCEL生成的临时文件
                if (filesInfo[i].Name.Contains(".xlsx") && !filesInfo[i].Name.Contains("~$"))
                {
                    pathStructs.Add(new pathStruct(filesInfo[i].FullName, true));
                }
                else if (filesInfo[i].Name.Contains(".xls") && !filesInfo[i].Name.Contains("~$"))
                {
                    pathStructs.Add(new pathStruct(filesInfo[i].FullName, false));
                }
            }
            return pathStructs.ToArray();
        }
        /// <summary>
        /// 将EXCEL数据拼接为类CSV（用^分割）字符串
        /// </summary>
        private string ExcelDataHandle(DataSet dataSet, string tableName)
        {
            DataTable data = dataSet.Tables[0];
            int columnValue = data.Columns.Count;
            //清除列 Remark备注
            for (int i = 0; i < columnValue; i++)
            {
                if (data.Rows[0][i].ToString().ToLower().Equals("remark"))
                {
                    data.Columns.RemoveAt(i);
                    columnValue = data.Columns.Count;
                    break;
                }
            }
            //清除 行备注
            data.Rows.Remove(data.Rows[1]);
            int rowValue = data.Rows.Count;
            StringBuilder stringBuilder = new StringBuilder();
            for (int i = 0; i < rowValue; i++)
            {
                for (int j = 0; j < columnValue; j++)
                {
                    stringBuilder.Append(data.Rows[i][j].ToString());
                    if (j != columnValue - 1)
                    {
                        stringBuilder.Append('^');
                    }
                }
                if (i != rowValue - 1)
                {
                    stringBuilder.Append("\r\n");
                }
            }
            return stringBuilder.ToString(); ;
        }
        public void ExcelDataRead()
        {
            //EXCEL路径处理
            pathStruct[] excelFilesPath = GetAllOfExcelFilesPath();
            DirectoryInfo directoryInfos = new DirectoryInfo(targetPath);
            //删除之前的文件
            FileInfo[] fileInfos = directoryInfos.GetFiles();
            for (int i = 0; i < fileInfos.Length; i++)
            {
                fileInfos[i].Delete();
            }
            Console.WriteLine("Delete previous txt file success: {0}", fileInfos.Length);

            for (int i = 0; i < excelFilesPath.Length; i++)
            {
                string _path;
                if (!excelFilesPath[i].isXlsx)
                {
                    _path = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelFilesPath[i].path + ";" + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
                }
                else
                {
                    _path = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + excelFilesPath[i].path + ";" + ";Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1\"";
                }
                OleDbConnection conn = new OleDbConnection(_path);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                DataSet dataSet;
                //EXCEL Table名称读取
                DataTable schemaTable = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                conn.Close();
                conn.Dispose();
                for (int j = 0; j < schemaTable.Rows.Count; j++)
                {
                    //EXCEL Table名称读取
                    string tableName = schemaTable.Rows[j]["Table_Name"].ToString().Replace("$", "").Trim();
                    //过滤掉默认名称表格
                    if (tableName.Equals("Sheet1") || tableName.Equals("Sheet2") || tableName.Equals("Sheet3"))
                    {
                        continue;
                    }
                    strExcel = "select * from [" + tableName + "$]";
                    myCommand = new OleDbDataAdapter(strExcel, _path);
                    dataSet = new DataSet();
                    myCommand.Fill(dataSet);
                    //将EXCEL转换为字符串
                    string targetJsonStr = ExcelDataHandle(dataSet, tableName);
                    //生成相应TXT文件
                    StreamWriter streamWriter = new StreamWriter(string.Format("{0}\\{1}.txt", targetPath, tableName));
                    streamWriter.Write(targetJsonStr);
                    streamWriter.Close();
                    streamWriter.Dispose();
                    Console.WriteLine("Generate new txt:  " + tableName);
                }
            }
        }
    }
}
