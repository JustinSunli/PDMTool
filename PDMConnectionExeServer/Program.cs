using System.IO;
using PDMConnection;

namespace PDMConnectionExeServer
{
    class Program
    {
        const int ARG_ERROR = -1000;
        const int EXCEPTION_ERROR = -1001;
        const string SIGN_VARIABLE_NAME = "标识";

        static int Main(string[] args)
        {
            if (args.Length < 1)
                return ARG_ERROR;
            
            try
            {
                string command = args[0];
                int ret;
                switch (command)
                {
                    case "-CheckConnection":
                        ret = CheckConnection(args);
                        break;
                    case "-CheckOut":
                        ret = CheckOut(args);
                        break;
                    case "-CheckIn":
                        ret = CheckIn(args);
                        break;
                    case "-Query":
                        ret = Query(args);
                        break;
                    case "-GetVaultPath":
                        ret = GetVaultPath(args);
                        break;
                    default:
                        ret = ARG_ERROR;
                        break;
                }

                string strLog = string.Empty;
                foreach (string arg in args)
                    strLog += '"' + arg + "\" ";
                strLog += "返回值:";
                strLog += ret.ToString();
                PDMConnection.Log.WriteLog(strLog);

                return ret;
            }
            catch (System.Exception ex)
            {
                PDMConnection.Log.WriteLog(ex.Message + "\r\n" + ex.StackTrace);
                return EXCEPTION_ERROR;
            }
        }

        static int CheckConnection(string[] args)
        {
            if (args.Length < 4)
                return ARG_ERROR;

            string vaultName = args[1];
            string userName = args[2];
            string password = args[3];
            return PDMConnectionTool.CheckConnection(vaultName, userName, password);
        }

        static int CheckOut(string[] args)
        {
            if (args.Length < 6)
                return ARG_ERROR;

            string vaultName = args[1];
            string userName = args[2];
            string password = args[3];
            bool flag = !(args[4] == "0");
            string files = args[5];
            string[] filePaths = files.Split('|');
            return PDMConnectionTool.CheckOutFiles(vaultName, userName, password, flag, filePaths);
        }

        static int CheckIn(string[] args)
        {
            if (args.Length < 6)
                return ARG_ERROR;

            string vaultName = args[1];
            string userName = args[2];
            string password = args[3];
            bool flag = !(args[4] == "0");
            string files = args[5];
            string[] filePaths = files.Split('|');
            return PDMConnectionTool.CheckInFiles(vaultName, userName, password, flag, filePaths);
        }

        static int Query(string[] args)
        {
            if (args.Length < 6)
                return ARG_ERROR;

            string vaultName = args[1];
            string userName = args[2];
            string password = args[3];
            string filePath = args[4];
            string signVariableValue = args[5];  //标识
            return PDMConnectionTool.QuerySignVariableChanged(vaultName, userName, password, filePath, SIGN_VARIABLE_NAME, signVariableValue);
        }

        static int GetVaultPath(string[] args)
        {
            if (args.Length < 5)
                return ARG_ERROR;

            string vaultName = args[1];
            string userName = args[2];
            string password = args[3];
            string outFilePath = args[4];

            string vaultPath;
            int ret = PDMConnectionTool.GetVaultRootPath(vaultName, userName, password, out vaultPath);
            if (ret == (int)PDMConnection.ResultCodes.LoginResult.OK)
            {
                string context = "[PDMInfo]\r\npath = " + vaultPath;
                File.WriteAllText(outFilePath, context);
            }
            return ret;
        }
    }
}
