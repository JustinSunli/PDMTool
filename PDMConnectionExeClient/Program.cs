using System.IO;
using PDMConnection;

namespace PDMConnectionExeClient
{
    class Program
    {
        const int ARG_ERROR = -1000;

        static int Main(string[] args)
        {
            if (args.Length < 1)
                return ARG_ERROR;

            string command = args[0];
            int ret;
            switch (command)
            {
                case "-Refresh":
                    ret = Refresh(args);
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

        static int Refresh(string[] args)
        {
            if (args.Length < 5)
                return ARG_ERROR;

            string vaultName = args[1];
            string userName = args[2];
            string password = args[3];
            string[] fileRelativePaths = args[4].Split('|');

            return PDMConnectionTool.RefreshFiles(vaultName, userName, password, fileRelativePaths);
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
