using System;
using System.IO;
using System.Collections.Generic;
using EPDM.Interop.epdm;
using EPDM.Interop.EPDMResultCode;
using PDMConnection.ResultCodes;

namespace PDMConnection
{
    enum SWFileType
    {
        PART = 0,
        ASSEMBLY,
        DRAWING,
        UNKNOWN
    }

    public class PDMConnectionTool
    {
        #region public

        /// <summary>
        /// 检查链接
        /// </summary>
        /// <param name="vaultName">库名</param>
        /// <param name="userName">用户名</param>
        /// <param name="password">密码</param>
        /// <returns>LoginError</returns>
        public static int CheckConnection(string vaultName, string userName, string password)
        {
            IEdmVault5 vault;
            int ret = (int)Login(vaultName, userName, password, out vault);
            return ret;
        }

        /// <summary>
        /// 检出文件
        /// </summary>
        /// <param name="vaultName">库名</param>
        /// <param name="userName">用户名</param>
        /// <param name="password">密码</param>
        /// <param name="fileRelativePaths">要检出的文件列表</param>
        /// <returns>FileCheckOutError</returns>
        public static int CheckOutFiles(string vaultName, string userName, string password, bool bNeedErrorCode, string[] fileRelativePaths)
        {
            IEdmVault5 vault;
            if (Login(vaultName, userName, password, out vault) != LoginResult.OK || !vault.IsLoggedIn)
                return (int)FileCheckOutResult.LOGIN_FAILED;

            int ret = (int)FileCheckOutResult.OK;
            if (bNeedErrorCode)
            {
                //获取所有引用
                List<string> referenceFiles = new List<string>();
                for (int i = 0; i < fileRelativePaths.Length; ++i)
                {
                    string fileFullPath = Path.Combine(vault.RootFolderPath, fileRelativePaths[i]);
                    IEdmFile5 edmFile;
                    IEdmFolder5 edmFolder;
                    if(!GetFileFromPath(vault, fileFullPath, out edmFile, out edmFolder))
                        return (int)FileCheckOutResult.FILE_NOT_FOUND;
                    if (!GetReferencedFiles(vault, null, fileFullPath, "", ref referenceFiles))
                        return (int)FileCheckOutResult.UNKNOWN_ERROR;
                }
                
                for (int i = 0; i < referenceFiles.Count; ++i)
                {
                    ret = (int)LockFile(vault, userName, referenceFiles[i]);
                    if (ret != (int)FileCheckOutResult.OK)
                        return ret;
                }
            }
            else
            {
                for (int i = 0; i < fileRelativePaths.Length; ++i)
                    fileRelativePaths[i] = Path.Combine(vault.RootFolderPath, fileRelativePaths[i]);

                bool bIncludeReference = true;
                ret = (int)LockFiles(vault, fileRelativePaths, bIncludeReference);
            }

            return ret;
        }

        /// <summary>
        /// 检入文件
        /// </summary>
        /// <param name="vaultName">库名</param>
        /// <param name="userName">用户名</param>
        /// <param name="password">密码</param>
        /// <param name="fileRelativePaths">要检入的文件列表</param>
        /// <returns>FileCheckInError</returns>
        public static int CheckInFiles(string vaultName, string userName, string password, bool bNeedErrorCode, string[] filePaths)
        {
            IEdmVault5 vault;
            if (Login(vaultName, userName, password, out vault) != LoginResult.OK || !vault.IsLoggedIn)
                return (int)FileCheckInResult.LOGIN_FAILED;
            
            int ret = (int)FileCheckInResult.OK;
            if (bNeedErrorCode)
            {
                foreach (string filePath in filePaths)
                {
                    ret = (int)UnlockFile(vault, filePath);
                    if (ret != (int)FileCheckInResult.OK)
                        return ret;
                }
            }
            else
            {
                //有异常o(╥﹏╥)o
                //ret = (int)UnlockFiles(vault, filePaths);

                foreach (string filePath in filePaths)
                {
                    ret = (int)UnlockFile(vault, filePath);
                    if (ret != (int)FileCheckInResult.OK)
                    {
                        ret = (int)FileCheckInResult.UNKNOWN_ERROR;
                        break;
                    }
                }
            }

            return ret;
        }

        /// <summary>
        /// 查询变量值是否改变（注意“标识”变量在所有配置下的值是一样的）
        /// </summary>
        /// <param name="vaultName">库名</param>
        /// <param name="userName">用户名</param>
        /// <param name="password">密码</param>
        /// <param name="fileRelativePath">查询文件相对路径</param>
        /// <param name="variableName">变量名</param>
        /// <param name="value">新值</param>
        /// <returns>查询结果</returns>
        public static int QuerySignVariableChanged(string vaultName, string userName, string password, string fileRelativePath, string variableName, string value)
        {
            IEdmVault5 vault;
            if (Login(vaultName, userName, password, out vault) != LoginResult.OK)
                return (int)QueryVariableChangeResult.LOGIN_FAILED;
            
            string filePath = Path.Combine(vault.RootFolderPath, fileRelativePath);
            IEdmFile5 edmFile;
            IEdmFolder5 edmFolder;
            if (!GetFileFromPath(vault, filePath, out edmFile, out edmFolder))
                return (int)QueryVariableChangeResult.FILE_NOT_FOUND;

            //只查询@配置，即自定义属性，而不是配置特定属性
            string variableValue;
            if (!TryGetVaribleValue(edmFile, "@", variableName, out variableValue))
                return (int)QueryVariableChangeResult.QUERY_VARIABLE_VALUE_FAILED;

            if (variableValue != value)
                return (int)QueryVariableChangeResult.VALUE_CHANGED;

            return (int)QueryVariableChangeResult.VALUE_DOSENT_CHANGED;
        }

        /// <summary>
        /// 更新文件到最新版本
        /// </summary>
        /// <param name="vaultName">库名</param>
        /// <param name="userName">用户名</param>
        /// <param name="password">密码</param>
        /// <param name="filePaths">文件路径</param>
        /// <returns></returns>
        public static int RefreshFiles(string vaultName, string userName, string password, string[] fileRelativePaths)
        {
            IEdmVault5 vault;
            if (Login(vaultName, userName, password, out vault) != LoginResult.OK)
                return (int)RefreshFileResult.LOGIN_FAILED;
            
            foreach (string fileRelativePath in fileRelativePaths)
            {
                string filePath = Path.Combine(vault.RootFolderPath, fileRelativePath);

                IEdmFile5 edmfile;
                IEdmFolder5 edmFolder;
                if (!GetFileFromPath(vault, filePath, out edmfile, out edmFolder))
                    return (int)RefreshFileResult.FILE_NOT_FOUND;
                
                if (!TryRefreshFile(edmfile))
                    return (int)RefreshFileResult.REFRESH_FAILED;
                
                if (!TryGetFileLocalCopy(edmfile))
                    return (int)RefreshFileResult.REFRESH_FAILED;
            }

            return (int)RefreshFileResult.OK;
        }

        /// <summary>
        /// 获取指定库的本地目录
        /// </summary>
        /// <param name="vaultName">库名</param>
        /// <param name="userName">用户名</param>
        /// <param name="password">密码</param>
        /// <param name="rootPath">返回库路径</param>
        /// <returns></returns>
        public static int GetVaultRootPath(string vaultName, string userName, string password, out string rootPath)
        {
            rootPath = null;
            IEdmVault5 vault;
            int ret = (int)Login(vaultName, userName, password, out vault);
            if (ret == (int)LoginResult.OK && vault != null)
                rootPath = vault.RootFolderPath;
            return ret;
        }

        #endregion

        private static LoginResult Login(string vaultName, string userName, string password, out IEdmVault5 vault)
        {
            LoginResult retCode = ResultCodes.LoginResult.OK;

            vault = null;
            try
            {
                vault = new EdmVault5();
                vault.Login(userName, password, vaultName);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                switch (ex.ErrorCode)
                {
                    case (int)EdmResultErrorCodes_e.E_EDM_LOGIN_FAILED:
                        retCode = ResultCodes.LoginResult.LOGIN_FAILED;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_CANT_OPEN_DATABASE:
                        retCode = ResultCodes.LoginResult.CANT_OPEN_DATABASE;
                        break;
                    case -2147221164:   //没有注册类
                        retCode = ResultCodes.LoginResult.PDM_CLIENT_DOESNT_EXIST;
                        break;
                    default:
                        retCode = ResultCodes.LoginResult.UNKNOWN_ERROR;
                        break;
                }
            }
            catch (Exception ex)
            {
                retCode = ResultCodes.LoginResult.UNKNOWN_ERROR;
            }

            return retCode;
        }

        private static bool GetFileFromPath(IEdmVault5 vault, string filePath, out IEdmFile5 file, out IEdmFolder5 parentFolder)
        {
            file = null;
            parentFolder = null;
            
            if (vault == null)
                return false;

            try
            {
                file = vault.GetFileFromPath(filePath, out parentFolder);
                return (file != null && parentFolder != null);
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static bool AddFile2Vault(IEdmVault5 vault, string filePath, out IEdmFile5 edmFile)
        {
            edmFile = null;
            try
            {
                IEdmFolder9 parentFolder = vault.GetFolderFromPath(Path.GetDirectoryName(filePath)) as IEdmFolder9;
                if (parentFolder == null)
                {
                    //文件夹也是新增的
                    string fileFolder = Path.GetDirectoryName(filePath).ToLower();
                    string vaultPath = vault.RootFolderPath.ToLower();
                    fileFolder = fileFolder.Replace(vault.RootFolderPath.ToLower(), "");
                    string[] subFolders = fileFolder.Split('/', '\\');
                    if (!RecursivelyAddFolder(vault, vault.RootFolder, subFolders, out parentFolder))
                        return false;
                }

                int addFileStatus;
                parentFolder.AddFile2(0, filePath, out addFileStatus, "", (int)EdmAddFlag.EdmAdd_Simple);
                IEdmFolder5 folder;
                return GetFileFromPath(vault, filePath, out edmFile, out folder);
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static bool RecursivelyAddFolder(IEdmVault5 vault, IEdmFolder5 rootFolder, string[] subFolders, out IEdmFolder9 outFolder)
        {
            outFolder = null;

            if (vault == null || rootFolder == null)
                return false;

            if (subFolders.Length == 0)
                return true;

            IEdmFolder5 tempFolder = rootFolder;

            try
            {
                for (int i = 0; i < subFolders.Length; ++i)
                {
                    IEdmFolder5 subFolder;
                    string folderPath = Path.Combine(tempFolder.LocalPath, subFolders[i]);
                    subFolder = vault.GetFolderFromPath(folderPath);
                    if (subFolder == null)
                    {
                        subFolder = tempFolder.AddFolder(0, subFolders[i]);
                        if (subFolder == null)
                            return false;
                    }
                    tempFolder = subFolder;
                }

                outFolder = tempFolder as IEdmFolder9;

                return outFolder != null;
            }
            catch(Exception ex)
            {
                return false;
            }
        }

        private static bool GetReferencedFiles(IEdmVault5 vault, IEdmReference10 reference, string filePath, string projectName, ref List<string> refFiles)
        {
            try
            {
                bool bTop = false;
                if (reference == null)
                {
                    bTop = true;
                    SWFileType type = GetSWFileType(filePath);
                    if (type != SWFileType.UNKNOWN)
                    {
                        refFiles.Add(filePath);
                        IEdmFile5 edmFile;
                        IEdmFolder5 edmFolder = null;
                        if (!GetFileFromPath(vault, filePath, out edmFile, out edmFolder))
                            return false;
                        reference = (IEdmReference10)edmFile.GetReferenceTree(edmFolder.ID);
                        if (type == SWFileType.ASSEMBLY)    //装配体递归
                            GetReferencedFiles(vault, reference, "", projectName, ref refFiles);
                    }
                }
                else
                {
                    IEdmPos5 pos = default(IEdmPos5);
                    pos = reference.GetFirstChildPosition3(projectName, bTop, true, (int)EdmRefFlags.EdmRef_File, "", 0);
                    while ((!pos.IsNull))
                    {
                        IEdmReference10  @ref = (IEdmReference10)reference.GetNextChild(pos);
                        SWFileType type = GetSWFileType(@ref.FoundPath);
                        if (type != SWFileType.UNKNOWN)
                        {
                            if (!refFiles.Contains(@ref.FoundPath))
                            {
                                refFiles.Add(@ref.FoundPath);
                                if (type == SWFileType.ASSEMBLY)    //装配体递归
                                    GetReferencedFiles(vault, @ref, "", projectName, ref refFiles);
                            }
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static SWFileType GetSWFileType(string filePath)
        {
            SWFileType ret = SWFileType.UNKNOWN;
            string ext = Path.GetExtension(filePath).ToLower();
            switch (ext)
            {
                case ".sldprt":
                    ret = SWFileType.PART;
                    break;
                case ".sldasm":
                    ret = SWFileType.ASSEMBLY;
                    break;
                case ".slddrw":
                    ret = SWFileType.DRAWING;
                    break;
                default:
                    break;
            }

            return ret;
        }

        #region 文件检出

        private static FileCheckOutResult LockFile(IEdmVault5 vault, string user, string fileFullPath)
        {
            IEdmFile5 file;
            IEdmFolder5 parentFolder;
            if (!GetFileFromPath(vault, fileFullPath, out file, out parentFolder))
                return FileCheckOutResult.FILE_NOT_FOUND;

            //已经检出
            if (file.IsLocked)
            {
                if (file.LockedByUser.Name.ToLower() == user.ToLower())
                    return FileCheckOutResult.OK;
                else
                    return FileCheckOutResult.FILE_IS_LOCKED_BY_OTHER_USER;
            }

            int? edmResultCode;
            if (TryLockFile(file, parentFolder, out edmResultCode))
                return FileCheckOutResult.OK;
            else
            {
                FileCheckOutResult retCode = FileCheckOutResult.UNKNOWN_ERROR;
                switch (edmResultCode)
                {
                    case (int)EdmResultErrorCodes_e.E_EDM_PERMISSION_DENIED:
                        retCode = ResultCodes.FileCheckOutResult.PERMISSION_DENIED;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_OPERATION_REFUSED_BY_PLUGIN:
                        retCode = ResultCodes.FileCheckOutResult.OPERATION_REFUSED_BY_PLUGIN;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_FILE_NOT_FOUND:
                        retCode = ResultCodes.FileCheckOutResult.FILE_NOT_FOUND;
                        break;
                    default:
                        retCode = ResultCodes.FileCheckOutResult.UNKNOWN_ERROR;
                        break;
                }
                return retCode;
            }
        }

        private static bool TryLockFile(IEdmFile5 file, IEdmFolder5 parentFolder, out int? edmResultCode)
        {
            try
            {
                file.LockFile(parentFolder.ID, 0, (int)EdmLockFlag.EdmLock_Simple);
                edmResultCode = null;
                return true;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                edmResultCode = ex.ErrorCode;
                return false;
            }
        }

        //Use IEdmBatchGet to check out several files, it is more efficient repeatedly calling LockFile.
        private static FileCheckOutResult LockFiles(IEdmVault5 vault, string[] filePaths, bool bLockReferenceFiles)
        {
            try
            {
                IEdmVault7 vault7 = vault as IEdmVault7;
                IEdmBatchGet batchGetter = vault7.CreateUtility(EdmUtility.EdmUtil_BatchGet);
                for (int i = 0; i < filePaths.Length; ++i)
                {
                    IEdmFile5 edmFile;
                    IEdmFolder5 edmFolder;
                    if (!GetFileFromPath(vault, filePaths[i], out edmFile, out edmFolder))
                        return FileCheckOutResult.UNKNOWN_ERROR;

                    batchGetter.AddSelectionEx((EdmVault5)vault, edmFile.ID, edmFolder.ID, null);
                }

                int cmdFlag = bLockReferenceFiles ? (int)(EdmGetCmdFlags.Egcf_LockReferencedFilesToo | EdmGetCmdFlags.Egcf_Lock) : (int)EdmGetCmdFlags.Egcf_Lock;
                batchGetter.CreateTree(0, cmdFlag);
                batchGetter.GetFiles(0);

                return FileCheckOutResult.OK;
            }
            catch (Exception ex)
            {
                return FileCheckOutResult.UNKNOWN_ERROR;
            }
        }

        #endregion 文件检出方法

        #region 文件检入

        private static FileCheckInResult UnlockFile(IEdmVault5 vault, string fileFullPath)
        {
            if (!File.Exists(fileFullPath))
                return FileCheckInResult.LOCAL_FILE_NOT_FOUND;
            
            IEdmFile5 file;
            IEdmFolder5 parentFolder;
            if (!GetFileFromPath(vault, fileFullPath, out file, out parentFolder))
            {
                //本地文件存在库文件不存在表示是新文件
                if (!AddFile2Vault(vault, fileFullPath, out file))
                    return FileCheckInResult.UNKNOWN_ERROR;
            }
            
            //如果是装配体，引用零件也要检查是否是新增的
            List<string> referenceFiles = new List<string>();
            if (!GetReferencedFiles(vault, null, fileFullPath, "", ref referenceFiles))
                return FileCheckInResult.UNKNOWN_ERROR;
            for (int i = 1; i < referenceFiles.Count; ++i)
            {
                IEdmFile5 refFile;
                IEdmFolder5 refFileFolder;
                if (!GetFileFromPath(vault, referenceFiles[i], out refFile, out refFileFolder))
                {
                    //本地文件存在库文件不存在表示是新文件
                    if (!AddFile2Vault(vault, referenceFiles[i], out refFile))
                        return FileCheckInResult.UNKNOWN_ERROR;
                }
            }

            if (!file.IsLocked)
                return FileCheckInResult.NOT_BEEN_LOCKED;
            
            int? edmResultCode;
            if (TryUnlockFile(file, out edmResultCode))
                return FileCheckInResult.OK;
            else
            {
                FileCheckInResult retCode = FileCheckInResult.UNKNOWN_ERROR;
                switch (edmResultCode)
                {
                    case (int)EdmResultErrorCodes_e.E_EDM_FILE_NOT_LOCKED_BY_YOU:
                        retCode = ResultCodes.FileCheckInResult.FILE_NOT_LOCKED_BY_YOU;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_LOCKED_ON_OTHER_COMPUTER:
                        retCode = ResultCodes.FileCheckInResult.LOCKED_ON_OTHER_COMPUTER;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_FILE_NOT_FOUND:
                        retCode = ResultCodes.FileCheckInResult.FILE_NOT_FOUND;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_LOCAL_FILE_NOT_FOUND:
                        retCode = ResultCodes.FileCheckInResult.LOCAL_FILE_NOT_FOUND;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_FILE_SHARE_ERROR:
                        retCode = ResultCodes.FileCheckInResult.FILE_SHARE_ERROR;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_CANCELLED_BY_USER:
                        retCode = ResultCodes.FileCheckInResult.CANCELLED_BY_USER;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_INVALID_FILE:
                        retCode = ResultCodes.FileCheckInResult.INVALID_FILE;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_MISSING_MANDATORY_VALUE:
                        retCode = ResultCodes.FileCheckInResult.MISSING_MANDATORY_VALUE;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_OPERATION_REFUSED_BY_PLUGIN:
                        retCode = ResultCodes.FileCheckInResult.OPERATION_REFUSED_BY_PLUGIN;
                        break;
                    case (int)EdmResultErrorCodes_e.E_EDM_VALUE_NOT_UNIQUE:
                        retCode = ResultCodes.FileCheckInResult.VALUE_NOT_UNIQUE;
                        break;
                    default:
                        retCode = ResultCodes.FileCheckInResult.UNKNOWN_ERROR;
                        break;
                }
                return retCode;
            }
        }

        private static bool TryUnlockFile(IEdmFile5 file, out int? edmResultCode)
        {
            try
            {
                file.UnlockFile(0, "");
                edmResultCode = null;
                return true;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                edmResultCode = ex.ErrorCode;
                return false;
            }
        }

        //more efficient than using TryUnlockFile for checking in multiple files.
        private static FileCheckInResult UnlockFiles(IEdmVault5 vault, string[] filePaths)
        {
            try
            {
                IEdmVault7 vault7 = vault as IEdmVault7;
                IEdmBatchUnlock batchUnlocker = (IEdmBatchUnlock)vault7.CreateUtility(EdmUtility.EdmUtil_BatchUnlock);

                EdmSelItem[] selItems = new EdmSelItem[filePaths.Length];
                for (int i = 0; i < filePaths.Length; ++i)
                {
                    IEdmFile5 edmFile;
                    IEdmFolder5 edmFolder;
                    if (!GetFileFromPath(vault, filePaths[i], out edmFile, out edmFolder))
                        return FileCheckInResult.UNKNOWN_ERROR;
                    
                    selItems[i].mlDocID = edmFile.ID;
                    selItems[i].mlProjID = edmFolder.ID;
                }

                batchUnlocker.AddSelection(vault as EdmVault5, ref selItems); //这里必然有异常o(╥﹏╥)o
                batchUnlocker.CreateTree(0, (int)EdmUnlockBuildTreeFlags.Eubtf_MayUnlock);
                batchUnlocker.UnlockFiles(0);

                return FileCheckInResult.OK;
            }
            catch (Exception ex)
            {
                return FileCheckInResult.UNKNOWN_ERROR;
            }
        }

        #endregion 文件检入方法
        
        private static bool TryGetConfigurations(IEdmFile5 file, out List<string> cfgs)
        {
            cfgs = new List<string>();

            try
            {
                EdmStrLst5 cfgList = file.GetConfigurations();
                IEdmPos5 pos = cfgList.GetHeadPosition();
                while (!pos.IsNull)
                {
                    string cfgName = cfgList.GetNext(pos);
                    cfgs.Add(cfgName);
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static bool TryGetVaribleValue(IEdmFile5 file, string cfgName, string variableName, out string varibleValue)
        {
            varibleValue = null;

            try
            {
                IEdmEnumeratorVariable5 vars = file.GetEnumeratorVariable();
                object retValue;
                if (!vars.GetVar(variableName, cfgName, out retValue))
                    return false;
                varibleValue = retValue.ToString();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static bool TryRefreshFile(IEdmFile5 file)
        {
            try
            {
                file.Refresh();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static bool TryGetFileLocalCopy(IEdmFile5 file)
        {
            try
            {
                object version = file.CurrentVersion;
                Log.WriteLog(version.ToString());
                file.GetFileCopy(0, ref version);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}