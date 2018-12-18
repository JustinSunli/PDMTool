namespace PDMConnection.ResultCodes
{
    public enum LoginResult
    {
        OK = 1,
        LOGIN_FAILED = -101,                //用户名或密码错误
        CANT_OPEN_DATABASE = -102,          //无法打开库
        PDM_CLIENT_DOESNT_EXIST = -103,     //未安装PDM客户端
        UNKNOWN_ERROR = -104
    }

    enum FileCheckOutResult
    {
        OK = 1,
        LOGIN_FAILED = -201,                //登录失败
        FILE_NOT_FOUND = -202,              //文件不存在
        FILE_IS_LOCKED_BY_OTHER_USER = -203,//被其他用户检出
        PERMISSION_DENIED = -204,           //用户无权限检出文件
        OPERATION_REFUSED_BY_PLUGIN = -205, //检出被其他插件阻止
        UNKNOWN_ERROR = -206
    }

    enum FileCheckInResult
    {
        OK = 1,
        NOT_BEEN_LOCKED = -300,             //没有被检出
        LOGIN_FAILED = -301,                //登录失败
        FILE_NOT_FOUND = -302,              //文件不存在
        FILE_NOT_LOCKED_BY_YOU = -303,      //文件由其他用户检出
        LOCKED_ON_OTHER_COMPUTER = -304,    //文件在其他电脑检出
        LOCAL_FILE_NOT_FOUND = -305,        //There is no copy of the file in the cache folder on the client machine. 
        FILE_SHARE_ERROR = -306,            //文件被占用
        CANCELLED_BY_USER = -307,           //用户终止
        INVALID_FILE = -308,                //不正确的文件类型
        MISSING_MANDATORY_VALUE = -309,     //文件缺少必要的卡片值
        OPERATION_REFUSED_BY_PLUGIN = -310, //检入被其他插件阻止
        VALUE_NOT_UNIQUE = -311,            //尝试重复保存相同的特殊值
        UNKNOWN_ERROR = -312,
    }

    enum QueryVariableChangeResult
    {
        VALUE_CHANGED = 1,
        VALUE_DOSENT_CHANGED = 0,
        FILE_NOT_FOUND = -1,                //PDM文件不存在（新文件）
        LOGIN_FAILED = -401,                //登录失败
        QUERY_VARIABLE_VALUE_FAILED = -402  //查询值失败
    }

    enum FileUpdateResult
    {
        OK = 1,
        LOGIN_FAILED = -501,                //登录失败
        FILE_LOCK_ERROR = -502,             //文件检出失败
        FILE_UNLOCK_ERROR = -503,           //文件检入失败
        PDM_FILE_NOT_FOUND = -504,          //PDM文件不存在
        LOCAL_FILE_NOT_FOUND = -505,        //本地文件不存在
        REPLACE_PDM_FILE_ERROR = -506,      //替换PDM文件失败 
        UNKNOWN_ERROR = -507
    }

    enum RefreshFileResult
    {
        OK = 1,
        LOGIN_FAILED = -601,                //登录失败
        FILE_NOT_FOUND = -602,
        REFRESH_FAILED = -603,
    }
}
