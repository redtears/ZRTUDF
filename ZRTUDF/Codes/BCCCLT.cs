﻿using System;
//using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
using BCHANDLE = System.IntPtr;

namespace ZRTUDF.Codes
{
    class BCCCLT
    {
        private static string sDrtpIP_Main = ZRTUDF.sDrtpIP_Main;
        private static int iDrtpPort_Main = ZRTUDF.iDrtpPort_Main;
        private static int iDrtpNode_Main = ZRTUDF.iDrtpNode_Main;
        private static int iMainFuncNo_Main = ZRTUDF.iMainFuncNo_Main;

        private static string sDrtpIP_Vip = ZRTUDF.sDrtpIP_Vip;
        private static int iDrtpPort_Vip = ZRTUDF.iDrtpPort_Vip;
        private static int iDrtpNode_Vip = ZRTUDF.iDrtpNode_Vip;
        private static int iMainFuncNo_Vip = ZRTUDF.iMainFuncNo_Vip;

        private static string sDrtpIP_Test = ZRTUDF.sDrtpIP_Test;
        private static int iDrtpPort_Test = ZRTUDF.iDrtpPort_Test;
        private static int iDrtpNode_Test = ZRTUDF.iDrtpNode_Test;
        private static int iMainFuncNo_Test = ZRTUDF.iMainFuncNo_Test;
        
        private const int iTimeout = 6000;  //单位毫秒

        // 描述  : 初试化本系统将直接连接多少个业务通讯平台节点个数
        //		Notice: 系统只能调用一次
        // 返回  : bool 
        // 参数  : int maxnodes=1 [IN]: 设置最多会有多少个连接节点
        [DllImport("bccclt.dll")]
        public static extern bool BCCCLTInit(int maxnodes);

        // 描述  : 增加一个要直接连接的业务通讯平台节点
        //       作为系统初试化部分，建议在调用BCCCLTInit后，在统一一个线程里一次性的设置完毕，供整个环境使用
        // 返回  : int  增加的业务通讯平台节点编号 >=0:成功；否则，可能没有正确调用BCCCLTInit()
        // 参数  : char * ip [IN]: 该业务通讯平台的IP地址
        // 参数  : int port [IN]: 该业务通讯平台提供的端口号port
        [DllImport("bccclt.dll")]
        public static extern int BCAddDrtpNode(string ip, int port);


        // 描述  : 构件一个用于XPack数据交换的数据区
        // 返回  : BCHANDLE : 返回该XPack数据区的句柄；==0(NULL): 失败
        // 参数  : const char * XpackDescribleFile [IN]: XPack结构定义文件名，不能超过1024字节长，如: "/xpacks/cpack.dat"
        [DllImport("bccclt.dll")]
        public static extern BCHANDLE BCNewHandle(string XpackDescribleFile);

        // 描述  : 将已经用BCNewHandle得到的数据区释放(卸构)
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        [DllImport("bccclt.dll")]
        public static extern bool BCDeleteHandle(BCHANDLE Handle);

        // 描述  : 重置数据区内容，主要用于准备首包请求时，填写具体请求数据字段之前调用
        // 返回  : bool : false - 往往为错误的Handle(==NULL的时候)
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        [DllImport("bccclt.dll")]
        public static extern bool BCResetHandle(BCHANDLE Handle);

        // 描述  : 将SourceHandle中的状态、格式、数据等信息复制到DestHandle中
        // 返回  : bool : false - 往往为错误的Handle(==NULL的时候)
        // 参数  : BCHANDLE SourceHandle [IN]: 被复制的原XPack数据区句柄
        // 参数  : BCHANDLE DestHandle   [OUT]: 被复制的目标XPack数据区句柄，内部数据被覆盖，但不会新派生一个句柄
        [DllImport("bccclt.dll")]
        public static extern bool BCCopyHandle(BCHANDLE SourceHandle, BCHANDLE DestHandle);
        //--------------------------------------------------------------------------------------------------
        //////////////////////////////////////////////////////////////////////////
        // 在通过BCNewHandle得到句柄后，就可以设置请求字段数据了：

        // 描述  : 设置指定数据区中第Row条记录中的对应字段的值
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Row [IN]: 本包中指定的记录号 [0..RecordCount-1]
        // 参数  : char * FieldName [IN]: 字段名称 如"vsmess"
        // 参数  : int Value [IN]: 数据值
        [DllImport("bccclt.dll")]
        public static extern bool BCSetIntFieldByName(BCHANDLE Handle, int Row, string FieldName, int Value);

        // 描述  : 设置指定数据区中第Row条记录中的对应字段的值
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Row [IN]: 本包中指定的记录号 [0..RecordCount-1]
        // 参数  : char * FieldName [IN]: 字段名称 如"vsmess"
        // 参数  : double Value [IN]: 数据值
        [DllImport("bccclt.dll")]
        public static extern bool BCSetDoubleFieldByName(BCHANDLE Handle, int Row, string FieldName, double Value);

        // 描述  : 设置指定数据区中第Row条记录中的对应字段的值
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Row [IN]: 本包中指定的记录号 [0..RecordCount-1]
        // 参数  : char * FieldName [IN]: 字段名称 如"vsmess"
        // 参数  : char * Value [IN]: 数据值，不可以超过255长度
        [DllImport("bccclt.dll")]
        public static extern bool BCSetStringFieldByName(BCHANDLE Handle, int Row, string FieldName, string Value);


        // 描述  : 设置指定记录为指定的RawData记录
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Row [IN]: 指定设置的记录号 [0..RecordCount-1]
        // 参数  : char * RawData [IN]: 要设置数据块
        // 参数  : int RawDataLength [IN]: 要设置数据块长度
        [DllImport("bccclt.dll")]
        public static extern bool BCSetRawRecord(BCHANDLE Handle, int Row, string RawData, int RawDataLength);


        // 描述  : 设置请求数据中的功能号 （在EncodeXpackForRequest之前一定要设置）
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int RequestType [IN]: 对应的业务功能号
        [DllImport("bccclt.dll")]
        public static extern bool BCSetRequestType(BCHANDLE Handle, int RequestType);


        // 描述  : 对于BCC类的服务器，客户端可以设置应答包每次返回的最多记录数
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int MaxRetCount [IN]: 指定的最多记录数，0则由应用服务器端 [0..100];
        [DllImport("bccclt.dll")]
        public static extern bool BCSetMaxRetCount(BCHANDLE Handle, int MaxRetCount);


        // 2014/1/22 16:44:28 新增设置原始请求方的Mac地址或IP地址（raw数据）
        // Function name: BCSetAddress
        // Author: CHENYH @ 2014/7/9 10:55:09
        // Description: 对请求包中设置指定的机器地址，以供服务器核查问题
        // Return type: bool 
        // Argument   : BCHANDLE Handle
        // Argument   : void *pAddress [IN]: 即最长六个字节，即Mac地址或IP地址
        [DllImport("bccclt.dll")]
        public static extern bool BCSetAddress(BCHANDLE Handle, IntPtr pAddress);

        //////////////////////////////////////////////////////////////////////////
        // 在设置了请求数据和请求功能号，接下来就可以用下列函数向服务器发请求了：

        // 描述  : 将前面设置好XPack请求数据打包发送给指定的通讯平台和应用服务器
        // 返回  : bool : true 成功发送，并且接收解码了应答数据；false 操作失败，具体错误，参见错误码
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int drtpno [IN]: 从那个直联的业务通讯平台节点发送和接收数据，参见BCAddDrtpNode()的返回值
        // 参数  : int branchno [IN]: 目标应用服务器连接的业务通讯平台节点编号（参见业务通讯平台的配置参数表）
        // 参数  : int function [IN]: 目标应用服务器注册在业务通讯平台的功能号 (参见BCC配置文件中的BASEFUNCNO)
        // 参数  : int timeout [IN]: 发送接收数据的超时时间，以毫秒计
        // 参数  : int* errcode [OUT]: 返回的错误码
        // 参数  : char * errormessage [OUT]: 返回的错误信息
        [DllImport("bccclt.dll")]
        public static extern bool BCCallRequest(BCHANDLE Handle, int drtpno, int branchno, int function, int timeout, ref int errcode, StringBuilder errormessage);


        // Function name: bool BCCallRequestEx
        // Author: CHENYH @ 2012/8/29 10:50:57
        // 描述  : 将前面设置好XPack请求数据打包发送给指定的通讯平台和应用服务器
        // 返回  : bool : true 成功发送，并且接收解码了应答数据；false 操作失败，具体错误，参见错误码
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int drtpno [IN]: 从那个直联的业务通讯平台节点发送和接收数据，参见BCAddDrtpNode()的返回值
        // 参数  : int branchno [IN]: 目标应用服务器连接的业务通讯平台节点编号（参见业务通讯平台的配置参数表）
        // 参数  : int function [IN]: 目标应用服务器注册在业务通讯平台的功能号 (参见BCC配置文件中的BASEFUNCNO)
        // 参数  : int timeout [IN]: 发送接收数据的超时时间，以毫秒计
        // 参数  : int* errcode [OUT]: 返回的错误码
        // 参数  : char * errormessage [OUT]: 返回的错误信息
        // Argument   : bool bCheckUserData [IN]: 指定是否采用userdata检查，true-则检查，否则false(0)不检查
        [DllImport("bccclt.dll")]
        public static extern bool BCCallRequestEx(BCHANDLE Handle, int drtpno, int branchno, int function, int timeout, ref int errcode, StringBuilder errormessage, bool bCheckUserData);

        // 函数名: BCPush0
        // 编程  : 陈永华 2010-8-16 16:39:34
        // 描述  : 将前面设置好的XPack请求数据打包以PM0（即无确认模式）推送给指定的通讯平台和服务端
        // 返回  : bool : true 推送成功；false 推送失败，具体错误参见errormessage
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int drtpno [IN]: 从那个直联的业务通讯平台节点发送和接收数据，参见BCAddDrtpNode()的返回值
        // 参数  : int branchno [IN]: 目标接收推送的服务端连接的业务通讯平台节点编号（参见业务通讯平台的配置参数表）
        // 参数  : int function [IN]: 目标接收推送的服务端注册在业务通讯平台的功能号 (参见BCC配置文件中的BASEFUNCNO)
        // 参数  : int *errcode [OUT]: 返回的错误码，>=0: 成功，否则(<0)失败
        // 参数  : char * errormessage [OUT]: 返回的错误信息
        [DllImport("bccclt.dll")]
        public static extern bool BCPush0(BCHANDLE Handle, int drtpno, int branchno, int function, ref int errcode, StringBuilder errormessage);


        //////////////////////////////////////////////////////////////////////////
        //在调用了BCCallRequest后，就要检查返回的数据内容，

        // 描述  : 在调用BCCallRequest后读取其返回码
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int* RetCode [OUT]: 返回收到数据包中的返回码
        [DllImport("bccclt.dll")]
        public static extern bool BCGetRetCode(BCHANDLE Handle, ref int RetCode);

        // 描述  : 在调用BCCallRequest/BCCallNext后读取本包中的记录数RecordCount
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int * RecordCount [OUT]: 返回解包后的记录数
        [DllImport("bccclt.dll")]
        public static extern bool BCGetRecordCount(BCHANDLE Handle, ref int RecordCount);

        // 描述  : 根据收得解码后的数据，从中得到属于可以获取后续应答数据包的应用服务器专用通讯功能号
        //          详细技术请参考KSBCC使用说明书
        // 返回  : int : >0 - 可获取后续数据的专用通讯功能号；否则失败
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        [DllImport("bccclt.dll")]
        public static extern int BCGetPrivateFunctionForNext(BCHANDLE Handle);

        // 描述  : 根据收得解码后的数据，从中得到属于可以获取后续应答数据包的目标通讯平台节点编号
        //          详细技术请参考KSMBCC使用说明书
        // 返回  : int >0: 获取后续应答数据包的目标通讯平台节点编号; 否则还是用原来的节点编号
        //                当该功能是用Transfer实现的时候，需要这个可能被变更了的节点号
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        [DllImport("bccclt.dll")]
        public static extern int BCGetBranchNoForNext(BCHANDLE Handle);

        // 编程  : 陈永华 2005-11-14 22:01:15
        // 描述  : 当解码后，通过本函数判断是否还有后续应答包
        // 返回  : bool : true-表示还有后续应答包；false－无后续包
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        [DllImport("bccclt.dll")]
        public static extern bool BCHaveNextPack(BCHANDLE Handle);

        // 描述  : 取下一个后续应答包
        //          在成功调用BCCallRequest后，如果检查到BCHaveNextPack()后，则用本功能获取下一个后续应答包
        // 返回  : bool ：true 成功发送，并且接收解码了应答数据；false 操作失败，具体错误，参见错误码
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int timeout [IN]: 发送接收数据的超时时间，以毫秒计
        // 参数  : int* errcode [OUT]: 返回的错误码
        // 参数  : char * errormessage [OUT]: 返回的错误信息
        [DllImport("bccclt.dll")]
        public static extern bool BCCallNext(BCHANDLE Handle, int timeout, ref int errcode, StringBuilder errormessage);


        // Function name: BCCallNextEx
        // Author: CHENYH @ 2012/8/29 10:53:34
        // 描述  : 取下一个后续应答包
        //          在成功调用BCCallRequest后，如果检查到BCHaveNextPack()后，则用本功能获取下一个后续应答包
        // 返回  : bool ：true 成功发送，并且接收解码了应答数据；false 操作失败，具体错误，参见错误码
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int timeout [IN]: 发送接收数据的超时时间，以毫秒计
        // 参数  : int* errcode [OUT]: 返回的错误码
        // 参数  : char * errormessage [OUT]: 返回的错误信息
        // Argument   : bool bCheckUserData [IN]: 指定是否采用userdata检查，true-则检查，否则false(0)不检查
        [DllImport("bccclt.dll")]
        public static extern bool BCCallNextEx(BCHANDLE Handle, int timeout, ref int errcode, StringBuilder errormessage, bool bCheckUserData);

        // 描述  : 查询当前数据块中数据记录类型
        // 返回  : int 返回数据记录类型 0-标准ST_PACK记录；1-ST_SDPACK的RawData记录
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        [DllImport("bccclt.dll")]
        public static extern int BCGetXPackType(BCHANDLE Handle);

        // 描述  : 当用BCGetXPackType返回为1即ST_SDPACK类的记录，则可以用本函数获取各个有效记录中的RawData记录
        // 返回  : int : >=0 - 成功返回RawData的数据长度(字节数); -1:RawData Size过小; -2: 非ST_SDPACK类记录或记录不存在; -3: 无效句柄Handle
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Row [IN]: 指定的记录号 [0..RecordCount-1]
        // 参数  : char * RawData [OUT]: 存放读取的RawData数据块 
        // 参数  : int RawDataSize [IN]: RawData数据块的可存放的最长长度
        [DllImport("bccclt.dll")]
        public static extern int BCGetRawRecord(BCHANDLE Handle, int Row, string RawData, int RawDataSize);

        // 描述  : 检查应答数据中，是否有指定字段的值
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Row [IN]: 指定的记录号 [0..RecordCount-1]
        // 参数  : char *FieldName [IN]: 字段名称 如"vsmess"
        [DllImport("bccclt.dll")]
        public static extern bool BCIsValidField(BCHANDLE Handle, int Row, string FieldName);

        //////////////////////////////////////////////////////////////////////////
        // 获取应答包中各条记录中的字段值

        // 描述  : 从数据区读取第Row条记录中的对应字段的值 (int)
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Row [IN]: 本包中指定的记录号 [0..RecordCount-1]
        // 参数  : char * FieldName [IN]: 字段名称 如"vsmess"
        // 参数  : int* Value [OUT]: 返回得到的值
        [DllImport("bccclt.dll")]
        public static extern bool BCGetIntFieldByName(BCHANDLE Handle, int Row, string FieldName, ref int Value);

        // 描述  : 从数据区读取第Row条记录中的对应字段的值 (double)
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Row [IN]: 本包中指定的记录号 [0..RecordCount-1]
        // 参数  : char * FieldName [IN]: 字段名称 如"vsmess"
        // 参数  : double* Value [OUT]: 返回得到的值
        [DllImport("bccclt.dll")]
        public static extern bool BCGetDoubleFieldByName(BCHANDLE Handle, int Row, string FieldName, Double Value);

        // 描述  : 从数据区读取第Row条记录中的对应字段的值 (double)
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Row [IN]: 本包中指定的记录号 [0..RecordCount-1]
        // 参数  : char * FieldName [IN]: 字段名称 如"vsmess"
        // 参数  : char * Value [OUT]: 返回得到的值
        // 参数  : int ValueBufferSize [IN]: Value缓存区的长度（即可以存放的最大长度） 一般可以为 <=256
        [DllImport("bccclt.dll")]
        public static extern bool BCGetStringFieldByName(BCHANDLE Handle, int Row, string FieldName, StringBuilder Value, int ValueBufferSize);

        // 函数名: BCGetPCharFieldByName
        // 编程  : 陈永华 2009-5-12 14:33:10
        // 描述  : 和BCGetStringFieldByName不同的在于：当字段类型为unsigned char的情况下 Value为原始值，不是用0x开始的可打印字符串
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Row [IN]: 本包中指定的记录号 [0..RecordCount-1]
        // 参数  : char * FieldName [IN]: 字段名称 如"vsmess"
        // 参数  : char * Value [OUT]: 返回得到的值
        // 参数  : int ValueBufferSize [IN]: Value缓存区的长度（即可以存放的最大长度） 一般可以为 <=256
        [DllImport("bccclt.dll")]
        public static extern bool BCGetPCharFieldByName(BCHANDLE Handle, int Row, string FieldName, StringBuilder Value, int ValueBufferSize);

        //////////////////////////////////////////////////////////////////////////
        // 

        // 描述  : 获取本格式的XPack中最大的有效字段索引值
        // 返回  : int 返回ST_PACK中的最大有效字段索引值
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        [DllImport("bccclt.dll")]
        public static extern int BCGetMaxColumn(BCHANDLE Handle);

        // 描述  : 根据字段名称，获得字段索引号
        // 返回  : int >=0 为本XPACK中有效字段，返回的值即为索引号；-1为本XPACK中没有这个字段; -2: 无效句柄
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : char * FieldName [IN]: 字段名称 如"vsmess"
        [DllImport("bccclt.dll")]
        public static extern int BCGetFieldColumn(BCHANDLE Handle, string FieldName);

        // 描述  : 取得对应字段的类型(与BCNewHandle中输入的结构定义文件相关)
        // 返回  : int  : 0 - 空字段; 1-char; 2-vs_char; 3-unsigned char; 4-int; 5-double 
        //                -1 字段索引不存在或Handle错误
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : char * FieldName [IN]: 指定的字段名称,如"vsmess"
        [DllImport("bccclt.dll")]
        public static extern int BCGetFieldTypeByName(BCHANDLE Handle, string FieldName);

        // 描述  : 读取指定字段的信息
        // 返回  : bool : true-为有效字段
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int Col [IN]: 指定的字段索引号
        // 参数  : char * FieldName [OUT]: 返回指定的字段名称
        // 参数  : int* FieldType [OUT]: 指定的字段类型
        // 参数  : int* FieldLength [OUT]: 指定的字段长度
        [DllImport("bccclt.dll")]
        public static extern bool BCGetFieldInfo(BCHANDLE Handle, int Col, StringBuilder FieldName, ref int FieldType, ref int FieldLength);

        // 描述  : 当需要在一次请求的时候，请求数据需要发送多记录的时候，为了获得一次能够发送多少记录，
        //         需要在第一条记录用SetxxxxFieldByName完整后，才可调用本函数。
        //          注意：所有记录使用的字段必须相同，即就是第一条记录设置的那些字段
        // 返回  : int ：返回最大可设置多少条记录
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        [DllImport("bccclt.dll")]
        public static extern int BCEmulateMaxRows(BCHANDLE Handle);

        // 描述  : 设置是否打开调试跟踪文件
        // 返回  : void 
        // 参数  : bool bDebug [IN]: 0(false) - off; 1(true) - on
        [DllImport("bccclt.dll")]
        public static extern void BCSetDebugSwitch(bool bDebug);


        // 描述  : 将信息记录到KLG文件中（只有用BCSetDebugSwitch(true) 后有效）
        // 返回  : bool  当BCSetDebugSwitch(true)是否写入成功
        // 参数  : char *szMsg : 需要记录的文本信息串
        [DllImport("bccclt.dll")]
        public static extern bool BCKLGMsg(string szMsg);

        // 描述  : 将前面设置好XPack请求数据向所有符合条件的服务端(destno:funcno)广播
        //          注意:由于接收端可能许多，这种模式如同BCC架构中的推送消息0模式处理
        //                接收端不要发送应答或确认消息；本端也不要接收
        //                本功能只能在DRTP4版本上实现；DRTP3一概返回false
        // 返回  : bool 
        // 参数  : BCHANDLE Handle [IN]: 指定的XPack数据区句柄
        // 参数  : int drtpno [IN]: 从那个直联的业务通讯平台节点广播数据，参见BCAddDrtpNode()的返回值
        // 参数  : int branchno [IN]: 目标应用服务连接的业务通讯平台节点编号 >=0;  =0: 则全网广播,否则就在指定节点上的广播
        // 参数  : int function [IN]: 目标应用服务注册在业务通讯平台的用于接收广播的功能号（常常为通用功能号或专门接收广播的功能号）
        // 参数  : int* errcode [OUT]: 返回的错误码
        // 参数  : char *errmsg [OUT]: 需要512字节空间，返回错误信息
        [DllImport("bccclt.dll")]
        public static extern bool BCBroad(BCHANDLE Handle, int drtpno, int destno, int funcno, ref int errcode, StringBuilder errmsg);
        

        /// <summary>
        /// 调用交易编码的统一入口
        /// </summary>
        /// <param name="hHandle">句柄</param>
        /// <param name="iFuncNo">主功能号</param>
        /// <param name="arlFields">参数名集合</param>
        /// <param name="arlParams">参数值集合</param>
        /// <param name="sTarget">目标节点  MAIN，VIP，TEST</param>
        /// <returns>大于0则执行成功</returns>

        public static int ExecuteCommand(BCHANDLE hHandle, int iFuncNo, ArrayList arlFields, ArrayList arlParams, string sTarget)
        {
            StringBuilder sbMsg = new StringBuilder(1024);
            StringBuilder sbValue = new StringBuilder(1024);           
            int iDrtp = 0;           
            int iErrCode = 0;
            string sDrtpIP = "";
            int iDrtpPort = 0;
            int iDrtpNode = 0;
            int iMainFuncNo = 0;

            BCSetDebugSwitch(true);


            switch (sTarget.ToUpper())
            {
                case "MAIN":
                    sDrtpIP = sDrtpIP_Main;
                    iDrtpPort = iDrtpPort_Main;
                    iDrtpNode = iDrtpNode_Main;
                    iMainFuncNo = iMainFuncNo_Main;
                    break;
                case "VIP":
                    sDrtpIP = sDrtpIP_Main;
                    iDrtpPort = iDrtpPort_Main;
                    iDrtpNode = iDrtpNode_Main;
                    iMainFuncNo = iMainFuncNo_Vip;
                    break;
                case "TEST":
                    sDrtpIP = sDrtpIP_Test;
                    iDrtpPort = iDrtpPort_Test;
                    iDrtpNode = iDrtpNode_Test;
                    iMainFuncNo = iMainFuncNo_Test;
                    break;
            }
      
            try
            {
                if (hHandle.Equals(IntPtr.Zero))
                {
                    return -1;
                }

                if (!BCCCLT.BCSetRequestType(hHandle, iFuncNo))
                {
                    return -2;
                }

                iDrtp = BCCCLT.BCAddDrtpNode(sDrtpIP, iDrtpPort);
                
                for (int i = 0; i < arlFields.Count; i++)
                {
                    if (!BCSetStringFieldByName(hHandle, 0, arlFields[i].ToString(), arlParams[i].ToString()))
                    {
                        return -3;
                    }
                }

                if (!BCCallRequest(hHandle, iDrtp, iDrtpNode, iMainFuncNo, iTimeout, ref iErrCode, sbMsg))
                {
                    return -4;
                }

                return 1;

            }
            catch (Exception ex)
            {
                BCKLGMsg("执行请求" + iFuncNo.ToString() + "失败!");
                BCKLGMsg(ex.ToString());
                return -99;
            }
        }
    }
}
