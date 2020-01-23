using System;
//using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using ZRTUDF.Codes;
using System.Collections;
using System.Data;
using System.Configuration;

namespace ZRTUDF
{
    [Guid("04A285B3-EE0B-474F-800D-B39AE231198F")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]

    public class ZRTUDF
    {
        
        private const int iTimeout = 6000;  //单位毫秒

        public static string sEmpCode;
        public static string sBranchNo;
        public static int iMaxNodes;
        public static string sDrtpIP_Main;
        public static int iDrtpPort_Main;
        public static int iDrtpNode_Main;
        public static int iMainFuncNo_Main;
        public static string sDrtpIP_Vip;
        public static int iDrtpPort_Vip;
        public static int iDrtpNode_Vip;
        public static int iMainFuncNo_Vip;
        public static string sDrtpIP_Test;
        public static int iDrtpPort_Test;
        public static int iDrtpNode_Test;
        public static int iMainFuncNo_Test;


        public ZRTUDF()
        {
            InitParams();

            if (!BCCCLT.BCCCLTInit(iMaxNodes))
            {
                throw new Exception();
            }
        }

        /// <summary>
        /// 初始化参数
        /// </summary>
        private void InitParams()
        {
            Config cfg = new Config();
            sEmpCode = cfg.sEmpCode;
            sBranchNo = cfg.sBranchNo;
            iMaxNodes = cfg.iMaxNodes;
            sDrtpIP_Main = cfg.sDrtpIP_Main;
            iDrtpPort_Main = cfg.iDrtpPort_Main;
            iDrtpNode_Main = cfg.iDrtpNode_Main;
            iMainFuncNo_Main = cfg.iMainFuncNo_Main;

            sDrtpIP_Vip = cfg.sDrtpIP_Vip;
            iDrtpPort_Vip = cfg.iDrtpPort_Vip;
            iDrtpNode_Vip = cfg.iDrtpNode_Vip;
            iMainFuncNo_Vip = cfg.iMainFuncNo_Vip;
        }
        
        /// <summary>
        /// 获取非交易委托数量（已报状态）
        /// </summary>
        /// <param name="cust_no">客户号</param>
        /// <param name="date">日期</param>
        /// <param name="market_code">市场代码</param>
        /// <param name="sec_code">证券代码</param>
        /// <param name="target">目标： MAIN, VIP, TEST</param>
        /// <returns></returns>
        public string GetNontradeStockPaybackVol(string cust_no, string date, string market_code, string sec_code, string target)
        {
            string sResult = "";
            string sTarget = "";

            if (target.Trim() == "快订")
            {
                sTarget = "VIP";
            }
            else if(target.Trim() == "TEST")
            {
                sTarget = "TEST";
            }
            else
            {
                sTarget = "MAIN";
            }

            switch (market_code)
            {
                case "1":
                    sResult = GetNontradeStockPaybackVol_sh(cust_no, date, sec_code, sTarget);
                    break;
                case "2":
                    sResult = GetNontradeStockPaybackVol_sz(cust_no, date, sec_code, sTarget);
                    break;
            }

            if (sResult == "")
            {
                return "0";
            }

            return sResult;
        }

        /// <summary>
        /// 取客户上海市场直接还券累计数量
        /// </summary>
        /// <param name="cust_no">客户号</param>
        /// <param name="date">交易日</param>
        /// <param name="sec_code">证券代码</param>
        /// <returns></returns>
        public string GetNontradeStockPaybackVol_sh(string cust_no, string date, string sec_code, string target)
        {
            
            ArrayList arlFields = new ArrayList();
            ArrayList arlParams = new ArrayList();            
            string scust_no = cust_no;
            string sdate0 = date;
            string sdate1 = date;
            StringBuilder damt0 = new StringBuilder();
            damt0.Append(sec_code.Substring(0, 3));
            damt0.Append(".");
            damt0.Append(sec_code.Substring(3));
            string susset6 = "02,";
            const string sstock_code = "799982";
            const string sstatus1 = "2"; //已报
            StringBuilder sbPaybackVol = new StringBuilder();
            StringBuilder sbErrMsg = new StringBuilder();
            DataTable dtResult = new DataTable();
            dtResult.Columns.Add("payback_vol");
            dtResult.Columns["payback_vol"].DataType = System.Type.GetType("System.Int32");
            int iErrCode = 0;
            int iVol0 = 0;
            int iRecordCount = 0;

            IntPtr handle = BCCCLT.BCNewHandle("C:\\ZRTUDF\\cpack.dat");  

            arlFields.Clear();
            arlParams.Clear();
            arlFields.Add("semp"); 
            arlParams.Add(sEmpCode);
            arlFields.Add("sbranch_code0");
            arlParams.Add(sBranchNo);
            arlFields.Add("scust_no");
            arlParams.Add(scust_no);
            arlFields.Add("sdate0");
            arlParams.Add(sdate0);
            arlFields.Add("sdate1");
            arlParams.Add(sdate1);
            arlFields.Add("damt0");
            arlParams.Add(damt0.ToString());
            arlFields.Add("damt1");
            arlParams.Add(damt0.ToString());
            arlFields.Add("sstock_code");
            arlParams.Add(sstock_code);
            arlFields.Add("sstatus1");
            arlParams.Add(sstatus1);
            arlFields.Add("usset6");
            arlParams.Add(susset6);

            if (BCCCLT.ExecuteCommand(handle, 498051, arlFields, arlParams, target) > 0)
            {
                try
                {
                    do
                    {   
                        if (BCCCLT.BCGetRetCode(handle, ref iErrCode))
                        {
                            if(iErrCode == 0)
                            {
                                if(!BCCCLT.BCGetRecordCount(handle, ref iRecordCount))
                                {
                                    return "获取记录条数失败！";
                                }

                                for (int i=0; i<iRecordCount; i++)
                                {
                                    BCCCLT.BCGetIntFieldByName(handle, i, "lvol0", ref iVol0);
                                    DataRow dr = dtResult.NewRow();
                                    dr[0] = iVol0;
                                    dtResult.Rows.Add(dr);
                                }
                                //BCCCLT.BCGetStringFieldByName(handle, 0, "lvol0", sbPaybackVol, 160);
                            }
                            
                        }

                        if (!BCCCLT.BCHaveNextPack(handle))
                        {
                            break;
                        }

                    } while (BCCCLT.BCCallNext(handle, iTimeout, ref iErrCode, sbErrMsg));
                }
                catch (Exception ex)
                {   
                    throw ex;
                }
            }

            BCCCLT.BCDeleteHandle(handle);

            return dtResult.Compute("Sum(payback_vol)", "").ToString();
        }

        /// <summary>
        /// 获取客户深市直接还券数量
        /// </summary>
        /// <param name="cust_no">客户号</param>
        /// <param name="date">日期</param>
        /// <param name="sec_code">证券代码</param>
        /// <param name="target">目标：MAIN, VIP, TEST</param>
        /// <returns></returns>
        public string GetNontradeStockPaybackVol_sz(string cust_no, string date, string sec_code, string target)
        {

            ArrayList arlFields = new ArrayList();
            ArrayList arlParams = new ArrayList();            
            string scust_no = cust_no;
            string sdate0 = date;
            string sdate1 = date;
            string susset6 = "0e,";
            string sstock_code = sec_code;
            const string sstatus1 = "2"; //已报
            StringBuilder sbPaybackVol = new StringBuilder();
            StringBuilder sbErrMsg = new StringBuilder();
            DataTable dtResult = new DataTable();
            dtResult.Columns.Add("payback_vol");
            dtResult.Columns["payback_vol"].DataType = System.Type.GetType("System.Int32");
            int iErrCode = 0;
            int iVol0 = 0;
            int iRecordCount = 0;

            IntPtr handle = BCCCLT.BCNewHandle("C:\\ZRTUDF\\cpack.dat");

            arlFields.Clear();
            arlParams.Clear();
            arlFields.Add("semp");
            arlParams.Add(sEmpCode);
            arlFields.Add("sbranch_code0");
            arlParams.Add(sBranchNo);
            arlFields.Add("scust_no");
            arlParams.Add(scust_no);
            arlFields.Add("sdate0");
            arlParams.Add(sdate0);
            arlFields.Add("sdate1");
            arlParams.Add(sdate1);
            arlFields.Add("sstock_code");
            arlParams.Add(sstock_code);
            arlFields.Add("sstatus1");
            arlParams.Add(sstatus1);
            arlFields.Add("usset6");
            arlParams.Add(susset6);

            if (BCCCLT.ExecuteCommand(handle, 110226, arlFields, arlParams, target) > 0)
            {
                try
                {
                    do
                    {
                        if (BCCCLT.BCGetRetCode(handle, ref iErrCode))
                        {
                            if (iErrCode == 0)
                            {
                                if (!BCCCLT.BCGetRecordCount(handle, ref iRecordCount))
                                {
                                    return "获取记录条数失败！";
                                }

                                for (int i = 0; i < iRecordCount; i++)
                                {
                                    BCCCLT.BCGetIntFieldByName(handle, i, "lvol0", ref iVol0);
                                    DataRow dr = dtResult.NewRow();
                                    dr[0] = iVol0;
                                    dtResult.Rows.Add(dr);
                                }
                                //BCCCLT.BCGetStringFieldByName(handle, 0, "lvol0", sbPaybackVol, 160);
                            }

                        }

                        if (!BCCCLT.BCHaveNextPack(handle))
                        {
                            break;
                        }
                    } while (BCCCLT.BCCallNext(handle, iTimeout, ref iErrCode, sbErrMsg));
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            BCCCLT.BCDeleteHandle(handle);

            return dtResult.Compute("Sum(payback_vol)", "").ToString();
        }

        /// <summary>
        /// 获取客户买券还券数量
        /// </summary>
        /// <param name="cust_no">客户号</param>
        /// <param name="market_code">市场代码</param>
        /// <param name="sec_code">证券代码</param>
        /// <param name="target">目标：MAIN, VIP, TEST</param>
        /// <returns></returns>
        public string GetBuyPaybackVol(string cust_no, string market_code, string sec_code, string target)
        {
            ArrayList arlFields = new ArrayList();
            ArrayList arlParams = new ArrayList();            
            string scust_no = cust_no;
            const string sstatus4 = "1"; //是否查融券
            const long lvol0 = 1; //汇总方式：1-按明细输出；2-按营业部+证券汇总；3-按客户+证券汇总
            const long lvol2 = 1; //是否输出合计 0-否， 1-是
            const long lvol12 = 2; //开放式基金 0-不包含 1-包含 2-按营业部参数
            StringBuilder sbPaybackVol = new StringBuilder();
            StringBuilder sbErrMsg = new StringBuilder();
            DataTable dtResult = new DataTable();
            dtResult.Columns.Add("payback_vol");
            dtResult.Columns["payback_vol"].DataType = System.Type.GetType("System.Int32");
            int iErrCode = 0;
            int iVol0 = 0;
            int iRecordCount = 0;
            string sTarget = "";

            IntPtr handle = BCCCLT.BCNewHandle("C:\\ZRTUDF\\cpack.dat");

            arlFields.Clear();
            arlParams.Clear();
            arlFields.Add("semp"); //职工代码
            arlParams.Add(sEmpCode);
            arlFields.Add("sbranch_code0"); //营业部代码
            arlParams.Add(sBranchNo);
            arlFields.Add("scust_no"); //客户号
            arlParams.Add(scust_no);
            arlFields.Add("lvol0"); //汇总方式：1-按明细输出；2-按营业部+证券汇总；3-按客户+证券汇总
            arlParams.Add(lvol0.ToString());
            arlFields.Add("sstock_code"); //证券代码
            arlParams.Add(sec_code);
            arlFields.Add("sstatus4"); //查融券
            arlParams.Add(sstatus4);
            arlFields.Add("lvol2"); //输出合计
            arlParams.Add(lvol2.ToString());
            arlFields.Add("lvol12"); //开放式基金 0-不包含 1-包含 2-按营业部参数
            arlParams.Add(lvol12.ToString());
            arlFields.Add("usset5"); //市场代码
            arlParams.Add(market_code + ",");


            if (target.Trim() == "快订")
            {
                sTarget = "VIP";
            }
            else if (target.Trim() == "TEST")
            {
                sTarget = "TEST";
            }
            else
            {
                sTarget = "MAIN";
            }

            if (BCCCLT.ExecuteCommand(handle, 200103, arlFields, arlParams, sTarget) > 0)
            {
                try
                {
                    do
                    {
                        if (BCCCLT.BCGetRetCode(handle, ref iErrCode))
                        {
                            if (iErrCode == 0)
                            {
                                if (!BCCCLT.BCGetRecordCount(handle, ref iRecordCount))
                                {
                                    return "获取记录条数失败！";
                                }

                                for (int i = 0; i < iRecordCount; i++)
                                {
                                    BCCCLT.BCGetIntFieldByName(handle, i, "lvol1", ref iVol0);
                                    DataRow dr = dtResult.NewRow();
                                    dr[0] = iVol0;
                                    dtResult.Rows.Add(dr);
                                }
                                //BCCCLT.BCGetStringFieldByName(handle, 0, "lvol0", sbPaybackVol, 160);
                            }

                        }

                        if (!BCCCLT.BCHaveNextPack(handle))
                        {
                            break;
                        }
                    } while (BCCCLT.BCCallNext(handle, iTimeout, ref iErrCode, sbErrMsg));
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            BCCCLT.BCDeleteHandle(handle);

            if (dtResult.Compute("Sum(payback_vol)", "").ToString() == "")
            {
                return "0";
            }

            return dtResult.Compute("Sum(payback_vol)", "").ToString();
        }

        /// <summary>
        /// 获取客户可直接还券数量，用于敞口监控
        /// </summary>
        /// <param name="cust_no">客户号</param>
        /// <param name="market_code">市场代码</param>
        /// <param name="sec_code">证券代码</param>
        /// <param name="target">目标：MAIN, VIP, TEST</param>
        /// <returns></returns>
        public string GetCanDirectPaybackVol(string cust_no, string market_code, string sec_code, string target)
        {

            ArrayList arlFields = new ArrayList();
            ArrayList arlParams = new ArrayList();
            string scust_no = cust_no;
            const string sstatus4 = "1"; //是否查融券
            const long lvol0 = 1; //汇总方式：1-按明细输出；2-按营业部+证券汇总；3-按客户+证券汇总
            const long lvol2 = 1; //是否输出合计 0-否， 1-是
            const long lvol12 = 2; //开放式基金 0-不包含 1-包含 2-按营业部参数
            StringBuilder sbPaybackVol = new StringBuilder();
            StringBuilder sbErrMsg = new StringBuilder();
            DataTable dtResult = new DataTable();
            dtResult.Columns.Add("payback_vol");
            dtResult.Columns["payback_vol"].DataType = System.Type.GetType("System.Int32");
            int iErrCode = 0;
            int iVol0 = 0;
            int iRecordCount = 0;
            string sTarget = "";

            IntPtr handle = BCCCLT.BCNewHandle("C:\\ZRTUDF\\cpack.dat");

            arlFields.Clear();
            arlParams.Clear();
            arlFields.Add("semp"); //职工代码
            arlParams.Add(sEmpCode);
            arlFields.Add("sbranch_code0"); //营业部代码
            arlParams.Add(sBranchNo);
            arlFields.Add("scust_no"); //客户号
            arlParams.Add(scust_no);
            arlFields.Add("lvol0"); //汇总方式：1-按明细输出；2-按营业部+证券汇总；3-按客户+证券汇总
            arlParams.Add(lvol0.ToString());
            arlFields.Add("sstock_code"); //证券代码
            arlParams.Add(sec_code);
            arlFields.Add("sstatus4"); //查融券
            arlParams.Add(sstatus4);
            arlFields.Add("lvol2"); //输出合计
            arlParams.Add(lvol2.ToString());
            arlFields.Add("lvol12"); //开放式基金 0-不包含 1-包含 2-按营业部参数
            arlParams.Add(lvol12.ToString());
            arlFields.Add("usset5"); //市场代码
            arlParams.Add(market_code + ",");

            if (target.Trim() == "快订")
            {
                sTarget = "VIP";
            }
            else if (target.Trim() == "TEST")
            {
                sTarget = "TEST";
            }
            else
            {
                sTarget = "MAIN";
            }

            if (BCCCLT.ExecuteCommand(handle, 200103, arlFields, arlParams, sTarget) > 0)
            {
                try
                {
                    do
                    {
                        if (BCCCLT.BCGetRetCode(handle, ref iErrCode))
                        {
                            if (iErrCode == 0)
                            {
                                if (!BCCCLT.BCGetRecordCount(handle, ref iRecordCount))
                                {
                                    return "获取记录条数失败！";
                                }

                                for (int i = 0; i < iRecordCount; i++)
                                {
                                    BCCCLT.BCGetIntFieldByName(handle, i, "lserial1", ref iVol0);
                                    DataRow dr = dtResult.NewRow();
                                    dr[0] = iVol0;
                                    dtResult.Rows.Add(dr);
                                }
                                //BCCCLT.BCGetStringFieldByName(handle, 0, "lvol0", sbPaybackVol, 160);
                            }

                        }

                        if (!BCCCLT.BCHaveNextPack(handle))
                        {
                            break;
                        }
                    } while (BCCCLT.BCCallNext(handle, iTimeout, ref iErrCode, sbErrMsg));
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            BCCCLT.BCDeleteHandle(handle);

            if (dtResult.Compute("Sum(payback_vol)", "").ToString() == "")
            {
                return "0";
            }

            return dtResult.Compute("Sum(payback_vol)", "").ToString();
        }

        /// <summary>
        /// 获取客户实时融券余额
        /// </summary>
        /// <param name="cust_no">客户号</param>
        /// <param name="market_code">市场代码</param>
        /// <param name="sec_code">证券代码</param>
        /// <param name="target">目标：MAIN, VIP, TEST</param>
        /// <returns></returns>
        public string GetRealTimeCreditVol(string cust_no, string market_code, string sec_code, string target)
        {

            ArrayList arlFields = new ArrayList();
            ArrayList arlParams = new ArrayList();
            string scust_no = cust_no;
            const string sstatus4 = "1"; //是否查融券
            const long lvol0 = 1; //汇总方式：1-按明细输出；2-按营业部+证券汇总；3-按客户+证券汇总
            const long lvol2 = 1; //是否输出合计 0-否， 1-是
            const long lvol12 = 2; //开放式基金 0-不包含 1-包含 2-按营业部参数
            StringBuilder sbPaybackVol = new StringBuilder();
            StringBuilder sbErrMsg = new StringBuilder();
            DataTable dtResult = new DataTable();
            dtResult.Columns.Add("realtime_credit_vol");
            dtResult.Columns["realtime_credit_vol"].DataType = System.Type.GetType("System.Int32");
            int iErrCode = 0;
            int iVol0 = 0;
            int iRecordCount = 0;
            string sTarget = "";

            IntPtr handle = BCCCLT.BCNewHandle("C:\\ZRTUDF\\cpack.dat");

            arlFields.Clear();
            arlParams.Clear();
            arlFields.Add("semp"); //职工代码
            arlParams.Add(sEmpCode);
            arlFields.Add("sbranch_code0"); //营业部代码
            arlParams.Add(sBranchNo);
            arlFields.Add("scust_no"); //客户号
            arlParams.Add(scust_no);
            arlFields.Add("lvol0"); //汇总方式：1-按明细输出；2-按营业部+证券汇总；3-按客户+证券汇总
            arlParams.Add(lvol0.ToString());
            arlFields.Add("sstock_code"); //证券代码
            arlParams.Add(sec_code);
            arlFields.Add("sstatus4"); //查融券
            arlParams.Add(sstatus4);
            arlFields.Add("lvol2"); //输出合计
            arlParams.Add(lvol2.ToString());
            arlFields.Add("lvol12"); //开放式基金 0-不包含 1-包含 2-按营业部参数
            arlParams.Add(lvol12.ToString());
            arlFields.Add("usset5"); //市场代码
            arlParams.Add(market_code + ",");

            if (target.Trim() == "快订")
            {
                sTarget = "VIP";
            }
            else if (target.Trim() == "TEST")
            {
                sTarget = "TEST";
            }
            else
            {
                sTarget = "MAIN";
            }

            if (BCCCLT.ExecuteCommand(handle, 200103, arlFields, arlParams, sTarget) > 0)
            {
                try
                {
                    do
                    {
                        if (BCCCLT.BCGetRetCode(handle, ref iErrCode))
                        {
                            if (iErrCode == 0)
                            {
                                if (!BCCCLT.BCGetRecordCount(handle, ref iRecordCount))
                                {
                                    return "获取记录条数失败！";
                                }

                                for (int i = 0; i < iRecordCount; i++)
                                {
                                    BCCCLT.BCGetIntFieldByName(handle, i, "lserial0", ref iVol0);
                                    DataRow dr = dtResult.NewRow();
                                    dr[0] = iVol0;
                                    dtResult.Rows.Add(dr);
                                }
                                //BCCCLT.BCGetStringFieldByName(handle, 0, "lvol0", sbPaybackVol, 160);
                            }

                        }

                        if (!BCCCLT.BCHaveNextPack(handle))
                        {
                            break;
                        }
                    } while (BCCCLT.BCCallNext(handle, iTimeout, ref iErrCode, sbErrMsg));
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            BCCCLT.BCDeleteHandle(handle);

            if (dtResult.Compute("Sum(realtime_credit_vol)", "").ToString() == "")
            {
                return "0";
            }

            return dtResult.Compute("Sum(realtime_credit_vol)", "").ToString();
        }

        #region COM Related

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            Registry.ClassesRoot.CreateSubKey(GetSubKeyName(type, "Programmable"));
            var key = Registry.ClassesRoot.OpenSubKey(GetSubKeyName(type, "InprocServer32"), true);
            key.SetValue("", Environment.SystemDirectory + @"\mscoree.dll", RegistryValueKind.String);
            //key.SetValue("", @"C:\Windows\System32\mscoree.dll", RegistryValueKind.String);
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            Registry.ClassesRoot.DeleteSubKey(GetSubKeyName(type, "Programmable"), false);
        }

        private static string GetSubKeyName(Type type, string subKeyName)
        {
            var s = new System.Text.StringBuilder();
            s.Append(@"CLSID\{");
            s.Append(type.GUID.ToString().ToUpper());
            s.Append(@"}\");
            s.Append(subKeyName);
            return s.ToString();
        }
        #endregion
    }
}
