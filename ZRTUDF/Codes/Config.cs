using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace ZRTUDF.Codes
{
    /// <summary>
    /// 配置文件类，用于从config文件中读取配置
    /// </summary>
    class Config
    {
        public string sEmpCode { get; private set;}
        public string sBranchNo { get; private set;}
        public int iMaxNodes { get; private set;}
        public string sDrtpIP_Main { get; private set;}
        public int iDrtpPort_Main { get; private set;}
        public int iDrtpNode_Main { get; private set;}
        public int iMainFuncNo_Main { get; private set;}
        public string sDrtpIP_Vip { get; private set;}
        public int iDrtpPort_Vip { get; private set;}
        public int iDrtpNode_Vip { get; private set;}
        public int iMainFuncNo_Vip { get; private set;}
        public string sDrtpIP_Test { get; private set;}
        public int iDrtpPort_Test { get; private set;}
        public int iDrtpNode_Test { get; private set;}
        public int iMainFuncNo_Test { get; private set;}


        private string GetParamters(string param)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(GetType().Assembly.Location);
            AppSettingsSection appSettings = (AppSettingsSection)config.GetSection("appSettings");
            return appSettings.Settings[param].Value;
        }

        private void LoadConfig()
        {
            sEmpCode = GetParamters("emp_code");
            sBranchNo = GetParamters("branch_code");
            iMaxNodes = Convert.ToInt32(GetParamters("maxnodes"));

            sDrtpIP_Main = GetParamters("drtp_main");
            iDrtpPort_Main = Convert.ToInt32(GetParamters("port_main"));
            iDrtpNode_Main = Convert.ToInt32(GetParamters("node_main"));
            iMainFuncNo_Main = Convert.ToInt32(GetParamters("funcno_main"));

            sDrtpIP_Vip = GetParamters("drtp_vip");
            iDrtpPort_Vip = Convert.ToInt32(GetParamters("port_vip"));
            iDrtpNode_Vip = Convert.ToInt32(GetParamters("node_vip"));
            iMainFuncNo_Vip = Convert.ToInt32(GetParamters("funcno_vip"));
        }

        public Config()
        {
            LoadConfig();
        }
    }
}
