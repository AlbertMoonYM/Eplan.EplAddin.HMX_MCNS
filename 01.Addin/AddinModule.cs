using Eplan.MCNS.Lib.Share_CS;
using Eplan.MCNS.Lib.UI_CS;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Gui;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace Eplan.EplAddin.HMX_MCNS
{
    public class AddinModule : IEplAddIn, IEplAddInShadowCopy
    {
        private String m_strOriginalAssemblyPath;

        public void OnBeforeInit(string strOriginalAssemblyPath)
        {
            m_strOriginalAssemblyPath = strOriginalAssemblyPath; // 원본 어셈블리 경로 저장
        }

        public String GetOriginalAssemblyPath()
        {
            return m_strOriginalAssemblyPath; // 저장된 경로 반환
        }


        public bool OnRegister(ref bool bLoadOnStart)
        {
            bLoadOnStart = false;
            
            var ribbonBar = new RibbonBar();
            var tab = ribbonBar.AddTab(AddinStatic.tabName);
            var cmdGroup = tab.AddCommandGroup(AddinStatic.cmdGroupName);
            cmdGroup.AddCommand(AddinStatic.cmdName, AddinStatic.actionName, new RibbonIcon(CommandIcon.Rectangle_M));

            return true;

        }
        public bool OnUnregister()
        {
            
            //언로드시 리본탭 삭제
            var ribbonBar = new RibbonBar();
            RibbonTab temp = null;

            for (int i = 0; i < ribbonBar.Tabs.Length; i++)
            {
                if (ribbonBar.Tabs[i].Name == AddinStatic.tabName)
                {
                    temp = ribbonBar.Tabs[i];
                }
            }

            temp.Remove();

            return true;
        }

        public bool OnExit()
        {
            return true;
        }

        public bool OnInit()
        {
            
            return true;
        }
        
        public bool OnInitGui()
        {

            // 어셈블리 경로 가져오기

            string orginalDllFilePath = GetOriginalAssemblyPath();
            string configFilePath = Path.Combine(Path.GetDirectoryName(orginalDllFilePath), "config.xml");
            string itemListFilePath = Path.Combine(Path.GetDirectoryName(orginalDllFilePath), "ItemList.xml");

            CS_PathData.ConfigFilePath = configFilePath;
            CS_PathData.ItemListFilePath = itemListFilePath;

            return true;
        }

        
        
    }
}
