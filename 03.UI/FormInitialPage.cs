using DevExpress.XtraEditors;
using Eplan.MCNS.Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Reflection;

namespace Eplan.EplAddin.HMX_MCNS
{

    public partial class FormInitialPage : DevExpress.XtraEditors.XtraForm
    {
        ToolTip tip = new ToolTip();

        

        public FormInitialPage()
        {
            InitializeComponent();
           
            ControlFormFunction();
            SetToolTip();

            //StandAlone 일때
            //StringUnits.strConfigFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "config.xml");// @"C:\Users\kr70009769\Desktop\01.Task\02. API 소스\01. ProtoType\Eplan.EplAddin.HMX_MCNS\bin\Debug\config.xml";
            //StringUnits.strItemListFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ItemList.xml");//@"C:\Users\kr70009769\Desktop\01.Task\02. API 소스\01. ProtoType\Eplan.EplAddin.HMX_MCNS\bin\Debug\ItemList.xml";

            // XML 파일을 로드합니다.
            XDocument configXml = XDocument.Load(StringUnits.strConfigFilePath);

            // 기초 파일 paths 가져오기
            StringUnits.strPrjFolderPath = configXml.Descendants("add").FirstOrDefault(x => (string)x.Attribute("key") == "ProjectSaveFolder")?.Attribute("value")?.Value;
            StringUnits.strBasicTempletFilePath = configXml.Descendants("add").FirstOrDefault(x => (string)x.Attribute("key") == "BasicTempletFilePath")?.Attribute("value")?.Value;
            StringUnits.strMacroFolderPath = configXml.Descendants("add").FirstOrDefault(x => (string)x.Attribute("key") == "MacroFolderPath")?.Attribute("value")?.Value;
            StringUnits.strIoListFilePath = configXml.Descendants("add").FirstOrDefault(x => (string)x.Attribute("key") == "IoListFilePath")?.Attribute("value")?.Value;
            StringUnits.strMccbFilePath = configXml.Descendants("add").FirstOrDefault(x => (string)x.Attribute("key") == "MccbFilePath")?.Attribute("value")?.Value;
        }

        private void ControlFormFunction()
        {

            lblSCcheckSheet.MouseClick += (o, e) =>
            {
                if (!IsValidPath(StringUnits.strPrjFolderPath))
                {
                    MessageBox.Show("프로젝트 폴더 경로가 올바르지 않습니다. 설정에서 경로를 설정하세요");
                    return;
                }

                if (!IsValidFile(StringUnits.strBasicTempletFilePath))
                {
                    MessageBox.Show("기본 프로젝트 경로가 올바르지 않습니다. 설정에서 경로를 설정하세요");
                    return;
                }

                if (!IsValidPath(StringUnits.strMacroFolderPath))
                {
                    MessageBox.Show("매크로 폴더 경로가 올바르지 않습니다. 설정에서 경로를 설정하세요");
                    return;
                }

                if (!IsValidFile(StringUnits.strIoListFilePath))
                {
                    MessageBox.Show("IO 템플릿 엑셀 경로가 올바르지 않습니다. 설정에서 경로를 설정하세요");
                    return;
                }

                if (!IsValidFile(StringUnits.strMccbFilePath))
                {
                    MessageBox.Show("IO 템플릿 엑셀 경로가 올바르지 않습니다. 설정에서 경로를 설정하세요");
                    return;
                }


                if (FormUnits.formConceptSheet == null || FormUnits.formConceptSheet.IsDisposed)
                {

                    FormUnits.formConceptSheet = new FormConceptSheet();
                    FormUnits.formConceptSheet.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
                }
                else
                {
                    // 기존 창이 이미 열려 있을 경우 해당 창으로 포커스 이동
                    FormUnits.formConceptSheet.Focus();
                }

                // 현재 폼 없앤다
                this.Hide();
                
            };
            picBoxSetting.MouseClick += (o, e) =>
            {
                if (FormUnits.formConfigPage == null || FormUnits.formConfigPage.IsDisposed)
                {
                    FormUnits.formConfigPage = new FormConfigPage();
                    FormUnits.formConfigPage.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
                }
                else
                {
                    FormUnits.formConfigPage.Focus(); // 기존 창에 포커스 이동

                }
            };

            
        }
        private bool IsValidFile(string path)
        {
            return File.Exists(path);
        }
        private bool IsValidPath(string path)
        {
            return !string.IsNullOrEmpty(path) && Directory.Exists(Path.GetDirectoryName(path));
        }
        private void SetToolTip()
        {
            tip.SetToolTip(lblLogo, "메인 메뉴");
            tip.SetToolTip(picBoxLogo, "메인 메뉴");

            tip.SetToolTip(picBoxSetting, "경로 셋팅");
        }
    }

}
