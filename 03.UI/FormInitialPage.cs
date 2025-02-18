using DevExpress.XtraEditors;
using Eplan.MCNS.Lib.Share_CS;
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

namespace Eplan.EplAddin.HMX_MCNS._03.UI
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
            //CS_PathData.ConfigFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "config.xml");// @"C:\Users\kr70009769\Desktop\01.Task\02. API 소스\01. ProtoType\Eplan.EplAddin.HMX_MCNS\bin\Debug\config.xml";
            //CS_PathData.ItemListFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ItemList.xml");//@"C:\Users\kr70009769\Desktop\01.Task\02. API 소스\01. ProtoType\Eplan.EplAddin.HMX_MCNS\bin\Debug\ItemList.xml";

            // XML 파일을 로드합니다.
            XDocument configXml = XDocument.Load(CS_PathData.ConfigFilePath);

            // 기초 파일 paths 가져오기
            CS_PathData.PrjFolderPath = configXml.Descendants("add").FirstOrDefault(x => (string)x.Attribute("key") == "ProjectSaveFolder")?.Attribute("value")?.Value;
            CS_PathData.BasicTempletFilePath = configXml.Descendants("add").FirstOrDefault(x => (string)x.Attribute("key") == "BasicTempletFilePath")?.Attribute("value")?.Value;
            CS_PathData.MacroFolderPath = configXml.Descendants("add").FirstOrDefault(x => (string)x.Attribute("key") == "MacroFolderPath")?.Attribute("value")?.Value;
            CS_PathData.IoListFilePath = configXml.Descendants("add").FirstOrDefault(x => (string)x.Attribute("key") == "IoListFilePath")?.Attribute("value")?.Value;
            CS_PathData.MccbFilePath = configXml.Descendants("add").FirstOrDefault(x => (string)x.Attribute("key") == "MccbFilePath")?.Attribute("value")?.Value;
        }

        private void ControlFormFunction()
        {
            //마우스 클릭 폼 이동
            pnlTap.MouseDown += (o, e) => { if (e.Button == MouseButtons.Left) { CS_StaticEtc.On = true; CS_StaticEtc.Pos = e.Location; } };
            pnlTap.MouseMove += (o, e) => { if (CS_StaticEtc.On) Location = new Point(Location.X + (e.X - CS_StaticEtc.Pos.X), Location.Y + (e.Y - CS_StaticEtc.Pos.Y)); };
            pnlTap.MouseUp += (o, e) => { if (e.Button == MouseButtons.Left) { CS_StaticEtc.On = false; CS_StaticEtc.Pos = e.Location; } };            

            lblSCcheckSheet.MouseClick += (o, e) =>
            {
                if (!IsValidPath(CS_PathData.PrjFolderPath))
                {
                    MessageBox.Show("프로젝트 폴더 경로가 올바르지 않습니다. 설정에서 경로를 설정하세요");
                    return;
                }

                if (!IsValidFile(CS_PathData.BasicTempletFilePath))
                {
                    MessageBox.Show("기본 프로젝트 경로가 올바르지 않습니다. 설정에서 경로를 설정하세요");
                    return;
                }

                if (!IsValidPath(CS_PathData.MacroFolderPath))
                {
                    MessageBox.Show("매크로 폴더 경로가 올바르지 않습니다. 설정에서 경로를 설정하세요");
                    return;
                }

                if (!IsValidFile(CS_PathData.IoListFilePath))
                {
                    MessageBox.Show("IO 템플릿 엑셀 경로가 올바르지 않습니다. 설정에서 경로를 설정하세요");
                    return;
                }

                if (!IsValidFile(CS_PathData.MccbFilePath))
                {
                    MessageBox.Show("IO 템플릿 엑셀 경로가 올바르지 않습니다. 설정에서 경로를 설정하세요");
                    return;
                }


                if (Ui_StaticForm.formConceptSheet == null || Ui_StaticForm.formConceptSheet.IsDisposed)
                {

                    Ui_StaticForm.formConceptSheet = new FormConceptSheet();
                    Ui_StaticForm.formConceptSheet.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
                }
                else
                {
                    // 기존 창이 이미 열려 있을 경우 해당 창으로 포커스 이동
                    Ui_StaticForm.formConceptSheet.Focus();
                }

                // 현재 폼 없앤다
                this.Hide();
                
            };
            picBoxSetting.MouseClick += (o, e) =>
            {
                if (Ui_StaticForm.formConfigPage == null || Ui_StaticForm.formConfigPage.IsDisposed)
                {
                    Ui_StaticForm.formConfigPage = new FormConfigPage();
                    Ui_StaticForm.formConfigPage.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
                }
                else
                {
                    Ui_StaticForm.formConfigPage.Focus(); // 기존 창에 포커스 이동

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
