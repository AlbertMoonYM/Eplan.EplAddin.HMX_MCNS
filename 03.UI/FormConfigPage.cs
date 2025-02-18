using DevExpress.XtraEditors;
using Eplan.MCNS.Lib.Share_CS;
using Eplan.MCNS.Lib;
using Eplan.MCNS.Lib.UI_CS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace Eplan.EplAddin.HMX_MCNS._03.UI
{
    public partial class FormConfigPage : DevExpress.XtraEditors.XtraForm
    {
        CS_Button cs_Button = new CS_Button();
        
        public FormConfigPage()
        {
            InitializeComponent();

            ControlFormFunction();

            //초기값
            cbGenPrjFolderPath.Text = CS_PathData.PrjFolderPath;
            cbBasicTempletFilePath.Text = CS_PathData.BasicTempletFilePath;
            cbIoExcelFilesPath.Text = CS_PathData.IoListFilePath;
            cbMacroFolderPath.Text = CS_PathData.MacroFolderPath;
            cbMccbFilePath.Text = CS_PathData.MccbFilePath;


            //경로 바꾸기 액션
            cs_Button.FolderFinder(btnGenPrjFolderPath, cbGenPrjFolderPath);
            cs_Button.FileFinder(btnBasicTempletFilePath, cbBasicTempletFilePath, CS_PathData.XmlFolderPath, "zw9 File (*.zw9)|*.zw9|All Files (*.*)|*.*");
            cs_Button.FileFinder(btnIoExcelFilesPath, cbIoExcelFilesPath, CS_PathData.XmlFolderPath, "Excel File (*.xlsx)|*.xlsx|All Files (*.*)|*.*");
            cs_Button.FolderFinder(btnMacroFolderPath, cbMacroFolderPath);
            cs_Button.FileFinder(btnMccbFilePath, cbMccbFilePath, CS_PathData.XmlFolderPath, "Excel File (*.xlsx)|*.xlsx|All Files (*.*)|*.*");


        }
        public void ControlFormFunction()
        {
            //마우스 클릭 폼 이동
            pnlTap.MouseDown += (o, e) => { if (e.Button == MouseButtons.Left) { CS_StaticEtc.On = true; CS_StaticEtc.Pos = e.Location; } };
            pnlTap.MouseMove += (o, e) => { if (CS_StaticEtc.On) Location = new Point(Location.X + (e.X - CS_StaticEtc.Pos.X), Location.Y + (e.Y - CS_StaticEtc.Pos.Y)); };
            pnlTap.MouseUp += (o, e) => { if (e.Button == MouseButtons.Left) { CS_StaticEtc.On = false; CS_StaticEtc.Pos = e.Location; } };
            this.FormClosing += (o, e) =>
            {
                // 잘못된 경로나 파일 경로를 저장할 변수
                string errActPathTxt = "";

                // 검증할 컨트롤 배열
                Control[] actPath = { cbGenPrjFolderPath, cbMacroFolderPath };
                Control[] actFile = { cbBasicTempletFilePath, cbIoExcelFilesPath, cbMccbFilePath };

                // 경로 검증
                foreach (ComboBoxEdit cb in actPath)
                {
                    if (!IsValidPath(cb.Text))
                    {
                        string labelText = cb.Parent.Controls.OfType<Label>().FirstOrDefault()?.Text ?? "알 수 없는 항목";
                        errActPathTxt += $"[{labelText}]";
                    }
                }

                // 파일 검증
                foreach (ComboBoxEdit cb in actFile)
                {
                    if (!IsValidFile(cb.Text))
                    {
                        string labelText = cb.Parent.Controls.OfType<Label>().FirstOrDefault()?.Text ?? "알 수 없는 항목";
                        errActPathTxt += $"[{labelText}]";
                    }
                }

                // 잘못된 경로나 파일이 있는 경우
                if (!string.IsNullOrEmpty(errActPathTxt))
                {
                    DialogResult result = MessageBox.Show(
                        $"{errActPathTxt} 경로가 올바르지 않습니다. 나가시겠습니까?",
                        "경고",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning
                    );

                    // 사용자가 "아니오"를 선택하면 닫기 취소
                    if (result == DialogResult.No)
                    {
                        e.Cancel = true;
                        return;
                    }
                }

                // 리소스 정리
                if (Ui_StaticForm.formConfigPage != null && !Ui_StaticForm.formConfigPage.IsDisposed)
                {
                    Ui_StaticForm.formConfigPage.Dispose();
                }
            };




            btnSaveConfig.Click += (o, e) =>
            {
                try
                {
                    // config 파일 경로
                    string configFilePath = CS_PathData.ConfigFilePath;

                    // XML 파일 로드
                    XDocument xdoc = XDocument.Load(configFilePath);

                    // 수정할 경로들
                    string newPrjFolderPath = cbGenPrjFolderPath.Text;
                    string newPrjTempletPath = cbBasicTempletFilePath.Text;
                    string newIoExcelFilesPath = cbIoExcelFilesPath.Text;
                    string newMacroFolderPath = cbMacroFolderPath.Text;
                    string newMccbFilePath = cbMccbFilePath.Text;


                    CS_PathData.PrjFolderPath = newPrjFolderPath;
                    CS_PathData.BasicTempletFilePath = newPrjTempletPath;
                    CS_PathData.IoListFilePath = newIoExcelFilesPath;
                    CS_PathData.MacroFolderPath = newMacroFolderPath;
                    CS_PathData.MccbFilePath = newMccbFilePath;


                    // XML 내용 수정
                    xdoc.Descendants("add")
                        .FirstOrDefault(x => (string)x.Attribute("key") == "ProjectSaveFolder")?.SetAttributeValue("value", newPrjFolderPath);

                    xdoc.Descendants("add")
                        .FirstOrDefault(x => (string)x.Attribute("key") == "BasicTempletFilePath")?.SetAttributeValue("value", newPrjTempletPath);

                    xdoc.Descendants("add")
                        .FirstOrDefault(x => (string)x.Attribute("key") == "IoListFilePath")?.SetAttributeValue("value", newIoExcelFilesPath);

                    xdoc.Descendants("add")
                        .FirstOrDefault(x => (string)x.Attribute("key") == "MacroFolderPath")?.SetAttributeValue("value", newMacroFolderPath);

                    xdoc.Descendants("add")
                        .FirstOrDefault(x => (string)x.Attribute("key") == "MccbFilePath")?.SetAttributeValue("value", newMccbFilePath);


                    // 수정된 XML 파일 저장
                    xdoc.Save(configFilePath);

                    MessageBox.Show("설정이 저장되었습니다.", "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"설정 저장 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        
    }
}