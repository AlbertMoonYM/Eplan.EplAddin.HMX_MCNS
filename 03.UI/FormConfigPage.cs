using DevExpress.XtraEditors;
using Eplan.MCNS.Lib;
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

namespace Eplan.EplAddin.HMX_MCNS
{
    public partial class FormConfigPage : DevExpress.XtraEditors.XtraForm
    {
        ButtonManager btnManager = new ButtonManager();
        
        public FormConfigPage()
        {
            InitializeComponent();

            ControlFormFunction();

            //초기값
            cbGenPrjFolderPath.Text = StringUnits.strPrjFolderPath;
            cbBasicTempletFilePath.Text = StringUnits.strBasicTempletFilePath;
            cbIoExcelFilesPath.Text = StringUnits.strIoListFilePath;
            cbMacroFolderPath.Text = StringUnits.strMacroFolderPath;
            cbMccbFilePath.Text = StringUnits.strMccbFilePath;


            //경로 바꾸기 액션
            btnManager.FolderFinder(btnGenPrjFolderPath, cbGenPrjFolderPath);
            btnManager.FileFinder(btnBasicTempletFilePath, cbBasicTempletFilePath, StringUnits.strXmlFolderPath, "zw9 File (*.zw9)|*.zw9|All Files (*.*)|*.*");
            btnManager.FileFinder(btnIoExcelFilesPath, cbIoExcelFilesPath, StringUnits.strXmlFolderPath, "Excel File (*.xlsx)|*.xlsx|All Files (*.*)|*.*");
            btnManager.FolderFinder(btnMacroFolderPath, cbMacroFolderPath);
            btnManager.FileFinder(btnMccbFilePath, cbMccbFilePath, StringUnits.strXmlFolderPath, "Excel File (*.xlsx)|*.xlsx|All Files (*.*)|*.*");


        }
        public void ControlFormFunction()
        {
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
                if (FormUnits.formConfigPage != null && !FormUnits.formConfigPage.IsDisposed)
                {
                    FormUnits.formConfigPage.Dispose();
                }
            };




            btnSaveConfig.Click += (o, e) =>
            {
                try
                {
                    // config 파일 경로
                    string configFilePath = StringUnits.strConfigFilePath;

                    // XML 파일 로드
                    XDocument xdoc = XDocument.Load(configFilePath);

                    // 수정할 경로들
                    string newPrjFolderPath = cbGenPrjFolderPath.Text;
                    string newPrjTempletPath = cbBasicTempletFilePath.Text;
                    string newIoExcelFilesPath = cbIoExcelFilesPath.Text;
                    string newMacroFolderPath = cbMacroFolderPath.Text;
                    string newMccbFilePath = cbMccbFilePath.Text;


                    StringUnits.strPrjFolderPath = newPrjFolderPath;
                    StringUnits.strBasicTempletFilePath = newPrjTempletPath;
                    StringUnits.strIoListFilePath = newIoExcelFilesPath;
                    StringUnits.strMacroFolderPath = newMacroFolderPath;
                    StringUnits.strMccbFilePath = newMccbFilePath;


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