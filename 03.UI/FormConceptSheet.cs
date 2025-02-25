using DevExpress.XtraEditors;
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

using Eplan.MCNS.Lib;
using System.Reflection.Emit;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid;
using DevExpress.XtraPrinting.Preview;
using System.Diagnostics;
using System.Xml.Linq;
using ClosedXML.Excel;
using DevExpress.XtraEditors.Repository;
using McnsSchemGenEngine.Models;
using OfficeOpenXml;
using DevExpress.XtraTab;
using System.Threading;
using DevExpress.Utils.Extensions;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;


namespace Eplan.EplAddin.HMX_MCNS
{
    public partial class FormConceptSheet : DevExpress.XtraEditors.XtraForm
    {
       

        //외부 CS 정의
        //외부 CS 정의
        LabelManager lblManager = new LabelManager();
        ComboBoxManager cbManager = new ComboBoxManager();
        DataTableManager dtManager = new DataTableManager();
        GridViewManager gvManager = new GridViewManager();
        GroupControlManager gcManager = new GroupControlManager();
        XtraTabControlManager xtcManager = new XtraTabControlManager();
        CheckboxManager ckbManager = new CheckboxManager();
        InterLockLibrary interLock = new InterLockLibrary();
        FilePathManager pathManager = new FilePathManager();
        FileSystemWatcher watcher; // FileSystemWatcher 선언
        ToolTip tip = new ToolTip();
        McnsSchemGenEngine.Controls.McnsControl mcnsControl = new McnsSchemGenEngine.Controls.McnsControl();
        

        public FormConceptSheet()
        {
            InitializeComponent();


            SetProjectObjecstSetting();
            SetModelObjectsSetting();
            SetMainSpecObjectsSetting();
            SetLayOutObjectsSetting();
            SetOptionObjectsSetting();
            SetFunctionObjectsSetting();
            SetPanel3D();
            SetToolTip();
            SetComboBoxItems();
            GetSetFunctionData();
            ControlFormFunction();

            ControlPlcFunction();
            LoadIoFromExcel();
            LoadMccbFromExcel();
            ActivateEplan();

            UpdateComboBoxItemList();
            Interlock();
        }
        private void SetPanel3D()
        {
            LabelControl[] labelControls = new LabelControl[] 
            {
                labelControl10, labelControl11, labelControl12, labelControl13, labelControl14, labelControl15, labelControl16,labelControl18,labelControl19,
                labelControl28,labelControl46,labelControl21,labelControl22, labelControl23, labelControl24,labelControl25, labelControl26, labelControl27,labelControl29,
                labelControl31,labelControl32,labelControl33,labelControl34,labelControl35,
                labelControl37,labelControl38,labelControl39,labelControl40,
                labelControl41,labelControl42,labelControl43, labelControl45
            };

            foreach(LabelControl lbc in labelControls)
            {
                lbc.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Style3D; // 3D 스타일
            }

        }
        private void SetProjectObjecstSetting()
        {
            // button을 Radio button으로 사용
            ckbManager.ChangeToRadioButton(ckbPRJdomestic, ckbPRJoverseas);
            // 프로젝트 ChangeToTextBox 설정
            cbManager.ChangeToTextBox(cbPRJnumber, "텍스트를 기입하세요.", false, TypeFlag.strFlag);
            cbManager.ChangeToTextBox(cbPRJname, "텍스트를 기입하세요.", false, TypeFlag.strFlag);
            cbManager.ChangeToTextBox(cbPRJwriter, "텍스트를 기입하세요.", false, TypeFlag.strFlag);
            // 프로젝트 SettingComboBox 설정
            cbManager.SettingComboBox(cbPRJyear, "년도", false);
            cbManager.SettingComboBox(cbPRJmonth, "월", false);
            cbManager.SettingComboBox(cbPRJday, "일", false);
            int currentYear = DateTime.Now.Year;
            // 연도 추가 (현재 연도를 기준으로 ±10년)
            cbPRJyear.Properties.Items.AddRange(
                Enumerable.Range(currentYear - 10, 21).Select(y => y.ToString()).ToArray());

            // 월 추가 (1~12)
            cbPRJmonth.Properties.Items.AddRange(
                Enumerable.Range(1, 12).Select(m => m.ToString("D2")).ToArray());

            // 일 추가 (1~31)
            cbPRJday.Properties.Items.AddRange(
                Enumerable.Range(1, 31).Select(d => d.ToString("D2")).ToArray());
        }
        private void SetModelObjectsSetting()
        {
            // 모델 ChangeToTextBox 설정
            cbManager.ChangeToTextBox(cbMODfullName, "-", false, TypeFlag.strFlag);
            cbManager.ChangeToTextBox(cbMODheight, "높이", false, TypeFlag.intFlag);
            cbManager.ChangeToTextBox(cbMODweight, "화물 중량", false, TypeFlag.intFlag);
            cbManager.ChangeToTextBox(cbMODfullName, "", false, TypeFlag.intFlag);
            // 모델 SettingComboBox 설정
            cbManager.SettingComboBox(cbMODname, "모델명", false);
            cbManager.SettingComboBox(cbMODoption1, "-", false);
            cbManager.SettingComboBox(cbMODoption2, "-", false);
            cbManager.SettingComboBox(cbMODoption3, "-", false);
            cbManager.SettingComboBox(cbMODoption4, "-", false);
        }
        private void SetMainSpecObjectsSetting()
        {
            // 주요 사양 ChangeToTextBox 설정
            cbManager.ChangeToTextBox(cbMSPpanelSizeW, "W", false, TypeFlag.intFlag);
            cbManager.ChangeToTextBox(cbMSPpanelSizeD, "D", false, TypeFlag.intFlag);
            cbManager.ChangeToTextBox(cbMSPpanelSizeH, "H", false, TypeFlag.intFlag);
            // 주요 사양 SettingComboBox 설정
            cbManager.SettingComboBox(cbMSPinputVolt, "V", true);
            cbManager.SettingComboBox(cbMSPinputHz, "Hz", true);
            cbManager.SettingComboBox(cbMSPpanelSize, "-", false);
            cbManager.SettingComboBox(cbMSPcontrollerSpec, "-", false);
            cbManager.SettingComboBox(cbMSPinverterMaker, "-", false);
            cbManager.SettingComboBox(cbMSPinverterSpec, "-", false);
        }
        private void SetLayOutObjectsSetting()
        {
            // 레이아웃 ChangeToTextBox 설정
            cbManager.ChangeToTextBox(cbLOUTtravLength, "M", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLOUTliftHeight, "M", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLOUTstationNum, "EA", true, TypeFlag.intFlag);

            // 레이아웃 화물 DataGridView 셋업
            dtManager.GetDataTable(DataTableUnits.dtLout, StringUnits.dArrDtLoutColums);
            // 레이아웃 화물 DataTable 3행 셋업
            DataTableUnits.dtLout.Rows.Add("화물1");
            DataTableUnits.dtLout.Rows.Add("화물2");
            DataTableUnits.dtLout.Rows.Add("화물3");
            DataTableUnits.dtLout.Rows.Add("화물4");
            gridLOUTcargo.DataSource = DataTableUnits.dtLout;
            gvManager.SetLoutCargo(gridViewCargo);
        }
        private void SetOptionObjectsSetting()
        {
            // 옵션 SettingComboBox 설정
            cbManager.SettingComboBox(cbOPmachineControl, "-", false);
            cbManager.SettingComboBox(cbOPremoteControl, "-", false);
            cbManager.SettingComboBox(cbOPemergencyPower, "-", false);
            cbManager.SettingComboBox(cbOPemergencyLocation, "-", false);
        }
        private void SetFunctionObjectsSetting()
        {
            // ELEQ 기능 그룹
            cbManager.ChangeToTextBox(cbEleqBrakeResistorKw, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbEleqBrakeResistorOhm, "Ω", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbEleqPowerKw, "kW", true);
            cbManager.SettingComboBox(cbEleqPowerA, "A", true);
            cbManager.SettingComboBox(cbEleqMainMccbSpec, "-", false);
            cbManager.SettingComboBox(cbEleqMainPowerCable, "SQ", true);
            cbManager.SettingComboBox(cbEleqMccbSpec, "-", false);
            cbManager.SettingComboBox(cbEleqPowerCable, "SQ", true);
            cbManager.SettingComboBox(cbMSPusingVoltage, "V", true);
            cbManager.SettingComboBox(cbEleqMccbModel, "-", false);
            cbManager.SettingComboBox(cbEleqSmpsModel, "-", false);
            cbManager.SettingComboBox(cbEleqCableModel, "-", false);
            cbManager.SettingComboBox(cbEleqHubModel, "-", false);
            cbManager.SettingComboBox(cbEleqFanQuantity, "-", false);
            cbManager.SettingComboBox(cbEleqTerminal, "-", false);
            cbManager.SettingComboBox(cbEleqPanel, "-", false);
            cbManager.SettingComboBox(cbEleqHmi, "-", false);
            cbManager.SettingComboBox(cbEleqOpt, "-", false);
            cbManager.SettingComboBox(cbEleqTowerLamp, "-", false);
            cbManager.SettingComboBox(cbEleqSafety, "-", false);
            cbManager.SettingComboBox(cbEleqSafetyQuantity, "-", false);
            cbManager.SettingComboBox(cbEleqSensorType, "-", false);
            cbManager.SettingComboBox(cbEleqModem, "-", false);
            cbManager.SettingComboBox(cbEleqInterLockSensorSide, "-", false);
            cbManager.SettingComboBox(cbEleqInterLockBit, "-", false);
            cbManager.SettingComboBox(cbEleqLocation, "-", false);
            cbManager.SettingComboBox(cbEleqType, "-", false);
            cbManager.SettingComboBox(cbEleqDt, "-", false);
            cbManager.SettingComboBox(cbEleqParts, "-", false);
            cbManager.SettingComboBox(cbEleqPoint, "-", false);
            cbManager.SettingComboBox(cbEleqSensorItem, "-", false);

            // LIFT 기능 그룹
            cbManager.SettingComboBox(cbLiftInverterKw, "kW", true);
            cbManager.SettingComboBox(cbLiftInverterA, "A", true);
            cbManager.ChangeToTextBox(cbLiftBrakeResistorKw, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftBrakeResistorOhm, "Ω", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbLiftPowerCable, "SQ", true);
            cbManager.SettingComboBox(cbLiftMccbSpec, "-", false);
            cbManager.ChangeToTextBox(cbLiftOutPut, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftSpeed, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftGearRatio, "i", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftRatedCurrent, "A", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftBkVoltage, "V", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbLiftBrakeOption, "-", false);
            cbManager.SettingComboBox(cbLiftMotorSpec, "-", false);
            cbManager.SettingComboBox(cbLiftMotorMaker, "-", false);
            cbManager.SettingComboBox(cbLiftMotorMethod, "-", false);
            cbManager.ChangeToTextBox(cbLiftMotorType, "텍스트 기입", false, TypeFlag.strFlag);
            cbManager.ChangeToTextBox(cbLiftMotorVoltage, "V", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftMotorHz, "Hz", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftMotorEncoderSpec, "텍스트 기입", false, TypeFlag.strFlag);
            cbManager.SettingComboBox(cbLiftAbsLocation, "-", false);
            cbManager.SettingComboBox(cbLiftRightPosition, "-", false);
            cbManager.SettingComboBox(cbLiftLimitSwitch, "-", false);
            cbManager.ChangeToTextBox(cbLiftNoneLoadHighSpeed, "MPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftNoneLoadRotationNum, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftNoneLoadAcceleration, "", false, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftLoadHighSpeed, "MPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftLoadRotationNum, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbLiftLoadAcceleration, "", false, TypeFlag.fltFlag);

            // TRAV1 기능 그룹
            cbManager.SettingComboBox(cbTrav1InverterKw, "kW", true);
            cbManager.SettingComboBox(cbTrav1InverterA, "A", true);
            cbManager.ChangeToTextBox(cbTrav1BrakeResistorKw, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1BrakeResistorOhm, "Ω", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbTrav1PowerCable, "SQ", true);
            cbManager.SettingComboBox(cbTrav1MccbSpec, "-", false);
            cbManager.ChangeToTextBox(cbTrav1OutPut, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1Speed, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1GearRatio, "i", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1RatedCurrent, "A", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1BkVoltage, "V", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbTrav1BrakeOption, "-", false);
            cbManager.SettingComboBox(cbTrav1MotorSpec, "-", false);
            cbManager.SettingComboBox(cbTrav1MotorMaker, "-", false);
            cbManager.SettingComboBox(cbTrav1MotorMethod, "-", false);
            cbManager.ChangeToTextBox(cbTrav1MotorType, "텍스트 기입", false, TypeFlag.strFlag);
            cbManager.ChangeToTextBox(cbTrav1MotorVoltage, "V", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1MotorHz, "Hz", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1MotorEncoderSpec, "텍스트 기입", false, TypeFlag.strFlag);
            cbManager.SettingComboBox(cbTrav1AbsLocation, "-", false);
            cbManager.SettingComboBox(cbTrav1RightPosition, "-", false);
            cbManager.SettingComboBox(cbTrav1LimitSwitch, "-", false);
            cbManager.ChangeToTextBox(cbTrav1NoneLoadHighSpeed, "MPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1NoneLoadRotationNum, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1NoneLoadAcceleration, "", false, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1LoadHighSpeed, "MPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1LoadRotationNum, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav1LoadAcceleration, "", false, TypeFlag.fltFlag);

            // Trav2 기능 그룹
            cbManager.SettingComboBox(cbTrav2InverterKw, "kW", true);
            cbManager.SettingComboBox(cbTrav2InverterA, "A", true);
            cbManager.ChangeToTextBox(cbTrav2BrakeResistorKw, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2BrakeResistorOhm, "Ω", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbTrav2PowerCable, "SQ", true);
            cbManager.SettingComboBox(cbTrav2MccbSpec, "-", false);
            cbManager.ChangeToTextBox(cbTrav2OutPut, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2Speed, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2GearRatio, "i", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2RatedCurrent, "A", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2BkVoltage, "V", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbTrav2BrakeOption, "-", false);
            cbManager.SettingComboBox(cbTrav2MotorSpec, "-", false);
            cbManager.SettingComboBox(cbTrav2MotorMaker, "-", false);
            cbManager.SettingComboBox(cbTrav2MotorMethod, "-", false);
            cbManager.ChangeToTextBox(cbTrav2MotorType, "텍스트 기입", false, TypeFlag.strFlag);
            cbManager.ChangeToTextBox(cbTrav2MotorVoltage, "V", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2MotorHz, "Hz", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2MotorEncoderSpec, "텍스트 기입", false, TypeFlag.strFlag);
            cbManager.SettingComboBox(cbTrav2AbsLocation, "-", false);
            cbManager.SettingComboBox(cbTrav2RightPosition, "-", false);
            cbManager.SettingComboBox(cbTrav2LimitSwitch, "-", false);
            cbManager.ChangeToTextBox(cbTrav2NoneLoadHighSpeed, "MPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2NoneLoadRotationNum, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2NoneLoadAcceleration, "", false, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2LoadHighSpeed, "MPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2LoadRotationNum, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbTrav2LoadAcceleration, "", false, TypeFlag.fltFlag);

            // FORK1 기능 그룹
            cbManager.SettingComboBox(cbFork1InverterKw, "kW", true);
            cbManager.SettingComboBox(cbFork1InverterA, "A", true);
            cbManager.ChangeToTextBox(cbFork1BrakeResistorKw, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1BrakeResistorOhm, "Ω", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbFork1PowerCable, "SQ", true);
            cbManager.SettingComboBox(cbFork1MccbSpec, "-", false);
            cbManager.ChangeToTextBox(cbFork1OutPut, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1Speed, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1GearRatio, "i", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1RatedCurrent, "A", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1BkVoltage, "V", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbFork1BrakeOption, "-", false);
            cbManager.SettingComboBox(cbFork1MotorSpec, "-", false);
            cbManager.SettingComboBox(cbFork1MotorMaker, "-", false);
            cbManager.SettingComboBox(cbFork1MotorMethod, "-", false);
            cbManager.ChangeToTextBox(cbFork1MotorType, "텍스트 기입", false, TypeFlag.strFlag);
            cbManager.ChangeToTextBox(cbFork1MotorVoltage, "V", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1MotorHz, "Hz", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1MotorEncoderSpec, "텍스트 기입", false, TypeFlag.strFlag);
            cbManager.SettingComboBox(cbFork1RightPosition, "-", false);
            cbManager.ChangeToTextBox(cbFork1NoneLoadHighSpeed, "MPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1NoneLoadRotationNum, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1NoneLoadAcceleration, "", false, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1LoadHighSpeed, "MPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1LoadRotationNum, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork1LoadAcceleration, "", false, TypeFlag.fltFlag);

            // FORK2 기능 그룹
            cbManager.SettingComboBox(cbFork2InverterKw, "kW", true);
            cbManager.SettingComboBox(cbFork2InverterA, "A", true);
            cbManager.ChangeToTextBox(cbFork2BrakeResistorKw, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2BrakeResistorOhm, "Ω", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbFork2PowerCable, "SQ", true);
            cbManager.SettingComboBox(cbFork2MccbSpec, "-", false);
            cbManager.ChangeToTextBox(cbFork2OutPut, "kW", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2Speed, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2GearRatio, "i", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2RatedCurrent, "A", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2BkVoltage, "V", true, TypeFlag.fltFlag);
            cbManager.SettingComboBox(cbFork2BrakeOption, "-", false);
            cbManager.SettingComboBox(cbFork2MotorSpec, "-", false);
            cbManager.SettingComboBox(cbFork2MotorMaker, "-", false);
            cbManager.SettingComboBox(cbFork2MotorMethod, "-", false);
            cbManager.ChangeToTextBox(cbFork2MotorType, "텍스트 기입", false, TypeFlag.strFlag);
            cbManager.ChangeToTextBox(cbFork2MotorVoltage, "V", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2MotorHz, "Hz", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2MotorEncoderSpec, "텍스트 기입", false, TypeFlag.strFlag);
            cbManager.SettingComboBox(cbFork2RightPosition, "-", false);
            cbManager.ChangeToTextBox(cbFork2NoneLoadHighSpeed, "MPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2NoneLoadRotationNum, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2NoneLoadAcceleration, "", false, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2LoadHighSpeed, "MPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2LoadRotationNum, "RPM", true, TypeFlag.fltFlag);
            cbManager.ChangeToTextBox(cbFork2LoadAcceleration, "", false, TypeFlag.fltFlag);

            // CARR 기능 그룹
            cbManager.SettingComboBox(cbCarrSensor, "-", false);
            cbManager.SettingComboBox(cbCarrDoubleInput, "-", false);
        }
        private void SetComboBoxItems()
        {
            // 모델 콤보 박스 리스트 설정
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMODName", cbMODname);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMODOption", cbMODoption1);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMODOption", cbMODoption2);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMODOption", cbMODoption3);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMODOption", cbMODoption4);

            // 주요 사양 콤보 박스 리스트 설정
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMSPinputVolt", cbMSPinputVolt);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMSPinputHz", cbMSPinputHz);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMSPpanelSize", cbMSPpanelSize);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMSPcontrollerSpec", cbMSPcontrollerSpec);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMSPinverterMaker", cbMSPinverterMaker);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listMSPinverterSpec", cbMSPinverterSpec);

            // 옵션 콤보 박스 리스트 설정
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listOPmachineControl", cbOPmachineControl);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listOPremoteControl", cbOPremoteControl);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listOPemergencyPower", cbOPemergencyPower);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listOPemergencyLocation", cbOPemergencyLocation);

            // ELEQ 콤보 박스 리스트 설정
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqUsingVoltage", cbMSPusingVoltage);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqMccbModel", cbEleqMccbModel);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqSmpsModel", cbEleqSmpsModel);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqCableModel", cbEleqCableModel);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqHubModel", cbEleqHubModel);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqFanQuantity", cbEleqFanQuantity);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqTerminal", cbEleqTerminal);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqPanel", cbEleqPanel);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqHmi", cbEleqHmi);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqOpt", cbEleqOpt);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqTowerLamp", cbEleqTowerLamp);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqSafety", cbEleqSafety);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqSafetyQuantity", cbEleqSafetyQuantity);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqSensorType", cbEleqSensorType);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqInterLockSensorSide", cbEleqInterLockSensorSide);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqInterLockBit", cbEleqInterLockBit);

            // LIFT 콤보 박스 리스트 설정
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftMotorSpec", cbLiftMotorSpec);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftMotorMaker", cbLiftMotorMaker);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftMotorMethod", cbLiftMotorMethod);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftLimitSwitch", cbLiftLimitSwitch);


            // TRAV 콤보 박스 리스트 설정
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravMotorSpec", cbTrav1MotorSpec);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravMotorMaker", cbTrav1MotorMaker);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravMotorMethod", cbTrav1MotorMethod);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravLimitSwitch", cbTrav1LimitSwitch);

            // TRAV2 콤보 박스 리스트 설정
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravMotorSpec", cbTrav2MotorSpec);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravMotorMaker", cbTrav2MotorMaker);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravMotorMethod", cbTrav2MotorMethod);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravLimitSwitch", cbTrav2LimitSwitch);

            // FORK 콤보 박스 리스트 설정
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkMotorSpec", cbFork1MotorSpec);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkMotorMaker", cbFork1MotorMaker);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkMotorMethod", cbFork1MotorMethod);

            // FORK2 콤보 박스 리스트 설정
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkMotorSpec", cbFork2MotorSpec);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkMotorMaker", cbFork2MotorMaker);
            pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkMotorMethod", cbFork2MotorMethod);

            

        }
        private void GetSetFunctionData()
        {
            // ModelPage ComboBox 배열을 포함하는 Dictionary를 정의
            DictionaryUnits.dicCtrlMod = new Dictionary<string, Control[]>
            {
                { "프로젝트", new Control[] { ckbPRJdomestic, ckbPRJoverseas, cbPRJnumber, cbPRJname, cbPRJwriter, cbPRJyear, cbPRJmonth,cbPRJday } },
                { "모델", new Control[] { cbMODname, cbMODheight, cbMODweight, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4 } },
                { "주요사양",new Control[]{ cbMSPinputVolt, cbMSPusingVoltage, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbMSPpanelSize , cbMSPpanelSizeW, cbMSPpanelSizeD, cbMSPpanelSizeH } },
                { "레이아웃", new Control[] { ckbLevelSame, ckbBayTrue, cbLOUTtravLength, cbLOUTliftHeight, cbLOUTstationNum} },
                { "옵션", new Control[] { cbOPmachineControl, cbOPremoteControl, cbOPemergencyPower, cbOPemergencyLocation, ckbVibrationControlTrue, ckbCctvTrue, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue } }
            };
            // FunctionPage ComboBox 배열을 포함하는 Dictionary를 정의
            DictionaryUnits.dicCtrlFunc = new Dictionary<string, Control[]>
            {
                { "ELEQ_파워", new Control[] { cbEleqMccbModel, cbEleqSmpsModel, cbEleqCableModel, cbEleqHubModel,
                    cbEleqMainMccbSpec, cbEleqMainPowerCable, cbEleqFanQuantity, cbEleqTerminal, 
                    cbEleqPanel, cbEleqHmi, cbEleqOpt, cbEleqTowerLamp, 
                    cbEleqSafety, cbEleqSafetyQuantity, cbEleqMccbSpec, cbEleqPowerCable,
                    cbEleqPowerKw, cbEleqPowerA, cbEleqBrakeResistorKw, cbEleqBrakeResistorOhm
                } },
                { "ELEQ_센서", new Control[] { cbEleqSensorType, cbEleqModem, 
                    cbEleqInterLockSensorSide, cbEleqInterLockBit, cbEleqSensorItem } },
                { "LIFT_인버터", new Control[] { cbLiftInverterKw, cbLiftInverterA, cbLiftBrakeResistorKw, cbLiftBrakeResistorOhm, 
                    cbLiftPowerCable, cbLiftMccbSpec, ckbLiftDdi } },
                { "LIFT_모터", new Control[] { cbLiftOutPut, cbLiftSpeed, cbLiftGearRatio, cbLiftRatedCurrent, 
                    cbLiftBkVoltage, cbLiftBrakeOption, ckbLiftCoolerFan,
                    cbLiftMotorSpec, cbLiftMotorMaker, cbLiftMotorMethod, cbLiftMotorType,
                    cbLiftMotorVoltage, cbLiftMotorHz, cbLiftMotorEncoderSpec} },
                { "LIFT_센서", new Control[] { ckbLiftRaser, ckbLiftBarcode, cbLiftAbsLocation, cbLiftRightPosition, cbLiftLimitSwitch } },
                { "LIFT_드라이브", new Control[] { cbLiftNoneLoadHighSpeed, cbLiftNoneLoadRotationNum, cbLiftNoneLoadAcceleration, 
                    cbLiftLoadHighSpeed, cbLiftLoadRotationNum, cbLiftLoadAcceleration } },
                { "TRAV1_인버터", new Control[] { cbTrav1InverterKw, cbTrav1InverterA, cbTrav1BrakeResistorKw, cbTrav1BrakeResistorOhm, 
                    cbTrav1PowerCable, cbTrav1MccbSpec, ckbTrav1Ddi } },
                { "TRAV1_모터", new Control[] { cbTrav1OutPut, cbTrav1Speed, cbTrav1GearRatio, cbTrav1RatedCurrent,
                    cbTrav1BkVoltage, cbTrav1BrakeOption, ckbTrav1CoolerFan,
                    cbTrav1MotorSpec, cbTrav1MotorMaker, cbTrav1MotorMethod, cbTrav1MotorType,
                    cbTrav1MotorVoltage, cbTrav1MotorHz, cbTrav1MotorEncoderSpec} },
                { "TRAV1_센서", new Control[] { ckbTrav1Raser, ckbTrav1Barcode, cbTrav1AbsLocation, cbTrav1RightPosition, cbTrav1LimitSwitch } },
                { "TRAV1_드라이브", new Control[] { cbTrav1NoneLoadHighSpeed, cbTrav1NoneLoadRotationNum, cbTrav1NoneLoadAcceleration, 
                    cbTrav1LoadHighSpeed, cbTrav1LoadRotationNum, cbTrav1LoadAcceleration } },
                { "TRAV2_인버터", new Control[] { cbTrav2InverterKw, cbTrav2InverterA, cbTrav2BrakeResistorKw, cbTrav2BrakeResistorOhm,
                    cbTrav2PowerCable, cbTrav2MccbSpec, ckbTrav2Ddi } },
                { "TRAV2_모터", new Control[] { cbTrav2OutPut, cbTrav2Speed, cbTrav2GearRatio, cbTrav2RatedCurrent,
                    cbTrav2BkVoltage, cbTrav2BrakeOption, ckbTrav2CoolerFan,
                    cbTrav2MotorSpec, cbTrav2MotorMaker, cbTrav2MotorMethod, cbTrav2MotorType,
                    cbTrav2MotorVoltage, cbTrav2MotorHz, cbTrav2MotorEncoderSpec} },
                { "TRAV2_센서", new Control[] { ckbTrav2Raser, ckbTrav2Barcode, cbTrav2AbsLocation, cbTrav2RightPosition, cbTrav2LimitSwitch } },
                { "TRAV2_드라이브", new Control[] { cbTrav2NoneLoadHighSpeed, cbTrav2NoneLoadRotationNum, cbTrav2NoneLoadAcceleration,
                    cbTrav2LoadHighSpeed, cbTrav2LoadRotationNum, cbTrav2LoadAcceleration } },
                { "FORK1_인버터", new Control[] { cbFork1InverterKw, cbFork1InverterA, cbFork1BrakeResistorKw, cbFork1BrakeResistorOhm,
                    cbFork1PowerCable, cbFork1MccbSpec, ckbFork1Ddi } },
                { "FORK1_모터", new Control[] { cbFork1OutPut, cbFork1Speed, cbFork1GearRatio, cbFork1RatedCurrent,
                    cbFork1BkVoltage, cbFork1BrakeOption, ckbFork1CoolerFan,
                    cbFork1MotorSpec, cbFork1MotorMaker, cbFork1MotorMethod, cbFork1MotorType,
                    cbFork1MotorVoltage, cbFork1MotorHz, cbFork1MotorEncoderSpec} },
                { "FORK1_센서", new Control[] { cbFork1RightPosition, ckbFork1PosTrue } },
                { "FORK1_드라이브", new Control[] { cbFork1NoneLoadHighSpeed, cbFork1NoneLoadRotationNum, cbFork1NoneLoadAcceleration,
                    cbFork1LoadHighSpeed, cbFork1LoadRotationNum, cbFork1LoadAcceleration } },
                { "FORK2_인버터", new Control[] { cbFork2InverterKw, cbFork2InverterA, cbFork2BrakeResistorKw, cbFork2BrakeResistorOhm,
                    cbFork2PowerCable, cbFork2MccbSpec, ckbFork2Ddi } },
                { "FORK2_모터", new Control[] { cbFork2OutPut, cbFork2Speed, cbFork2GearRatio, cbFork2RatedCurrent,
                    cbFork2BkVoltage, cbFork2BrakeOption, ckbFork2CoolerFan,
                    cbFork2MotorSpec, cbFork2MotorMaker, cbFork2MotorMethod, cbFork2MotorType,
                    cbFork2MotorVoltage, cbFork2MotorHz, cbFork2MotorEncoderSpec} },
                { "FORK2_센서", new Control[] { cbFork2RightPosition, ckbFork2PosTrue } },
                { "FORK2_드라이브", new Control[] { cbFork2NoneLoadHighSpeed, cbFork2NoneLoadRotationNum, cbFork2NoneLoadAcceleration,
                    cbFork2LoadHighSpeed, cbFork2LoadRotationNum, cbFork2LoadAcceleration } },
                { "CARR_센서", new Control[] { cbCarrSensor, cbCarrDoubleInput } }
            };
            // FunctionPage GroupControl 배열을 포함하는 Dictionary를 정의
            DictionaryUnits.dicGrpSrmFunc = new Dictionary<string, GroupControl[]>
            {
                { "Eleq", new GroupControl[] { grpEleqPower, grpEleqSensor,grpEleqPlc } },
                { "Lift", new GroupControl[] { grpLiftInverter, grpLiftMotor, grpLiftSensor, grpLiftDrive, grpLiftSensorList } },
                { "Trav1", new GroupControl[] { grpTrav1Inverter, grpTrav1Motor, grpTrav1Sensor, grpTrav1Drive, grpTrav1SensorList } },
                { "Trav2", new GroupControl[] { grpTrav2Inverter, grpTrav2Motor, grpTrav2Sensor, grpTrav2Drive, grpTrav2SensorList } },
                { "Fork1", new GroupControl[] { grpFork1Inverter, grpFork1Motor, grpFork1Sensor, grpFork1Drive, grpFork1SensorList } },
                { "Fork2", new GroupControl[] { grpFork2Inverter, grpFork2Motor, grpFork2Sensor, grpFork2Drive, grpFork2SensorList } },
                { "Carr", new GroupControl[] { grpCarrSensor, grpCarrSensorList } }
            };
            // FunctionPage GridControl 배열을 정의
            ControlUnits.arrGcFunc = new GridControl[]
            {
                gridEleq,
                gridLift,
                gridTrav1,
                gridTrav2,
                gridFork1,
                gridFork2,
                gridCarr
            };
            // FunctionPage GridView 배열을 정의
            ControlUnits.arrGvFunc = new GridView[]
            {
                gridViewEleq,
                gridViewLift,
                gridViewTrav1,
                gridViewTrav2,
                gridViewFork1,
                gridViewFork2,
                gridViewCarr
            };

            // FunctionPage ComboBox 그룹 배열을 생성
            ControlUnits.arrCtrlGrpSrmFunc = new Control[][]
            {
                DictionaryUnits.dicCtrlFunc.Where(kvp => kvp.Key.Contains("ELEQ")).SelectMany(kvp => kvp.Value).ToArray(),
                DictionaryUnits.dicCtrlFunc.Where(kvp => kvp.Key.Contains("LIFT")).SelectMany(kvp => kvp.Value).ToArray(),
                DictionaryUnits.dicCtrlFunc.Where(kvp => kvp.Key.Contains("TRAV1")).SelectMany(kvp => kvp.Value).ToArray(),
                DictionaryUnits.dicCtrlFunc.Where(kvp => kvp.Key.Contains("TRAV2")).SelectMany(kvp => kvp.Value).ToArray(),
                DictionaryUnits.dicCtrlFunc.Where(kvp => kvp.Key.Contains("FORK1")).SelectMany(kvp => kvp.Value).ToArray(),
                DictionaryUnits.dicCtrlFunc.Where(kvp => kvp.Key.Contains("FORK2")).SelectMany(kvp => kvp.Value).ToArray(),
                DictionaryUnits.dicCtrlFunc.Where(kvp => kvp.Key.Contains("CARR")).SelectMany(kvp => kvp.Value).ToArray()
            };
            // FunctionPage DataTable 그룹 배열을 생성
            DataTableUnits.arrDtSrmFunc = new DataTable[DictionaryUnits.dicGrpSrmFunc.Count];
            // DataTable 그룹 배열의 값을 거치할 BindingSource 배열 생성
            ControlUnits.bindingSource = new BindingSource[DataTableUnits.arrDtSrmFunc.Length];
            // FunctionPage GroupControl 그룹 배열을 생성
            ControlUnits.arrGrpAll = DictionaryUnits.dicGrpSrmFunc.Values.SelectMany(arr => arr).ToArray();

            // 각 DataTable 생성 및 DataGridView 설정
            for (int i = 0; i < DictionaryUnits.dicGrpSrmFunc.Count; i++)
            {
                DataTableUnits.arrDtSrmFunc[i] = new DataTable();

                dtManager.GetDataTable(DataTableUnits.arrDtSrmFunc[i], StringUnits.dArrDtColums);

                // DataTable에 데이터 추가
                for (int j = 0; j < ControlUnits.arrCtrlGrpSrmFunc[i].Length; j++)
                {
                    // 각 콤보박스 항목에서 정보 추출
                    string pageName = "기능"; // 페이지 이름, 필요에 따라 수정
                    string funcName = ControlUnits.arrCtrlGrpSrmFunc[i][j].Parent.Parent.Parent.Parent.Parent.Text;
                    string grpName = ControlUnits.arrCtrlGrpSrmFunc[i][j].Parent.Parent.Parent.Text;
                    string lblName = ControlUnits.arrCtrlGrpSrmFunc[i][j].Parent.Controls.OfType<LabelControl>().FirstOrDefault()?.Text;
                    string objectType = ControlUnits.arrCtrlGrpSrmFunc[i][j].GetType().ToString();

                    // DataTable에 행 추가
                    DataTableUnits.arrDtSrmFunc[i].Rows.Add(j + 1, pageName, funcName, grpName, lblName, "", objectType);
                }
                ControlUnits.arrGcFunc[i].DataSource = DataTableUnits.arrDtSrmFunc[i];

                // GridView 설정
                gvManager.SetGridView(ControlUnits.arrGvFunc[i]);
            }

            // 콤보박스 타이핑
            for (int i = 0; i < DictionaryUnits.dicGrpSrmFunc.Count; i++)
            {
                int index = i; // 지역 변수로 캡처
                ControlUnits.arrGvFunc[index].FocusedRowChanged += (o, e) =>
                {
                    if (e.FocusedRowHandle >= 0)
                    {
                        int rowIndex = e.FocusedRowHandle; // 선택된 행의 인덱스를 가져옴

                        // 행 인덱스에 맞는 콤보박스를 선택하고 포커스
                        if (rowIndex >= 0 && rowIndex < ControlUnits.arrCtrlGrpSrmFunc[index].Length)
                        {
                            // 콤보박스에 포커스를 주고 선택 상태로 만듭니다.
                            ControlUnits.arrCtrlGrpSrmFunc[index][rowIndex].Focus();
                        }
                    }
                };

                string[] initialValues = new string[ControlUnits.arrCtrlGrpSrmFunc[index].Length];
                // 각 콤보박스와 체크박스에 KeyUp 이벤트 핸들러를 등록
                for (int j = 0; j < ControlUnits.arrCtrlGrpSrmFunc[index].Length; j++)
                {
                    // 지역 변수로 캡처
                    int indexRow = j;

                    // arrCbGrpFunc[index][indexRow]가 ComboBox인지 확인
                    if (ControlUnits.arrCtrlGrpSrmFunc[index][indexRow] is ComboBoxEdit cb)
                    {
                        // 초기값 저장
                        initialValues[indexRow] = cb.Text;

                        // TextChanged 이벤트 핸들러 등록
                        cb.TextChanged += (o, e) =>
                        {
                            // 텍스트가 초기값과 같으면 공란으로 설정
                            DataTableUnits.arrDtSrmFunc[index].Rows[indexRow]["Data"] = cb.Text == initialValues[indexRow] ? "" : cb.Text;
                        };
                    }
                    else if (ControlUnits.arrCtrlGrpSrmFunc[index][indexRow] is CheckEdit ckb)
                    {
                        // 초기값 저장
                        initialValues[indexRow] = ckb.Text;

                        // CheckedChanged 이벤트 핸들러 등록
                        ckb.CheckedChanged += (o, e) =>
                        {
                            // 체크 상태에 따라 Data를 설정
                            DataTableUnits.arrDtSrmFunc[index].Rows[indexRow]["Data"] = ckb.Checked ? ckb.Text : "";
                        };
                    }
                }

            }

            // 각 GroupControl에 Spread Event 적용
            for (int i = 0; i < ControlUnits.arrGrpAll.Length; i++)
            {
                gcManager.SpreadGroupControl(ControlUnits.arrGrpAll[i], Eplan.EplAddin.HMX_MCNS.Properties.Resources.CaretBelow, Eplan.EplAddin.HMX_MCNS.Properties.Resources.CaretRight);
            }

            //Event에 사용하기 위해 static으로 저장
            DictionaryUnits.dicCtrlSrmAll = DictionaryUnits.dicCtrlMod.Concat(DictionaryUnits.dicCtrlFunc).ToDictionary(x => x.Key, x => x.Value);
        }
        private void SetToolTip()
        {
            tip.SetToolTip(lblLogo, "메인 메뉴");
            tip.SetToolTip(picBoxLogo, "메인 메뉴");

            tip.SetToolTip(picBoxItems, "항목 관리");
            tip.SetToolTip(picBoxLoad, "가져오기");
            tip.SetToolTip(picBoxSave, "저장하기");
        }
        private void ControlFormFunction()
        {
            // 초기 설정
            lblLogo.ForeColor = ColorUtility.colors[Ecolor.Logo];
            lblSRM.ForeColor = ColorUtility.colors[Ecolor.HmxGreen];
            // xtraTabControl 헤더 숨기기
            xtraTabControlLarge.ShowTabHeader = DevExpress.Utils.DefaultBoolean.False;
            // 각 Label에 Event 적용
            lblManager.HoverLabel(lblTAPmodel, ColorUtility.colors[Ecolor.HmxGreen]);
            lblManager.HoverLabel(lblTAPfunction, ColorUtility.colors[Ecolor.HmxYellow]);
            lblManager.HoverLabel(lblTAPexport, ColorUtility.colors[Ecolor.HmxBlue]);
            // 기능페이지 패널 기능 추가
            xtcManager.AddPanelToTabPage(xtraTabControlFunction, StringUnits.dArrStrFunc);

            this.FormClosing += (o, e) =>
            {
                // 종료 확인 메시지 표시
                DialogResult result = MessageBox.Show(
                    "정말 종료하시겠습니까?",
                    "종료 확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    FormUnits.formInitialPage.Close();
                }
                else if(result == DialogResult.No)
                {
                    e.Cancel = true; // 종료 취소
                    return;
                }
            };
            
            picBoxLogo.MouseClick += (o, e) =>
            {
                // 메시지 상자를 생성하고 표시합니다.
                DialogResult result = MessageBox.Show(
                    "메인 화면으로 돌아가시겠습니까?\n(작성한 내용은 저장되지 않습니다.)",
                    "확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

                // 메시지 상자의 결과를 확인합니다.
                if (result == DialogResult.Yes)
                {
                    // formConceptSheet가 null 또는 해제된 상태가 아닌지 확인
                    if (FormUnits.formConceptSheet != null && !FormUnits.formConceptSheet.IsDisposed)
                    {
                        FormUnits.formConceptSheet.Dispose();
                    }

                    // formInitialPage가 이미 열려 있는지 확인 후 표시
                    if (FormUnits.formInitialPage == null || FormUnits.formInitialPage.IsDisposed)
                    {
                        FormUnits.formInitialPage.Close();
                        FormUnits.formInitialPage = new FormInitialPage();
                    }

                    FormUnits.formInitialPage.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
                }
            };
            picBoxSave.MouseClick += (o, e) =>
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.InitialDirectory = StringUnits.strXmlFolderPath;
                    saveFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
                    saveFileDialog.Title = "xml 데이터 저장";
                    saveFileDialog.FileName = String.Concat(cbPRJnumber.Text, "_", cbMODfullName.Text); // 기본 파일 이름 설정


                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            StringUnits.strXmlFilePath = saveFileDialog.FileName;

                            XElement root = new XElement("Data",
                                new XElement("Controls", DictionaryUnits.dicCtrlSrmAll.Select(dic =>
                                    new XElement("ControlGroup", new XAttribute("Key", dic.Key), dic.Value.Select((ctrl, index) =>
                                    {
                                        if (ctrl is ComboBoxEdit cb)
                                        {

                                            string text = cb.BackColor == ColorUtility.colors[Ecolor.Active] ? cb.EditValue?.ToString() : null;
                                            string labelText = cb.Parent.Controls.OfType<LabelControl>().FirstOrDefault()?.Text ?? "No Label";

                                            return new XElement("ComboBox",
                                                new XAttribute("Index", index),
                                                new XAttribute("LabelText", labelText),
                                                new XElement("Text", text));
                                        }
                                        else if (ctrl is CheckEdit chk)
                                        {
                                            string labelText = chk.Parent.Controls.OfType<LabelControl>().FirstOrDefault()?.Text ?? "No Label";

                                            return new XElement("CheckBox",
                                                new XAttribute("Index", index),
                                                new XAttribute("LabelText", labelText),
                                                new XElement("Checked", chk.Checked));
                                        }
                                        return null;
                                    }).Where(x => x != null)))),
                                new XElement("DataTable_Lout",
                                DataTableUnits.dtLout != null
                                ? DataTableUnits.dtLout.AsEnumerable().Select(row =>
                                new XElement("Row",
                                row.ItemArray.Select((item, index) =>
                                new XElement($"Column{index}", item?.ToString() ?? string.Empty))))
                                : Enumerable.Empty<XElement>()),

                                new XElement("DataTable_SensorIo",
                                DataTableUnits.dtSensorIo != null
                                ? DataTableUnits.dtSensorIo.AsEnumerable().Select(row =>
                                new XElement("Row",
                                row.ItemArray.Select((item, index) =>
                                new XElement($"Column{index}", item?.ToString() ?? string.Empty))))
                                : Enumerable.Empty<XElement>()),

                                new XElement("DataTable_SensorCopyIo",
                                DataTableUnits.dtSensorCopyIo != null
                                ? DataTableUnits.dtSensorCopyIo.AsEnumerable().Select(row =>
                                new XElement("Row",
                                row.ItemArray.Select((item, index) =>
                                new XElement($"Column{index}", item?.ToString() ?? string.Empty))))
                                : Enumerable.Empty<XElement>()),

                                new XElement("DataTable_logicIo",
                                DataTableUnits.dtLogicIo != null
                                ? DataTableUnits.dtLogicIo.AsEnumerable().Select(row =>
                                new XElement("Row",
                                row.ItemArray.Select((item, index) =>
                                new XElement($"Column{index}", item?.ToString() ?? string.Empty))))
                                : Enumerable.Empty<XElement>()),

                                new XElement("DataTable_UniqueIo",
                                DataTableUnits.dtUniqueIo != null
                                ? DataTableUnits.dtUniqueIo.AsEnumerable().Select(row =>
                                new XElement("Row",
                                row.ItemArray.Select((item, index) =>
                                new XElement($"Column{index}", item?.ToString() ?? string.Empty))))
                                : Enumerable.Empty<XElement>()));

                            root.Save(StringUnits.strXmlFilePath);
                            MessageBox.Show("파일이 저장되었습니다.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"파일 저장 중 오류가 발생했습니다: {ex.Message}");
                        }
                    }
                }
            };
            picBoxLoad.MouseClick += (o, e) =>
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = StringUnits.strXmlFolderPath;
                    openFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
                    openFileDialog.Title = "xml 데이터 로드";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            StringUnits.strXmlFilePath = openFileDialog.FileName;
                            XElement root = XElement.Load(StringUnits.strXmlFilePath);

                            xtraTabControlLarge.SelectedTabPageIndex = 0;

                            XElement mccbElement = null; // Index=18 (MCCB 사양) 저장
                            XElement mainBreakerElement = null; // Index=9 (메인 차단기 사양) 저장

                            // ComboBox와 CheckBox 상태 복원
                            foreach (var group in root.Element("Controls").Elements("ControlGroup"))
                            {
                                string key = group.Attribute("Key").Value;
                                if (DictionaryUnits.dicCtrlSrmAll.TryGetValue(key, out var controls))
                                {
                                    foreach (var controlElement in group.Elements())
                                    {
                                        int index = int.Parse(controlElement.Attribute("Index").Value);
                                        string labelText = controlElement.Attribute("LabelText")?.Value;

                                        // MCCB 사양 (Index=18) 저장
                                        if (index == 14 && labelText == "MCCB 사양")
                                        {
                                            mccbElement = controlElement;
                                            continue;
                                        }
                                        // 메인 차단기 사양 (Index=9) 저장
                                        if (index == 4 && labelText == "메인 차단기 사양")
                                        {
                                            mainBreakerElement = controlElement;
                                            continue;
                                        }

                                        if (index < controls.Length)
                                        {
                                            var control = controls[index];


                                            if (control is ComboBoxEdit cb && controlElement.Name.LocalName == "ComboBox")
                                            {
                                                string text = controlElement.Element("Text")?.Value;
                                                if (!string.IsNullOrEmpty(text))
                                                {
                                                    cb.Text = text;
                                                }
                                            }
                                            else if (control is CheckEdit chk && controlElement.Name.LocalName == "CheckBox")
                                            {
                                                bool isChecked = bool.Parse(controlElement.Element("Checked").Value);
                                                chk.Checked = isChecked;
                                            }
                                        }
                                    }
                                }
                            }
                            // 1. MCCB 사양 (Index=18) 적용
                            if (mccbElement != null)
                            {
                                string key = mccbElement.Parent.Attribute("Key").Value;
                                if (DictionaryUnits.dicCtrlSrmAll.TryGetValue(key, out var controls))
                                {
                                    int index = int.Parse(mccbElement.Attribute("Index").Value);
                                    if (index < controls.Length && controls[index] is ComboBoxEdit cb)
                                    {
                                        string text = mccbElement.Element("Text")?.Value;
                                        if (!string.IsNullOrEmpty(text))
                                        {
                                            cb.Text = text; // 먼저 적용
                                        }
                                    }
                                }
                            }

                            // 2. 메인 차단기 사양 (Index=9) 적용 (가장 마지막에)
                            if (mainBreakerElement != null)
                            {
                                string key = mainBreakerElement.Parent.Attribute("Key").Value;
                                if (DictionaryUnits.dicCtrlSrmAll.TryGetValue(key, out var controls))
                                {
                                    int index = int.Parse(mainBreakerElement.Attribute("Index").Value);
                                    if (index < controls.Length && controls[index] is ComboBoxEdit cb)
                                    {
                                        string text = mainBreakerElement.Element("Text")?.Value;
                                        if (!string.IsNullOrEmpty(text))
                                        {
                                            cb.Text = text; // 마지막에 적용
                                        }
                                    }
                                }
                            }
                            // DataTable_Lout 데이터 로드
                            if (DataTableUnits.dtLout == null)
                            {
                                DataTableUnits.dtLout = new DataTable();
                            }
                            else
                            {
                                DataTableUnits.dtLout.Clear();
                            }

                            foreach (XElement rowElement in root.Element("DataTable_Lout")?.Elements("Row") ?? Enumerable.Empty<XElement>())
                            {
                                DataRow row = DataTableUnits.dtLout.NewRow();
                                foreach (DataColumn col in DataTableUnits.dtLout.Columns)
                                {
                                    string cellValue = rowElement.Element($"Column{DataTableUnits.dtLout.Columns.IndexOf(col)}")?.Value;
                                    row[col.ColumnName] = string.IsNullOrEmpty(cellValue) ? (object)DBNull.Value : cellValue;
                                }
                                DataTableUnits.dtLout.Rows.Add(row);
                            }

                            // sensorIoDt, sensorCopyIoDt 데이터 로드
                            if (DataTableUnits.dtSensorIo == null)
                            {
                                DataTableUnits.dtSensorIo = new DataTable();
                            }
                            else
                            {
                                DataTableUnits.dtSensorIo.Clear();
                            }

                            foreach (XElement rowElement in root.Element("DataTable_SensorIo")?.Elements("Row") ?? Enumerable.Empty<XElement>())
                            {
                                DataRow row = DataTableUnits.dtSensorIo.NewRow();
                                foreach (DataColumn col in DataTableUnits.dtSensorIo.Columns)
                                {
                                    string cellValue = rowElement.Element($"Column{DataTableUnits.dtSensorIo.Columns.IndexOf(col)}")?.Value;
                                    row[col.ColumnName] = string.IsNullOrEmpty(cellValue) ? (object)DBNull.Value : cellValue;
                                }
                                DataTableUnits.dtSensorIo.Rows.Add(row);
                            }

                            if (DataTableUnits.dtSensorCopyIo == null)
                            {
                                DataTableUnits.dtSensorCopyIo = new DataTable();
                            }
                            else
                            {
                                DataTableUnits.dtSensorCopyIo.Clear();
                            }

                            foreach (XElement rowElement in root.Element("DataTable_SensorCopyIo")?.Elements("Row") ?? Enumerable.Empty<XElement>())
                            {
                                DataRow row = DataTableUnits.dtSensorCopyIo.NewRow();
                                foreach (DataColumn col in DataTableUnits.dtSensorCopyIo.Columns)
                                {
                                    string cellValue = rowElement.Element($"Column{DataTableUnits.dtSensorCopyIo.Columns.IndexOf(col)}")?.Value;
                                    row[col.ColumnName] = string.IsNullOrEmpty(cellValue) ? (object)DBNull.Value : cellValue;
                                }
                                DataTableUnits.dtSensorCopyIo.Rows.Add(row);
                            }

                            // dtUniqueIo 데이터 로드
                            if (DataTableUnits.dtLogicIo == null)
                            {
                                DataTableUnits.dtLogicIo = new DataTable();
                            }
                            else
                            {
                                DataTableUnits.dtLogicIo.Clear();
                            }

                            foreach (XElement rowElement in root.Element("DataTable_logicIo")?.Elements("Row") ?? Enumerable.Empty<XElement>())
                            {
                                DataRow row = DataTableUnits.dtLogicIo.NewRow();
                                foreach (DataColumn col in DataTableUnits.dtLogicIo.Columns)
                                {
                                    string cellValue = rowElement.Element($"Column{DataTableUnits.dtLogicIo.Columns.IndexOf(col)}")?.Value;
                                    row[col.ColumnName] = string.IsNullOrEmpty(cellValue) ? (object)DBNull.Value : cellValue;
                                }
                                DataTableUnits.dtLogicIo.Rows.Add(row);
                            }

                            if (DataTableUnits.dtUniqueIo == null)
                            {
                                DataTableUnits.dtUniqueIo = new DataTable();
                            }
                            else
                            {
                                DataTableUnits.dtUniqueIo.Clear();
                            }

                            foreach (XElement rowElement in root.Element("DataTable_UniqueIo")?.Elements("Row") ?? Enumerable.Empty<XElement>())
                            {
                                DataRow row = DataTableUnits.dtUniqueIo.NewRow();
                                foreach (DataColumn col in DataTableUnits.dtUniqueIo.Columns)
                                {
                                    string cellValue = rowElement.Element($"Column{DataTableUnits.dtUniqueIo.Columns.IndexOf(col)}")?.Value;
                                    row[col.ColumnName] = string.IsNullOrEmpty(cellValue) ? (object)DBNull.Value : cellValue;
                                }
                                DataTableUnits.dtUniqueIo.Rows.Add(row);
                            }

                            MessageBox.Show("파일이 성공적으로 로드되었습니다.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"파일 로드 중 오류가 발생했습니다: {ex.Message}");
                        }
                    }
                }
            };
            picBoxItems.MouseClick += (o, e) =>
            {
                FormUnits.formItemsList = new FormItemsList();
                FormUnits.formItemsList.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
            };
            picBoxPDFexport.MouseClick += (o, e) =>
            {
                // SaveFileDialog 생성
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf";  // PDF 파일 형식 필터
                    saveFileDialog.Title = "PDF 파일로 저장";
                    saveFileDialog.FileName = String.Concat(cbPRJnumber.Text, "_", cbMODfullName.Text,"_엔지니어링시트"); // 기본 파일 이름 설정

                    // 대화상자가 열리고 사용자가 경로를 선택하면
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            // 선택한 경로에 PDF 저장
                            string filePath = saveFileDialog.FileName;
                            FormUnits.xtraReport1.ExportToPdf(filePath);

                            // PDF가 성공적으로 저장되면 메시지를 표시
                            MessageBox.Show("PDF로 저장되었습니다: " + filePath, "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            // 오류가 발생했을 경우 사용자에게 알림
                            MessageBox.Show("PDF 저장 중 오류가 발생했습니다: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            };
            //Label Action
            lblTAPmodel.MouseClick += (e, o) => { xtraTabPageLargeModel.Show(); lblSRM.ForeColor = ColorUtility.colors[Ecolor.HmxGreen]; };
            lblTAPfunction.MouseClick += (e, o) => { xtraTabPageLargeFunction.Show(); lblSRM.ForeColor = ColorUtility.colors[Ecolor.HmxYellow]; };
            lblTAPexport.MouseClick += (e, o) =>
            {
                xtraTabPageLargeGenerating.Show();
                lblSRM.ForeColor = ColorUtility.colors[Ecolor.HmxBlue];

                // XtraReport1 인스턴스 생성
                FormUnits.xtraReport1 = new XtraReport1
                {
                    PaperKind = System.Drawing.Printing.PaperKind.A4, // A4 크기 설정
                    Margins = new System.Drawing.Printing.Margins(90, 90, 0, 0), // 마진 설정
                };

                // 문서 생성 및 페이지 너비에 맞추기
                FormUnits.xtraReport1.CreateDocument();
                FormUnits.xtraReport1.PrintingSystem.Document.AutoFitToPagesWidth = 1;

                // DocumentViewer에 바인딩
                documentViewer1.DocumentSource = FormUnits.xtraReport1;
            };
            lblLogo.MouseClick += (o, e) =>
            {
                // 메시지 상자를 생성하고 표시합니다.
                DialogResult result = MessageBox.Show(
                    "메인 화면으로 돌아가시겠습니까?\n(작성한 내용은 저장되지 않습니다.)",
                    "확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

                // 메시지 상자의 결과를 확인합니다.
                if (result == DialogResult.Yes)
                {
                    // formConceptSheet가 null 또는 해제된 상태가 아닌지 확인
                    if (FormUnits.formConceptSheet != null && !FormUnits.formConceptSheet.IsDisposed)
                    {
                        FormUnits.formConceptSheet.Dispose();
                    }

                    // formInitialPage가 이미 열려 있는지 확인 후 표시
                    if (FormUnits.formInitialPage == null || FormUnits.formInitialPage.IsDisposed)
                    {
                        FormUnits.formInitialPage = new FormInitialPage();
                    }

                    FormUnits.formInitialPage.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
                }
            };

        }
        private void ControlPlcFunction()
        {
            Dictionary<string, List<string>> funcSensorList = new Dictionary<string, List<string>>()
                {
                    { "ELEQ", new List<string>() },
                    { "LIFT", new List<string>() },
                    { "TRAV", new List<string>() },
                    { "TRAV2", new List<string>() },
                    { "FORK", new List<string>() },
                    { "FORK2", new List<string>() },
                    { "CARR", new List<string>() },
                    { "ETC", new List<string>() },
                };
            cbMSPcontrollerSpec.TextChanged += (o, e) =>
            {
                LoadIoFromExcel();
                ClearSensorList();
                UpdateMatching();
                UpdateSensorList();
                UpdateGridControl();
                UpdateComboBox();
                UpdatePlcCard();
            };

            cbMODfullName.VisibleChanged += (o, e) =>
            {
                if (!cbMODfullName.Visible)
                {
                    gridControl1.DataSource = null;
                }
            };

            cbMODfullName.TextChanged += (o, e) =>
            {

                LoadIoFromExcel();
                ClearSensorList();
                UpdateMatching();
                UpdateSensorList();
                UpdateGridControl();
                UpdateComboBox();
                UpdatePlcCard();
            };
            cbEleqLocation.TextChanged += (o, e) =>
            {
                // cbEleqLocation의 선택 항목을 가져옴
                string selectedLocation = cbEleqLocation.SelectedItem?.ToString();
                cbEleqDt.Properties.Items.Clear(); // cbEleqDt의 기존 항목 제거

                if (!string.IsNullOrEmpty(selectedLocation))
                {
                    try
                    {
                        // 선택한 LOCATION 값에 따른 DT 값 중 최대값 가져오기
                        var maxDtRow = DataTableUnits.dtUniqueIo.AsEnumerable()
                            .Where(row => row.Field<string>("LOCATION") == selectedLocation)
                            .OrderByDescending(row => row.Field<string>("DT"))
                            .FirstOrDefault();

                        if (maxDtRow != null)
                        {
                            // 최대 DT 값에서 "KE" 뒤의 숫자 추출하여 다음 값 설정
                            string maxDtValue = maxDtRow.Field<string>("DT");
                            if (int.TryParse(maxDtValue.Replace("KE", ""), out int maxNumber))
                            {
                                string nextValue = "KE" + (maxNumber + 1); // 다음 값 설정
                                cbEleqDt.Text = nextValue; // cbEleqDt에 설정
                                cbEleqDt.BackColor = ColorUtility.colors[Ecolor.Active];
                                cbEleqDt.ForeColor = ColorUtility.colors[Ecolor.TextBlack];
                            }
                        }
                    }
                    catch (Exception)
                    {
                        return;
                    }

                }

            };
            cbEleqType.TextChanged += UpdatePartsItems;
            cbEleqPoint.TextChanged += UpdatePartsItems;
            ckbRegenerativeUnitTrue.CheckedChanged += (o, e) =>
            {
                UpdateMatching();
                UpdateSensorList();
                UpdateGridControl();
                UpdateComboBox();
                UpdatePlcCard();
            };
            ckbFork1PosTrue.CheckedChanged += (o, e) =>
            {
                UpdateMatching();
                UpdateSensorList();
                UpdateGridControl();
                UpdateComboBox();
                UpdatePlcCard();
            };
            ckbFork2PosTrue.CheckedChanged += (o, e) =>
            {
                UpdateMatching();
                UpdateSensorList();
                UpdateGridControl();
                UpdateComboBox();
                UpdatePlcCard();
            };
            ckbCarrDoubleCarriageGOXS.CheckedChanged += (o, e) =>
            {
                UpdateMatching();
                UpdateSensorList();
                UpdateGridControl();
                UpdateComboBox();
                UpdatePlcCard();
            };
            ckbCarrDoubleCarriageGOXM.CheckedChanged += (o, e) =>
            {
                UpdateMatching();
                UpdateSensorList();
                UpdateGridControl();
                UpdateComboBox();
                UpdatePlcCard();
            };
            ckbCarrDoubleCarriageGOXH.CheckedChanged += (o, e) =>
            {
                UpdateMatching();
                UpdateSensorList();
                UpdateGridControl();
                UpdateComboBox();
                UpdatePlcCard();
            };
            btnEleqPlcPlus.Click += (o, e) =>
            {
                if (cbEleqLocation.BackColor != ColorUtility.colors[Ecolor.Active] || cbEleqType.BackColor != ColorUtility.colors[Ecolor.Active] || cbEleqPoint.BackColor != ColorUtility.colors[Ecolor.Active] || cbEleqParts.BackColor != ColorUtility.colors[Ecolor.Active] || cbEleqDt.BackColor != ColorUtility.colors[Ecolor.Active]
                || cbEleqLocation.Text == "" || cbEleqType.Text == "" || cbEleqPoint.Text == "" || cbEleqParts.Text == "" || cbEleqDt.Text == "")
                {
                    MessageBox.Show("모델 및 컨트롤러 사양을 확정하세요.");
                    return;
                }
                // CS_StaticSensor.dtUniqueIo의 DataTable 객체가 존재하는지 확인
                if (DataTableUnits.dtUniqueIo != null)
                {
                    // NO 컬럼의 최대값 구하기
                    int maxNo = 0;
                    if (DataTableUnits.dtUniqueIo.Rows.Count > 0)
                    {
                        maxNo = DataTableUnits.dtUniqueIo.AsEnumerable()
                                                          .Max(row => row.Field<int>("NO"));
                    }

                    // 새 행 생성
                    DataRow newRow = DataTableUnits.dtUniqueIo.NewRow();

                    // NO 컬럼에 최대값 + 1 설정
                    newRow["NO"] = maxNo + 1;

                    // 기타 열에 기본 값 설정 (필요에 따라 값 설정)
                    newRow["LOCATION"] = cbEleqLocation.Text;
                    newRow["TYPE1"] = cbEleqType.Text;
                    newRow["POINT"] = cbEleqPoint.Text;
                    newRow["PARTS"] = cbEleqParts.Text;
                    newRow["DT"] = cbEleqDt.Text;
                    newRow["IFB1"] = false;
                    newRow["IFB2"] = false;
                    newRow["IFB3"] = false;
                    newRow["IFB4"] = false;
                    newRow["8BIT"] = false;


                    // DataTable에 새 행 추가
                    DataTableUnits.dtUniqueIo.Rows.Add(newRow);
                    gridControl1.DataSource = DataTableUnits.dtUniqueIo;

                }

                if (DataTableUnits.dtSensorIo != null)
                {
                    // cbEleqPoint.Text가 숫자인지 확인하고, 숫자라면 그 값을 사용하여 행 추가
                    if (int.TryParse(cbEleqPoint.Text, out int numberOfRowsToAdd))
                    {
                        for (int i = 0; i < numberOfRowsToAdd; i++)
                        {
                            // 새 행 생성
                            DataRow newRow = DataTableUnits.dtSensorIo.NewRow();

                            // 기타 열에 기본 값 설정
                            newRow["LOCATION"] = cbEleqLocation.Text;
                            newRow["TYPE1"] = cbEleqType.Text;
                            newRow["TYPE2"] = "DI1";
                            newRow["PARTS"] = cbEleqParts.Text;
                            newRow["DT"] = cbEleqDt.Text;
                            newRow["기능"] = "";
                            newRow["SIGNAL"] = "";
                            newRow["DESCRIPTION"] = "";

                            // DataTable에 새 행 추가
                            DataTableUnits.dtSensorIo.Rows.Add(newRow);
                        }
                    }
                    else
                    {
                        // cbEleqPoint.Text가 숫자가 아닌 경우, 오류 메시지 표시
                        MessageBox.Show("cbEleqPoint.Text에 유효한 숫자를 입력하세요.");
                    }
                }

                // cbEleqDt.Text에서 숫자 부분을 증가시키는 로직을 마지막에 수행
                string prefix = cbEleqDt.Text.Substring(0, 2);
                int numberPart;
                if (int.TryParse(cbEleqDt.Text.Substring(2), out numberPart))
                {
                    // 숫자 부분을 증가시킴
                    numberPart++;

                    // 새로운 DT 값 생성
                    string newDt = prefix + numberPart.ToString();

                    // cbEleqDt.Text 값 업데이트
                    cbEleqDt.Text = newDt;
                }
            };
            btnEleqPlcReset.Click += (o, e) =>
            {
                LoadIoFromExcel();
                ClearSensorList();
                UpdateMatching();
                UpdateSensorList();
                UpdateGridControl();
                UpdateComboBox();
                UpdatePlcCard();
            };


            gridView1.CustomRowCellEdit += (o, e) =>
            {
                if (e.Column.FieldName == "IFB1" || e.Column.FieldName == "IFB3")
                {

                    bool bitValue = Convert.ToBoolean(gridView1.GetRowCellValue(e.RowHandle, "8BIT"));

                    if (bitValue == true)
                    {
                        // 현재 셀 값이 체크(true) 상태인지 확인
                        bool isChecked = Convert.ToBoolean(gridView1.GetRowCellValue(e.RowHandle, e.Column.FieldName));

                        if (isChecked)
                        {
                            // 체크된 상태를 해제(false)로 변경
                            gridView1.SetRowCellValue(e.RowHandle, e.Column.FieldName, false);
                        }
                        RepositoryItemCheckEdit checkEdit = new RepositoryItemCheckEdit();
                        checkEdit.ValueChecked = true;
                        checkEdit.ValueUnchecked = false;
                        checkEdit.ReadOnly = true;
                        e.RepositoryItem = checkEdit;

                    }

                }
            };
            gridView1.RowCellStyle += (o, e) =>
            {
                // IFB1, IFB2, IFB3, IFB4 열에 대해 스타일 변경
                if (e.Column.FieldName == "IFB1" || e.Column.FieldName == "IFB2" || e.Column.FieldName == "IFB3" || e.Column.FieldName == "IFB4" || e.Column.FieldName == "8BIT")
                {
                    bool isChecked = Convert.ToBoolean(gridView1.GetRowCellValue(e.RowHandle, e.Column.FieldName));

                    // 체크되지 않으면 빨간색 배경으로 설정
                    if (!isChecked)
                    {
                        e.Appearance.BackColor = Color.Red;
                        e.Appearance.ForeColor = Color.White; // 글자색 흰색으로 설정
                    }

                }

            };

            btnPlcDetailSheet.MouseClick += (o, e) =>
            {
                if (cbMODfullName.BackColor != ColorUtility.colors[Ecolor.Active])
                {
                    MessageBox.Show("모델명이 확정되지 않았습니다.");
                    return;
                }

                try
                {
                    // Excel 파일을 열고 입력된 텍스트에 맞는 워크시트 로드
                    using (var workbook = new XLWorkbook(StringUnits.strIoListFilePath))
                    {
                        if (!workbook.Worksheets.Contains(cbMSPcontrollerSpec.Text))
                        {
                            MessageBox.Show("컨트롤러 명칭의 IO 템플릿 워크시트가 없습니다.");
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 파일을 읽을 수 없는 경우 예외 처리
                    MessageBox.Show($"IO 템플릿을 읽는 중 오류 발생: {ex.Message}");
                    return;
                }

                if (DataTableUnits.dtSensorCopyIo == DataTableUnits.dtSensorIo)
                {
                    return;
                }
                else
                {
                    DataTableUnits.dtSensorCopyIo = DataTableUnits.dtSensorIo.Copy();
                }

                FormUnits.formIoList = new FormIoList();
                FormUnits.formIoList.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
            };

            void UpdateSensorList()
            {
                if (cbMODname.BackColor != ColorUtility.colors[Ecolor.Active] || cbMODheight.BackColor != ColorUtility.colors[Ecolor.Active] || cbMODweight.BackColor != Color.White ||
                     cbMODname.Text == "" || cbMODheight.Text == "" || cbMODweight.Text == "")
                {
                    ClearSensorList();
                }

                foreach (var key in funcSensorList.Keys.ToList())
                {
                    if (funcSensorList[key] != null)
                    {
                        funcSensorList[key] = funcSensorList[key].Distinct().ToList();
                    }
                }

                UpdateSensorListToLabel(funcSensorList, "LIFT", fpnl1, new Font("맑은 고딕", 8, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "LIFT", fpnlLiftSensorList, new Font("맑은 고딕", 9, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "TRAV", fpnl2, new Font("맑은 고딕", 8, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "TRAV", fpnlTrav1SensorList, new Font("맑은 고딕", 9, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "TRAV2", fpnl6, new Font("맑은 고딕", 8, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "TRAV2", fpnlTrav2SensorList, new Font("맑은 고딕", 9, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "FORK", fpnl3, new Font("맑은 고딕", 8, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "FORK", fpnlFork1SensorList, new Font("맑은 고딕", 9, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "FORK2", fpnl4, new Font("맑은 고딕", 8, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "FORK2", fpnlFork2SensorList, new Font("맑은 고딕", 9, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "CARR", fpnl5, new Font("맑은 고딕", 8, FontStyle.Regular));
                UpdateSensorListToLabel(funcSensorList, "CARR", fpnlCarrSensorList, new Font("맑은 고딕", 9, FontStyle.Regular));

                StringUnits.strLiftSensor = UpdateSensorListToString(funcSensorList, "LIFT");
                StringUnits.strTrav1Sensor = UpdateSensorListToString(funcSensorList, "TRAV");
                StringUnits.strTrav2Sensor = UpdateSensorListToString(funcSensorList, "TRAV2");
                StringUnits.strFork1Sensor = UpdateSensorListToString(funcSensorList, "FORK");
                StringUnits.strFork2Sensor = UpdateSensorListToString(funcSensorList, "FORK2");
                StringUnits.strCarrSensor = UpdateSensorListToString(funcSensorList, "CARR");

                DictionaryUnits.dicfuncSensor = funcSensorList;
            }
            void UpdateSensorListToLabel(Dictionary<string, List<string>> dict, string function, Panel pnl, Font font)
            {


                if (dict.TryGetValue(function, out var values))
                {
                    // fpnl1의 기존 라벨을 초기화
                    pnl.Controls.Clear();

                    // LIFT 값들을 라벨로 추가
                    foreach (var signal in values)
                    {
                        LabelControl label = new LabelControl
                        {
                            Text = signal, // 신호 텍스트 설정
                            AutoSize = true,    // 라벨 크기를 텍스트에 맞춤
                            ForeColor = Color.Blue, // 라벨 텍스트 색상 설정
                            Font = font,
                        };
                        pnl.Controls.Add(label); // fpnl1에 라벨 추가
                    }
                }
            }
            string UpdateSensorListToString(Dictionary<string, List<string>> dict, string function)
            {
                if (dict.TryGetValue(function, out var liftValues))
                {
                    return string.Join(", ", liftValues);
                }
                else
                {
                    return string.Empty;
                }
            }
            void UpdatePartsItems(object sender, EventArgs e)
            {
                // dtUniqueIo가 null인지 확인하고, null인 경우 메서드를 종료
                if (DataTableUnits.dtUniqueIo == null)
                    return;

                string selectedType = cbEleqType.Text;
                string selectedPoint = cbEleqPoint.Text;

                // cbEleqParts 초기화
                cbEleqParts.Properties.Items.Clear();

                // TYPE 및 POINT 값과 일치하는 PARTS 항목 필터링 및 중복 제거
                var filteredParts = DataTableUnits.dtUniqueIo.AsEnumerable()
                    .Where(row => row.Field<string>("TYPE1") == selectedType &&
                                  row.Field<int>("POINT").ToString() == selectedPoint)
                    .Select(row => row.Field<string>("PARTS"))
                    .Distinct();

                // 필터링된 PARTS 항목을 cbEleqParts에 추가
                cbEleqParts.Properties.Items.AddRange(filteredParts.ToArray());
            }
            void UpdateGridControl()
            {
                try
                {
                    // 초기화
                    gridControl1.DataSource = null;

                    // DataTable 확인 및 복사
                    if (DataTableUnits.dtSensorIo == null)
                        return;

                    DataTable copyDt = DataTableUnits.dtSensorIo.Copy();

                    // 고유 데이터 테이블 생성
                    DataTableUnits.dtLogicIo = new DataView(copyDt).ToTable(true, "LOCATION", "TYPE1", "TYPE2", "PARTS", "DT");
                    // NO 열 추가 및 순번 할당
                    DataTableUnits.dtLogicIo.Columns.Add("NO", typeof(int)).SetOrdinal(0);
                    for (int i = 0; i < DataTableUnits.dtLogicIo.Rows.Count; i++)
                        DataTableUnits.dtLogicIo.Rows[i]["NO"] = i + 1;

                    // 추가 열 생성
                    DataTableUnits.dtLogicIo.Columns.Add("POINT", typeof(int)).SetOrdinal(4);
                    string[] boolColumns = { "IFB1", "IFB2", "IFB3", "IFB4", "8BIT" };
                    foreach (string colName in boolColumns)
                        DataTableUnits.dtLogicIo.Columns.Add(colName, typeof(bool));

                    // 기본 값 설정
                    foreach (DataRow row in DataTableUnits.dtLogicIo.Rows)
                    {
                        foreach (string colName in boolColumns)
                            row[colName] = true;
                    }

                    foreach (DataRow row in DataTableUnits.dtLogicIo.Rows)
                    {
                        string type2Value = row.Field<string>("TYPE2");
                        string type1Value = row.Field<string>("TYPE1");
                        string locationValue = row.Field<string>("LOCATION");
                        string dtValue = row.Field<string>("DT");

                        row["POINT"] = copyDt.AsEnumerable()
                                             .Count(r => r.Field<string>("LOCATION") == locationValue &&
                                                         r.Field<string>("DT") == dtValue);
                        if (type2Value != null)
                        {
                            // TYPE2 값에 'DI'가 포함된 경우
                            if (type2Value.Contains("DI"))
                            {
                                row["IFB3"] = false;
                                row["IFB4"] = false;
                            }

                            // TYPE2 값에 'DO'가 포함된 경우
                            if (type2Value.Contains("DO"))
                            {
                                row["IFB1"] = false;
                                row["IFB2"] = false;
                            }
                        }
                        // POINT 값이 32인 경우 모든 IFB 값을 true로 설정
                        int? pointValue = row.Field<int?>("POINT");
                        if (pointValue.HasValue && pointValue.Value == 32)
                        {
                            row["IFB1"] = true;
                            row["IFB2"] = true;
                            row["IFB3"] = true;
                            row["IFB4"] = true;
                        }

                        if (type1Value != null)
                        {
                            // TYPE2 값에 'DI'가 포함된 경우
                            if (type1Value.Contains("DIO"))
                            {
                                row["IFB1"] = false;
                                row["IFB2"] = false;
                                row["IFB3"] = false;
                                row["IFB4"] = false;
                                                        }
                        }
                    }



                    
                    


                    DataTableUnits.dtUniqueIo = new DataView(copyDt).ToTable(true, "LOCATION", "TYPE1", "PARTS", "DT");
                    DataTableUnits.dtUniqueIo.Columns.Add("NO", typeof(int)).SetOrdinal(0);
                    for (int i = 0; i < DataTableUnits.dtUniqueIo.Rows.Count; i++)
                        DataTableUnits.dtUniqueIo.Rows[i]["NO"] = i + 1;

                    // 추가 열 생성
                    DataTableUnits.dtUniqueIo.Columns.Add("POINT", typeof(int)).SetOrdinal(4);
                    // POINT 값 설정
                    foreach (DataRow row in DataTableUnits.dtUniqueIo.Rows)
                    {
                        string locationValue = row.Field<string>("LOCATION");
                        string dtValue = row.Field<string>("DT");

                        row["POINT"] = copyDt.AsEnumerable()
                                             .Count(r => r.Field<string>("LOCATION") == locationValue &&
                                                         r.Field<string>("DT") == dtValue);
                    }

                    string[] boolColumns1 = { "IFB1", "IFB2", "IFB3", "IFB4", "8BIT" };
                    foreach (string colName in boolColumns1)
                        DataTableUnits.dtUniqueIo.Columns.Add(colName, typeof(bool));

                    // 기본 값 설정
                    foreach (DataRow row in DataTableUnits.dtUniqueIo.Rows)
                    {
                        foreach (string colName in boolColumns1)
                            row[colName] = false;
                    }



                    foreach (DataRow logicRow in DataTableUnits.dtLogicIo.Rows)
                    {
                        // LOCATION과 DT 값을 확인
                        string logicLocation = logicRow.Field<string>("LOCATION");
                        string logicDt = logicRow.Field<string>("DT");

                        bool logicIfb1 = logicRow.Field<bool>("IFB1");
                        bool logicIfb2 = logicRow.Field<bool>("IFB2");
                        bool logicIfb3 = logicRow.Field<bool>("IFB3");
                        bool logicIfb4 = logicRow.Field<bool>("IFB4");
                        bool logic8bit = logicRow.Field<bool>("8BIT");


                        foreach (DataRow uniqueRow in DataTableUnits.dtUniqueIo.Rows)
                        {
                            string uniqueLocation = uniqueRow.Field<string>("LOCATION");
                            string uniqueDt = uniqueRow.Field<string>("DT");

                            // LOCATION과 DT가 일치하면 IFB1부터 IFB4까지 값을 설정
                            if (logicLocation == uniqueLocation && logicDt == uniqueDt)
                            {
                                // IFB1부터 IFB4까지 FALSE인 경우, uniqueRow의 값을 FALSE로 설정
                                if (logicIfb1)
                                    uniqueRow["IFB1"] = true;
                                if (logicIfb2)
                                    uniqueRow["IFB2"] = true;
                                if (logicIfb3)
                                    uniqueRow["IFB3"] = true;
                                if (logicIfb4)
                                    uniqueRow["IFB4"] = true;
                                if (logic8bit)
                                    uniqueRow["8BIT"] = false;
                            }
                        }
                    }

                    //test
                    gridControl1.DataSource = DataTableUnits.dtUniqueIo;

                    // GridView 설정
                    ConfigureGridView();


                }
                catch (Exception ex)
                {
                    // 예외 처리 (로그나 사용자 알림 등 추가 가능)
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }

            void ConfigureGridView()
            {
                // GridView 옵션 설정
                gridView1.OptionsView.ShowIndicator = false;
                gridView1.OptionsView.ShowGroupPanel = false;
                gridView1.OptionsView.ShowAutoFilterRow = false;
                gridView1.OptionsCustomization.AllowGroup = false;
                gridView1.OptionsCustomization.AllowFilter = false;
                gridView1.OptionsCustomization.AllowColumnMoving = false;
                gridView1.OptionsCustomization.AllowSort = false;
                gridView1.OptionsBehavior.Editable = true;

                // 컬럼 설정
                for (int i = 0; i < 6; i++)
                {
                    gridView1.Columns[i].OptionsColumn.AllowEdit = false;
                    gridView1.Columns[i].BestFit();
                }
                for (int i = 6; i < 11; i++)
                {
                    gridView1.Columns[i].ColumnEdit = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
                    gridView1.Columns[i].OptionsColumn.AllowEdit = true;
                    gridView1.Columns[i].Width = 30;
                }


            }

            void UpdateComboBox()
            {
                // ComboBox 초기화
                cbEleqLocation.Properties.Items.Clear();
                cbEleqType.Properties.Items.Clear();
                cbEleqDt.Properties.Items.Clear();
                cbEleqParts.Properties.Items.Clear();
                cbEleqPoint.Properties.Items.Clear();

                cbEleqLocation.SelectedIndex = -1;
                cbEleqType.SelectedIndex = -1;
                cbEleqDt.SelectedIndex = -1;
                cbEleqParts.SelectedIndex = -1;
                cbEleqPoint.SelectedIndex = -1;

                cbManager.SettingComboBox(cbEleqLocation, "-", false);
                cbManager.SettingComboBox(cbEleqType, "-", false);
                cbManager.SettingComboBox(cbEleqDt, "-", false);
                cbManager.SettingComboBox(cbEleqParts, "-", false);
                cbManager.SettingComboBox(cbEleqPoint, "-", false);

                // 기본 LOCATION, TYPE, POINT 값 추가
                cbEleqLocation.Properties.Items.AddRange(new string[] { "MP", "SB" });
                cbEleqType.Properties.Items.AddRange(new string[] { "DIO", "RIO" });
                cbEleqPoint.Properties.Items.AddRange(new string[] { "16", "32" });

            }
            void ClearSensorList()
            {
                // 리스트 클리어
                foreach (var key in funcSensorList.Keys.ToList())
                {
                    funcSensorList[key].Clear(); // 리스트를 비움
                                                 // 여기서 필요한 경우 새로운 항목을 추가하세요.
                }
            }
            void UpdateMatching()
            {
                try
                {

                    // 사용자가 제공한 텍스트 (예시 텍스트로 초기화)
                    string userText = cbMODfullName.Text; // 실제 입력값으로 변경
                    string[] userTextParts = userText.Split('-'); // 텍스트를 '-'로 분리하여 배열로 저장

                    DataTable tempDt = DataTableUnits.dtExcelIo.Copy();

                    // 공통 시그널 항목 추가
                    foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "공통"))
                    {
                        var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                        var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                        // 해당 기능의 리스트에 시그널 값 추가
                        if (funcSensorList.ContainsKey(funcValue))
                        {
                            funcSensorList[funcValue].Add(signalValue);
                        }
                    }

                    // 공통 시그널 항목 추가
                    if (ckbFork1PosTrue.Checked)
                    {
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "3POS_1"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능의 리스트에 시그널 값 추가
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Add(signalValue);
                            }
                        }
                    }
                    else
                    {
                        // "3POS_1"에 해당하는 항목만 제거
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "3POS_1"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능에서 "3POS_1" 시그널 값만 제거
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Remove(signalValue);
                            }
                        }

                    }

                    // 공통 시그널 항목 추가
                    if (ckbFork2PosTrue.Checked)
                    {
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "3POS_2"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능의 리스트에 시그널 값 추가
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Add(signalValue);
                            }
                        }
                    }
                    else
                    {
                        // "3POS_1"에 해당하는 항목만 제거
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "3POS_2"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능에서 "3POS_1" 시그널 값만 제거
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Remove(signalValue);
                            }
                        }

                    }

                    if (ckbRegenerativeUnitTrue.Checked)
                    {
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "회생 유닛"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능의 리스트에 시그널 값 추가
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Add(signalValue);
                            }
                        }
                    }
                    else
                    {
                        // "3POS_1"에 해당하는 항목만 제거
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "회생 유닛"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능에서 "3POS_1" 시그널 값만 제거
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Remove(signalValue);
                            }
                        }

                    }

                    if (ckbCarrDoubleCarriageGOXS.Checked)
                    {
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "GOXS"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능의 리스트에 시그널 값 추가
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Add(signalValue);
                            }
                        }
                    }
                    else
                    {
                        // "3POS_1"에 해당하는 항목만 제거
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "GOXS"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능에서 "3POS_1" 시그널 값만 제거
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Remove(signalValue);
                            }
                        }

                    }
                    if (ckbCarrDoubleCarriageGOXM.Checked)
                    {
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "GOXM"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능의 리스트에 시그널 값 추가
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Add(signalValue);
                            }
                        }
                    }
                    else
                    {
                        // "3POS_1"에 해당하는 항목만 제거
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "GOXM"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능에서 "3POS_1" 시그널 값만 제거
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Remove(signalValue);
                            }
                        }

                    }
                    if (ckbCarrDoubleCarriageGOXH.Checked)
                    {
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "GOXH"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능의 리스트에 시그널 값 추가
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Add(signalValue);
                            }
                        }
                    }
                    else
                    {
                        // "3POS_1"에 해당하는 항목만 제거
                        foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "GOXH"))
                        {
                            var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                            var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값

                            // 해당 기능에서 "3POS_1" 시그널 값만 제거
                            if (funcSensorList.ContainsKey(funcValue))
                            {
                                funcSensorList[funcValue].Remove(signalValue);
                            }
                        }

                    }

                    foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "미사용"))
                    {
                        var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값
                        var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값


                        // SIGNAL과 DESCRIPTION 값 삭제
                        row.SetField<string>("SIGNAL", null); // 또는 빈 문자열 ""
                        row.SetField<string>("DESCRIPTION", null); // 또는 빈 문자열 ""
                        row.SetField<string>("타입", null);
                    }

                    // dt 전체에서 "모델명" 구분의 행을 찾고 시그널 항목 추가
                    foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "모델명"))
                    {
                        var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값 저장
                        var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값 저장
                        var includeCondition = row.Field<string>("포함조건"); // 포함조건 열의 값 저장
                        var exceptCondition = row.Field<string>("제외조건"); // 제외조건 열의 값 저장

                        // 기능이나 시그널이 공란이면 다음 반복으로 넘어가기
                        if (string.IsNullOrWhiteSpace(funcValue) || string.IsNullOrWhiteSpace(signalValue))
                        {
                            continue; // 공란인 경우 이 반복을 건너뜁니다.
                        }

                        // 포함조건 배열로 나누기: 쉼표로 구분하여 배열 생성, 공백 요소는 제외
                        var includeConditions = includeCondition.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                        // 제외조건 배열로 나누기: 쉼표로 구분하여 배열 생성, 공백 요소는 제외
                        var exceptConditions = exceptCondition.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                        // 포함조건 체크: userTextParts[0]에 포함조건이 존재하거나, 포함조건이 비어있을 경우
                        bool isIncluded = includeConditions.All(condition => userTextParts[0].Contains(condition.Trim())) || string.IsNullOrWhiteSpace(includeCondition);

                        // 제외조건 체크: userTextParts[0]에 제외조건이 존재하는지 확인
                        bool isExcluded = exceptConditions.Any(condition => userTextParts[0].Contains(condition.Trim()));

                        if (!funcSensorList.ContainsKey(funcValue))
                        {
                            MessageBox.Show(string.Concat(funcValue, " : 해당 기능은 존재하지 않습니다. \n엑셀 IO 템플릿에서 해당 컨트롤러의 기능을 수정해주세요."), "경고", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // 포함조건이 만족되면 해당 기능의 시그널을 리스트에 추가
                        if (isIncluded)
                        {
                            funcSensorList[funcValue].Add(signalValue); // 조건을 만족하면 시그널 추가
                        }

                        // 제외조건이 만족되면 해당 기능의 시그널을 리스트에서 제거
                        if (isExcluded)
                        {
                            funcSensorList[funcValue].Remove(signalValue); // 조건을 만족하면 시그널 제거
                        }


                    }

                    // dt 전체에서 "옵션" 구분의 행을 찾고 시그널 항목 추가
                    foreach (var row in tempDt.AsEnumerable().Where(row => row.Field<string>("구분") == "옵션"))
                    {
                        var funcValue = row.Field<string>("기능"); // 기능 컬럼의 값 저장
                        var signalValue = row.Field<string>("SIGNAL"); // 시그널 컬럼의 값 저장
                        var includeCondition = row.Field<string>("포함조건"); // 포함조건 열의 값 저장
                        var exceptCondition = row.Field<string>("제외조건"); // 제외조건 열의 값 저장

                        // 기능이나 시그널이 공란이면 다음 반복으로 넘어가기
                        if (string.IsNullOrWhiteSpace(funcValue) || string.IsNullOrWhiteSpace(signalValue))
                        {
                            continue; // 공란인 경우 이 반복을 건너뜁니다.
                        }

                        // 포함조건 배열로 나누기: 쉼표로 구분하여 배열 생성, 공백 요소는 제외
                        var includeConditions = includeCondition.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                                .Select(condition => condition.Trim()).ToArray();

                        // 제외조건 배열로 나누기: 쉼표로 구분하여 배열 생성, 공백 요소는 제외
                        var exceptConditions = exceptCondition.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                              .Select(condition => condition.Trim()).ToArray();

                        bool isIncluded = false;
                        bool isExcluded = false;

                        // userTextParts 배열의 길이 체크
                        if (userTextParts.Length > 0 && !string.IsNullOrWhiteSpace(userTextParts[1]))
                        {
                            // 포함조건 체크
                            isIncluded = includeConditions.All(condition => userTextParts[1].Contains(condition)) ||
                                         (userTextParts.Length < 3 && string.IsNullOrWhiteSpace(includeCondition));

                            // 제외조건 체크
                            isExcluded = exceptConditions.Any(condition => userTextParts[1].Contains(condition));
                        }

                        if (userTextParts.Length > 2 && !string.IsNullOrWhiteSpace(userTextParts[2]))
                        {
                            // 포함조건 체크
                            isIncluded = isIncluded || includeConditions.All(condition => userTextParts[2].Contains(condition)) ||
                                         (userTextParts.Length < 3 && string.IsNullOrWhiteSpace(includeCondition));

                            // 제외조건 체크
                            isExcluded = isExcluded || exceptConditions.Any(condition => userTextParts[2].Contains(condition));
                        }

                        // 포함조건이 만족되면 해당 기능의 시그널을 리스트에 추가
                        if (isIncluded)
                        {
                            if (!funcSensorList[funcValue].Contains(signalValue))
                            {
                                funcSensorList[funcValue].Add(signalValue); // 조건을 만족하면 시그널 추가
                            }
                        }

                        // 제외조건이 만족되면 해당 기능의 시그널을 리스트에서 제거
                        if (isExcluded)
                        {
                            funcSensorList[funcValue].Remove(signalValue); // 조건을 만족하면 시그널 제거
                        }


                    }

                    foreach (var row in tempDt.AsEnumerable())
                    {
                        var funcValue = row.Field<string>("기능");    // "기능" 컬럼 값
                        var signalValue = row.Field<string>("SIGNAL"); // "SIGNAL" 컬럼 값

                        // Null 체크
                        if (!string.IsNullOrEmpty(funcValue) && !string.IsNullOrEmpty(signalValue))
                        {
                            // `funcSensorList`에서 해당 기능 키가 있는지 확인
                            if (funcSensorList.TryGetValue(funcValue, out var signals))
                            {
                                // SIGNAL 값이 없으면 DESCRIPTION과 SIGNAL 초기화
                                if (!signals.Contains(signalValue))
                                {
                                    row.SetField<string>("DESCRIPTION", null); // DESCRIPTION 초기화
                                    row.SetField<string>("SIGNAL", null);      // SIGNAL 초기화
                                    row.SetField<string>("타입", null);
                                }
                            }
                        }
                    }

                    // 조건에 맞는 값을 미리 정의한 HashSet에 넣어두고, Contains로 확인
                    var excludedValues = new HashSet<string> { "공통", "미사용", "모델명", "옵션", "GOXS", "GOXM", "GOXH", "회생 유닛", "3POS_1", "3POS_2" };

                    foreach (var row in tempDt.AsEnumerable().Where(row => !excludedValues.Contains(row.Field<string>("구분"))))
                    {
                        row.SetField<string>("DESCRIPTION", null); // DESCRIPTION 초기화
                        row.SetField<string>("SIGNAL", null);      // SIGNAL 초기화
                        row.SetField<string>("타입", null);
                    }

                    // 필요한 열만 선택하여 새 DataTable로 변환
                    DataTableUnits.dtSensorIo = tempDt.DefaultView.ToTable(false, "PARTS", "LOCATION", "DT", "TYPE1", "TYPE2", "CN.", "NO.", "ADD.", "SIGNAL", "DESCRIPTION", "구분", "포함조건", "제외조건", "기능", "타입");


                    // LOCATION 및 DT로 그룹화하여 SIGNAL 또는 DESCRIPTION이 없는 그룹 필터링
                    var filteredRows = DataTableUnits.dtSensorIo.AsEnumerable()
                        .GroupBy(row => new { LOCATION = row.Field<string>("LOCATION"), DT = row.Field<string>("DT") })
                        .Where(group =>
                            group.Any(row =>
                                !string.IsNullOrWhiteSpace(row.Field<string>("SIGNAL")) ||
                                !string.IsNullOrWhiteSpace(row.Field<string>("DESCRIPTION"))))
                        .SelectMany(group => group);

                    // 결과를 새로운 DataTable로 변환
                    DataTableUnits.dtSensorIo = filteredRows.Any()
                        ? filteredRows.CopyToDataTable()
                        : DataTableUnits.dtSensorIo.Clone(); // 결과가 없으면 빈 테이블 생성

                }
                catch (Exception ex) { }

            }
            void UpdatePlcCard()
            {
                if (cbMSPcontrollerSpec.Text == "HMX_MICOM")
                {
                    // gridView1의 모든 행을 순회
                    for (int i = 0; i < gridView1.RowCount; i++)
                    {
                        string dtValue = gridView1.GetRowCellValue(i, "DT") as string;
                        string locationValue = gridView1.GetRowCellValue(i, "LOCATION") as string;

                        // "DT" 열의 값이 "KE2"인 경우
                        if (dtValue == "KE2" && locationValue == "MP")
                        {
                            // 해당 행의 "8BIT" 열을 체크 상태로 설정 (true)
                            gridView1.SetRowCellValue(i, "8BIT", true); // 체크박스 체크
                        }
                    }
                }
                if (cbMSPcontrollerSpec.Text == "MIT_Q_IOLINK"
                    || cbMSPcontrollerSpec.Text == "MIT_Q_AJ" 
                    || cbMSPcontrollerSpec.Text == "MIT_R_IOLINK"
                    || cbMSPcontrollerSpec.Text == "MIT_R_AJ"
                    )
                {
                    // gridView1의 모든 행을 순회
                    for (int i = 0; i < gridView1.RowCount; i++)
                    {
                        string dtValue = gridView1.GetRowCellValue(i, "DT") as string;
                        string locationValue = gridView1.GetRowCellValue(i, "LOCATION") as string;

                        // "DT" 열의 값이 "KE2"인 경우
                        if (dtValue == "KE8" && locationValue == "MP")
                        {
                            // 해당 행의 "8BIT" 열을 체크 상태로 설정 (true)
                            gridView1.SetRowCellValue(i, "8BIT", true); // 체크박스 체크
                        }
                    }
                }
                if (cbMSPcontrollerSpec.Text == "SIE_ET200SP"
                    || cbMSPcontrollerSpec.Text == "SIE_S71500"
                    )
                {
                    // gridView1의 모든 행을 순회
                    for (int i = 0; i < gridView1.RowCount; i++)
                    {
                        string dtValue = gridView1.GetRowCellValue(i, "DT") as string;
                        string locationValue = gridView1.GetRowCellValue(i, "LOCATION") as string;

                        // "DT" 열의 값이 "KE2"인 경우
                        if (dtValue == "KE4" && locationValue == "MP")
                        {
                            // 해당 행의 "8BIT" 열을 체크 상태로 설정 (true)
                            gridView1.SetRowCellValue(i, "8BIT", true); // 체크박스 체크
                        }
                        if (dtValue == "KE5" && locationValue == "MP")
                        {
                            // 해당 행의 "8BIT" 열을 체크 상태로 설정 (true)
                            gridView1.SetRowCellValue(i, "8BIT", true); // 체크박스 체크
                        }
                    }
                }
            }

        }
        private void LoadIoFromExcel()
        {
            // DataTable 초기화 (데이터 및 컬럼 제거)
            DataTableUnits.dtExcelIo.Clear();
            DataTableUnits.dtExcelIo.Columns.Clear(); // 컬럼도 제거하여 중복 방지

            try
            {


                // Excel 파일을 열고 입력된 텍스트에 맞는 워크시트 로드
                using (var workbook = new XLWorkbook(StringUnits.strIoListFilePath))
                {
                    // 필요한 변수 선언 (이전 값 저장용)
                    string previousParts = string.Empty;
                    string previousLocation = string.Empty;
                    string previousDt = string.Empty;
                    string previousType1 = string.Empty;
                    string previousType2 = string.Empty;

                    var worksheet = workbook.Worksheet(cbMSPcontrollerSpec.Text);

                    // 첫 번째 행에서 컬럼 이름을 가져옴
                    bool firstRow = true;
                    foreach (var row in worksheet.RowsUsed())
                    {
                        if (firstRow)
                        {
                            // 첫 번째 행은 컬럼 이름으로 사용
                            foreach (var cell in row.Cells())
                            {
                                DataTableUnits.dtExcelIo.Columns.Add(cell.Value.ToString()); // 컬럼 추가
                            }
                            firstRow = false;
                        }
                        else
                        {
                            // 데이터 행 처리
                            DataRow dataRow = DataTableUnits.dtExcelIo.NewRow();
                            int i = 0;
                            foreach (var cell in row.Cells())
                            {
                                string cellValue = cell.Value.ToString();

                                // 공란 필드가 있는 경우 이전 행의 값을 가져와 채움
                                if (DataTableUnits.dtExcelIo.Columns[i].ColumnName == "PARTS")
                                {
                                    dataRow[i] = string.IsNullOrWhiteSpace(cellValue) ? previousParts : cellValue;
                                    previousParts = dataRow[i].ToString(); // 현재 값 저장
                                }
                                else if (DataTableUnits.dtExcelIo.Columns[i].ColumnName == "LOCATION")
                                {
                                    dataRow[i] = string.IsNullOrWhiteSpace(cellValue) ? previousLocation : cellValue;
                                    previousLocation = dataRow[i].ToString(); // 현재 값 저장
                                }
                                else if (DataTableUnits.dtExcelIo.Columns[i].ColumnName == "DT")
                                {
                                    dataRow[i] = string.IsNullOrWhiteSpace(cellValue) ? previousDt : cellValue;
                                    previousDt = dataRow[i].ToString(); // 현재 값 저장
                                }
                                else if (DataTableUnits.dtExcelIo.Columns[i].ColumnName == "TYPE1")
                                {
                                    dataRow[i] = string.IsNullOrWhiteSpace(cellValue) ? previousType1 : cellValue;
                                    previousType1 = dataRow[i].ToString(); // 현재 값 저장
                                }
                                else if (DataTableUnits.dtExcelIo.Columns[i].ColumnName == "TYPE2")
                                {
                                    dataRow[i] = string.IsNullOrWhiteSpace(cellValue) ? previousType2 : cellValue;
                                    previousType2 = dataRow[i].ToString(); // 현재 값 저장
                                }
                                else
                                {
                                    // 공란이 아닌 필드일 경우 그대로 값을 설정
                                    dataRow[i] = cellValue;
                                }

                                i++;
                            }

                            // DataTable에 데이터 추가
                            DataTableUnits.dtExcelIo.Rows.Add(dataRow);
                        }
                    }
                }
            }

            catch (Exception ex)
            {

            }
        }
        private void LoadMccbFromExcel()
        {
            // DataTable 초기화 (데이터 및 컬럼 제거)
            DataTableUnits.dtExcelMccb.Clear();
            DataTableUnits.dtExcelMccb.Columns.Clear(); // 컬럼도 제거하여 중복 방지


            try
            {
                // Excel 파일 열기
                using (var workbook = new XLWorkbook(StringUnits.strMccbFilePath))
                {
                    // 특정 워크시트 선택
                    var worksheet = workbook.Worksheet("차단기");

                    if (worksheet != null) // 워크시트가 존재하는 경우만 실행
                    {
                        // 첫 번째 행에서 컬럼명 읽기 (중복 컬럼명 방지)
                        if (DataTableUnits.dtExcelMccb.Columns.Count == 0)
                        {
                            var firstRow = worksheet.FirstRowUsed();
                            foreach (var cell in firstRow.CellsUsed())
                            {
                                DataTableUnits.dtExcelMccb.Columns.Add(cell.Value.ToString()); // 컬럼 추가
                            }
                        }

                        // 데이터 읽기 (첫 번째 행 이후부터)
                        foreach (var row in worksheet.RowsUsed().Skip(1))
                        {
                            var dataRow = DataTableUnits.dtExcelMccb.NewRow();
                            int columnIndex = 0;

                            foreach (var cell in row.CellsUsed())
                            {
                                if (columnIndex < DataTableUnits.dtExcelMccb.Columns.Count)
                                    dataRow[columnIndex] = cell.Value;
                                columnIndex++;
                            }

                            DataTableUnits.dtExcelMccb.Rows.Add(dataRow);
                        }
                    }
                }
            }

            catch (Exception ex)
            {

            }

            Control[] ctrls = { ckbPRJdomestic, ckbPRJoverseas, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqMccbModel };
            ComboBoxEdit[] liftCtrls = { cbLiftInverterKw, cbLiftInverterA, cbLiftPowerCable, cbLiftMccbSpec };
            ComboBoxEdit[] trav1Ctrls = { cbTrav1InverterKw, cbTrav1InverterA, cbTrav1PowerCable, cbTrav1MccbSpec };
            ComboBoxEdit[] trav2Ctrls = { cbTrav2InverterKw, cbTrav2InverterA, cbTrav2PowerCable, cbTrav2MccbSpec };
            ComboBoxEdit[] fork1Ctrls = { cbFork1InverterKw, cbFork1InverterA, cbFork1PowerCable, cbFork1MccbSpec };
            ComboBoxEdit[] fork2Ctrls = { cbFork2InverterKw, cbFork2InverterA, cbFork2PowerCable, cbFork2MccbSpec };

            void SelectInverter(Control[] controls, ComboBoxEdit[] funcControls)
            {
                string strPrjTarget = "";
                string strInverterMaker = "";
                string strInverterSpec = "";
                string strEleqMccbModel = "";

                string strInverterKw = "";
                string strInverterA = "";
                string strPowerCable = "";
                string strMccbSpec = "";
                string strOutPut = "";

                
                foreach (Control ctrl in controls)
                {
                    if(ctrl is CheckEdit ckb)
                    {
                        ckb.CheckedChanged += (o, e) =>
                        {
                            strPrjTarget = ckbPRJdomestic.Checked ? ckbPRJdomestic.Text : ckbPRJoverseas.Checked ? ckbPRJoverseas.Text : string.Empty;
                            strInverterMaker = cbMSPinverterMaker.Text;
                            strInverterSpec = cbMSPinverterSpec.Text;
                            strEleqMccbModel = cbEleqMccbModel.Text;

                            foreach (ComboBoxEdit funcCtrl in funcControls)
                            {
                                funcCtrl.SelectedIndex = -1;

                                List<string> listInverterKw = DataTableUnits.dtExcelMccb.AsEnumerable()
                                    .Where(row => row.Field<string>("타겟") == strPrjTarget &&
                                                row.Field<string>("인버터 제조사") == strInverterMaker &&
                                                row.Field<string>("인버터") == strInverterSpec &&
                                                row.Field<string>("차단기 제조사") == strEleqMccbModel)
                                    .Select(row => row.Field<string>("인버터 용량(kW)") ?? "")
                                    .Where(inverterKw =>
                                    {
                                        // "인버터 용량(Kw)"를 double로 변환하여 정렬을 위해 반환
                                        double inverterKwValue;
                                        return double.TryParse(inverterKw, out inverterKwValue);
                                    })
                                    .Select(inverterKw =>
                                    {
                                        // "인버터 용량(Kw)"를 double로 변환하여 정렬을 위해 반환
                                        double inverterKwValue;
                                        double.TryParse(inverterKw, out inverterKwValue);
                                        return new { inverterKw, inverterKwValue };
                                    })
                                    .OrderBy(x => x.inverterKwValue)  // 작은 순서대로 정렬
                                    .Select(x => x.inverterKw)  // 정렬된 값을 다시 string 리스트로 변환
                                    .Distinct()
                                    .ToList();

                                funcControls[0].Properties.Items.Clear();
                                funcControls[0].Properties.Items.AddRange(listInverterKw);
                            }

                        };
                    }
                    else if(ctrl is ComboBoxEdit cb)
                    {
                        cb.TextChanged += (o, e) =>
                        {
                            strPrjTarget = ckbPRJdomestic.Checked ? ckbPRJdomestic.Text : ckbPRJoverseas.Checked ? ckbPRJoverseas.Text : string.Empty;
                            strInverterMaker = cbMSPinverterMaker.Text;
                            strInverterSpec = cbMSPinverterSpec.Text;
                            strEleqMccbModel = cbEleqMccbModel.Text;
                            
                            foreach(ComboBoxEdit funcCtrl in funcControls)
                            {
                                funcCtrl.SelectedIndex = -1;

                                List<string> listInverterKw = DataTableUnits.dtExcelMccb.AsEnumerable()
                                    .Where(row => row.Field<string>("타겟") == strPrjTarget &&
                                                row.Field<string>("인버터 제조사") == strInverterMaker &&
                                                row.Field<string>("인버터") == strInverterSpec &&
                                                row.Field<string>("차단기 제조사") == strEleqMccbModel)
                                    .Select(row => row.Field<string>("인버터 용량(kW)") ?? "")
                                    .Where(inverterKw =>
                                    {
                                        // "인버터 용량(Kw)"를 double로 변환하여 정렬을 위해 반환
                                        double inverterKwValue;
                                        return double.TryParse(inverterKw, out inverterKwValue);
                                    })
                                    .Select(inverterKw =>
                                    {
                                        // "인버터 용량(Kw)"를 double로 변환하여 정렬을 위해 반환
                                        double inverterKwValue;
                                        double.TryParse(inverterKw, out inverterKwValue);
                                        return new { inverterKw, inverterKwValue };
                                    })
                                    .OrderBy(x => x.inverterKwValue)  // 작은 순서대로 정렬
                                    .Select(x => x.inverterKw)  // 정렬된 값을 다시 string 리스트로 변환
                                    .Distinct()
                                    .ToList();

                                funcControls[0].Properties.Items.Clear();
                                funcControls[0].Properties.Items.AddRange(listInverterKw);
                            }
                        };
                    }
                }

                funcControls[0].TextChanged += (o, e) =>
                {
                    strInverterKw = funcControls[0].Text;
                    
                    List<string> listInverterA = DataTableUnits.dtExcelMccb.AsEnumerable()
                        .Where(row => row.Field<string>("타겟") == strPrjTarget &&
                                row.Field<string>("인버터 제조사") == strInverterMaker &&
                                row.Field<string>("인버터") == strInverterSpec &&
                                row.Field<string>("차단기 제조사") == strEleqMccbModel &&
                                row.Field<string>("인버터 용량(kW)") == strInverterKw)
                        .Select(row => row.Field<string>("인버터 용량(A)") ?? "")
                        .Distinct()
                        .ToList();

                    funcControls[1].Properties.Items.Clear();
                    funcControls[1].Properties.Items.AddRange(listInverterA);
                    funcControls[1].SelectedIndex = 0;
                    strInverterA = funcControls[1].EditValue?.ToString();
                    
                    List<string> listPowerCable = DataTableUnits.dtExcelMccb.AsEnumerable()
                        .Where(row => row.Field<string>("타겟") == strPrjTarget &&
                                row.Field<string>("인버터 제조사") == strInverterMaker &&
                                row.Field<string>("인버터") == strInverterSpec &&
                                row.Field<string>("차단기 제조사") == strEleqMccbModel &&
                                row.Field<string>("인버터 용량(kW)") == strInverterKw &&
                                row.Field<string>("인버터 용량(A)") == strInverterA)
                        .Select(row => row.Field<string>("Power Cable(SQ)") ?? "")
                        .Distinct()
                        .ToList();

                    funcControls[2].Properties.Items.Clear();
                    funcControls[2].Properties.Items.AddRange(listPowerCable);
                    funcControls[2].SelectedIndex = 0;
                    strPowerCable = funcControls[2].EditValue?.ToString();

                    List<string> listMccbSpec = DataTableUnits.dtExcelMccb.AsEnumerable()
                        .Where(row => row.Field<string>("타겟") == strPrjTarget &&
                                row.Field<string>("인버터 제조사") == strInverterMaker &&
                                row.Field<string>("인버터") == strInverterSpec &&
                                row.Field<string>("차단기 제조사") == strEleqMccbModel &&
                                row.Field<string>("인버터 용량(kW)") == strInverterKw &&
                                row.Field<string>("인버터 용량(A)") == strInverterA &&
                                row.Field<string>("Power Cable(SQ)") == strPowerCable)
                        .Select(row => row.Field<string>("MCCB사양")??"")
                        .Distinct()
                        .ToList();
                   

                    funcControls[3].Properties.Items.Clear();
                    funcControls[3].Properties.Items.AddRange(listMccbSpec);
                    funcControls[3].SelectedIndex = 0;
                    
                };
            }

            ComboBoxEdit[] cbInverterAmps = { cbLiftInverterA, cbTrav1InverterA, cbTrav2InverterA, cbMSPinputVolt, cbLiftMotorVoltage, cbTrav1MotorVoltage, cbTrav2MotorVoltage };
            ComboBoxEdit[] cbPowerCables = { cbEleqMainMccbSpec, cbEleqMccbSpec };
            void SelectMainInverter(Control[] inverterAmps, Control[] powerCables)
            {
                double liftAmp = 0;
                double trav1Amp = 0;
                double trav2Amp = 0;

                double mspInputVolt = 0;

                double liftMotorVolt = 0;
                double trav1MotorVolt = 0;
                double trav2MotorVolt = 0;

                double sumAmp = 0;

                string strPrjTarget = "";
                string strInverterMaker = "";
                string strInverterSpec = "";
                string strEleqMccbModel = "";


                List<string> listInverterA = new List<string>();

                foreach (Control cb in inverterAmps) 
                {
                    
                    cb.TextChanged += (o, e) =>
                    {
                        strPrjTarget = ckbPRJdomestic.Checked ? ckbPRJdomestic.Text : ckbPRJoverseas.Checked ? ckbPRJoverseas.Text : string.Empty;
                        strInverterMaker = cbMSPinverterMaker.Text;
                        strInverterSpec = cbMSPinverterSpec.Text;
                        strEleqMccbModel = cbEleqMccbModel.Text;

                        double.TryParse(cbLiftInverterA.EditValue?.ToString() ?? "", out liftAmp);
                        double.TryParse(cbTrav1InverterA.EditValue?.ToString() ?? "", out trav1Amp);
                        double.TryParse(cbTrav2InverterA.EditValue?.ToString() ?? "", out trav2Amp);
                        double.TryParse(cbMSPinputVolt.EditValue?.ToString() ?? "", out mspInputVolt);
                        double.TryParse(cbLiftMotorVoltage.EditValue?.ToString() ?? "", out liftMotorVolt);
                        double.TryParse(cbTrav1MotorVoltage.EditValue?.ToString() ?? "", out trav1MotorVolt);
                        double.TryParse(cbTrav2MotorVoltage.EditValue?.ToString() ?? "", out trav2MotorVolt);
                        sumAmp =
                        (((liftAmp * liftMotorVolt / mspInputVolt) + (trav1Amp * trav1MotorVolt / mspInputVolt) + (trav2Amp * trav2MotorVolt / mspInputVolt)) * 1.4) + 5;


                        listInverterA = DataTableUnits.dtExcelMccb.AsEnumerable()
                                .Where(row => row.Field<string>("타겟") == strPrjTarget &&
                                            row.Field<string>("인버터 제조사") == strInverterMaker &&
                                            row.Field<string>("인버터") == strInverterSpec &&
                                            row.Field<string>("차단기 제조사") == strEleqMccbModel)
                                .Select(row => new
                                {
                                    InverterAmp = row.Field<string>("인버터 용량(A)") ?? "",
                                    MccbSpec = row.Field<string>("MCCB사양") ?? ""  // MCCB 사양 열을 추가
                                })
                                .Where(x =>
                                {
                                    // "인버터 용량(A)"를 double로 변환하여 sumAmp와 비교
                                    double inverterAmp;
                                    return double.TryParse(x.InverterAmp, out inverterAmp) && inverterAmp > sumAmp;
                                })
                                .Where(x => !string.IsNullOrWhiteSpace(x.MccbSpec))  // MCCB 사양이 공란이 아닌 경우만 필터링
                                .Select(x =>
                                {
                                    // "인버터 용량(A)"를 double로 변환하여 정렬을 위해 반환
                                    double inverterAmp;
                                    double.TryParse(x.InverterAmp, out inverterAmp);
                                    return new { x.InverterAmp, inverterAmp, x.MccbSpec };  // MCCB 사양도 포함
                                })
                                .OrderBy(x => x.inverterAmp)  // 암페어 용량으로 오름차순 정렬
                                .Select(x => x.MccbSpec)  // MCCB 사양만 반환
                                .Distinct()  // 중복 제거
                                .ToList();


                        if (cbMSPinverterMaker.Text == "SEW" && cbMSPinverterSpec.Text == "MODULAR")
                        {
                            cbEleqMccbSpec.Properties.Items.Clear();
                            cbEleqMccbSpec.Properties.Items.AddRange(listInverterA);
                            cbEleqMccbSpec.SelectedIndex = 0;

                            cbEleqMainMccbSpec.Properties.Items.Clear();
                            cbEleqMainMccbSpec.Properties.Items.AddRange(listInverterA);
                            cbEleqMainMccbSpec.SelectedIndex = 1;

                        }
                        else if(cbMSPinverterMaker.Text == "SEW" && cbMSPinverterSpec.Text == "SYSTEM")
                        {
                            cbEleqMainMccbSpec.Properties.Items.Clear();
                            cbEleqMainMccbSpec.Properties.Items.AddRange(listInverterA);
                            cbEleqMainMccbSpec.SelectedIndex = 0;
                        }
                        

                    };

                }

                foreach (Control cb in powerCables) 
                {
                    cb.TextChanged += (o, e) =>
                    {
                        strPrjTarget = ckbPRJdomestic.Checked ? ckbPRJdomestic.Text : ckbPRJoverseas.Checked ? ckbPRJoverseas.Text : string.Empty;
                        strInverterMaker = cbMSPinverterMaker.Text;
                        strInverterSpec = cbMSPinverterSpec.Text;
                        strEleqMccbModel = cbEleqMccbModel.Text;

                        List<string> listMainPowerCable = DataTableUnits.dtExcelMccb.AsEnumerable()
                        .Where(row => row.Field<string>("타겟") == strPrjTarget &&
                                row.Field<string>("인버터 제조사") == strInverterMaker &&
                                row.Field<string>("인버터") == strInverterSpec &&
                                row.Field<string>("차단기 제조사") == strEleqMccbModel &&
                                row.Field<string>("MCCB사양") == cbEleqMainMccbSpec.EditValue?.ToString())
                        .Select(row => row.Field<string>("Power Cable(SQ)") ?? "")
                        .Distinct()
                        .ToList();

                        List<string> listPowerCable = DataTableUnits.dtExcelMccb.AsEnumerable()
                        .Where(row => row.Field<string>("타겟") == strPrjTarget &&
                                row.Field<string>("인버터 제조사") == strInverterMaker &&
                                row.Field<string>("인버터") == strInverterSpec &&
                                row.Field<string>("차단기 제조사") == strEleqMccbModel &&
                                row.Field<string>("MCCB사양") == cbEleqMccbSpec.EditValue?.ToString())
                        .Select(row => row.Field<string>("Power Cable(SQ)") ?? "")
                        .Distinct()
                        .ToList();

                        cbEleqMainPowerCable.Properties.Items.Clear();
                        cbEleqMainPowerCable.Properties.Items.AddRange(listMainPowerCable);
                        cbEleqMainPowerCable.SelectedIndex = 0;

                        cbEleqPowerCable.Properties.Items.Clear();
                        cbEleqPowerCable.Properties.Items.AddRange(listPowerCable);
                        cbEleqPowerCable.SelectedIndex = 0;

                        List<string> listPowerKw = DataTableUnits.dtExcelMccb.AsEnumerable()
                        .Where(row => row.Field<string>("타겟") == strPrjTarget &&
                                row.Field<string>("인버터 제조사") == strInverterMaker &&
                                row.Field<string>("인버터") == strInverterSpec &&
                                row.Field<string>("차단기 제조사") == strEleqMccbModel &&
                                row.Field<string>("MCCB사양") == cbEleqMccbSpec.EditValue?.ToString() &&
                                row.Field<string>("Power Cable(SQ)") == cbEleqPowerCable.EditValue?.ToString())
                        .Select(row => row.Field<string>("인버터 용량(kW)") ?? "")
                        .Distinct()
                        .ToList();

                        cbEleqPowerKw.Properties.Items.Clear();
                        cbEleqPowerKw.Properties.Items.AddRange(listPowerKw);
                        cbEleqPowerKw.SelectedIndex = 0;

                        List<string> listPowerA = DataTableUnits.dtExcelMccb.AsEnumerable()
                        .Where(row => row.Field<string>("타겟") == strPrjTarget &&
                                row.Field<string>("인버터 제조사") == strInverterMaker &&
                                row.Field<string>("인버터") == strInverterSpec &&
                                row.Field<string>("차단기 제조사") == strEleqMccbModel &&
                                row.Field<string>("MCCB사양") == cbEleqMccbSpec.EditValue?.ToString() &&
                                row.Field<string>("Power Cable(SQ)") == cbEleqPowerCable.EditValue?.ToString() &&
                                row.Field<string>("인버터 용량(kW)") == cbEleqPowerKw.EditValue?.ToString())
                        .Select(row => row.Field<string>("인버터 용량(A)") ?? "")
                        .Distinct()
                        .ToList();

                        cbEleqPowerA.Properties.Items.Clear();
                        cbEleqPowerA.Properties.Items.AddRange(listPowerA);
                        cbEleqPowerA.SelectedIndex = 0;

                    };
                };
            }
            SelectMainInverter(cbInverterAmps, cbPowerCables);

            SelectInverter(ctrls, liftCtrls);
            SelectInverter(ctrls, trav1Ctrls);
            SelectInverter(ctrls, trav2Ctrls);
            SelectInverter(ctrls, fork1Ctrls);
            SelectInverter(ctrls, fork2Ctrls);

            cbEleqMainMccbSpec.Click += (o, e) =>
            {
                if(cbLiftInverterKw.BackColor != ColorUtility.colors[Ecolor.Active] || cbTrav1InverterKw.BackColor != ColorUtility.colors[Ecolor.Active])
                {
                    MessageBox.Show("LIFT, TRAV 인버터 선정을 해주세요.","주의",MessageBoxButtons.OK,MessageBoxIcon.Information);
                }
            };
            cbEleqMccbSpec.Click += (o, e) =>
            {
                if(cbLiftInverterKw.BackColor != ColorUtility.colors[Ecolor.Active] || cbTrav1InverterKw.BackColor != ColorUtility.colors[Ecolor.Active])
                {
                    MessageBox.Show("LIFT, TRAV 인버터 선정을 해주세요.", "주의", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
        }
        private void ActivateEplan()
        {
            DataTableUnits.dtBrakeOption.Rows.Clear();
            DataTableUnits.dtBrakeOption.Columns.Clear();
            DataTableUnits.dtBrakeOption.Columns.Add("functionTxt", typeof(string));
            DataTableUnits.dtBrakeOption.Columns.Add("valueBool", typeof(bool));

            picBoxProjectGenerate.Click += (o, e) =>
            {
                bool trav2Flag = false;
                bool fork2Flag = false;
                bool coldSelected = false;
                bool mainTenanceSelected = false;
                int brakeOptionBme15Count = 0;
                int liftCount = 0;
                int trav1Count = 0;
                int trav2Count = 0;
                int fork1Count = 0;
                int fork2Count = 0;
                Control[] brakeOptionControl = { cbLiftBrakeOption, cbTrav1BrakeOption, cbTrav2BrakeOption, cbFork1BrakeOption, cbFork2BrakeOption };

                DataTableUnits.dtBrakeOption.Clear();
                foreach (Control control in brakeOptionControl)
                {
                    string funcTxt = control.Parent.Parent.Parent.Parent.Parent.Parent.Text ?? "";
                    bool valueBool = control.Text == "BME 1.5" ? true : false;
                    DataTableUnits.dtBrakeOption.Rows.Add(funcTxt, valueBool);
                }

                if (ckbTravDoubleMotorTrue.Checked)
                {
                    trav2Flag = true;
                }
                if (ckbMODforkoption.Checked || cbMODoption1.Text == "D(v)" || cbMODoption2.Text == "D(v)" || cbMODoption3.Text == "D(v)" || cbMODoption4.Text == "D(v)")
                {
                    fork2Flag = true;
                }
                if (cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C")
                {
                    coldSelected = true;
                }
                if (cbMODoption1.Text == "M" || cbMODoption2.Text == "M" || cbMODoption3.Text == "M" || cbMODoption4.Text == "M")
                {
                    mainTenanceSelected = true;
                }

                InstallSiteType installSiteType =
                    ckbPRJdomestic.Checked ? InstallSiteType.DOMESTIC :
                    ckbPRJoverseas.Checked ? InstallSiteType.OVERSEAS
                    : InstallSiteType.ETC;

                PowerDpType powerDpType =
                    cbEleqTerminal.Text == "PDB" ? PowerDpType.PDB :
                    cbEleqTerminal.Text == "GENERAL" ? PowerDpType.GENERAL
                    : PowerDpType.ETC;

                InverterRegenType inverterRegenType =
                    ckbRegenerativeUnitTrue.Checked ? InverterRegenType.REGENERATIVE :
                    InverterRegenType.GENERAL;

                InverterMakerType inverterMakerType =
                    cbMSPinverterMaker.Text == "SEW" ? InverterMakerType.SEW :
                    cbMSPinverterMaker.Text == "SIE" ? InverterMakerType.SIEMENS :
                    InverterMakerType.ETC;

                InverterType inverterType =
                    cbMSPinverterSpec.Text == "MODULAR" ? InverterType.MODULAR :
                    cbMSPinverterSpec.Text == "SYSTEM" ? InverterType.SYSTEM :
                    InverterType.ETC;

                BrakeOptionType brakeOptionAcFlag = BrakeOptionType.Non;
                BrakeOptionType brakeOptionBitFlag = BrakeOptionType.Non;
                

                BrakeOptionType brakeOptionTypeLift =
                    cbLiftBrakeOption.Text == "BMKB 1.5" ? BrakeOptionType.BMBK15 :
                    cbLiftBrakeOption.Text == "BME 1.5" ? BrakeOptionType.BME15 :
                    cbLiftBrakeOption.Text == "BMV 5" ? BrakeOptionType.BMV5 :
                    cbLiftBrakeOption.Text == "BMH 1.5" ? BrakeOptionType.BMH15 :
                    BrakeOptionType.ETC;

                BrakeOptionType brakeOptionTypeTrav1 =
                    cbTrav1BrakeOption.Text == "BMKB 1.5" ? BrakeOptionType.BMBK15 :
                    cbTrav1BrakeOption.Text == "BME 1.5" ? BrakeOptionType.BME15 :
                    cbTrav1BrakeOption.Text == "BMV 5" ? BrakeOptionType.BMV5 :
                    cbTrav1BrakeOption.Text == "BMH 1.5" ? BrakeOptionType.BMH15 :
                    BrakeOptionType.ETC;

                BrakeOptionType brakeOptionTypeTrav2 =
                    cbTrav2BrakeOption.Text == "BMKB 1.5" ? BrakeOptionType.BMBK15 :
                    cbTrav2BrakeOption.Text == "BME 1.5" ? BrakeOptionType.BME15 :
                    cbTrav2BrakeOption.Text == "BMV 5" ? BrakeOptionType.BMV5 :
                    cbTrav2BrakeOption.Text == "BMH 1.5" ? BrakeOptionType.BMH15 :
                    BrakeOptionType.ETC;

                BrakeOptionType brakeOptionTypeFork1 =
                    cbFork1BrakeOption.Text == "BMKB 1.5" ? BrakeOptionType.BMBK15 :
                    cbFork1BrakeOption.Text == "BME 1.5" ? BrakeOptionType.BME15 :
                    cbFork1BrakeOption.Text == "BMV 5" ? BrakeOptionType.BMV5 :
                    cbFork1BrakeOption.Text == "BMH 1.5" ? BrakeOptionType.BMH15 :
                    BrakeOptionType.ETC;

                BrakeOptionType brakeOptionTypeFork2 =
                    cbFork2BrakeOption.Text == "BMKB 1.5" ? BrakeOptionType.BMBK15 :
                    cbFork2BrakeOption.Text == "BME 1.5" ? BrakeOptionType.BME15 :
                    cbFork2BrakeOption.Text == "BMV 5" ? BrakeOptionType.BMV5 :
                    cbFork2BrakeOption.Text == "BMH 1.5" ? BrakeOptionType.BMH15 :
                    BrakeOptionType.ETC;

                ControlVoltageType controlVoltageType =
                    cbMSPusingVoltage.Text == "120 V" ? ControlVoltageType.V120 :
                    cbMSPusingVoltage.Text == "220 V" ? ControlVoltageType.V220 :
                    ControlVoltageType.ETC;

                ForkType forkType =
                    ckbMODforkoption.Checked || cbMODoption1.Text == "D(v)" || cbMODoption2.Text == "D(v)" || cbMODoption3.Text == "D(v)" || cbMODoption4.Text == "D(v)" ? ForkType.FORK2 :
                    ForkType.FORK1;

                TravelType travelType =
                    ckbTravDoubleMotorTrue.Checked ? TravelType.TRAV2 :
                    TravelType.TRAV1;

                ControllerType controllerType =
                    cbMSPcontrollerSpec.Text == "HMX_MICOM" ? ControllerType.MICOM :
                    cbMSPcontrollerSpec.Text == "MIT_Q_AJ" ? ControllerType.MITSUBISHI_Q :
                    cbMSPcontrollerSpec.Text == "MIT_R_AJ" ? ControllerType.MITSUBISHI_R :
                    cbMSPcontrollerSpec.Text == "MIT_Q_IOLINK" ? ControllerType.MITSUBISHI_Q_IOLINK :
                    cbMSPcontrollerSpec.Text == "MIT_R_IOLINK" ? ControllerType.MITSUBISHI_R_IOLINK :
                    cbMSPcontrollerSpec.Text == "SIE_ET200SP" ? ControllerType.SIEMENS_ET200SP :
                    cbMSPcontrollerSpec.Text == "SIE_S71500" ? ControllerType.SIEMENS_S7_1500 :
                    ControllerType.ETC;

                MotorCableType motorCableTypeLift =
                    ckbLiftDdi.Checked ? MotorCableType.DDI : MotorCableType.GENERAL;
                MotorCableType motorCableTypeTrav1 =
                    ckbTrav1Ddi.Checked ? MotorCableType.DDI : MotorCableType.GENERAL;
                MotorCableType motorCableTypeTrav2 =
                    ckbTrav2Ddi.Checked ? MotorCableType.DDI : MotorCableType.GENERAL;
                MotorCableType motorCableTypeFork1 =
                    ckbFork1Ddi.Checked ? MotorCableType.DDI : MotorCableType.GENERAL;
                MotorCableType motorCableTypeFork2 =
                    ckbFork2Ddi.Checked ? MotorCableType.DDI : MotorCableType.GENERAL;

                EncoderType encoderTypeLift =
                    ckbLiftRaser.Checked ? EncoderType.LASER : EncoderType.BARCODE;
                EncoderType encoderTypeTrav1 =
                    ckbTrav1Raser.Checked ? EncoderType.LASER : EncoderType.BARCODE;
                EncoderType encoderTypeTrav2 =
                    ckbTrav2Raser.Checked ? EncoderType.LASER : EncoderType.BARCODE;
                EncoderType encoderTypeFork1 = EncoderType.ETC;
                EncoderType encoderTypeFork2 = EncoderType.ETC;

                SensorOutputType sensorOutputType =
                    cbEleqSensorType.Text == "NPN" ? SensorOutputType.NPN :
                    cbEleqSensorType.Text == "PNP" ? SensorOutputType.PNP :
                    SensorOutputType.ETC;

                FluorescentType fluorescentType =
                    cbEleqPanel.Text == "GENERAL" ? FluorescentType.GENERAL :
                    cbEleqPanel.Text == "RITTAL" ? FluorescentType.RITTAL :
                    FluorescentType.ETC;

                TowerLampType towerLampType =
                    cbEleqTowerLamp.Text == "3-COLOR" ? TowerLampType.COLOR3 :
                    cbEleqTowerLamp.Text == "4-COLOR" ? TowerLampType.COLOR4 :
                    TowerLampType.ETC;


                foreach (Control ctrl in brakeOptionControl)
                {
                    if (ctrl.Text == "BME 1.5" || ctrl.Text == "BMH 1.5")
                    {
                        brakeOptionAcFlag = BrakeOptionType.BME15;
                        brakeOptionBme15Count++;
                    }
                }

                foreach (Control ctrl in brakeOptionControl)
                {
                    if (ctrl.Text == "BME 1.5")
                    {
                        brakeOptionBitFlag = BrakeOptionType.BME15;
                    }
                }

                if (trav2Flag && fork2Flag)
                {
                    liftCount = 1;
                    trav1Count = 2;
                    trav2Count = 3;
                    fork1Count = 4;
                    fork2Count = 5;
                }
                else if (trav2Flag && !fork2Flag)
                {
                    liftCount = 1;
                    trav1Count = 2;
                    trav2Count = 3;
                    fork1Count = 4;
                    fork2Count = 0;
                }
                else if (!trav2Flag && fork2Flag)
                {
                    liftCount = 1;
                    trav1Count = 2;
                    trav2Count = 0;
                    fork1Count = 3;
                    fork2Count = 4;
                }
                else if (!trav2Flag && !fork2Flag)
                {
                    liftCount = 1;
                    trav1Count = 2;
                    trav2Count = 0;
                    fork1Count = 3;
                    fork2Count = 0;
                }

                if (!CheckBeforeGenerating())
                    return;
                GeneratePageMacro();




                bool CheckBeforeGenerating()
                {

                    // ✅ 유효성 검사할 ComboBoxEdit 컨트롤 리스트
                    List<(ComboBoxEdit cb, bool checkFlag)> comboBoxList = new List<(ComboBoxEdit, bool)>
                {
                    (cbPRJnumber, true),
                    (cbMODfullName, true),
                    (cbMSPcontrollerSpec, true),
                    (cbMSPinverterMaker, true),
                    (cbMSPinverterSpec, true),
                    (cbEleqTerminal, true),
                    (cbLiftBrakeOption, true),
                    (cbTrav1BrakeOption, true),
                    (cbFork1BrakeOption, true),
                    (cbTrav2BrakeOption, trav2Flag),
                    (cbFork2BrakeOption, fork2Flag)
                };

                    string errActCbTxt = "";

                    foreach (var (cb, checkFlag) in comboBoxList)
                    {
                        if (checkFlag && (cb.BackColor != ColorUtility.colors[Ecolor.Active] || string.IsNullOrEmpty(cb.Text)))
                        {
                            // 레이블 텍스트 찾기 (부모 컨트롤에서 검색)
                            var labelText = cb.Parent.Controls.OfType<LabelControl>().FirstOrDefault()?.Text;
                            if (labelText != null)
                            {
                                errActCbTxt += $"[{labelText}] "; // 추가할 레이블 텍스트
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(errActCbTxt))
                    {
                        MessageBox.Show($"다음 필드를 기입해주세요. : {errActCbTxt}", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false; // 유효성 검사 실패
                    }

                    return true; // 유효성 검사 성공
                }

                void GeneratePageMacro()
                {

                    mcnsControl.SetMacroRootDirectory(StringUnits.strMacroFolderPath);

                    string elkName = string.Concat(cbPRJnumber.Text, "_", cbMODfullName.Text);
                    string prjFullFilePath = Path.Combine(StringUnits.strPrjFolderPath, elkName + ".elk");

                    // 프로젝트 파일이 이미 존재하는지 확인
                    if (File.Exists(prjFullFilePath))
                    {
                        // 덮어쓰기 여부를 묻는 메시지 박스
                        DialogResult dialogResult = MessageBox.Show(string.Concat("이미 생성된 프로젝트가 있습니다.", "\n경로 : ", StringUnits.strPrjFolderPath, "\n파일이름 : ", elkName), "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        using (new LockingStep())
                        {
                            Progress oProgress = new Progress("SimpleProgress");
                            oProgress.ShowImmediately();

                            //part 1
                            oProgress.BeginPart(13.5, "");
                            oProgress.SetActionText("프로젝트 자동 생성");
                            oProgress.SetNeededSteps(1);
                            oProgress.Step(1); this.mcnsControl.CreateAndOpenProject(prjFullFilePath, StringUnits.strBasicTempletFilePath);
                            oProgress.EndPart(false);

                            //part 2
                            oProgress.BeginPart(38.5, "");
                            oProgress.SetActionText("매크로 삽입");
                            oProgress.SetNeededSteps(1);
                            void CheckMcnsEngineFunction(ResponseModel responseModel)
                            {
                                if (responseModel.Success == false)
                                    MessageBox.Show(responseModel.Message);
                            }

                            CheckMcnsEngineFunction(this.mcnsControl.InsertACPowerEmpMacro(installSiteType, powerDpType, inverterRegenType, inverterMakerType, inverterType, controlVoltageType, brakeOptionAcFlag, brakeOptionBme15Count));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertDCPowerEmpMacro(installSiteType, powerDpType, inverterMakerType, inverterType, controllerType));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertInverterPublicEmpMacro(installSiteType, inverterMakerType, inverterType, inverterRegenType, powerDpType, controllerType, forkType, travelType, controlVoltageType));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertSLDPageMacro(installSiteType, inverterMakerType, inverterType, controllerType));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertCfgPageMacro(inverterMakerType, controllerType));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertSystemEmpMacro(sensorOutputType,controllerType,controlVoltageType,installSiteType));


                            int brakeCount = 0;
                            int nobrakeCount = 0;

                            for (int i =0; i < brakeOptionControl.Count(); i++)
                            {
                                string funcTxt = brakeOptionControl[i].Parent.Parent.Parent.Parent.Parent.Parent.Text;
                                string brakeTxt = brakeOptionControl[i].Text;

                                FunctionType functionType = new FunctionType();
                                MotorCableType motorCableType = new MotorCableType();
                                EncoderType encoderType = new EncoderType();
                                BrakeOptionType brakeOptionType = new BrakeOptionType();
                                int count = 0;
                                

                                if(funcTxt == "LIFT")
                                {
                                    functionType = FunctionType.LIFT;
                                    motorCableType = motorCableTypeLift;
                                    encoderType = encoderTypeLift;
                                    brakeOptionType = brakeOptionTypeLift;
                                    count = liftCount;
                                }
                                else if(funcTxt == "TRAV")
                                {
                                    functionType = FunctionType.TRAV1;
                                    motorCableType = motorCableTypeTrav1;
                                    encoderType = encoderTypeTrav1;
                                    brakeOptionType = brakeOptionTypeTrav1;
                                    count = trav1Count;
                                }
                                else if (funcTxt == "TRAV2")
                                {
                                    if (!trav2Flag)
                                    {
                                        continue;
                                    }
                                    functionType = FunctionType.TRAV2;
                                    motorCableType = motorCableTypeTrav2;
                                    encoderType = encoderTypeTrav2;
                                    brakeOptionType = brakeOptionTypeTrav2;
                                    count = trav2Count;
                                }
                                else if (funcTxt == "FORK")
                                {
                                    functionType = FunctionType.FORK1;
                                    motorCableType = motorCableTypeFork1;
                                    encoderType = encoderTypeFork1;
                                    brakeOptionType = brakeOptionTypeFork1;
                                    count = fork1Count;
                                }
                                else if (funcTxt == "FORK2")
                                {
                                    if (!fork2Flag)
                                    {
                                        continue;
                                    }
                                    functionType = FunctionType.FORK2;
                                    motorCableType = motorCableTypeFork2;
                                    encoderType = encoderTypeFork2;
                                    brakeOptionType = brakeOptionTypeFork2;
                                    count = fork2Count;
                                }

                                if(brakeTxt == "BME 1.5")
                                {
                                    brakeCount++;
                                    CheckMcnsEngineFunction(this.mcnsControl.InsertMotorPublicEmpMacro(functionType, motorCableType, encoderType, installSiteType, inverterMakerType, inverterType, inverterRegenType, powerDpType, controllerType, forkType, travelType, brakeOptionType, count, brakeCount));
                                }
                                else
                                {
                                    CheckMcnsEngineFunction(this.mcnsControl.InsertMotorPublicEmpMacro(functionType, motorCableType, encoderType, installSiteType, inverterMakerType, inverterType, inverterRegenType, powerDpType, controllerType, forkType, travelType, brakeOptionType, count, nobrakeCount));
                                }

                            }

                            if (!coldSelected)
                            {
                                CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorFan(int.Parse(cbEleqFanQuantity.Text)));
                            }
                            CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorFluorenscentLamp(fluorescentType, installSiteType, controlVoltageType));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorHMI(cbEleqHmi.Text));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorHub(cbEleqHubModel.Text));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorOPT(installSiteType, controllerType, DataTableUnits.dtUniqueIo, cbEleqOpt.Text));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorSafetyEmergency(installSiteType, int.Parse(cbEleqSafetyQuantity.Text), forkType, travelType, cbEleqSafety.Text));
                            if (ckbPRJoverseas.Checked)
                            {
                                CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacroSafetyRelay(cbEleqSafety.Text));
                                CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacroSafetyReset(cbEleqSafety.Text));
                            }

                            CheckMcnsEngineFunction(this.mcnsControl.InsertControllerMacro(controllerType, DataTableUnits.dtUniqueIo, DataTableUnits.dtSensorIo));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertControllerBitIOMacro(controllerType, DataTableUnits.dtUniqueIo, DataTableUnits.dtSensorIo, brakeOptionBitFlag, DataTableUnits.dtBrakeOption));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertPlcIOSwitchWindowMacro(installSiteType, controllerType, int.Parse(cbEleqSafetyQuantity.Text)));
                            CheckMcnsEngineFunction(this.mcnsControl.InsertPlcTowerLampWindowMacro(controllerType, towerLampType));

                            if (mainTenanceSelected)
                            {
                                CheckMcnsEngineFunction(this.mcnsControl.InsertMaintenanceOption(coldSelected));
                            }

                            if (coldSelected)
                            {
                                CheckMcnsEngineFunction(this.mcnsControl.InsertColdStorageOption(installSiteType, controllerType));
                            }

                            if (ckbCctvTrue.Checked)
                            {
                                CheckMcnsEngineFunction(this.mcnsControl.InsertOptionCCTV());
                            }

                            if (!fork2Flag && cbMSPinverterMaker.Text != "SIE")
                            {
                                CheckMcnsEngineFunction(this.mcnsControl.DeleteSEWModularIINH(FunctionType.FORK1));
                            }
                            oProgress.EndPart(false);

                            //part3
                            oProgress.BeginPart(48.0, "");
                            oProgress.SetActionText("PDF 생성");
                            oProgress.SetNeededSteps(1);
                            CheckMcnsEngineFunction(this.mcnsControl.ApplyWirePlaceHolder(installSiteType));
                            this.mcnsControl.SetProjectProperty("EPLAN.Project.UserSupplementaryField15", cbPRJname.Text);
                            this.mcnsControl.SetProjectProperty("EPLAN.Project.UserSupplementaryField9", elkName);
                            this.mcnsControl.SetProjectProperty("EPLAN.Project.UserSupplementaryField21", cbMSPinputVolt.Text);
                            this.mcnsControl.SetProjectProperty("EPLAN.Project.UserSupplementaryField22", cbMSPinputHz.Text);
                            this.mcnsControl.SetProjectProperty("EPLAN.Project.UserSupplementaryField23", cbEleqPowerA.Text);
                            this.mcnsControl.SetProjectProperty("10232", cbPRJwriter.Text);

                           

                            using (ExcelPackage package = new ExcelPackage())
                            {
                                // 워크시트 추가
                                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                                // DataTable 데이터를 워크시트에 로드
                                worksheet.Cells["A1"].LoadFromDataTable(DataTableUnits.dtSensorIo, true);

                                // 파일 저장
                                FileInfo file = new FileInfo(Path.Combine(StringUnits.strPrjFolderPath, elkName + ".xlsx"));
                                package.SaveAs(file);
                            }
                            this.mcnsControl.GeneratePdf();
                            oProgress.EndPart(true);

                            //MessageBox.Show(elkName + ": 프로젝트 생성 완료");
                        }

                    }


                }
            };
        }
        private void Interlock()
        {
            

            interLock.UpdateFullText(
                cbMODfullName,
                new Control[] { cbMODname, cbMODheight, cbMODweight, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4 });
            
            interLock.SplitTextByDelimiter(
                cbMSPpanelSize,
                "기타",
                new ComboBoxEdit[] { cbMSPpanelSizeW, cbMSPpanelSizeD, cbMSPpanelSizeH },
                '*');

            interLock.CheckSwitchByText(
                new Control[] { cbMODname },
                new string[] { "UCX" },
                ckbTravDoubleMotorTrue
                );
            interLock.CheckSwitchByText(
                new Control[] { cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4 },
                new string[] { "C" },
                ckbRegenerativeUnitTrue
                );

            interLock.ActivatePageByText(
                new Control[] { ckbTravDoubleMotorTrue },
                new string[] { "적용" },
                xtraTabControlFunction,
                3
                );
            interLock.ActivatePageByText(
                new Control[] { ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4 },
                new string[] { "x2", "D(v)" },
                xtraTabControlFunction,
                5);
            interLock.ActivatePageByText(
                new Control[] { cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4 },
                new string[] { "M" },
                xtraTabControlFunction,
                7);
            interLock.ActivatePageByText(
                new Control[] { ckbTravDoubleMotorTrue },
                new string[] { "적용" },
                xtraTabControlFunction,
                8
                );

            interLock.AlramToFunctionByText(rtbxEleq, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxEleq, new Control[] { cbMODname , ckbMODforkoption, cbMODoption1 , cbMODoption2 , cbMODoption3 , cbMODoption4 , cbMSPinputVolt , cbMSPinputHz , cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec , cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxLift, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxTrav1, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxTrav2, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxFork1, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxFork2, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxCarr, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });

            //패널 사이즈에 따른 팬 수량
            UpdateFanQuantityByPanel();
            cbMSPpanelSizeW.TextChanged += (o, e) =>
            {
                UpdateFanQuantityByPanel();
            };
            void UpdateFanQuantityByPanel()
            {
                if (int.TryParse(cbMSPpanelSizeW.Text, out int value))
                {
                    if (value < 800)
                    {
                        cbEleqFanQuantity.Text = "4";
                    }
                    else
                    {
                        cbEleqFanQuantity.Text = "6";
                    }
                }
                else
                {
                    cbEleqFanQuantity.SelectedIndex = -1;
                }
            }
            
            //프로젝트 옵션 콤보 박스 visble 인터락
            List<string> opItemsList = cbMODoption1.Properties.Items.Cast<string>().ToList();
            cbMODoption2.Hide();
            cbMODoption3.Hide();
            cbMODoption4.Hide();
            cbMODoption1.TextChanged += (o, e) =>
            {
                FilterComboBox(cbMODoption1, cbMODoption2, opItemsList);
                UpdateComboBoxVisibility();
            };
            cbMODoption2.TextChanged += (o, e) =>
            {
                FilterComboBox(cbMODoption2, cbMODoption3, opItemsList, cbMODoption1);
                UpdateComboBoxVisibility();
            };
            cbMODoption3.TextChanged += (o, e) =>
            {
                FilterComboBox(cbMODoption3, cbMODoption4, opItemsList, cbMODoption1, cbMODoption2);
                UpdateComboBoxVisibility();
            };
            void FilterComboBox(ComboBoxEdit currentCb, ComboBoxEdit nextCb, List<string> items, params ComboBoxEdit[] previousCbs)
            {
                // White 배경일 때만 필터링을 적용하고 다음 ComboBox를 보여줍니다.
                if (currentCb.BackColor == ColorUtility.colors[Ecolor.Active])
                {
                    // 모든 이전 ComboBox와 현재 ComboBox에서 선택된 최대 인덱스를 가져옵니다.
                    int maxIndex = previousCbs
                        .Append(currentCb)
                        .Select(cb => items.IndexOf(cb.Text))
                        .Where(index => index >= 0)
                        .DefaultIfEmpty(-1)
                        .Max();

                    // 필터링된 리스트 설정
                    nextCb.Show();
                    nextCb.Properties.Items.Clear();
                    nextCb.Properties.Items.AddRange(items
                        .Where((_, i) => i > maxIndex)
                        .ToArray());
                }
            }
            void UpdateComboBoxVisibility()
            {
                // cbMODoption1의 배경이 흰색이 아닌 경우 2, 3, 4번 콤보박스를 숨기고 텍스트를 지웁니다.
                if (cbMODoption1.BackColor != ColorUtility.colors[Ecolor.Active])
                {
                    cbMODoption2.Hide();
                    cbMODoption2.Text = ""; // 텍스트 지우기
                    cbMODoption3.Hide();
                    cbMODoption3.Text = ""; // 텍스트 지우기
                    cbMODoption4.Hide();
                    cbMODoption4.Text = ""; // 텍스트 지우기
                }
                // cbMODoption2의 배경이 흰색이 아닌 경우 3, 4번 콤보박스를 숨기고 텍스트를 지웁니다.
                else if (cbMODoption2.BackColor != ColorUtility.colors[Ecolor.Active])
                {
                    cbMODoption3.Hide();
                    cbMODoption3.Text = ""; // 텍스트 지우기
                    cbMODoption4.Hide();
                    cbMODoption4.Text = ""; // 텍스트 지우기
                }
                // cbMODoption3의 배경이 흰색이 아닌 경우 4번 콤보박스를 숨기고 텍스트를 지웁니다.
                else if (cbMODoption3.BackColor != ColorUtility.colors[Ecolor.Active])
                {
                    cbMODoption4.Hide();
                    cbMODoption4.Text = ""; // 텍스트 지우기
                }
                // 모든 ComboBox가 흰색일 경우 모두 보이도록 설정하고 텍스트는 지우지 않습니다.
                else
                {
                    cbMODoption2.Show();
                    cbMODoption3.Show();
                    cbMODoption4.Show();
                }
            }
            


            //화물 셀 기입 인터락
            gridViewCargo.CellValueChanged += (o, e) =>
            {
                //int rowCount = 3; // 1~3행만 검사 (필요시 rowCount를 조정)
                int columnCount = gridViewCargo.Columns.Count;

                // 각 행이 모두 채워졌는지 확인하는 플래그
                bool isFirstRowFilled = true;
                bool isSecondRowFilled = true;
                bool isThirdRowFilled = true;
                bool isFourthRowFilled = true;

                // 1행의 각 셀을 확인
                for (int colIndex = 0; colIndex < columnCount; colIndex++)
                {
                    var cellValue = gridViewCargo.GetRowCellValue(0, gridViewCargo.Columns[colIndex]); // 1행
                    if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                    {
                        isFirstRowFilled = false;
                        break;
                    }
                }

                // 2행의 각 셀을 확인
                for (int colIndex = 0; colIndex < columnCount; colIndex++)
                {
                    var cellValue = gridViewCargo.GetRowCellValue(1, gridViewCargo.Columns[colIndex]); // 2행
                    if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                    {
                        isSecondRowFilled = false;
                        break;
                    }
                }

                // 3행의 각 셀을 확인
                for (int colIndex = 0; colIndex < columnCount; colIndex++)
                {
                    var cellValue = gridViewCargo.GetRowCellValue(2, gridViewCargo.Columns[colIndex]); // 3행
                    if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                    {
                        isThirdRowFilled = false;
                        break;
                    }
                }

                // 4행의 각 셀을 확인
                for (int colIndex = 0; colIndex < columnCount; colIndex++)
                {
                    var cellValue = gridViewCargo.GetRowCellValue(3, gridViewCargo.Columns[colIndex]); // 4행
                    if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                    {
                        isFourthRowFilled = false;
                        break;
                    }
                }

                // 체크박스 제어
                if (isFirstRowFilled && isSecondRowFilled)
                {
                    // 2행이 채워졌다면 ckbCarrDoubleCarriageGOXH 활성화 및 체크
                    ckbCarrDoubleCarriageGOXH.Enabled = true;
                    ckbCarrDoubleCarriageGOXH.Checked = true;
                }
                else
                {
                    // 2행이 채워지지 않았다면 비활성화 및 체크 해제
                    ckbCarrDoubleCarriageGOXH.Checked = false;
                    ckbCarrDoubleCarriageGOXH.Enabled = false;
                }

                if (isFirstRowFilled && isSecondRowFilled && isThirdRowFilled)
                {
                    // 2, 3행이 채워졌다면 ckbCarrDoubleCarriageGOXM 활성화 및 체크
                    ckbCarrDoubleCarriageGOXM.Enabled = true;
                    ckbCarrDoubleCarriageGOXM.Checked = true;
                }
                else
                {
                    // 2, 3행이 채워지지 않았다면 비활성화 및 체크 해제
                    ckbCarrDoubleCarriageGOXM.Checked = false;
                    ckbCarrDoubleCarriageGOXM.Enabled = false;
                }

                if (isFirstRowFilled && isSecondRowFilled && isThirdRowFilled && isFourthRowFilled)
                {
                    // 2, 3, 4행이 모두 채워졌다면 ckbCarrDoubleCarriageGOXS 활성화 및 체크
                    ckbCarrDoubleCarriageGOXS.Enabled = true;
                    ckbCarrDoubleCarriageGOXS.Checked = true;
                }
                else
                {
                    // 1, 2, 3행 중 하나라도 비어 있으면 비활성화 및 체크 해제
                    ckbCarrDoubleCarriageGOXS.Checked = false;
                    ckbCarrDoubleCarriageGOXS.Enabled = false;
                }
            };

            // button을 Radio button으로 사용
            ckbManager.ChangeToRadioButton(ckbLiftRaser, ckbLiftBarcode);
            // button을 Radio button으로 사용
            ckbManager.ChangeToRadioButton(ckbTrav1Raser, ckbTrav1Barcode);
            // button을 Radio button으로 사용
            ckbManager.ChangeToRadioButton(ckbTrav2Raser, ckbTrav2Barcode);

            //NPN,PNP,COLD 타입에 따른 목록 인터락
            cbMODoption1.TextChanged += ColdTypeOption_TextChanged;
            cbMODoption2.TextChanged += ColdTypeOption_TextChanged;
            cbMODoption3.TextChanged += ColdTypeOption_TextChanged;
            cbMODoption4.TextChanged += ColdTypeOption_TextChanged;
            void ColdTypeOption_TextChanged(object sender, EventArgs e)
            {
                bool coldFlag = false;
                if (cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C")
                {
                    coldFlag = true;
                }

                // "C"가 포함되어 있는지 확인
                if (coldFlag)
                {
                    cbOPmachineControl.SelectedIndex = 1;
                    cbEleqFanQuantity.Enabled = false;

                    cbEleqModem.Properties.Items.Clear();
                    cbEleqSensorItem.Properties.Items.Clear();
                    cbLiftAbsLocation.Properties.Items.Clear();
                    cbTrav1AbsLocation.Properties.Items.Clear();
                    cbTrav2AbsLocation.Properties.Items.Clear();
                    cbLiftBrakeOption.Properties.Items.Clear();
                    cbTrav1BrakeOption.Properties.Items.Clear();
                    cbTrav2BrakeOption.Properties.Items.Clear();
                    cbFork1BrakeOption.Properties.Items.Clear();
                    cbFork2BrakeOption.Properties.Items.Clear();

                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listColdEleqModem", cbEleqModem);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listColdEleqSensorItem", cbEleqSensorItem);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listColdLiftBrakeOption", cbLiftBrakeOption);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listColdTravBrakeOption", cbTrav1BrakeOption);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listColdTravBrakeOption", cbTrav2BrakeOption);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listColdForkBrakeOption", cbFork1BrakeOption);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listColdForkBrakeOption", cbFork2BrakeOption);

                    if (ckbLiftRaser.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listRaserColdLiftAbsLocation", cbLiftAbsLocation);
                    }
                    else if (ckbLiftBarcode.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listBarcodeColdLiftAbsLocation", cbLiftAbsLocation);
                    }
                    if (ckbTrav1Raser.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listRaserColdTravAbsLocation", cbTrav1AbsLocation);
                    }
                    else if (ckbTrav1Barcode.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listBarcodeColdTravAbsLocation", cbTrav1AbsLocation);
                    }
                    if (ckbTrav2Raser.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listRaserColdTravAbsLocation", cbTrav2AbsLocation);
                    }
                    else if (ckbTrav2Barcode.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listBarcodeColdTravAbsLocation", cbTrav2AbsLocation);
                    }

                    cbEleqModem.SelectedIndex = -1;
                    cbEleqSensorItem.SelectedIndex = -1;
                    cbLiftAbsLocation.SelectedIndex = -1;
                    cbTrav1AbsLocation.SelectedIndex = -1;
                    cbTrav2AbsLocation.SelectedIndex = -1;
                    cbLiftBrakeOption.SelectedIndex = 0;
                    cbTrav1BrakeOption.SelectedIndex = 0;
                    cbTrav2BrakeOption.SelectedIndex = 0;
                    cbFork1BrakeOption.SelectedIndex = 0;
                    cbFork2BrakeOption.SelectedIndex = 0;
                }
                else
                {
                    cbOPmachineControl.SelectedIndex = 0;
                    cbEleqFanQuantity.Enabled = true;

                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqModem", cbEleqModem);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftBrakeOption", cbLiftBrakeOption);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravBrakeOption", cbTrav1BrakeOption);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravBrakeOption", cbTrav2BrakeOption);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkBrakeOption", cbFork1BrakeOption);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkBrakeOption", cbFork2BrakeOption);

                    if (cbEleqSensorType.Text == "NPN")
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqNpnSensorItem", cbEleqSensorItem);
                    }
                    else if (cbEleqSensorType.Text == "PNP")
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqPnpSensorItem", cbEleqSensorItem);
                    }
                    if (ckbLiftRaser.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftRaserAbsLocation", cbLiftAbsLocation);
                    }
                    else if (ckbLiftBarcode.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftBarcodeAbsLocation", cbLiftAbsLocation);
                    }
                    if (ckbTrav1Raser.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravRaserAbsLocation", cbTrav1AbsLocation);
                    }
                    else if (ckbTrav1Barcode.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravBarcodeAbsLocation", cbTrav1AbsLocation);
                    }
                    if (ckbTrav2Raser.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravRaserAbsLocation", cbTrav2AbsLocation);
                    }
                    else if (ckbTrav2Barcode.Checked)
                    {
                        pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravBarcodeAbsLocation", cbTrav2AbsLocation);
                    }
                }
            }
            ColdTypeOption_TextChanged(this, EventArgs.Empty);

            cbEleqSensorType.TextChanged += (o, e) =>
            {
                bool coldFlag = false;
                bool npnFlag = false;
                bool pnpFlag = false;
                if(cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C")
                {
                    coldFlag = true;
                }
                if (cbEleqSensorType.Text == "NPN")
                {
                    npnFlag = true;
                }
                if (cbEleqSensorType.Text == "PNP")
                {
                    pnpFlag = true;
                }

                cbEleqSensorItem.Properties.Items.Clear();
                cbLiftRightPosition.Properties.Items.Clear();
                cbTrav1RightPosition.Properties.Items.Clear();
                cbTrav2RightPosition.Properties.Items.Clear();
                cbFork1RightPosition.Properties.Items.Clear();
                cbFork2RightPosition.Properties.Items.Clear();
                cbCarrSensor.Properties.Items.Clear();
                cbCarrDoubleInput.Properties.Items.Clear();

                cbEleqSensorItem.SelectedIndex = -1;
                cbLiftRightPosition.SelectedIndex = -1;
                cbTrav1RightPosition.SelectedIndex = -1;
                cbTrav2RightPosition.SelectedIndex = -1;
                cbFork1RightPosition.SelectedIndex = -1;
                cbFork2RightPosition.SelectedIndex = -1;
                cbCarrSensor.SelectedIndex = -1;
                cbCarrDoubleInput.SelectedIndex = -1;

                if (coldFlag)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listColdEleqModem", cbEleqModem);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listColdEleqSensorItem", cbEleqSensorItem);
                }
                else if(!coldFlag && npnFlag)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqNpnSensorItem", cbEleqSensorItem);
                }
                else if (!coldFlag && pnpFlag)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listEleqPnpSensorItem", cbEleqSensorItem);
                }

                if (npnFlag)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftNpnRightPosition", cbLiftRightPosition);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravNpnRightPosition", cbTrav1RightPosition);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravNpnRightPosition", cbTrav2RightPosition);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkNpnRightPosition", cbFork1RightPosition);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkNpnRightPosition", cbFork2RightPosition);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listCarrNpnSensor", cbCarrSensor);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listCarrNpnDoubleInput", cbCarrDoubleInput);

                }
                else if (pnpFlag)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftPnpRightPosition", cbLiftRightPosition);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravPnpRightPosition", cbTrav1RightPosition);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravPnpRightPosition", cbTrav2RightPosition);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkPnpRightPosition", cbFork1RightPosition);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listForkPnpRightPosition", cbFork2RightPosition);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listCarrPnpSensor", cbCarrSensor);
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listCarrPnpDoubleInput", cbCarrDoubleInput);

                }
            };
            ckbLiftRaser.CheckedChanged += (o, e) =>
            {
                cbLiftAbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbLiftRaser.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listRaserColdLiftAbsLocation", cbLiftAbsLocation);
                }
                else if (ckbLiftRaser.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftRaserAbsLocation", cbLiftAbsLocation);
                }
            };
            ckbLiftBarcode.CheckedChanged += (o, e) =>
            {
                cbLiftAbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbLiftBarcode.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listBarcodeColdLiftAbsLocation", cbLiftAbsLocation);
                }
                else if (ckbLiftBarcode.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listLiftBarcodeAbsLocation", cbLiftAbsLocation);
                }
            };
            ckbTrav1Raser.CheckedChanged += (o, e) =>
            {
                cbTrav1AbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbTrav1Raser.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listRaserColdTravAbsLocation", cbTrav1AbsLocation);
                }
                else if (ckbTrav1Raser.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravRaserAbsLocation", cbTrav1AbsLocation);
                }
            };
            ckbTrav1Barcode.CheckedChanged += (o, e) =>
            {
                cbTrav1AbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbTrav1Barcode.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listBarcodeColdTravAbsLocation", cbTrav1AbsLocation);
                }
                else if (ckbTrav1Barcode.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravBarcodeAbsLocation", cbTrav1AbsLocation);
                }
            };
            ckbTrav2Raser.CheckedChanged += (o, e) =>
            {
                cbTrav2AbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbTrav2Raser.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listRaserColdTravAbsLocation", cbTrav2AbsLocation);
                }
                else if (ckbTrav2Raser.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravRaserAbsLocation", cbTrav2AbsLocation);
                }
            };
            ckbTrav2Barcode.CheckedChanged += (o, e) =>
            {
                cbTrav2AbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbTrav2Barcode.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listBarcodeColdTravAbsLocation", cbTrav2AbsLocation);
                }
                else if (ckbTrav2Barcode.Checked)
                {
                    pathManager.LoadListFromXmlToComboBox(StringUnits.strItemListFilePath, "listTravBarcodeAbsLocation", cbTrav2AbsLocation);
                }
            };
            
            ComboBoxEdit[] cbModular = { cbEleqPowerKw, cbEleqPowerA, cbEleqBrakeResistorKw, cbEleqBrakeResistorOhm, cbEleqMccbSpec,cbEleqPowerCable };
            ComboBoxEdit[] cbSystem =
                {
                    cbLiftBrakeResistorKw, cbLiftBrakeResistorOhm, cbLiftMccbSpec,
                    cbTrav1BrakeResistorKw, cbTrav1BrakeResistorOhm, cbTrav1MccbSpec,
                    cbTrav2BrakeResistorKw, cbTrav2BrakeResistorOhm, cbTrav2MccbSpec,
                    cbFork1BrakeResistorKw, cbFork1BrakeResistorOhm, cbFork1MccbSpec,
                    cbFork2BrakeResistorKw, cbFork2BrakeResistorOhm, cbFork2MccbSpec
                };
            ComboBoxEdit[] cbRegenerative = new ComboBoxEdit[] 
                {
                    cbEleqBrakeResistorKw, cbEleqBrakeResistorOhm,
                    cbLiftBrakeResistorKw, cbLiftBrakeResistorOhm,
                    cbTrav1BrakeResistorKw, cbTrav1BrakeResistorOhm,
                    cbFork1BrakeResistorKw, cbFork1BrakeResistorOhm,
                    cbFork2BrakeResistorKw, cbFork2BrakeResistorOhm,
                };
            
            cbMODoption1.TextChanged += (o, e) => ChangeEnableState();
            cbMODoption2.TextChanged += (o, e) => ChangeEnableState();
            cbMODoption3.TextChanged += (o, e) => ChangeEnableState();
            cbMODoption4.TextChanged += (o, e) => ChangeEnableState();
            ckbTravDoubleMotorTrue.CheckedChanged += (o, e) => ChangeEnableState();
            ckbMODforkoption.CheckedChanged += (o, e) => ChangeEnableState();
            ckbRegenerativeUnitTrue.CheckedChanged += (o, e) => ChangeEnableState();
            cbMSPinverterSpec.TextChanged += (o, e) => ChangeEnableState();
            void ChangeEnableState()
            {
                bool trav2Flag = false;
                bool fork2Flag = false;
                bool coldFlag = false;
                bool mainTenanceFlag = false;
                bool regenerativeFlag = false;

                if (ckbTravDoubleMotorTrue.Checked)
                {
                    trav2Flag = true;
                }
                if (ckbMODforkoption.Checked || cbMODoption1.Text == "D(v)" || cbMODoption2.Text == "D(v)" || cbMODoption3.Text == "D(v)" || cbMODoption4.Text == "D(v)")
                {
                    fork2Flag = true;
                }
                if (cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C")
                {
                    coldFlag = true;
                }
                if (cbMODoption1.Text == "M" || cbMODoption2.Text == "M" || cbMODoption3.Text == "M" || cbMODoption4.Text == "M")
                {
                    mainTenanceFlag = true;
                }
                if (ckbRegenerativeUnitTrue.Checked)
                {
                    regenerativeFlag = true;
                }

                ChangeStateByKey(trav2Flag, "TRAV2");
                ChangeStateByKey(fork2Flag, "FORK2");
                ChangeStateByText(cbMSPinverterSpec, "MODULAR", cbSystem);
                ChangeStateByText(cbMSPinverterSpec, "SYSTEM", cbModular);
                ChangeStateByCheckEdit(regenerativeFlag, cbRegenerative);
                SetFuncDefault();

                
            }

            SetMainSpecDefault();
            SetOptionDefault();
            SetDateDefault();
            SetFuncDefault();

            btnMSPdefault.Click += (o, e) => SetMainSpecDefault();
            btnOPdefault.Click += (o, e) => SetOptionDefault();
            btnPrjDateNow.Click += (o, e) => SetDateDefault();

            void SetMainSpecDefault()
            {
                //주요사양 Default값
                foreach (var comboBox in new[] { cbMSPinputVolt, cbMSPusingVoltage, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbMSPpanelSize })
                {
                    comboBox.SelectedIndex = -1;
                    comboBox.Text = string.Empty;
                    comboBox.SelectedIndex = 0;
                }
            }
            void SetOptionDefault()
            {
                //옵션 Default값
                foreach (var comboBox in new[] { cbOPmachineControl, cbOPremoteControl, cbOPemergencyPower, cbOPemergencyLocation })
                {
                    comboBox.SelectedIndex = -1;
                    comboBox.Text = string.Empty;
                    comboBox.SelectedIndex = 0;
                }
            }
            void SetDateDefault()
            {
                cbPRJyear.Text = DateTime.Now.ToString("yyyy");
                cbPRJmonth.Text = DateTime.Now.ToString("MM");
                cbPRJday.Text = DateTime.Now.ToString("dd");

            }
            void SetFuncDefault()
            {
                foreach (ComboBoxEdit comboBox in new[] { cbEleqHubModel, cbEleqFanQuantity, cbEleqTerminal, cbEleqPanel, cbEleqHmi, cbEleqOpt, cbEleqTowerLamp, cbEleqSafety, cbEleqSafetyQuantity, cbEleqSensorType })
                {
                    comboBox.SelectedIndex = -1;
                    comboBox.Text = string.Empty;
                    comboBox.SelectedIndex = 0;
                }
                foreach (ComboBoxEdit comboBox in new[] { cbEleqMccbModel, cbEleqSmpsModel, cbEleqCableModel, cbEleqHubModel, cbLiftBrakeOption, cbTrav1BrakeOption, cbFork1BrakeOption, cbTrav2BrakeOption, cbFork2BrakeOption })
                {
                    comboBox.SelectedIndex = -1;
                    comboBox.Text = string.Empty;
                    comboBox.SelectedIndex = 0;
                }
            }
            void ChangeStateByKey(bool ckbState, string strFunc)
            {
                foreach (var key in DictionaryUnits.dicCtrlFunc.Keys)
                {
                    if (key.Contains(strFunc)) // strFunc가 포함된 키 찾기
                    {
                        foreach (var control in DictionaryUnits.dicCtrlFunc[key])
                        {
                            control.Enabled = ckbState; // 컨트롤 활성화
                        }
                    }
                }
                
            }
            void ChangeStateByCheckEdit(bool ckbState, Control[] targetCtrls)
            {
                if (ckbState)
                {
                    foreach (Control ctrl in targetCtrls)
                    {
                        ctrl.Enabled = !ckbState;
                    }
                }
                
                
            }
            void ChangeStateByText(Control ctrlName, string matchingName, Control[] targetCtrls)
            {
                bool enableState = ctrlName.Text == matchingName;
                
                foreach (Control ctrl in targetCtrls)
                {
                    ctrl.Enabled = !enableState;
                }
            }

        }
        private void UpdateComboBoxItemList()
        {
            if (watcher == null) // watcher가 아직 생성되지 않은 경우에만 생성
            {
                watcher = new FileSystemWatcher();
                watcher.Path = Path.GetDirectoryName(StringUnits.strItemListFilePath);
                watcher.Filter = Path.GetFileName(StringUnits.strItemListFilePath);
                watcher.NotifyFilter = NotifyFilters.LastWrite;

                // 변경 이벤트 핸들러 추가
                watcher.Changed += (o, e) =>
                {
                    if (e.ChangeType == WatcherChangeTypes.Changed)
                    {
                        // UI 스레드에서 실행되도록 Invoke 사용
                        this.Invoke((MethodInvoker)delegate
                        {
                            try
                            {
                                // XML 변경 감지 시 해당 그리드 업데이트
                                SetComboBoxItems();

                                // UI 갱신
                                Application.DoEvents();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"데이터 업데이트 중 오류 발생: {ex.Message}");
                            }
                        });
                    }
                };
                watcher.EnableRaisingEvents = true; // 이벤트 활성화
            }
        }

        
    }
}