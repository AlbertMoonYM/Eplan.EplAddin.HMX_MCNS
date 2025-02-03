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
using Eplan.MCNS.Lib.Share_CS;
using Eplan.MCNS.Lib.UI_CS;
using System.Reflection.Emit;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid;
using DevExpress.XtraPrinting.Preview;
using Eplan.EplAddin.HMX_MCNS._03.UI;
using System.Diagnostics;
using System.Xml.Linq;
using ClosedXML.Excel;
using DevExpress.XtraEditors.Repository;
using McnsSchemGenEngine.Models;
using OfficeOpenXml;
using DevExpress.XtraTab;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;



namespace Eplan.EplAddin.HMX_MCNS
{
    public partial class FormConceptSheet : DevExpress.XtraEditors.XtraForm
    {
       

        //외부 CS 정의
        //외부 CS 정의
        CS_Label cs_Label = new CS_Label();
        CS_ComboBox cs_ComboBox = new CS_ComboBox();
        CS_DataTable cs_DataTable = new CS_DataTable();
        CS_DataGridView cs_DataGrid = new CS_DataGridView();
        CS_GroupControl cs_GroupControl = new CS_GroupControl();
        CS_XtraTabControl cs_XtraTabControl = new CS_XtraTabControl();
        CS_CheckBox cs_CheckBox = new CS_CheckBox();
        CS_InterLock interLock = new CS_InterLock();
        CS_ListItems cs_ListItems = new CS_ListItems();
        FileSystemWatcher watcher; // FileSystemWatcher 선언
        ToolTip tip = new ToolTip();
        McnsSchemGenEngine.Controls.McnsControl mcnsControl = new McnsSchemGenEngine.Controls.McnsControl();

        //전역 변수 설정
        private DataTable excelDt = new DataTable();

        public FormConceptSheet()
        {
            InitializeComponent();
            

            SetPanel3D();
            SetComboBoxFunction();
            SetComboBoxItems();
            SetFunctionPageData();
            SetToolTip();

            ControlFormFunction();
            ControlPlcFunction();
            LoadIoFromExcel();
            ActivateEplan();

            UpdateComboBoxItemList();
            SetComboBoxDefaultValue();
            Interlock();
        }
        public void SetComboBoxFunction()
        {

            // button을 Radio button으로 사용
            cs_CheckBox.ChangeToRadioButton(ckbPRJdomestic, ckbPRJoverseas);
            // 프로젝트 ChangeToTextBox 설정
            cs_ComboBox.ChangeToTextBox(cbPRJnumber, TypeFlag.strFlag, 30, "텍스트를 기입하세요.","");
            cs_ComboBox.ChangeToTextBox(cbPRJname, TypeFlag.strFlag, 30, "텍스트를 기입하세요.", "");
            cs_ComboBox.ChangeToTextBox(cbPRJwriter, TypeFlag.strFlag, 10, "텍스트를 기입하세요.", "");
            // 프로젝트 SettingComboBox 설정
            cs_ComboBox.SettingComboBox(cbPRJyear, "-");
            cs_ComboBox.SettingComboBox(cbPRJmonth, "-");
            cs_ComboBox.SettingComboBox(cbPRJday, "-");
            int currentYear = DateTime.Now.Year;
            // 연도 추가 (현재 연도를 기준으로 ±10년)
            cbPRJyear.Properties.Items.AddRange(
                Enumerable.Range(currentYear - 10, 21).Select(y => y.ToString()).ToArray()
            );

            // 월 추가 (1~12)
            cbPRJmonth.Properties.Items.AddRange(
                Enumerable.Range(1, 12).Select(m => m.ToString("D2")).ToArray()
            );

            // 일 추가 (1~31)
            cbPRJday.Properties.Items.AddRange(
                Enumerable.Range(1, 31).Select(d => d.ToString("D2")).ToArray());

            // 모델 ChangeToTextBox 설정
            cs_ComboBox.ChangeToTextBox(cbMODheight, TypeFlag.intFlag, 10, "높이", "");
            cs_ComboBox.ChangeToTextBox(cbMODweight, TypeFlag.intFlag, 10, "화물 중량", "");
            cs_ComboBox.ChangeToTextBox(cbMODfullName, TypeFlag.strFlag, 0, "", "");
            // 모델 SettingComboBox 설정
            cs_ComboBox.SettingComboBox(cbMODname, "모델명");
            cs_ComboBox.SettingComboBox(cbMODoption1, "-");
            cs_ComboBox.SettingComboBox(cbMODoption2, "-");
            cs_ComboBox.SettingComboBox(cbMODoption3, "-");
            cs_ComboBox.SettingComboBox(cbMODoption4, "-");


            // 주요 사양 ChangeToTextBox 설정
            cs_ComboBox.ChangeToTextBox(cbMSPpanelSizeW, TypeFlag.intFlag, 4, "-", "");
            cs_ComboBox.ChangeToTextBox(cbMSPpanelSizeD, TypeFlag.intFlag, 4, "-", "");
            cs_ComboBox.ChangeToTextBox(cbMSPpanelSizeH, TypeFlag.intFlag, 4, "-", "");
            // 주요 사양 SettingComboBox 설정
            cs_ComboBox.SettingComboBox(cbMSPinputVolt, "-");
            cs_ComboBox.SettingComboBox(cbMSPinputHz, "-");
            cs_ComboBox.SettingComboBox(cbMSPpanelSize, "-");
            cs_ComboBox.SettingComboBox(cbMSPcontrollerSpec, "-");
            cs_ComboBox.SettingComboBox(cbMSPinverterMaker, "-");
            cs_ComboBox.SettingComboBox(cbMSPinverterSpec, "-");

            // 레이아웃 ChangeToTextBox 설정
            cs_ComboBox.ChangeToTextBox(cbLOUTtravLength, TypeFlag.fltFlag, 10, "실수 기입", "M");
            cs_ComboBox.ChangeToTextBox(cbLOUTliftHeight, TypeFlag.fltFlag, 10, "실수 기입", "M");
            cs_ComboBox.ChangeToTextBox(cbLOUTstationNum, TypeFlag.intFlag, 1, "정수 기입", "");

            // 레이아웃 화물 DataGridView 셋업
            cs_DataTable.GetDataTable(CS_StaticUnit.dtLout, CS_StaticString.dArrDtLoutColums);
            // 레이아웃 화물 DataTable 3행 셋업
            CS_StaticUnit.dtLout.Rows.Add("화물1");
            CS_StaticUnit.dtLout.Rows.Add("화물2");
            CS_StaticUnit.dtLout.Rows.Add("화물3");
            CS_StaticUnit.dtLout.Rows.Add("화물4");
            gridLOUTcargo.DataSource = CS_StaticUnit.dtLout;
            cs_DataGrid.SetLoutCargo(gridViewCargo);

            // 옵션 SettingComboBox 설정
            cs_ComboBox.SettingComboBox(cbOPmachineControl, "-");
            cs_ComboBox.SettingComboBox(cbOPremoteControl, "-");
            cs_ComboBox.SettingComboBox(cbOPemergencyPower, "-");
            cs_ComboBox.SettingComboBox(cbOPemergencyLocation, "-");


            // ELEQ 기능 그룹
            // ChangeToTextBox 호출
            cs_ComboBox.ChangeToTextBox(cbEleqPowerKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbEleqPowerA, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbEleqBrakeResistorKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbEleqBrakeResistorOhm, TypeFlag.fltFlag, 10, "실수 기입", "Ω");
            cs_ComboBox.ChangeToTextBox(cbEleqPowerCable, TypeFlag.fltFlag, 10, "실수 기입", "SQ");
            cs_ComboBox.ChangeToTextBox(cbEleqMccbSpec, TypeFlag.strFlag, 10, "텍스트 기입", "");

            // SettingComboBox 호출
            cs_ComboBox.SettingComboBox(cbMSPusingVoltage, "-");
            cs_ComboBox.SettingComboBox(cbEleqMccbModel, "-");
            cs_ComboBox.SettingComboBox(cbEleqSmpsModel, "-");
            cs_ComboBox.SettingComboBox(cbEleqCableModel, "-");
            cs_ComboBox.SettingComboBox(cbEleqHubModel, "-");
            cs_ComboBox.SettingComboBox(cbEleqFanQuantity, "-");
            cs_ComboBox.SettingComboBox(cbEleqTerminal, "-");

            cs_ComboBox.SettingComboBox(cbEleqPanel, "-");
            cs_ComboBox.SettingComboBox(cbEleqHmi, "-");
            cs_ComboBox.SettingComboBox(cbEleqOpt, "-");
            cs_ComboBox.SettingComboBox(cbEleqTowerLamp, "-");
            cs_ComboBox.SettingComboBox(cbEleqSafety, "-");
            cs_ComboBox.SettingComboBox(cbEleqSafetyQuantity, "-");

            cs_ComboBox.SettingComboBox(cbEleqSensorType, "-");
            cs_ComboBox.SettingComboBox(cbEleqModem, "-");
            cs_ComboBox.SettingComboBox(cbEleqInterLockSensorSide, "-");
            cs_ComboBox.SettingComboBox(cbEleqInterLockBit, "-");
            cs_ComboBox.SettingComboBox(cbEleqLocation, "-");
            cs_ComboBox.SettingComboBox(cbEleqType, "-");
            cs_ComboBox.SettingComboBox(cbEleqDt, "-");
            cs_ComboBox.SettingComboBox(cbEleqParts, "-");
            cs_ComboBox.SettingComboBox(cbEleqPoint, "-");
            cs_ComboBox.SettingComboBox(cbEleqSensorItem, "-");


            // LIFT 기능 그룹
            cs_ComboBox.ChangeToTextBox(cbLiftInverterKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbLiftInverterA, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbLiftBrakeResistorKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbLiftBrakeResistorOhm, TypeFlag.fltFlag, 10, "실수 기입", "Ω");
            cs_ComboBox.ChangeToTextBox(cbLiftPowerCable, TypeFlag.fltFlag, 10, "실수 기입", "SQ");
            cs_ComboBox.ChangeToTextBox(cbLiftMccbSpec, TypeFlag.strFlag, 6, "텍스트 기입", "");

            cs_ComboBox.ChangeToTextBox(cbLiftOutPut, TypeFlag.fltFlag, 10, "실수 기입", "kW");
            cs_ComboBox.ChangeToTextBox(cbLiftSpeed, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbLiftGearRatio, TypeFlag.fltFlag, 10, "실수 기입", "i");
            cs_ComboBox.ChangeToTextBox(cbLiftRatedCurrent, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbLiftBkVoltage, TypeFlag.fltFlag, 10, "실수 기입", "V");
            cs_ComboBox.SettingComboBox(cbLiftBrakeOption, "-");
            cs_ComboBox.SettingComboBox(cbLiftMotorSpec, "-");
            cs_ComboBox.SettingComboBox(cbLiftMotorMaker, "-");
            cs_ComboBox.SettingComboBox(cbLiftMotorMethod, "-");
            cs_ComboBox.ChangeToTextBox(cbLiftMotorType, TypeFlag.strFlag, 10, "텍스트 기입", "");
            cs_ComboBox.ChangeToTextBox(cbLiftMotorVoltage, TypeFlag.fltFlag, 10, "실수 기입", "V");
            cs_ComboBox.ChangeToTextBox(cbLiftMotorHz, TypeFlag.fltFlag, 10, "실수 기입", "Hz");
            cs_ComboBox.ChangeToTextBox(cbLiftMotorEncoderSpec, TypeFlag.strFlag, 10, "텍스트 기입", "");

            cs_ComboBox.SettingComboBox(cbLiftAbsLocation, "-");
            cs_ComboBox.SettingComboBox(cbLiftRightPosition, "-");
            cs_ComboBox.SettingComboBox(cbLiftLimitSwitch, "-");

            cs_ComboBox.ChangeToTextBox(cbLiftNoneLoadHighSpeed, TypeFlag.fltFlag, 10, "실수 기입", "MPM");
            cs_ComboBox.ChangeToTextBox(cbLiftNoneLoadRotationNum, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbLiftNoneLoadAcceleration, TypeFlag.fltFlag, 10, "실수 기입", "");
            cs_ComboBox.ChangeToTextBox(cbLiftLoadHighSpeed, TypeFlag.fltFlag, 10, "실수 기입", "MPM");
            cs_ComboBox.ChangeToTextBox(cbLiftLoadRotationNum, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbLiftLoadAcceleration, TypeFlag.fltFlag, 10, "실수 기입", "");


            // TRAV1 기능 그룹
            cs_ComboBox.ChangeToTextBox(cbTrav1InverterKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbTrav1InverterA, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbTrav1BrakeResistorKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbTrav1BrakeResistorOhm, TypeFlag.fltFlag, 10, "실수 기입", "Ω");
            cs_ComboBox.ChangeToTextBox(cbTrav1PowerCable, TypeFlag.fltFlag, 10, "실수 기입", "SQ");
            cs_ComboBox.ChangeToTextBox(cbTrav1MccbSpec, TypeFlag.strFlag, 6, "텍스트 기입", "");

            cs_ComboBox.ChangeToTextBox(cbTrav1OutPut, TypeFlag.fltFlag, 10, "실수 기입", "kW");
            cs_ComboBox.ChangeToTextBox(cbTrav1Speed, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbTrav1GearRatio, TypeFlag.fltFlag, 10, "실수 기입", "i");
            cs_ComboBox.ChangeToTextBox(cbTrav1RatedCurrent, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbTrav1BkVoltage, TypeFlag.fltFlag, 10, "실수 기입", "V");
            cs_ComboBox.SettingComboBox(cbTrav1BrakeOption, "-");
            cs_ComboBox.SettingComboBox(cbTrav1MotorSpec, "-");
            cs_ComboBox.SettingComboBox(cbTrav1MotorMaker, "-");
            cs_ComboBox.SettingComboBox(cbTrav1MotorMethod, "-");
            cs_ComboBox.ChangeToTextBox(cbTrav1MotorType, TypeFlag.strFlag, 10, "텍스트 기입", "");
            cs_ComboBox.ChangeToTextBox(cbTrav1MotorVoltage, TypeFlag.fltFlag, 10, "실수 기입", "V");
            cs_ComboBox.ChangeToTextBox(cbTrav1MotorHz, TypeFlag.fltFlag, 10, "실수 기입", "Hz");
            cs_ComboBox.ChangeToTextBox(cbTrav1MotorEncoderSpec, TypeFlag.strFlag, 10, "텍스트 기입", "");

            cs_ComboBox.SettingComboBox(cbTrav1AbsLocation, "-");
            cs_ComboBox.SettingComboBox(cbTrav1RightPosition, "-");
            cs_ComboBox.SettingComboBox(cbTrav1LimitSwitch, "-");

            cs_ComboBox.ChangeToTextBox(cbTrav1NoneLoadHighSpeed, TypeFlag.fltFlag, 10, "실수 기입", "MPM");
            cs_ComboBox.ChangeToTextBox(cbTrav1NoneLoadRotationNum, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbTrav1NoneLoadAcceleration, TypeFlag.fltFlag, 10, "실수 기입", "");
            cs_ComboBox.ChangeToTextBox(cbTrav1LoadHighSpeed, TypeFlag.fltFlag, 10, "실수 기입", "MPM");
            cs_ComboBox.ChangeToTextBox(cbTrav1LoadRotationNum, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbTrav1LoadAcceleration, TypeFlag.fltFlag, 10, "실수 기입", "");




            // Trav2 기능 그룹
            cs_ComboBox.ChangeToTextBox(cbTrav2InverterKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbTrav2InverterA, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbTrav2BrakeResistorKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbTrav2BrakeResistorOhm, TypeFlag.fltFlag, 10, "실수 기입", "Ω");
            cs_ComboBox.ChangeToTextBox(cbTrav2PowerCable, TypeFlag.fltFlag, 10, "실수 기입", "SQ");
            cs_ComboBox.ChangeToTextBox(cbTrav2MccbSpec, TypeFlag.strFlag, 6, "텍스트 기입", "");

            cs_ComboBox.ChangeToTextBox(cbTrav2OutPut, TypeFlag.fltFlag, 10, "실수 기입", "kW");
            cs_ComboBox.ChangeToTextBox(cbTrav2Speed, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbTrav2GearRatio, TypeFlag.fltFlag, 10, "실수 기입", "i");
            cs_ComboBox.ChangeToTextBox(cbTrav2RatedCurrent, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbTrav2BkVoltage, TypeFlag.fltFlag, 10, "실수 기입", "V");
            cs_ComboBox.SettingComboBox(cbTrav2BrakeOption, "-");
            cs_ComboBox.SettingComboBox(cbTrav2MotorSpec, "-");
            cs_ComboBox.SettingComboBox(cbTrav2MotorMaker, "-");
            cs_ComboBox.SettingComboBox(cbTrav2MotorMethod, "-");
            cs_ComboBox.ChangeToTextBox(cbTrav2MotorType, TypeFlag.strFlag, 10, "텍스트 기입", "");
            cs_ComboBox.ChangeToTextBox(cbTrav2MotorVoltage, TypeFlag.fltFlag, 10, "실수 기입", "V");
            cs_ComboBox.ChangeToTextBox(cbTrav2MotorHz, TypeFlag.fltFlag, 10, "실수 기입", "Hz");
            cs_ComboBox.ChangeToTextBox(cbTrav2MotorEncoderSpec, TypeFlag.strFlag, 10, "텍스트 기입", "");

            cs_ComboBox.SettingComboBox(cbTrav2AbsLocation, "-");
            cs_ComboBox.SettingComboBox(cbTrav2RightPosition, "-");
            cs_ComboBox.SettingComboBox(cbTrav2LimitSwitch, "-");

            cs_ComboBox.ChangeToTextBox(cbTrav2NoneLoadHighSpeed, TypeFlag.fltFlag, 10, "실수 기입", "MPM");
            cs_ComboBox.ChangeToTextBox(cbTrav2NoneLoadRotationNum, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbTrav2NoneLoadAcceleration, TypeFlag.fltFlag, 10, "실수 기입", "");
            cs_ComboBox.ChangeToTextBox(cbTrav2LoadHighSpeed, TypeFlag.fltFlag, 10, "실수 기입", "MPM");
            cs_ComboBox.ChangeToTextBox(cbTrav2LoadRotationNum, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbTrav2LoadAcceleration, TypeFlag.fltFlag, 10, "실수 기입", "");



            // FORK1 기능 그룹
            cs_ComboBox.ChangeToTextBox(cbFork1InverterKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbFork1InverterA, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbFork1BrakeResistorKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbFork1BrakeResistorOhm, TypeFlag.fltFlag, 10, "실수 기입", "Ω");
            cs_ComboBox.ChangeToTextBox(cbFork1PowerCable, TypeFlag.fltFlag, 10, "실수 기입", "SQ");
            cs_ComboBox.ChangeToTextBox(cbFork1MccbSpec, TypeFlag.strFlag, 6, "텍스트 기입", "");

            cs_ComboBox.ChangeToTextBox(cbFork1OutPut, TypeFlag.fltFlag, 10, "실수 기입", "kW");
            cs_ComboBox.ChangeToTextBox(cbFork1Speed, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbFork1GearRatio, TypeFlag.fltFlag, 10, "실수 기입", "i");
            cs_ComboBox.ChangeToTextBox(cbFork1RatedCurrent, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbFork1BkVoltage, TypeFlag.fltFlag, 10, "실수 기입", "V");
            cs_ComboBox.SettingComboBox(cbFork1BrakeOption, "-");
            cs_ComboBox.SettingComboBox(cbFork1MotorSpec, "-");
            cs_ComboBox.SettingComboBox(cbFork1MotorMaker, "-");
            cs_ComboBox.SettingComboBox(cbFork1MotorMethod, "-");
            cs_ComboBox.ChangeToTextBox(cbFork1MotorType, TypeFlag.strFlag, 10, "텍스트 기입", "");
            cs_ComboBox.ChangeToTextBox(cbFork1MotorVoltage, TypeFlag.fltFlag, 10, "실수 기입", "V");
            cs_ComboBox.ChangeToTextBox(cbFork1MotorHz, TypeFlag.fltFlag, 10, "실수 기입", "Hz");
            cs_ComboBox.ChangeToTextBox(cbFork1MotorEncoderSpec, TypeFlag.strFlag, 10, "텍스트 기입", "");

            cs_ComboBox.SettingComboBox(cbFork1RightPosition, "-");

            cs_ComboBox.ChangeToTextBox(cbFork1NoneLoadHighSpeed, TypeFlag.fltFlag, 10, "실수 기입", "MPM");
            cs_ComboBox.ChangeToTextBox(cbFork1NoneLoadRotationNum, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbFork1NoneLoadAcceleration, TypeFlag.fltFlag, 10, "실수 기입", "");
            cs_ComboBox.ChangeToTextBox(cbFork1LoadHighSpeed, TypeFlag.fltFlag, 10, "실수 기입", "MPM");
            cs_ComboBox.ChangeToTextBox(cbFork1LoadRotationNum, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbFork1LoadAcceleration, TypeFlag.fltFlag, 10, "실수 기입", "");

            // FORK2 기능 그룹
            // ChangeToTextBox 호출
            cs_ComboBox.ChangeToTextBox(cbFork2InverterKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbFork2InverterA, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbFork2BrakeResistorKw, TypeFlag.fltFlag, 10, "실수 기입", "Kw");
            cs_ComboBox.ChangeToTextBox(cbFork2BrakeResistorOhm, TypeFlag.fltFlag, 10, "실수 기입", "Ω");
            cs_ComboBox.ChangeToTextBox(cbFork2PowerCable, TypeFlag.fltFlag, 10, "실수 기입", "SQ");
            cs_ComboBox.ChangeToTextBox(cbFork2MccbSpec, TypeFlag.strFlag, 6, "텍스트 기입", "" );

            cs_ComboBox.ChangeToTextBox(cbFork2OutPut, TypeFlag.fltFlag, 10, "실수 기입", "kW");
            cs_ComboBox.ChangeToTextBox(cbFork2Speed, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbFork2GearRatio, TypeFlag.fltFlag, 10, "실수 기입", "i");
            cs_ComboBox.ChangeToTextBox(cbFork2RatedCurrent, TypeFlag.fltFlag, 10, "실수 기입", "A");
            cs_ComboBox.ChangeToTextBox(cbFork2BkVoltage, TypeFlag.fltFlag, 10, "실수 기입", "V");
            cs_ComboBox.SettingComboBox(cbFork2BrakeOption, "-");
            cs_ComboBox.SettingComboBox(cbFork2MotorSpec, "-");
            cs_ComboBox.SettingComboBox(cbFork2MotorMaker, "-");
            cs_ComboBox.SettingComboBox(cbFork2MotorMethod, "-");
            cs_ComboBox.ChangeToTextBox(cbFork2MotorType, TypeFlag.strFlag, 10, "텍스트 기입", "");
            cs_ComboBox.ChangeToTextBox(cbFork2MotorVoltage, TypeFlag.fltFlag, 10, "실수 기입", "V");
            cs_ComboBox.ChangeToTextBox(cbFork2MotorHz, TypeFlag.fltFlag, 10, "실수 기입", "Hz");
            cs_ComboBox.ChangeToTextBox(cbFork2MotorEncoderSpec, TypeFlag.strFlag, 10, "텍스트 기입", "");

            cs_ComboBox.SettingComboBox(cbFork2RightPosition, "-");

            cs_ComboBox.ChangeToTextBox(cbFork2NoneLoadHighSpeed, TypeFlag.fltFlag, 10, "실수 기입", "MPM");
            cs_ComboBox.ChangeToTextBox(cbFork2NoneLoadRotationNum, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbFork2NoneLoadAcceleration, TypeFlag.fltFlag, 10, "실수 기입", "");
            cs_ComboBox.ChangeToTextBox(cbFork2LoadHighSpeed, TypeFlag.fltFlag, 10, "실수 기입", "MPM");
            cs_ComboBox.ChangeToTextBox(cbFork2LoadRotationNum, TypeFlag.fltFlag, 10, "실수 기입", "RPM");
            cs_ComboBox.ChangeToTextBox(cbFork2LoadAcceleration, TypeFlag.fltFlag, 10, "실수 기입", "");

            // CARR 기능 그룹
            // SettingComboBox 호출
            cs_ComboBox.SettingComboBox(cbCarrSensor, "-");
            cs_ComboBox.SettingComboBox(cbCarrDoubleInput, "-");




        }
        public void SetPanel3D()
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
        public void SetComboBoxItems()
        {
            // 모델 콤보 박스 리스트 설정
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMODName", cbMODname);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMODOption", cbMODoption1);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMODOption", cbMODoption2);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMODOption", cbMODoption3);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMODOption", cbMODoption4);

            // 주요 사양 콤보 박스 리스트 설정
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMSPinputVolt", cbMSPinputVolt);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMSPinputHz", cbMSPinputHz);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMSPpanelSize", cbMSPpanelSize);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMSPcontrollerSpec", cbMSPcontrollerSpec);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMSPinverterMaker", cbMSPinverterMaker);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listMSPinverterSpec", cbMSPinverterSpec);

            // 옵션 콤보 박스 리스트 설정
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listOPmachineControl", cbOPmachineControl);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listOPremoteControl", cbOPremoteControl);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listOPemergencyPower", cbOPemergencyPower);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listOPemergencyLocation", cbOPemergencyLocation);

            // ComboBox List 설정
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqUsingVoltage", cbMSPusingVoltage);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqMccbModel", cbEleqMccbModel);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqSmpsModel", cbEleqSmpsModel);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqCableModel", cbEleqCableModel);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqHubModel", cbEleqHubModel);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqFanQuantity", cbEleqFanQuantity);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqTerminal", cbEleqTerminal);

            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqPanel", cbEleqPanel);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqHmi", cbEleqHmi);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqOpt", cbEleqOpt);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqTowerLamp", cbEleqTowerLamp);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqSafety", cbEleqSafety);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqSafetyQuantity", cbEleqSafetyQuantity);

            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqSensorType", cbEleqSensorType);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqInterLockSensorSide", cbEleqInterLockSensorSide);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqInterLockBit", cbEleqInterLockBit);


            // ComboBox List 설정
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftMotorSpec", cbLiftMotorSpec);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftMotorMaker", cbLiftMotorMaker);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftMotorMethod", cbLiftMotorMethod);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftLimitSwitch", cbLiftLimitSwitch);
           

            // ComboBox List 설정
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravMotorSpec", cbTrav1MotorSpec);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravMotorMaker", cbTrav1MotorMaker);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravMotorMethod", cbTrav1MotorMethod);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravLimitSwitch", cbTrav1LimitSwitch);

            // ComboBox List 설정
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravMotorSpec", cbTrav2MotorSpec);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravMotorMaker", cbTrav2MotorMaker);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravMotorMethod", cbTrav2MotorMethod);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravLimitSwitch", cbTrav2LimitSwitch);

            // ComboBox List 설정
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkMotorSpec", cbFork1MotorSpec);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkMotorMaker", cbFork1MotorMaker);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkMotorMethod", cbFork1MotorMethod);


            // ComboBox List 설정
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkMotorSpec", cbFork2MotorSpec);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkMotorMaker", cbFork2MotorMaker);
            cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkMotorMethod", cbFork2MotorMethod);

            cbEleqSensorType.TextChanged += (o, e) =>
            {
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

                if (cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C")
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listColdEleqModem", cbEleqModem);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listColdEleqSensorItem", cbEleqSensorItem);
                }
                else if (cbEleqSensorType.Text == "NPN")
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqNpnSensorItem", cbEleqSensorItem);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftNpnRightPosition", cbLiftRightPosition);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravNpnRightPosition", cbTrav1RightPosition);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravNpnRightPosition", cbTrav2RightPosition);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkNpnRightPosition", cbFork1RightPosition);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkNpnRightPosition", cbFork2RightPosition);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listCarrNpnSensor", cbCarrSensor);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listCarrNpnDoubleInput", cbCarrDoubleInput);

                }
                else if (cbEleqSensorType.Text == "PNP")
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqPnpSensorItem", cbEleqSensorItem);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftPnpRightPosition", cbLiftRightPosition);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravPnpRightPosition", cbTrav1RightPosition);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravPnpRightPosition", cbTrav2RightPosition);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkPnpRightPosition", cbFork1RightPosition);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkPnpRightPosition", cbFork2RightPosition);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listCarrPnpSensor", cbCarrSensor);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listCarrPnpDoubleInput", cbCarrDoubleInput);

                }
            };

            // 콜드 타입 로직
            cbMODoption1.TextChanged += ColdTypeOption_TextChanged;
            cbMODoption2.TextChanged += ColdTypeOption_TextChanged;
            cbMODoption3.TextChanged += ColdTypeOption_TextChanged;
            cbMODoption4.TextChanged += ColdTypeOption_TextChanged;

            ckbLiftRaser.CheckedChanged += (o, e) =>
            {
                cbLiftAbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C")&& ckbLiftRaser.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listRaserColdLiftAbsLocation", cbLiftAbsLocation);
                }
                else if (ckbLiftRaser.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftRaserAbsLocation", cbLiftAbsLocation);
                }
            };
            ckbLiftBarcode.CheckedChanged += (o, e) =>
            {
                cbLiftAbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbLiftBarcode.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listBarcodeColdLiftAbsLocation", cbLiftAbsLocation);
                }
                else if (ckbLiftBarcode.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftBarcodeAbsLocation", cbLiftAbsLocation);
                }
            };
            ckbTrav1Raser.CheckedChanged += (o, e) =>
            {
                cbTrav1AbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbTrav1Raser.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listRaserColdTravAbsLocation", cbTrav1AbsLocation);
                }
                else if (ckbTrav1Raser.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravRaserAbsLocation", cbTrav1AbsLocation);
                }
            };
            ckbTrav1Barcode.CheckedChanged += (o, e) =>
            {
                cbTrav1AbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbTrav1Barcode.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listBarcodeColdTravAbsLocation", cbTrav1AbsLocation);
                }
                else if (ckbTrav1Barcode.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravBarcodeAbsLocation", cbTrav1AbsLocation);
                }
            };
            ckbTrav2Raser.CheckedChanged += (o, e) =>
            {
                cbTrav2AbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbTrav2Raser.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listRaserColdTravAbsLocation", cbTrav2AbsLocation);
                }
                else if (ckbTrav2Raser.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravRaserAbsLocation", cbTrav2AbsLocation);
                }
            };
            ckbTrav2Barcode.CheckedChanged += (o, e) =>
            {
                cbTrav2AbsLocation.SelectedIndex = -1;

                if ((cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C") && ckbTrav2Barcode.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listBarcodeColdTravAbsLocation", cbTrav2AbsLocation);
                }
                else if (ckbTrav2Barcode.Checked)
                {
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravBarcodeAbsLocation", cbTrav2AbsLocation);
                }
            };


            void ColdTypeOption_TextChanged(object sender, EventArgs e)
            {
                // "C"가 포함되어 있는지 확인
                if (cbMODoption1.Text == "C" || cbMODoption2.Text == "C" || cbMODoption3.Text == "C" || cbMODoption4.Text == "C")
                {
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

                    cbEleqModem.SelectedIndex = -1;
                    cbEleqSensorItem.SelectedIndex = -1;
                    cbLiftAbsLocation.SelectedIndex = -1;
                    cbTrav1AbsLocation.SelectedIndex = -1;
                    cbTrav2AbsLocation.SelectedIndex = -1;
                    cbLiftBrakeOption.SelectedIndex = -1;
                    cbTrav1BrakeOption.SelectedIndex = -1;
                    cbTrav2BrakeOption.SelectedIndex = -1;
                    cbFork1BrakeOption.SelectedIndex = -1;
                    cbFork2BrakeOption.SelectedIndex = -1;


                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listColdEleqModem", cbEleqModem);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listColdEleqSensorItem", cbEleqSensorItem);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listColdLiftBrakeOption", cbLiftBrakeOption);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listColdTravBrakeOption", cbTrav1BrakeOption);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listColdTravBrakeOption", cbTrav2BrakeOption);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listColdForkBrakeOption", cbFork1BrakeOption);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listColdForkBrakeOption", cbFork2BrakeOption);

                    if (ckbLiftRaser.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listRaserColdLiftAbsLocation", cbLiftAbsLocation);
                    }
                    else if (ckbLiftBarcode.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listBarcodeColdLiftAbsLocation", cbLiftAbsLocation);
                    }
                    if (ckbTrav1Raser.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listRaserColdTravAbsLocation", cbTrav1AbsLocation);
                    }
                    else if (ckbTrav1Barcode.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listBarcodeColdTravAbsLocation", cbTrav1AbsLocation);
                    }
                    if (ckbTrav2Raser.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listRaserColdTravAbsLocation", cbTrav2AbsLocation);
                    }
                    else if (ckbTrav2Barcode.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listBarcodeColdTravAbsLocation", cbTrav2AbsLocation);
                    }
                }
                else
                {

                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqModem", cbEleqModem);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftBrakeOption", cbLiftBrakeOption);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravBrakeOption", cbTrav1BrakeOption);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravBrakeOption", cbTrav2BrakeOption);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkBrakeOption", cbFork1BrakeOption);
                    cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listForkBrakeOption", cbFork2BrakeOption);

                    if (cbEleqSensorType.Text == "NPN")
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqNpnSensorItem", cbEleqSensorItem);
                    }
                    else if (cbEleqSensorType.Text == "PNP")
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listEleqPnpSensorItem", cbEleqSensorItem);
                    }
                    if (ckbLiftRaser.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftRaserAbsLocation", cbLiftAbsLocation);
                    }
                    else if (ckbLiftBarcode.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listLiftBarcodeAbsLocation", cbLiftAbsLocation);
                    }
                    if (ckbTrav1Raser.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravRaserAbsLocation", cbTrav1AbsLocation);
                    }
                    else if (ckbTrav1Barcode.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravBarcodeAbsLocation", cbTrav1AbsLocation);
                    }
                    if (ckbTrav2Raser.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravRaserAbsLocation", cbTrav2AbsLocation);
                    }
                    else if (ckbTrav2Barcode.Checked)
                    {
                        cs_ListItems.LoadListFromXmlToComboBox(CS_PathData.ItemListFilePath, "listTravBarcodeAbsLocation", cbTrav2AbsLocation);
                    }
                }
            }

        }
        public void SetComboBoxDefaultValue()
        {
            SetMainSpecDefault();
            SetOptionDefault();
            SetDateDefault();
            SetFuncDefault();

            btnMSPdefault.Click += (o, e) =>
            {
                SetMainSpecDefault();
            };
            btnOPdefault.Click += (o, e) =>
            {
                SetOptionDefault();
            };
            btnPrjDateNow.Click += (o, e) =>
            {
                SetDateDefault();
            };

            void SetMainSpecDefault()
            {
                //주요사양 Default값
                foreach (var comboBox in new[] { cbMSPinputVolt, cbMSPusingVoltage, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbMSPpanelSize })
                {
                    comboBox.SelectedIndex = -1;
                    comboBox.Text = string.Empty;
                    comboBox.SelectedIndex = 0;
                    comboBox.BackColor = Color.White;
                    comboBox.ForeColor = Color.Black;
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
                    comboBox.BackColor = Color.White;
                    comboBox.ForeColor = Color.Black;
                }
            }
            void SetDateDefault()
            {
                cbPRJyear.Text = DateTime.Now.ToString("yyyy");
                cbPRJmonth.Text = DateTime.Now.ToString("MM");
                cbPRJday.Text = DateTime.Now.ToString("dd");

                foreach (ComboBoxEdit comboBox in new[] { cbPRJyear, cbPRJmonth, cbPRJday })
                {
                    comboBox.BackColor = Color.White;
                    comboBox.ForeColor = Color.Black;
                }
            }
            void SetFuncDefault()
            {
                foreach (ComboBoxEdit comboBox in new[] { cbMSPusingVoltage, cbEleqHubModel, cbEleqFanQuantity, cbEleqTerminal, cbEleqPanel, cbEleqHmi, cbEleqOpt, cbEleqTowerLamp, cbEleqSafety, cbEleqSafetyQuantity, cbEleqSensorType })
                {
                    comboBox.SelectedIndex = -1;
                    comboBox.Text = string.Empty;
                    comboBox.SelectedIndex = 0;
                    comboBox.BackColor = Color.White;
                    comboBox.ForeColor = Color.Black;
                }

            }
        }
        private void SetFunctionPageData()
        {

            // ModelPage ComboBox 배열을 포함하는 Dictionary를 정의
            Dictionary<string, Control[]> dicCtrlMod = new Dictionary<string, Control[]>
            {
                { "프로젝트", new Control[] { ckbPRJdomestic, ckbPRJoverseas, cbPRJnumber, cbPRJname, cbPRJwriter, cbPRJyear, cbPRJmonth,cbPRJday } },
                { "모델", new Control[] { cbMODname, cbMODheight, cbMODweight, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4 } },
                { "주요사양",new Control[]{ cbMSPinputVolt, cbMSPusingVoltage, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbMSPpanelSize , cbMSPpanelSizeW, cbMSPpanelSizeD, cbMSPpanelSizeH } },
                { "레이아웃", new Control[] { ckbLevelSame, ckbBayTrue, cbLOUTtravLength, cbLOUTliftHeight, cbLOUTstationNum} },
                { "옵션", new Control[] { cbOPmachineControl, cbOPremoteControl, cbOPemergencyPower, cbOPemergencyLocation, ckbVibrationControlTrue, ckbCctvTrue, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue } }
            };
            // FunctionPage ComboBox 배열을 포함하는 Dictionary를 정의
            Dictionary<string, Control[]> dicCtrlFunc = new Dictionary<string, Control[]>
            {
                { "ELEQ_파워", new Control[] { cbEleqMccbModel, cbEleqSmpsModel, cbEleqCableModel, cbEleqHubModel, 
                    cbEleqPowerKw, cbEleqPowerA, cbEleqBrakeResistorKw, cbEleqBrakeResistorOhm, 
                    cbEleqPowerCable, cbEleqMccbSpec, cbEleqFanQuantity, cbEleqTerminal, 
                    cbEleqPanel, cbEleqHmi, cbEleqOpt, cbEleqTowerLamp, 
                    cbEleqSafety, cbEleqSafetyQuantity} },
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
            Dictionary<string, GroupControl[]> dicGrpFunc = new Dictionary<string, GroupControl[]>
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
            GridControl[] arrGcFunc = new GridControl[]
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
            GridView[] arrGvFunc = new GridView[]
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
            Control[][] arrCtrlGrpFunc = new Control[][]
            {
                dicCtrlFunc.Where(kvp => kvp.Key.Contains("ELEQ")).SelectMany(kvp => kvp.Value).ToArray(),
                dicCtrlFunc.Where(kvp => kvp.Key.Contains("LIFT")).SelectMany(kvp => kvp.Value).ToArray(),
                dicCtrlFunc.Where(kvp => kvp.Key.Contains("TRAV1")).SelectMany(kvp => kvp.Value).ToArray(),
                dicCtrlFunc.Where(kvp => kvp.Key.Contains("TRAV2")).SelectMany(kvp => kvp.Value).ToArray(),
                dicCtrlFunc.Where(kvp => kvp.Key.Contains("FORK1")).SelectMany(kvp => kvp.Value).ToArray(),
                dicCtrlFunc.Where(kvp => kvp.Key.Contains("FORK2")).SelectMany(kvp => kvp.Value).ToArray(),
                dicCtrlFunc.Where(kvp => kvp.Key.Contains("CARR")).SelectMany(kvp => kvp.Value).ToArray()
            };
            // FunctionPage DataTable 그룹 배열을 생성
            DataTable[] arrDtFunc = new DataTable[dicGrpFunc.Count];
            // DataTable 그룹 배열의 값을 거치할 BindingSource 배열 생성
            BindingSource[] bindingSource = new BindingSource[arrDtFunc.Length];
            // FunctionPage GroupControl 그룹 배열을 생성
            GroupControl[] arrGrpAll = dicGrpFunc.Values.SelectMany(arr => arr).ToArray();


            // 각 DataTable 생성 및 DataGridView 설정
            for (int i = 0; i < dicGrpFunc.Count; i++)
            {
                arrDtFunc[i] = new DataTable();

                cs_DataTable.GetDataTable(arrDtFunc[i], CS_StaticString.dArrDtColums);

                // DataTable에 데이터 추가
                for (int j = 0; j < arrCtrlGrpFunc[i].Length; j++)
                {
                    // 각 콤보박스 항목에서 정보 추출
                    string pageName = "기능"; // 페이지 이름, 필요에 따라 수정
                    string funcName = arrCtrlGrpFunc[i][j].Parent.Parent.Parent.Parent.Parent.Text;
                    string grpName = arrCtrlGrpFunc[i][j].Parent.Parent.Parent.Text;
                    string lblName = arrCtrlGrpFunc[i][j].Parent.Controls.OfType<LabelControl>().FirstOrDefault()?.Text;
                    string objectType = arrCtrlGrpFunc[i][j].GetType().ToString();

                    // DataTable에 행 추가
                    arrDtFunc[i].Rows.Add(j + 1, pageName, funcName, grpName, lblName, "", objectType);
                }
                arrGcFunc[i].DataSource = arrDtFunc[i];

                // GridView 설정
                cs_DataGrid.SetGridView(arrGvFunc[i]);
            }


            // 콤보박스 타이핑
            for (int i = 0; i < dicGrpFunc.Count; i++)
            {
                int index = i; // 지역 변수로 캡처
                arrGvFunc[index].FocusedRowChanged += (o, e) =>
                {
                    if (e.FocusedRowHandle >= 0)
                    {
                        int rowIndex = e.FocusedRowHandle; // 선택된 행의 인덱스를 가져옴

                        // 행 인덱스에 맞는 콤보박스를 선택하고 포커스
                        if (rowIndex >= 0 && rowIndex < arrCtrlGrpFunc[index].Length)
                        {
                            // 콤보박스에 포커스를 주고 선택 상태로 만듭니다.
                            arrCtrlGrpFunc[index][rowIndex].Focus();
                        }
                    }
                };

                string[] initialValues = new string[arrCtrlGrpFunc[index].Length];
                // 각 콤보박스와 체크박스에 KeyUp 이벤트 핸들러를 등록
                for (int j = 0; j < arrCtrlGrpFunc[index].Length; j++)
                {
                    // 지역 변수로 캡처
                    int indexRow = j;

                    // arrCbGrpFunc[index][indexRow]가 ComboBox인지 확인
                    if (arrCtrlGrpFunc[index][indexRow] is ComboBoxEdit cb)
                    {
                        // 초기값 저장
                        initialValues[indexRow] = cb.Text;

                        // TextChanged 이벤트 핸들러 등록
                        cb.TextChanged += (o, e) =>
                        {
                            // 텍스트가 초기값과 같으면 공란으로 설정
                            arrDtFunc[index].Rows[indexRow]["Data"] = cb.Text == initialValues[indexRow] ? "" : cb.Text;
                        };
                    }
                    else if (arrCtrlGrpFunc[index][indexRow] is CheckEdit ckb)
                    {
                        // 초기값 저장
                        initialValues[indexRow] = ckb.Text;

                        // CheckedChanged 이벤트 핸들러 등록
                        ckb.CheckedChanged += (o, e) =>
                        {
                            // 체크 상태에 따라 Data를 설정
                            arrDtFunc[index].Rows[indexRow]["Data"] = ckb.Checked ? ckb.Text : "";
                        };
                    }
                }

            }


            // 각 GroupControl에 Spread Event 적용
            for (int i = 0; i < arrGrpAll.Length; i++)
            {
                cs_GroupControl.SpreadGroupControl(arrGrpAll[i], Properties.Resources.CaretBelow, Properties.Resources.CaretRight);
            }


            //Event에 사용하기 위해 static으로 저장
            CS_StaticUnit.dicCtrlSrmAll = dicCtrlMod.Concat(dicCtrlFunc).ToDictionary(x => x.Key, x => x.Value);
            CS_StaticUnit.dicCtrlMod = dicCtrlMod;
            CS_StaticUnit.dicCtrlFunc = dicCtrlFunc;
            CS_StaticUnit.dicGrpSrmFunc = dicGrpFunc;
            CS_StaticUnit.arrCtrlGrpSrmFunc = arrCtrlGrpFunc;
            CS_StaticUnit.arrDtSrmFunc = arrDtFunc;

            // 모든 컨트롤의 TabIndex를 -1로 설정하여 탭 순서에서 제외
            foreach (Control control in this.Controls)
            {
                control.TabIndex = 0;  // TabIndex를 -1로 설정하여 탭 순서에서 제외
            }

            // 탭 순서 부여를 위한 초기 탭 인덱스 설정
            int tabIndex = 1;

            // dicCtrlSrmAll의 컨트롤에 대해 TabIndex와 TabStop 설정
            foreach (var keyValuePair in CS_StaticUnit.dicCtrlSrmAll)
            {
                foreach (var control in keyValuePair.Value)
                {
                    // TabStop을 지원하는 컨트롤에만 TabIndex와 TabStop 설정
                    if (control is Control)
                    {
                        control.TabIndex = tabIndex++;  // TabIndex를 순서대로 증가
                        control.TabStop = true;         // TabStop을 true로 설정하여 탭 순서에 포함
                    }
                }
            }
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
            // 로고 색상 변경
            lblLogo.ForeColor = CS_StaticEtc.colors[4];
            // xtraTabControl 헤더 숨기기
            xtraTabControlLarge.ShowTabHeader = DevExpress.Utils.DefaultBoolean.False;
            lblSRM.ForeColor = CS_StaticEtc.colors[1];
            // 각 Label에 Event 적용
            cs_Label.HoverLabel(lblTAPmodel, CS_StaticEtc.colors[1]);
            cs_Label.HoverLabel(lblTAPfunction, CS_StaticEtc.colors[2]);
            cs_Label.HoverLabel(lblTAPexport, CS_StaticEtc.colors[3]);
            // 기능페이지 패널 기능 추가
            cs_XtraTabControl.AddPanelToTabPage(xtraTabControlFunction, CS_StaticString.dArrStrFunc);

            //Panel Action
            pnlTap.MouseDown += (o, e) => { if (e.Button == MouseButtons.Left) { CS_StaticEtc.On = true; CS_StaticEtc.Pos = e.Location; } };
            pnlTap.MouseMove += (o, e) => { if (CS_StaticEtc.On) Location = new Point(Location.X + (e.X - CS_StaticEtc.Pos.X), Location.Y + (e.Y - CS_StaticEtc.Pos.Y)); };
            pnlTap.MouseUp += (o, e) => { if (e.Button == MouseButtons.Left) { CS_StaticEtc.On = false; CS_StaticEtc.Pos = e.Location; } };
            
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
                    Ui_StaticForm.formInitialPage.Close();
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
                    if (Ui_StaticForm.formConceptSheet != null && !Ui_StaticForm.formConceptSheet.IsDisposed)
                    {
                        Ui_StaticForm.formConceptSheet.Dispose();
                    }

                    // formInitialPage가 이미 열려 있는지 확인 후 표시
                    if (Ui_StaticForm.formInitialPage == null || Ui_StaticForm.formInitialPage.IsDisposed)
                    {
                        Ui_StaticForm.formInitialPage.Close();
                        Ui_StaticForm.formInitialPage = new FormInitialPage();
                    }

                    Ui_StaticForm.formInitialPage.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
                }
            };
            picBoxSave.MouseClick += (o, e) =>
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.InitialDirectory = CS_PathData.XmlFolderPath;
                    saveFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
                    saveFileDialog.Title = "xml 데이터 저장";
                    saveFileDialog.FileName = String.Concat(cbPRJnumber.Text, "_", cbMODfullName.Text); // 기본 파일 이름 설정


                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            CS_PathData.XmlFilePath = saveFileDialog.FileName;

                            XElement root = new XElement("Data",
                                new XElement("Controls", CS_StaticUnit.dicCtrlSrmAll.Select(dic =>
                                    new XElement("ControlGroup", new XAttribute("Key", dic.Key), dic.Value.Select((ctrl, index) =>
                                    {
                                        if (ctrl is ComboBoxEdit cb)
                                        {

                                            cb.Focus();
                                            string text = cb.BackColor == System.Drawing.Color.White ? cb.Text : null;
                                            string labelText = cb.Parent.Controls.OfType<LabelControl>().FirstOrDefault()?.Text ?? "No Label";

                                            return new XElement("ComboBox",
                                                new XAttribute("Index", index),
                                                new XAttribute("LabelText", labelText),
                                                new XElement("Text", text));
                                        }
                                        else if (ctrl is CheckEdit chk)
                                        {
                                            chk.Focus();
                                            string labelText = chk.Parent.Controls.OfType<LabelControl>().FirstOrDefault()?.Text ?? "No Label";

                                            return new XElement("CheckBox",
                                                new XAttribute("Index", index),
                                                new XAttribute("LabelText", labelText),
                                                new XElement("Checked", chk.Checked));
                                        }
                                        return null;
                                    }).Where(x => x != null)))),
                                new XElement("DataTable_Lout",
                                CS_StaticUnit.dtLout != null
                                ? CS_StaticUnit.dtLout.AsEnumerable().Select(row =>
                                new XElement("Row",
                                row.ItemArray.Select((item, index) =>
                                new XElement($"Column{index}", item?.ToString() ?? string.Empty))))
                                : Enumerable.Empty<XElement>()),

                                new XElement("DataTable_SensorIo",
                                CS_StaticSensor.sensorIoDt != null
                                ? CS_StaticSensor.sensorIoDt.AsEnumerable().Select(row =>
                                new XElement("Row",
                                row.ItemArray.Select((item, index) =>
                                new XElement($"Column{index}", item?.ToString() ?? string.Empty))))
                                : Enumerable.Empty<XElement>()),

                                new XElement("DataTable_SensorCopyIo",
                                CS_StaticSensor.sensorCopyIoDt != null
                                ? CS_StaticSensor.sensorCopyIoDt.AsEnumerable().Select(row =>
                                new XElement("Row",
                                row.ItemArray.Select((item, index) =>
                                new XElement($"Column{index}", item?.ToString() ?? string.Empty))))
                                : Enumerable.Empty<XElement>()),

                                new XElement("DataTable_logicIo",
                                CS_StaticSensor.logicIoDt != null
                                ? CS_StaticSensor.logicIoDt.AsEnumerable().Select(row =>
                                new XElement("Row",
                                row.ItemArray.Select((item, index) =>
                                new XElement($"Column{index}", item?.ToString() ?? string.Empty))))
                                : Enumerable.Empty<XElement>()),

                                new XElement("DataTable_UniqueIo",
                                CS_StaticSensor.uniqueIoDt != null
                                ? CS_StaticSensor.uniqueIoDt.AsEnumerable().Select(row =>
                                new XElement("Row",
                                row.ItemArray.Select((item, index) =>
                                new XElement($"Column{index}", item?.ToString() ?? string.Empty))))
                                : Enumerable.Empty<XElement>()));

                            root.Save(CS_PathData.XmlFilePath);
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
                    openFileDialog.InitialDirectory = CS_PathData.XmlFolderPath;
                    openFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
                    openFileDialog.Title = "xml 데이터 로드";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            CS_PathData.XmlFilePath = openFileDialog.FileName;
                            XElement root = XElement.Load(CS_PathData.XmlFilePath);

                            xtraTabControlLarge.SelectedTabPageIndex = 0;

                            // ComboBox와 CheckBox 상태 복원
                            foreach (var group in root.Element("Controls").Elements("ControlGroup"))
                            {
                                string key = group.Attribute("Key").Value;
                                if (CS_StaticUnit.dicCtrlSrmAll.TryGetValue(key, out var controls))
                                {
                                    foreach (var controlElement in group.Elements())
                                    {
                                        int index = int.Parse(controlElement.Attribute("Index").Value);
                                        if (index < controls.Length)
                                        {
                                            var control = controls[index];


                                            if (control is ComboBoxEdit cb && controlElement.Name.LocalName == "ComboBox")
                                            {
                                                string text = controlElement.Element("Text")?.Value;
                                                if (!string.IsNullOrEmpty(text))
                                                {
                                                    cb.Focus();
                                                    cb.Select();
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

                            // DataTable_Lout 데이터 로드
                            if (CS_StaticUnit.dtLout == null)
                            {
                                CS_StaticUnit.dtLout = new DataTable();
                            }
                            else
                            {
                                CS_StaticUnit.dtLout.Clear();
                            }

                            foreach (XElement rowElement in root.Element("DataTable_Lout")?.Elements("Row") ?? Enumerable.Empty<XElement>())
                            {
                                DataRow row = CS_StaticUnit.dtLout.NewRow();
                                foreach (DataColumn col in CS_StaticUnit.dtLout.Columns)
                                {
                                    string cellValue = rowElement.Element($"Column{CS_StaticUnit.dtLout.Columns.IndexOf(col)}")?.Value;
                                    row[col.ColumnName] = string.IsNullOrEmpty(cellValue) ? (object)DBNull.Value : cellValue;
                                }
                                CS_StaticUnit.dtLout.Rows.Add(row);
                            }

                            // sensorIoDt, sensorCopyIoDt 데이터 로드
                            if (CS_StaticSensor.sensorIoDt == null)
                            {
                                CS_StaticSensor.sensorIoDt = new DataTable();
                            }
                            else
                            {
                                CS_StaticSensor.sensorIoDt.Clear();
                            }

                            foreach (XElement rowElement in root.Element("DataTable_SensorIo")?.Elements("Row") ?? Enumerable.Empty<XElement>())
                            {
                                DataRow row = CS_StaticSensor.sensorIoDt.NewRow();
                                foreach (DataColumn col in CS_StaticSensor.sensorIoDt.Columns)
                                {
                                    string cellValue = rowElement.Element($"Column{CS_StaticSensor.sensorIoDt.Columns.IndexOf(col)}")?.Value;
                                    row[col.ColumnName] = string.IsNullOrEmpty(cellValue) ? (object)DBNull.Value : cellValue;
                                }
                                CS_StaticSensor.sensorIoDt.Rows.Add(row);
                            }

                            if (CS_StaticSensor.sensorCopyIoDt == null)
                            {
                                CS_StaticSensor.sensorCopyIoDt = new DataTable();
                            }
                            else
                            {
                                CS_StaticSensor.sensorCopyIoDt.Clear();
                            }

                            foreach (XElement rowElement in root.Element("DataTable_SensorCopyIo")?.Elements("Row") ?? Enumerable.Empty<XElement>())
                            {
                                DataRow row = CS_StaticSensor.sensorCopyIoDt.NewRow();
                                foreach (DataColumn col in CS_StaticSensor.sensorCopyIoDt.Columns)
                                {
                                    string cellValue = rowElement.Element($"Column{CS_StaticSensor.sensorCopyIoDt.Columns.IndexOf(col)}")?.Value;
                                    row[col.ColumnName] = string.IsNullOrEmpty(cellValue) ? (object)DBNull.Value : cellValue;
                                }
                                CS_StaticSensor.sensorCopyIoDt.Rows.Add(row);
                            }

                            // uniqueIoDt 데이터 로드
                            if (CS_StaticSensor.logicIoDt == null)
                            {
                                CS_StaticSensor.logicIoDt = new DataTable();
                            }
                            else
                            {
                                CS_StaticSensor.logicIoDt.Clear();
                            }

                            foreach (XElement rowElement in root.Element("DataTable_logicIo")?.Elements("Row") ?? Enumerable.Empty<XElement>())
                            {
                                DataRow row = CS_StaticSensor.logicIoDt.NewRow();
                                foreach (DataColumn col in CS_StaticSensor.logicIoDt.Columns)
                                {
                                    string cellValue = rowElement.Element($"Column{CS_StaticSensor.logicIoDt.Columns.IndexOf(col)}")?.Value;
                                    row[col.ColumnName] = string.IsNullOrEmpty(cellValue) ? (object)DBNull.Value : cellValue;
                                }
                                CS_StaticSensor.logicIoDt.Rows.Add(row);
                            }

                            if (CS_StaticSensor.uniqueIoDt == null)
                            {
                                CS_StaticSensor.uniqueIoDt = new DataTable();
                            }
                            else
                            {
                                CS_StaticSensor.uniqueIoDt.Clear();
                            }

                            foreach (XElement rowElement in root.Element("DataTable_UniqueIo")?.Elements("Row") ?? Enumerable.Empty<XElement>())
                            {
                                DataRow row = CS_StaticSensor.uniqueIoDt.NewRow();
                                foreach (DataColumn col in CS_StaticSensor.uniqueIoDt.Columns)
                                {
                                    string cellValue = rowElement.Element($"Column{CS_StaticSensor.uniqueIoDt.Columns.IndexOf(col)}")?.Value;
                                    row[col.ColumnName] = string.IsNullOrEmpty(cellValue) ? (object)DBNull.Value : cellValue;
                                }
                                CS_StaticSensor.uniqueIoDt.Rows.Add(row);
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
                Ui_StaticForm.formItemsList = new FormItemsList();
                Ui_StaticForm.formItemsList.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
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
                            Ui_StaticForm.xtraReport1.ExportToPdf(filePath);

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
            lblTAPmodel.MouseClick += (e, o) => { xtraTabPageLargeModel.Show(); lblSRM.ForeColor = CS_StaticEtc.colors[1]; };
            lblTAPfunction.MouseClick += (e, o) => { xtraTabPageLargeFunction.Show(); lblSRM.ForeColor = CS_StaticEtc.colors[2]; };
            lblTAPexport.MouseClick += (e, o) =>
            {
                xtraTabPageLargeGenerating.Show();
                lblSRM.ForeColor = CS_StaticEtc.colors[3];

                // XtraReport1 인스턴스 생성
                Ui_StaticForm.xtraReport1 = new XtraReport1
                {
                    PaperKind = System.Drawing.Printing.PaperKind.A4, // A4 크기 설정
                    Margins = new System.Drawing.Printing.Margins(90, 90, 0, 0), // 마진 설정
                };

                // 문서 생성 및 페이지 너비에 맞추기
                Ui_StaticForm.xtraReport1.CreateDocument();
                Ui_StaticForm.xtraReport1.PrintingSystem.Document.AutoFitToPagesWidth = 1;

                // DocumentViewer에 바인딩
                documentViewer1.DocumentSource = Ui_StaticForm.xtraReport1;


                //gridControl2.DataSource = CS_StaticSensor.logicIoDt;
                //gridControl3.DataSource = CS_StaticSensor.uniqueIoDt;
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
                    if (Ui_StaticForm.formConceptSheet != null && !Ui_StaticForm.formConceptSheet.IsDisposed)
                    {
                        Ui_StaticForm.formConceptSheet.Dispose();
                    }

                    // formInitialPage가 이미 열려 있는지 확인 후 표시
                    if (Ui_StaticForm.formInitialPage == null || Ui_StaticForm.formInitialPage.IsDisposed)
                    {
                        Ui_StaticForm.formInitialPage = new FormInitialPage();
                    }

                    Ui_StaticForm.formInitialPage.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
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
                        var maxDtRow = CS_StaticSensor.uniqueIoDt.AsEnumerable()
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
                                cbEleqDt.BackColor = Color.White;
                                cbEleqDt.ForeColor = Color.Black;
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
                if (cbEleqLocation.BackColor != Color.White || cbEleqType.BackColor != Color.White || cbEleqPoint.BackColor != Color.White || cbEleqParts.BackColor != Color.White || cbEleqDt.BackColor != Color.White
                || cbEleqLocation.Text == "" || cbEleqType.Text == "" || cbEleqPoint.Text == "" || cbEleqParts.Text == "" || cbEleqDt.Text == "")
                {
                    MessageBox.Show("모델 및 컨트롤러 사양을 확정하세요.");
                    return;
                }
                // CS_StaticSensor.uniqueIoDt의 DataTable 객체가 존재하는지 확인
                if (CS_StaticSensor.uniqueIoDt != null)
                {
                    // NO 컬럼의 최대값 구하기
                    int maxNo = 0;
                    if (CS_StaticSensor.uniqueIoDt.Rows.Count > 0)
                    {
                        maxNo = CS_StaticSensor.uniqueIoDt.AsEnumerable()
                                                          .Max(row => row.Field<int>("NO"));
                    }

                    // 새 행 생성
                    DataRow newRow = CS_StaticSensor.uniqueIoDt.NewRow();

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
                    CS_StaticSensor.uniqueIoDt.Rows.Add(newRow);
                    gridControl1.DataSource = CS_StaticSensor.uniqueIoDt;

                }

                if (CS_StaticSensor.sensorIoDt != null)
                {
                    // cbEleqPoint.Text가 숫자인지 확인하고, 숫자라면 그 값을 사용하여 행 추가
                    if (int.TryParse(cbEleqPoint.Text, out int numberOfRowsToAdd))
                    {
                        for (int i = 0; i < numberOfRowsToAdd; i++)
                        {
                            // 새 행 생성
                            DataRow newRow = CS_StaticSensor.sensorIoDt.NewRow();

                            // 기타 열에 기본 값 설정
                            newRow["LOCATION"] = cbEleqLocation.Text;
                            newRow["TYPE1"] = cbEleqType.Text;
                            newRow["PARTS"] = cbEleqParts.Text;
                            newRow["DT"] = cbEleqDt.Text;
                            newRow["기능"] = "";
                            newRow["SIGNAL"] = "";
                            newRow["DESCRIPTION"] = "";

                            // DataTable에 새 행 추가
                            CS_StaticSensor.sensorIoDt.Rows.Add(newRow);
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
                if (e.Column.FieldName == "IFB1" || e.Column.FieldName == "IFB2" || e.Column.FieldName == "IFB3" || e.Column.FieldName == "IFB4")
                {
                    string type1Check = Convert.ToString(gridView1.GetRowCellValue(e.RowHandle, "TYPE1"));

                    if (type1Check == "DIO")
                    {
                        gridView1.SetRowCellValue(e.RowHandle, e.Column.FieldName, false);
                    }
                    RepositoryItemCheckEdit checkEdit = new RepositoryItemCheckEdit();
                    checkEdit.ValueChecked = true;
                    checkEdit.ValueUnchecked = false;
                    e.RepositoryItem = checkEdit; // 수정된 RepositoryItem을 셀에 적용
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
                if (cbMODfullName.BackColor != Color.White)
                {
                    MessageBox.Show("모델명이 확정되지 않았습니다.");
                    return;
                }

                try
                {
                    // Excel 파일을 열고 입력된 텍스트에 맞는 워크시트 로드
                    using (var workbook = new XLWorkbook(CS_PathData.IoListFilePath))
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

                if (CS_StaticSensor.sensorCopyIoDt == CS_StaticSensor.sensorIoDt)
                {
                    return;
                }
                else
                {
                    CS_StaticSensor.sensorCopyIoDt = CS_StaticSensor.sensorIoDt.Copy();
                }

                Ui_StaticForm.formIoList = new FormIoList();
                Ui_StaticForm.formIoList.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));
            };

            void UpdateSensorList()
            {
                if (cbMODname.BackColor != Color.White || cbMODheight.BackColor != Color.White || cbMODweight.BackColor != Color.White ||
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

                CS_StaticSensor.listLiftSensor = UpdateSensorListToString(funcSensorList, "LIFT");
                CS_StaticSensor.listTrav1Sensor = UpdateSensorListToString(funcSensorList, "TRAV");
                CS_StaticSensor.listTrav2Sensor = UpdateSensorListToString(funcSensorList, "TRAV2");
                CS_StaticSensor.listFork1Sensor = UpdateSensorListToString(funcSensorList, "FORK");
                CS_StaticSensor.listFork2Sensor = UpdateSensorListToString(funcSensorList, "FORK2");
                CS_StaticSensor.listCarrSensor = UpdateSensorListToString(funcSensorList, "CARR");

                CS_StaticSensor.funcSensorDict = funcSensorList;
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
                // uniqueIoDt가 null인지 확인하고, null인 경우 메서드를 종료
                if (CS_StaticSensor.uniqueIoDt == null)
                    return;

                string selectedType = cbEleqType.Text;
                string selectedPoint = cbEleqPoint.Text;

                // cbEleqParts 초기화
                cbEleqParts.Properties.Items.Clear();

                // TYPE 및 POINT 값과 일치하는 PARTS 항목 필터링 및 중복 제거
                var filteredParts = CS_StaticSensor.uniqueIoDt.AsEnumerable()
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
                    if (CS_StaticSensor.sensorIoDt == null)
                        return;

                    DataTable copyDt = CS_StaticSensor.sensorIoDt.Copy();

                    // 고유 데이터 테이블 생성
                    CS_StaticSensor.logicIoDt = new DataView(copyDt).ToTable(true, "LOCATION", "TYPE1", "TYPE2", "PARTS", "DT");
                    // NO 열 추가 및 순번 할당
                    CS_StaticSensor.logicIoDt.Columns.Add("NO", typeof(int)).SetOrdinal(0);
                    for (int i = 0; i < CS_StaticSensor.logicIoDt.Rows.Count; i++)
                        CS_StaticSensor.logicIoDt.Rows[i]["NO"] = i + 1;

                    // 추가 열 생성
                    CS_StaticSensor.logicIoDt.Columns.Add("POINT", typeof(int)).SetOrdinal(4);
                    string[] boolColumns = { "IFB1", "IFB2", "IFB3", "IFB4", "8BIT" };
                    foreach (string colName in boolColumns)
                        CS_StaticSensor.logicIoDt.Columns.Add(colName, typeof(bool));

                    // 기본 값 설정
                    foreach (DataRow row in CS_StaticSensor.logicIoDt.Rows)
                    {
                        foreach (string colName in boolColumns)
                            row[colName] = true;
                    }

                    foreach (DataRow row in CS_StaticSensor.logicIoDt.Rows)
                    {
                        string type2Value = row.Field<string>("TYPE2");

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


                    }



                    // POINT 값 설정
                    foreach (DataRow row in CS_StaticSensor.logicIoDt.Rows)
                    {
                        string locationValue = row.Field<string>("LOCATION");
                        string dtValue = row.Field<string>("DT");

                        row["POINT"] = copyDt.AsEnumerable()
                                             .Count(r => r.Field<string>("LOCATION") == locationValue &&
                                                         r.Field<string>("DT") == dtValue);

                        // POINT 값이 32인 경우 모든 IFB 값을 true로 설정
                        int? pointValue = row.Field<int?>("POINT");
                        if (pointValue.HasValue && pointValue.Value == 32)
                        {
                            row["IFB1"] = true;
                            row["IFB2"] = true;
                            row["IFB3"] = true;
                            row["IFB4"] = true;
                        }
                    }



                    CS_StaticSensor.uniqueIoDt = new DataView(copyDt).ToTable(true, "LOCATION", "TYPE1", "PARTS", "DT");
                    CS_StaticSensor.uniqueIoDt.Columns.Add("NO", typeof(int)).SetOrdinal(0);
                    for (int i = 0; i < CS_StaticSensor.uniqueIoDt.Rows.Count; i++)
                        CS_StaticSensor.uniqueIoDt.Rows[i]["NO"] = i + 1;

                    // 추가 열 생성
                    CS_StaticSensor.uniqueIoDt.Columns.Add("POINT", typeof(int)).SetOrdinal(4);
                    // POINT 값 설정
                    foreach (DataRow row in CS_StaticSensor.uniqueIoDt.Rows)
                    {
                        string locationValue = row.Field<string>("LOCATION");
                        string dtValue = row.Field<string>("DT");

                        row["POINT"] = copyDt.AsEnumerable()
                                             .Count(r => r.Field<string>("LOCATION") == locationValue &&
                                                         r.Field<string>("DT") == dtValue);
                    }

                    string[] boolColumns1 = { "IFB1", "IFB2", "IFB3", "IFB4", "8BIT" };
                    foreach (string colName in boolColumns1)
                        CS_StaticSensor.uniqueIoDt.Columns.Add(colName, typeof(bool));

                    // 기본 값 설정
                    foreach (DataRow row in CS_StaticSensor.uniqueIoDt.Rows)
                    {
                        foreach (string colName in boolColumns1)
                            row[colName] = false;
                    }



                    foreach (DataRow logicRow in CS_StaticSensor.logicIoDt.Rows)
                    {
                        // LOCATION과 DT 값을 확인
                        string logicLocation = logicRow.Field<string>("LOCATION");
                        string logicDt = logicRow.Field<string>("DT");

                        bool logicIfb1 = logicRow.Field<bool>("IFB1");
                        bool logicIfb2 = logicRow.Field<bool>("IFB2");
                        bool logicIfb3 = logicRow.Field<bool>("IFB3");
                        bool logicIfb4 = logicRow.Field<bool>("IFB4");
                        bool logic8bit = logicRow.Field<bool>("8BIT");


                        foreach (DataRow uniqueRow in CS_StaticSensor.uniqueIoDt.Rows)
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
                    gridControl1.DataSource = CS_StaticSensor.uniqueIoDt;

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

                cs_ComboBox.SettingComboBox(cbEleqLocation, "-");
                cs_ComboBox.SettingComboBox(cbEleqType, "-");
                cs_ComboBox.SettingComboBox(cbEleqDt, "-");
                cs_ComboBox.SettingComboBox(cbEleqParts, "-");
                cs_ComboBox.SettingComboBox(cbEleqPoint, "-");

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

                    DataTable tempDt = excelDt.Copy();

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
                    CS_StaticSensor.sensorIoDt = tempDt.DefaultView.ToTable(false, "PARTS", "LOCATION", "DT", "TYPE1", "TYPE2", "CN.", "NO.", "ADD.", "SIGNAL", "DESCRIPTION", "구분", "포함조건", "제외조건", "기능", "타입");


                    // LOCATION 및 DT로 그룹화하여 SIGNAL 또는 DESCRIPTION이 없는 그룹 필터링
                    var filteredRows = CS_StaticSensor.sensorIoDt.AsEnumerable()
                        .GroupBy(row => new { LOCATION = row.Field<string>("LOCATION"), DT = row.Field<string>("DT") })
                        .Where(group =>
                            group.Any(row =>
                                !string.IsNullOrWhiteSpace(row.Field<string>("SIGNAL")) ||
                                !string.IsNullOrWhiteSpace(row.Field<string>("DESCRIPTION"))))
                        .SelectMany(group => group);

                    // 결과를 새로운 DataTable로 변환
                    CS_StaticSensor.sensorIoDt = filteredRows.Any()
                        ? filteredRows.CopyToDataTable()
                        : CS_StaticSensor.sensorIoDt.Clone(); // 결과가 없으면 빈 테이블 생성

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
                if (cbMSPcontrollerSpec.Text == "MIT_Q_IOLINK")
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
                if (cbMSPcontrollerSpec.Text == "SIE_ET200SP")
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
            excelDt.Clear();
            excelDt.Columns.Clear(); // 컬럼도 제거하여 중복 방지

            try
            {


                // Excel 파일을 열고 입력된 텍스트에 맞는 워크시트 로드
                using (var workbook = new XLWorkbook(CS_PathData.IoListFilePath))
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
                                excelDt.Columns.Add(cell.Value.ToString()); // 컬럼 추가
                            }
                            firstRow = false;
                        }
                        else
                        {
                            // 데이터 행 처리
                            DataRow dataRow = excelDt.NewRow();
                            int i = 0;
                            foreach (var cell in row.Cells())
                            {
                                string cellValue = cell.Value.ToString();

                                // 공란 필드가 있는 경우 이전 행의 값을 가져와 채움
                                if (excelDt.Columns[i].ColumnName == "PARTS")
                                {
                                    dataRow[i] = string.IsNullOrWhiteSpace(cellValue) ? previousParts : cellValue;
                                    previousParts = dataRow[i].ToString(); // 현재 값 저장
                                }
                                else if (excelDt.Columns[i].ColumnName == "LOCATION")
                                {
                                    dataRow[i] = string.IsNullOrWhiteSpace(cellValue) ? previousLocation : cellValue;
                                    previousLocation = dataRow[i].ToString(); // 현재 값 저장
                                }
                                else if (excelDt.Columns[i].ColumnName == "DT")
                                {
                                    dataRow[i] = string.IsNullOrWhiteSpace(cellValue) ? previousDt : cellValue;
                                    previousDt = dataRow[i].ToString(); // 현재 값 저장
                                }
                                else if (excelDt.Columns[i].ColumnName == "TYPE1")
                                {
                                    dataRow[i] = string.IsNullOrWhiteSpace(cellValue) ? previousType1 : cellValue;
                                    previousType1 = dataRow[i].ToString(); // 현재 값 저장
                                }
                                else if (excelDt.Columns[i].ColumnName == "TYPE2")
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
                            excelDt.Rows.Add(dataRow);
                        }
                    }
                }
            }

            catch (Exception ex)
            {

            }
        }
        
        private void ActivateEplan()
        {

            picBoxProjectGenerate.Click += (o, e) =>
            {
                if (!CheckBeforeGenerating())
                    return;
                GeneratePageMacro();
            };

            bool CheckBeforeGenerating()
            {
                for (int i = 0; i < xtraTabControlLarge.TabPages.Count; i++)
                {
                    xtraTabControlLarge.SelectedTabPageIndex = i;
                }

                Control[] checkControls =
                    {
                    cbPRJnumber,
                    cbMODfullName,
                    cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec,
                    cbEleqTerminal
                };

                string errActCbTxt = "";

                foreach (ComboBoxEdit cb in checkControls)
                {
                    if (cb.BackColor != Color.White || string.IsNullOrEmpty(cb.Text))
                    {
                        // 레이블 텍스트 수집
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

                ForkType forkType =
                    ckbMODforkoption.Checked || cbMODoption1.Text == "D(v)" || cbMODoption2.Text == "D(v)" || cbMODoption3.Text == "D(v)" || cbMODoption4.Text == "D(v)" ? ForkType.FORK2 :
                    ForkType.FORK1;

                TravelType travelType =
                    ckbTravDoubleMotorTrue.Checked ? TravelType.TRAV2 :
                    TravelType.TRAV1;

                ControllerType controllerType =
                    cbMSPcontrollerSpec.Text == "HMX_MICOM" ? ControllerType.MICOM :
                    cbMSPcontrollerSpec.Text == "MIT_Q" ? ControllerType.MITSUBISHI_Q :
                    cbMSPcontrollerSpec.Text == "MIT_R" ? ControllerType.MITSUBISHI_R :
                    cbMSPcontrollerSpec.Text == "MIT_Q_IOLINK" ? ControllerType.MITSUBISHI_Q_IOLINK :
                    cbMSPcontrollerSpec.Text == "MIT_R_IOLINK" ? ControllerType.MITSUBISHI_R_IOLINK :
                    cbMSPcontrollerSpec.Text == "SIE_ET200SP" ? ControllerType.SIEMENS_ET200SP :
                    ControllerType.ETC;

                MotorCableType motorCableTypeLift =
                    ckbLiftDdi.Checked ? MotorCableType.DDI : MotorCableType.GENERAL;
                MotorCableType motorCableTypeTrav =
                    ckbTrav1Ddi.Checked ? MotorCableType.DDI : MotorCableType.GENERAL;
                MotorCableType motorCableTypeFork1 =
                    ckbFork1Ddi.Checked ? MotorCableType.DDI : MotorCableType.GENERAL;
                MotorCableType motorCableTypeFork2 =
                    ckbFork2Ddi.Checked ? MotorCableType.DDI : MotorCableType.GENERAL;

                EncoderType encoderTypeLift =
                    ckbLiftRaser.Checked ? EncoderType.LASER : EncoderType.BARCODE;
                EncoderType encoderTypeTrav =
                    ckbTrav1Raser.Checked ? EncoderType.LASER : EncoderType.BARCODE;
                EncoderType encoderTypeFork = EncoderType.ETC;

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

                mcnsControl.SetMacroRootDirectory(CS_PathData.MacroFolderPath);

                string elkName = string.Concat(cbPRJnumber.Text, "_", cbMODfullName.Text);
                string prjFullFilePath = Path.Combine(CS_PathData.PrjFolderPath, elkName + ".elk");
                
                // 프로젝트 파일이 이미 존재하는지 확인
                if (File.Exists(prjFullFilePath))
                {
                    // 덮어쓰기 여부를 묻는 메시지 박스
                    DialogResult dialogResult = MessageBox.Show(string.Concat("이미 생성된 프로젝트가 있습니다.", "\n경로 : ", CS_PathData.PrjFolderPath, "\n파일이름 : ", elkName), "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                        oProgress.Step(1);
                        this.mcnsControl.CreateAndOpenProject(prjFullFilePath, CS_PathData.BasicTempletFilePath);
                        oProgress.EndPart(false);

                        //part 2
                        oProgress.BeginPart(38.5, "");
                        oProgress.SetActionText("매크로 삽입");
                        oProgress.SetNeededSteps(1);
                        CheckMcnsEngineFunction(this.mcnsControl.InsertACPowerEmpMacro(installSiteType, powerDpType, inverterRegenType));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertDCPowerEmpMacro(installSiteType, powerDpType, inverterMakerType, inverterType, controllerType));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertInverterPublicEmpMacro(installSiteType, inverterMakerType, inverterType, inverterRegenType, powerDpType, controllerType, forkType, travelType));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertMotorPublicEmpMacro(FunctionType.LIFT, motorCableTypeLift, encoderTypeLift, installSiteType, inverterMakerType, inverterType, inverterRegenType, powerDpType, controllerType, forkType, travelType));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertMotorPublicEmpMacro(FunctionType.TRAV1, motorCableTypeTrav, encoderTypeTrav, installSiteType, inverterMakerType, inverterType, inverterRegenType, powerDpType, controllerType, forkType, travelType));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertMotorPublicEmpMacro(FunctionType.FORK1, motorCableTypeFork1, encoderTypeFork, installSiteType, inverterMakerType, inverterType, inverterRegenType, powerDpType, controllerType, forkType, travelType));
                        if (ckbTravDoubleMotorTrue.Checked)
                        {
                            CheckMcnsEngineFunction(this.mcnsControl.InsertMotorPublicEmpMacro(FunctionType.TRAV2, motorCableTypeTrav, encoderTypeTrav, installSiteType, inverterMakerType, inverterType, inverterRegenType, powerDpType, controllerType, forkType, travelType));
                        }
                        if (ckbMODforkoption.Checked || cbMODoption1.Text == "D(v)" || cbMODoption2.Text == "D(v)" || cbMODoption3.Text == "D(v)" || cbMODoption4.Text == "D(v)")
                        {
                            CheckMcnsEngineFunction(this.mcnsControl.InsertMotorPublicEmpMacro(FunctionType.FORK2, motorCableTypeFork2, encoderTypeFork, installSiteType, inverterMakerType, inverterType, inverterRegenType, powerDpType, controllerType, forkType, travelType));
                        }
                        CheckMcnsEngineFunction(this.mcnsControl.InsertSystemEmpMacro(sensorOutputType, controllerType));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorFan(int.Parse(cbEleqFanQuantity.Text)));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorFluorenscentLamp(fluorescentType, installSiteType));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorHMI(cbEleqHmi.Text));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorHub(cbEleqHubModel.Text));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorOPT(installSiteType, controllerType, cbEleqOpt.Text));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacorSafetyEmergency(installSiteType, int.Parse(cbEleqSafetyQuantity.Text), forkType, travelType, cbEleqSafety.Text));
                        //if (ckbPRJoverseas.Checked)
                        //{
                        //    CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacroSafetyRelay(cbEleqSafetyRelay.Text));
                        //    CheckMcnsEngineFunction(this.mcnsControl.InsertWindowMacroSafetyReset(cbEleqSafetyReset.Text));
                        //}

                        CheckMcnsEngineFunction(this.mcnsControl.InsertControllerMacro(controllerType, CS_StaticSensor.uniqueIoDt, CS_StaticSensor.sensorIoDt));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertControllerBitIOMacro(controllerType, CS_StaticSensor.uniqueIoDt, CS_StaticSensor.sensorIoDt));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertPlcIOSwitchWindowMacro(installSiteType, controllerType, int.Parse(cbEleqSafetyQuantity.Text)));
                        CheckMcnsEngineFunction(this.mcnsControl.InsertPlcTowerLampWindowMacro(controllerType, towerLampType));
                        oProgress.EndPart(false);


                        //part 4
                        oProgress.BeginPart(48.0, "");
                        oProgress.SetActionText("PDF 생성");
                        oProgress.SetNeededSteps(1);
                        CheckMcnsEngineFunction(this.mcnsControl.ApplyWirePlaceHolder(installSiteType));
                        this.mcnsControl.SetProjectProperty("EPLAN.Project.UserSupplementaryField50", cbPRJname.Text);
                        this.mcnsControl.SetProjectProperty("EPLAN.Project.UserSupplementaryField9", elkName);
                        this.mcnsControl.SetProjectProperty("EPLAN.Project.UserSupplementaryField10", cbMSPinputVolt.Text);
                        this.mcnsControl.SetProjectProperty("EPLAN.Project.UserSupplementaryField11", cbMSPinputHz.Text);
                        this.mcnsControl.SetProjectProperty("EPLAN.Project.UserSupplementaryField12", cbEleqPowerA.Text);
                        using (ExcelPackage package = new ExcelPackage())
                        {
                            // 워크시트 추가
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                            // DataTable 데이터를 워크시트에 로드
                            worksheet.Cells["A1"].LoadFromDataTable(CS_StaticSensor.sensorIoDt, true);

                            // 파일 저장
                            FileInfo file = new FileInfo(Path.Combine(CS_PathData.PrjFolderPath, elkName + ".xlsx"));
                            package.SaveAs(file);
                        }
                        this.mcnsControl.GeneratePdf();
                        oProgress.EndPart(true);

                        MessageBox.Show(elkName + ": 프로젝트 생성 완료");

                        void CheckMcnsEngineFunction(ResponseModel responseModel)
                        {
                            if (responseModel.Success == false)
                                MessageBox.Show(responseModel.Message);
                        }
                    }
                        
                }
                
                
            }


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

            interLock.ActivateControlSwitch(
                ckbRegenerativeUnitTrue,
                new ComboBoxEdit[]
                {
                    cbEleqBrakeResistorKw, cbEleqBrakeResistorOhm,
                    cbLiftBrakeResistorKw, cbLiftBrakeResistorOhm,
                    cbTrav1BrakeResistorKw, cbTrav1BrakeResistorOhm,
                    cbFork1BrakeResistorKw, cbFork1BrakeResistorOhm,
                    cbFork2BrakeResistorKw, cbFork2BrakeResistorOhm,
                });

            interLock.AlramToFunctionByText(rtbxEleq, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxEleq, new Control[] { cbMODname , ckbMODforkoption, cbMODoption1 , cbMODoption2 , cbMODoption3 , cbMODoption4 , cbMSPinputVolt , cbMSPinputHz , cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec , cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxLift, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxTrav1, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxTrav2, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxFork1, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxFork2, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });
            interLock.AlramToFunctionByText(rtbxCarr, new Control[] { cbMODname, ckbMODforkoption, cbMODoption1, cbMODoption2, cbMODoption3, cbMODoption4, cbMSPinputVolt, cbMSPinputHz, cbMSPcontrollerSpec, cbMSPinverterMaker, cbMSPinverterSpec, cbEleqSensorType, ckbTravDoubleMotorTrue, ckbRegenerativeUnitTrue });


            cbMSPpanelSizeW.TextChanged += (o, e) =>
            {
                UpdateFanQuantityByPanel();
            };
            UpdateFanQuantityByPanel();
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
                if (currentCb.BackColor == Color.White)
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
                if (cbMODoption1.BackColor != Color.White)
                {
                    cbMODoption2.Hide();
                    cbMODoption2.Text = ""; // 텍스트 지우기
                    cbMODoption3.Hide();
                    cbMODoption3.Text = ""; // 텍스트 지우기
                    cbMODoption4.Hide();
                    cbMODoption4.Text = ""; // 텍스트 지우기
                }
                // cbMODoption2의 배경이 흰색이 아닌 경우 3, 4번 콤보박스를 숨기고 텍스트를 지웁니다.
                else if (cbMODoption2.BackColor != Color.White)
                {
                    cbMODoption3.Hide();
                    cbMODoption3.Text = ""; // 텍스트 지우기
                    cbMODoption4.Hide();
                    cbMODoption4.Text = ""; // 텍스트 지우기
                }
                // cbMODoption3의 배경이 흰색이 아닌 경우 4번 콤보박스를 숨기고 텍스트를 지웁니다.
                else if (cbMODoption3.BackColor != Color.White)
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
            cs_CheckBox.ChangeToRadioButton(ckbLiftRaser, ckbLiftBarcode);
            // button을 Radio button으로 사용
            cs_CheckBox.ChangeToRadioButton(ckbTrav1Raser, ckbTrav1Barcode);
            // button을 Radio button으로 사용
            cs_CheckBox.ChangeToRadioButton(ckbTrav2Raser, ckbTrav2Barcode);

            //interLock.BlockCtrlsByInverter(cbMSPinverterMaker, cbMSPinverterSpec,"SEW","MODULAR" ,new Control[] { cbLiftMccbSpec, cbTrav1MccbSpec, cbTrav2MccbSpec, cbFork1MccbSpec, cbFork2MccbSpec, cbLiftBrakeResistorKw, cbLiftBrakeResistorOhm, cbTrav1BrakeResistorKw, cbTrav1BrakeResistorOhm, cbTrav2BrakeResistorKw, cbTrav2BrakeResistorOhm, cbFork1BrakeResistorKw, cbFork1BrakeResistorOhm, cbFork2BrakeResistorKw, cbFork2BrakeResistorOhm });
            //interLock.BlockCtrlsByInverter(cbMSPinverterMaker, cbMSPinverterSpec, "SEW", "SYSTEM", new Control[] { cbEleqPowerKw, cbEleqBrakeResistorKw, cbEleqBrakeResistorOhm });
        }
        private void UpdateComboBoxItemList()
        {
            if (watcher == null) // watcher가 아직 생성되지 않은 경우에만 생성
            {
                watcher = new FileSystemWatcher();
                watcher.Path = Path.GetDirectoryName(CS_PathData.ItemListFilePath);
                watcher.Filter = Path.GetFileName(CS_PathData.ItemListFilePath);
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