using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DevExpress.ClipboardSource.SpreadsheetML;
using Eplan.MCNS.Lib;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Xml.Linq;

namespace Eplan.EplAddin.HMX_MCNS
{
    public partial class FormItemsList : Form
    {
        GridViewManager gvManager = new GridViewManager();
        FilePathManager pathManager = new FilePathManager();
        

        // 컨트롤 DPI 스케일링 조정
      


        public FormItemsList()
        {
            InitializeComponent();
            
            LoadFromXmlData();

            btnSaveItems.Click += (o, e) =>
            {
                try
                {
                    // 데이터를 XML 파일로 저장
                    SaveToXmlData();


                    MessageBox.Show("데이터가 성공적으로 저장되었습니다. \n 프로그램을 다시 시작해주세요");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("저장 중 오류가 발생했습니다: " + ex.Message);
                }
            };

            SetGridView();

            
        }
        private void LoadFromXmlData()
        {
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listMODName", gridControl1);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listMODOption", gridControl2);

            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listMSPinputVolt", gridControl3);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listMSPinputHz", gridControl4);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listMSPcontrollerMaker", gridControl48);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listMSPcontrollerSpec", gridControl5);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listMSPinverterMaker", gridControl49);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listMSPinverterSpec", gridControl6);

            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listOPmachineControl", gridControl7);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listOPremoteControl", gridControl8);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listOPemergencyPower", gridControl9);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listOPemergencyLocation", gridControl10);

            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqUsingVoltage", gridControl56);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqMccbModel", gridControl11);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqSmpsModel", gridControl12);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqCableModel", gridControl13);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqHubModel", gridControl14);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqFanQuantity", gridControl46);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqTerminal", gridControl47);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqPanel", gridControl50);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqHmi", gridControl51);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqOpt", gridControl52);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqTowerLamp", gridControl55);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqSafety", gridControl53);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqSafetyQuantity", gridControl54);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqModem", gridControl15);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqInterLockSensorSide", gridControl16);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqInterLockBit", gridControl17);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqNpnSensorItem", gridControl18);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listEleqPnpSensorItem", gridControl19);

            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listLiftBrakeOption", gridControl23);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listLiftMotorSpec", gridControl20);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listLiftMotorMaker", gridControl26);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listLiftMotorMethod", gridControl60);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listLiftRaserAbsLocation", gridControl21);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listLiftBarcodeAbsLocation", gridControl22);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listLiftNpnRightPosition", gridControl25);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listLiftPnpRightPosition", gridControl37);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listLiftLimitSwitch", gridControl24);

            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listTravBrakeOption", gridControl30);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listTravMotorSpec", gridControl27);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listTravMotorMaker", gridControl64);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listTravMotorMethod", gridControl65);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listTravRaserAbsLocation", gridControl28);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listTravBarcodeAbsLocation", gridControl29);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listTravNpnRightPosition", gridControl31);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listTravPnpRightPosition", gridControl45);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listTravLimitSwitch", gridControl32);

            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listForkBrakeOption", gridControl38);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listForkMotorSpec", gridControl33);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listForkMotorMaker", gridControl34);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listForkMotorMethod", gridControl35);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listForkNpnRightPosition", gridControl39);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listForkPnpRightPosition", gridControl40);

            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listCarrNpnSensor", gridControl41);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listCarrPnpSensor", gridControl42);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listCarrNpnDoubleInput", gridControl43);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listCarrPnpDoubleInput", gridControl44);

            //콜드 타입
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listColdEleqModem", gridControl57);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listColdEleqSensorItem", gridControl58);

            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listRaserColdLiftAbsLocation", gridControl59);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listBarcodeColdLiftAbsLocation", gridControl61);

            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listRaserColdTravAbsLocation", gridControl62);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listBarcodeColdTravAbsLocation", gridControl63);

            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listColdLiftBrakeOption", gridControl36);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listColdTravBrakeOption", gridControl66);
            pathManager.LoadListFromXmlToDataTable(StringUnits.strItemListFilePath, "listColdForkBrakeOption", gridControl67);




        }

        private void SaveToXmlData()
        {
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listMODName", gridControl1);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listMODOption", gridControl2);

            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listMSPinputVolt", gridControl3);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listMSPinputHz", gridControl4);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listMSPcontrollerMaker", gridControl48);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listMSPcontrollerSpec", gridControl5);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listMSPinverterMaker", gridControl49);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listMSPinverterSpec", gridControl6);

            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listOPmachineControl", gridControl7);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listOPremoteControl", gridControl8);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listOPemergencyPower", gridControl9);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listOPemergencyLocation", gridControl10);

            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqUsingVoltage", gridControl56);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqMccbModel", gridControl11);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqSmpsModel", gridControl12);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqCableModel", gridControl13);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqHubModel", gridControl14);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqFanQuantity", gridControl46);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqTerminal", gridControl47);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqPanel", gridControl50);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqHmi", gridControl51);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqOpt", gridControl52);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqTowerLamp", gridControl55);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqSafetyEmo", gridControl53);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqEmoQuantity", gridControl54);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqModem", gridControl15);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqInterLockSensorSide", gridControl16);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqInterLockBit", gridControl17);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqNpnSensorItem", gridControl18);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listEleqPnpSensorItem", gridControl19);

            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listLiftBrakeOption", gridControl23);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listLiftMotorSpec", gridControl20);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listLiftMotorMaker", gridControl26);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listLiftMotorMethod", gridControl60);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listLiftRaserAbsLocation", gridControl21);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listLiftBarcodeAbsLocation", gridControl22);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listLiftNpnRightPosition", gridControl25);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listLiftPnpRightPosition", gridControl37);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listLiftLimitSwitch", gridControl24);

            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listTravBrakeOption", gridControl30);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listTravMotorSpec", gridControl27);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listTravMotorMaker", gridControl64);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listTravMotorMethod", gridControl65);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listTravRaserAbsLocation", gridControl28);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listTravBarcodeAbsLocation", gridControl29);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listTravNpnRightPosition", gridControl31);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listTravPnpRightPosition", gridControl45);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listTravLimitSwitch", gridControl32);

            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listForkBrakeOption", gridControl38);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listForkMotorSpec", gridControl33);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listForkMotorMaker", gridControl34);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listForkMotorMethod", gridControl35);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listForkNpnRightPosition", gridControl39);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listForkPnpRightPosition", gridControl40);

            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listCarrNpnSensor", gridControl41);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listCarrPnpSensor", gridControl42);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listCarrNpnDoubleInput", gridControl43);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listCarrPnpDoubleInput", gridControl44);

            //콜드 타입
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listColdEleqModem", gridControl57);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listColdEleqSensorItem", gridControl58);


            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listRaserColdLiftAbsLocation", gridControl59);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listBarcodeColdLiftAbsLocation", gridControl61);

            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listRaserColdTravAbsLocation", gridControl62);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listBarcodeColdTravAbsLocation", gridControl63);

            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listColdLiftBrakeOption", gridControl36);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listColdTravBrakeOption", gridControl66);
            pathManager.SaveListFromDataTableToXml(StringUnits.strItemListFilePath, "listColdForkBrakeOption", gridControl67);

        }

        private void SetGridView()
        {
            gvManager.SetItemListGridView(gridView1);
            gvManager.SetItemListGridView(gridView2);
            gvManager.SetItemListGridView(gridView3);
            gvManager.SetItemListGridView(gridView4);
            gvManager.SetItemListGridView(gridView5);
            gvManager.SetItemListGridView(gridView6);
            gvManager.SetItemListGridView(gridView7);
            gvManager.SetItemListGridView(gridView8);
            gvManager.SetItemListGridView(gridView9);
            gvManager.SetItemListGridView(gridView10);
            gvManager.SetItemListGridView(gridView11);
            gvManager.SetItemListGridView(gridView12);
            gvManager.SetItemListGridView(gridView13);
            gvManager.SetItemListGridView(gridView14);
            gvManager.SetItemListGridView(gridView15);
            gvManager.SetItemListGridView(gridView16);
            gvManager.SetItemListGridView(gridView17);
            gvManager.SetItemListGridView(gridView18);
            gvManager.SetItemListGridView(gridView19);
            gvManager.SetItemListGridView(gridView20);
            gvManager.SetItemListGridView(gridView21);
            gvManager.SetItemListGridView(gridView22);
            gvManager.SetItemListGridView(gridView23);
            gvManager.SetItemListGridView(gridView24);
            gvManager.SetItemListGridView(gridView27);
            gvManager.SetItemListGridView(gridView28);
            gvManager.SetItemListGridView(gridView29);
            gvManager.SetItemListGridView(gridView30);
            gvManager.SetItemListGridView(gridView31);
            gvManager.SetItemListGridView(gridView45);
            gvManager.SetItemListGridView(gridView32);
            gvManager.SetItemListGridView(gridView33);
            gvManager.SetItemListGridView(gridView34);
            gvManager.SetItemListGridView(gridView35);
            gvManager.SetItemListGridView(gridView38);
            gvManager.SetItemListGridView(gridView39);
            gvManager.SetItemListGridView(gridView40);
            gvManager.SetItemListGridView(gridView41);
            gvManager.SetItemListGridView(gridView42);
            gvManager.SetItemListGridView(gridView43);
            gvManager.SetItemListGridView(gridView44);
            gvManager.SetItemListGridView(gridView46);
            gvManager.SetItemListGridView(gridView47);
            gvManager.SetItemListGridView(gridView48);
            gvManager.SetItemListGridView(gridView49);
            gvManager.SetItemListGridView(gridView50);
            gvManager.SetItemListGridView(gridView51);
            gvManager.SetItemListGridView(gridView52);
            gvManager.SetItemListGridView(gridView53);
            gvManager.SetItemListGridView(gridView54);
            gvManager.SetItemListGridView(gridView55);
            gvManager.SetItemListGridView(gridView56);

            gvManager.SetItemListGridView(gridView57);
            gvManager.SetItemListGridView(gridView58);
            gvManager.SetItemListGridView(gridView59);
            gvManager.SetItemListGridView(gridView61);
            gvManager.SetItemListGridView(gridView62);
            gvManager.SetItemListGridView(gridView63);
            gvManager.SetItemListGridView(gridView26);
            gvManager.SetItemListGridView(gridView60);
            gvManager.SetItemListGridView(gridView25);
            gvManager.SetItemListGridView(gridView37);
            gvManager.SetItemListGridView(gridView64);
            gvManager.SetItemListGridView(gridView65);

            gvManager.SetItemListGridView(gridView36);
            gvManager.SetItemListGridView(gridView66);
            gvManager.SetItemListGridView(gridView67);
        }
    }
}
