using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DevExpress.ClipboardSource.SpreadsheetML;
using Eplan.MCNS.Lib;
using Eplan.MCNS.Lib.Share_CS;
using Eplan.MCNS.Lib.UI_CS;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Xml.Linq;

namespace Eplan.EplAddin.HMX_MCNS
{
    public partial class FormItemsList : Form
    {
        CS_DataGridView cs_DataGridView = new CS_DataGridView();
        CS_ListItems cs_ListItems = new CS_ListItems();
        

        // 컨트롤 DPI 스케일링 조정
      


        public FormItemsList()
        {
            InitializeComponent();
            
            EventMainForm();

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
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listMODName", gridControl1);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listMODOption", gridControl2);

            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listMSPinputVolt", gridControl3);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listMSPinputHz", gridControl4);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listMSPcontrollerMaker", gridControl48);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listMSPcontrollerSpec", gridControl5);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listMSPinverterMaker", gridControl49);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listMSPinverterSpec", gridControl6);

            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listOPmachineControl", gridControl7);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listOPremoteControl", gridControl8);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listOPemergencyPower", gridControl9);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listOPemergencyLocation", gridControl10);

            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqUsingVoltage", gridControl56);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqMccbModel", gridControl11);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqSmpsModel", gridControl12);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqCableModel", gridControl13);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqHubModel", gridControl14);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqFanQuantity", gridControl46);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqTerminal", gridControl47);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqPanel", gridControl50);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqHmi", gridControl51);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqOpt", gridControl52);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqTowerLamp", gridControl55);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqSafety", gridControl53);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqSafetyQuantity", gridControl54);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqModem", gridControl15);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqInterLockSensorSide", gridControl16);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqInterLockBit", gridControl17);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqNpnSensorItem", gridControl18);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listEleqPnpSensorItem", gridControl19);

            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listLiftBrakeOption", gridControl23);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listLiftMotorSpec", gridControl20);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listLiftMotorMaker", gridControl26);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listLiftMotorMethod", gridControl60);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listLiftRaserAbsLocation", gridControl21);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listLiftBarcodeAbsLocation", gridControl22);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listLiftNpnRightPosition", gridControl25);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listLiftPnpRightPosition", gridControl37);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listLiftLimitSwitch", gridControl24);

            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listTravBrakeOption", gridControl30);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listTravMotorSpec", gridControl27);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listTravMotorMaker", gridControl64);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listTravMotorMethod", gridControl65);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listTravRaserAbsLocation", gridControl28);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listTravBarcodeAbsLocation", gridControl29);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listTravNpnRightPosition", gridControl31);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listTravPnpRightPosition", gridControl45);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listTravLimitSwitch", gridControl32);

            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listForkBrakeOption", gridControl38);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listForkMotorSpec", gridControl33);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listForkMotorMaker", gridControl34);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listForkMotorMethod", gridControl35);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listForkNpnRightPosition", gridControl39);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listForkPnpRightPosition", gridControl40);

            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listCarrNpnSensor", gridControl41);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listCarrPnpSensor", gridControl42);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listCarrNpnDoubleInput", gridControl43);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listCarrPnpDoubleInput", gridControl44);

            //콜드 타입
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listColdEleqModem", gridControl57);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listColdEleqSensorItem", gridControl58);

            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listRaserColdLiftAbsLocation", gridControl59);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listBarcodeColdLiftAbsLocation", gridControl61);

            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listRaserColdTravAbsLocation", gridControl62);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listBarcodeColdTravAbsLocation", gridControl63);

            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listColdLiftBrakeOption", gridControl36);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listColdTravBrakeOption", gridControl66);
            cs_ListItems.LoadListFromXmlToDataTable(CS_PathData.ItemListFilePath, "listColdForkBrakeOption", gridControl67);




        }

        private void SaveToXmlData()
        {
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listMODName", gridControl1);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listMODOption", gridControl2);

            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listMSPinputVolt", gridControl3);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listMSPinputHz", gridControl4);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listMSPcontrollerMaker", gridControl48);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listMSPcontrollerSpec", gridControl5);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listMSPinverterMaker", gridControl49);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listMSPinverterSpec", gridControl6);

            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listOPmachineControl", gridControl7);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listOPremoteControl", gridControl8);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listOPemergencyPower", gridControl9);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listOPemergencyLocation", gridControl10);

            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqUsingVoltage", gridControl56);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqMccbModel", gridControl11);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqSmpsModel", gridControl12);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqCableModel", gridControl13);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqHubModel", gridControl14);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqFanQuantity", gridControl46);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqTerminal", gridControl47);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqPanel", gridControl50);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqHmi", gridControl51);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqOpt", gridControl52);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqTowerLamp", gridControl55);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqSafetyEmo", gridControl53);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqEmoQuantity", gridControl54);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqModem", gridControl15);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqInterLockSensorSide", gridControl16);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqInterLockBit", gridControl17);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqNpnSensorItem", gridControl18);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listEleqPnpSensorItem", gridControl19);

            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listLiftBrakeOption", gridControl23);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listLiftMotorSpec", gridControl20);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listLiftMotorMaker", gridControl26);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listLiftMotorMethod", gridControl60);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listLiftRaserAbsLocation", gridControl21);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listLiftBarcodeAbsLocation", gridControl22);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listLiftNpnRightPosition", gridControl25);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listLiftPnpRightPosition", gridControl37);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listLiftLimitSwitch", gridControl24);

            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listTravBrakeOption", gridControl30);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listTravMotorSpec", gridControl27);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listTravMotorMaker", gridControl64);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listTravMotorMethod", gridControl65);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listTravRaserAbsLocation", gridControl28);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listTravBarcodeAbsLocation", gridControl29);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listTravNpnRightPosition", gridControl31);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listTravPnpRightPosition", gridControl45);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listTravLimitSwitch", gridControl32);

            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listForkBrakeOption", gridControl38);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listForkMotorSpec", gridControl33);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listForkMotorMaker", gridControl34);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listForkMotorMethod", gridControl35);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listForkNpnRightPosition", gridControl39);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listForkPnpRightPosition", gridControl40);

            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listCarrNpnSensor", gridControl41);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listCarrPnpSensor", gridControl42);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listCarrNpnDoubleInput", gridControl43);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listCarrPnpDoubleInput", gridControl44);

            //콜드 타입
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listColdEleqModem", gridControl57);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listColdEleqSensorItem", gridControl58);


            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listRaserColdLiftAbsLocation", gridControl59);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listBarcodeColdLiftAbsLocation", gridControl61);

            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listRaserColdTravAbsLocation", gridControl62);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listBarcodeColdTravAbsLocation", gridControl63);

            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listColdLiftBrakeOption", gridControl36);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listColdTravBrakeOption", gridControl66);
            cs_ListItems.SaveListFromDataTableToXml(CS_PathData.ItemListFilePath, "listColdForkBrakeOption", gridControl67);

        }

        private void SetGridView()
        {
            cs_DataGridView.SetItemListGridView(gridView1);
            cs_DataGridView.SetItemListGridView(gridView2);
            cs_DataGridView.SetItemListGridView(gridView3);
            cs_DataGridView.SetItemListGridView(gridView4);
            cs_DataGridView.SetItemListGridView(gridView5);
            cs_DataGridView.SetItemListGridView(gridView6);
            cs_DataGridView.SetItemListGridView(gridView7);
            cs_DataGridView.SetItemListGridView(gridView8);
            cs_DataGridView.SetItemListGridView(gridView9);
            cs_DataGridView.SetItemListGridView(gridView10);
            cs_DataGridView.SetItemListGridView(gridView11);
            cs_DataGridView.SetItemListGridView(gridView12);
            cs_DataGridView.SetItemListGridView(gridView13);
            cs_DataGridView.SetItemListGridView(gridView14);
            cs_DataGridView.SetItemListGridView(gridView15);
            cs_DataGridView.SetItemListGridView(gridView16);
            cs_DataGridView.SetItemListGridView(gridView17);
            cs_DataGridView.SetItemListGridView(gridView18);
            cs_DataGridView.SetItemListGridView(gridView19);
            cs_DataGridView.SetItemListGridView(gridView20);
            cs_DataGridView.SetItemListGridView(gridView21);
            cs_DataGridView.SetItemListGridView(gridView22);
            cs_DataGridView.SetItemListGridView(gridView23);
            cs_DataGridView.SetItemListGridView(gridView24);
            cs_DataGridView.SetItemListGridView(gridView27);
            cs_DataGridView.SetItemListGridView(gridView28);
            cs_DataGridView.SetItemListGridView(gridView29);
            cs_DataGridView.SetItemListGridView(gridView30);
            cs_DataGridView.SetItemListGridView(gridView31);
            cs_DataGridView.SetItemListGridView(gridView45);
            cs_DataGridView.SetItemListGridView(gridView32);
            cs_DataGridView.SetItemListGridView(gridView33);
            cs_DataGridView.SetItemListGridView(gridView34);
            cs_DataGridView.SetItemListGridView(gridView35);
            cs_DataGridView.SetItemListGridView(gridView38);
            cs_DataGridView.SetItemListGridView(gridView39);
            cs_DataGridView.SetItemListGridView(gridView40);
            cs_DataGridView.SetItemListGridView(gridView41);
            cs_DataGridView.SetItemListGridView(gridView42);
            cs_DataGridView.SetItemListGridView(gridView43);
            cs_DataGridView.SetItemListGridView(gridView44);
            cs_DataGridView.SetItemListGridView(gridView46);
            cs_DataGridView.SetItemListGridView(gridView47);
            cs_DataGridView.SetItemListGridView(gridView48);
            cs_DataGridView.SetItemListGridView(gridView49);
            cs_DataGridView.SetItemListGridView(gridView50);
            cs_DataGridView.SetItemListGridView(gridView51);
            cs_DataGridView.SetItemListGridView(gridView52);
            cs_DataGridView.SetItemListGridView(gridView53);
            cs_DataGridView.SetItemListGridView(gridView54);
            cs_DataGridView.SetItemListGridView(gridView55);
            cs_DataGridView.SetItemListGridView(gridView56);

            cs_DataGridView.SetItemListGridView(gridView57);
            cs_DataGridView.SetItemListGridView(gridView58);
            cs_DataGridView.SetItemListGridView(gridView59);
            cs_DataGridView.SetItemListGridView(gridView61);
            cs_DataGridView.SetItemListGridView(gridView62);
            cs_DataGridView.SetItemListGridView(gridView63);
            cs_DataGridView.SetItemListGridView(gridView26);
            cs_DataGridView.SetItemListGridView(gridView60);
            cs_DataGridView.SetItemListGridView(gridView25);
            cs_DataGridView.SetItemListGridView(gridView37);
            cs_DataGridView.SetItemListGridView(gridView64);
            cs_DataGridView.SetItemListGridView(gridView65);

            cs_DataGridView.SetItemListGridView(gridView36);
            cs_DataGridView.SetItemListGridView(gridView66);
            cs_DataGridView.SetItemListGridView(gridView67);
        }

        public void EventMainForm()
        {
            // 마우스 클릭 폼 이동
            pnlTap.MouseDown += (o, e) => { if (e.Button == MouseButtons.Left) { CS_StaticEtc.On = true; CS_StaticEtc.Pos = e.Location; } };
            pnlTap.MouseMove += (o, e) => { if (CS_StaticEtc.On) Location = new Point(Location.X + (e.X - CS_StaticEtc.Pos.X), Location.Y + (e.Y - CS_StaticEtc.Pos.Y)); };
            pnlTap.MouseUp += (o, e) => { if (e.Button == MouseButtons.Left) { CS_StaticEtc.On = false; CS_StaticEtc.Pos = e.Location; } };
            
        }
    }
}
