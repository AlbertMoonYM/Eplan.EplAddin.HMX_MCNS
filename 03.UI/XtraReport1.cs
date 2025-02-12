using DevExpress.Diagram.Core.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraPrinting.Preview;
using DevExpress.XtraReports.UI;
using DocumentFormat.OpenXml.Drawing.Charts;
using Eplan.MCNS.Lib;
using Eplan.MCNS.Lib.UI_CS;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Eplan.EplAddin.HMX_MCNS._03.UI
{
    public partial class XtraReport1 : DevExpress.XtraReports.UI.XtraReport
    {
        public XtraReport1()
        {
            InitializeComponent();


            string[] keys = CS_StaticUnit.dicCtrlMod.Keys.ToArray();

            Control[] prjCtrls = CS_StaticUnit.dicCtrlMod[keys[0]];
            Control[] mspCtrls = CS_StaticUnit.dicCtrlMod[keys[2]];
            Control[] loutCtrls = CS_StaticUnit.dicCtrlMod[keys[3]];
            Control[] opCtrls = CS_StaticUnit.dicCtrlMod[keys[4]];

            Control[] modelPageCtrls = prjCtrls.Concat(mspCtrls).Concat(loutCtrls).Concat(opCtrls).ToArray();

            // XRTableCell 배열 초기화
            XRTableCell[] prjCells = new XRTableCell[] { xrtPrj01, xrtPrj02, xrtPrj03, xrtPrj04, xrtPrj05, xrtPrj07, xrtPrj08, xrtPrj09 };
            XRTableCell[] mspCells = new XRTableCell[] { xrtMsp01, xrtMsp02, xrtMsp03, xrtMsp04, xrtMsp05, xrtMsp06, xrtMsp07, xrtMsp08, xrtMsp09, xrtMsp10};
            XRTableCell[] loutCells = new XRTableCell[] { xrtLout01,xrtLout02,xrtLout03,xrtLout04,xrtLout05};
            XRTableCell[] opCells = new XRTableCell[] { xrtOp01, xrtOp02, xrtOp03, xrtOp04, xrtOp05, xrtOp06, xrtOp07, xrtOp08};

            XRTableCell[] modCells = prjCells.Concat(mspCells).Concat(loutCells).Concat(opCells).ToArray();
            // 모든 XRTableCell의 텍스트를 초기화
            foreach (XRTableCell cell in modCells)
            {
                cell.Text = ""; // 모든 셀의 텍스트를 빈 문자열로 설정
            }

            for (int i = 0; i < modelPageCtrls.Length; i++)
            {
                // 컨트롤이 ComboBox인지 확인
                if (modelPageCtrls[i] is ComboBoxEdit comboBox && comboBox.BackColor == Color.White)
                {
                    modCells[i].Text = comboBox.Text; // 해당 셀의 텍스트 업데이트
                }
                // 컨트롤이 CheckBox인지 확인
                else if (modelPageCtrls[i] is CheckEdit checkBox)
                {
                    modCells[i].Text = checkBox.Checked ? checkBox.Text : ""; // Checked 상태에 따라 텍스트 설정
                }
            }

            xrtMod01.Text = CS_StaticUnit.strModFullName;

            XRTableCell[] eleqCells = new XRTableCell[]
            {
                xrtEleq100, xrtEleq110, xrtEleq120, xrtEleq130,
                xrtEleq200, xrtEleq210, xrtEleq220, xrtEleq230,
                xrtEleq300, xrtEleq310, xrtEleq320, xrtEleq330,
                xrtEleq400,xrtEleq410,xrtEleq420,xrtEleq430,
                xrtEleq500,xrtEleq510,xrtEleq520,xrtEleq530,
                xrtEleq600, xrtEleq610, xrtEleq620, xrtEleq630,xrtEleq640
            };
            XRTableCell[] liftCells = new XRTableCell[]
            {
                xrtLift100, xrtLift110, xrtLift120, xrtLift130, 
                xrtLift200, xrtLift210, xrtLift220, 
                xrtLift300, xrtLift310, xrtLift320, xrtLift330,
                xrtLift400, xrtLift410, xrtLift420,
                xrtLift500, xrtLift510, xrtLift520, xrtLift530,
                xrtLift600, xrtLift610, xrtLift620,
                xrtLift700, xrtLift710, xrtLift720, xrtLift730, xrtLift740,
                xrtLift800, xrtLift810, xrtLift820,
                xrtLift900, xrtLift910, xrtLift920,
            };
            XRTableCell[] travCells = new XRTableCell[]
            {
                xrtTrav100, xrtTrav110, xrtTrav120, xrtTrav130,
                xrtTrav200, xrtTrav210, xrtTrav220,
                xrtTrav300, xrtTrav310, xrtTrav320, xrtTrav330,
                xrtTrav400, xrtTrav410, xrtTrav420,
                xrtTrav500, xrtTrav510, xrtTrav520, xrtTrav530,
                xrtTrav600, xrtTrav610, xrtTrav620,
                xrtTrav700, xrtTrav710, xrtTrav720, xrtTrav730, xrtTrav740,
                xrtTrav800, xrtTrav810, xrtTrav820,
                xrtTrav900, xrtTrav910, xrtTrav920,
            };
            XRTableCell[] trav2Cells = new XRTableCell[]
            {
                xrt2Trav100, xrt2Trav110, xrt2Trav120, xrt2Trav130,
                xrt2Trav200, xrt2Trav210, xrt2Trav220,
                xrt2Trav300, xrt2Trav310, xrt2Trav320, xrt2Trav330,
                xrt2Trav400, xrt2Trav410, xrt2Trav420,
                xrt2Trav500, xrt2Trav510, xrt2Trav520, xrt2Trav530,
                xrt2Trav600, xrt2Trav610, xrt2Trav620,
                xrt2Trav700, xrt2Trav710, xrt2Trav720, xrt2Trav730, xrt2Trav740,
                xrt2Trav800, xrt2Trav810, xrt2Trav820,
                xrt2Trav900, xrt2Trav910, xrt2Trav920,
            };
            XRTableCell[] forkCells = new XRTableCell[]
            {
                xrtFork100, xrtFork110, xrtFork120, xrtFork130,
                xrtFork200, xrtFork210, xrtFork220,
                xrtFork300, xrtFork310, xrtFork320, xrtFork330,
                xrtFork400, xrtFork410, xrtFork420,
                xrtFork500, xrtFork510, xrtFork520, xrtFork530,
                xrtFork600, xrtFork610, xrtFork620,
                xrtFork700, xrtFork710,
                xrtFork800, xrtFork810, xrtFork820,
                xrtFork900, xrtFork910, xrtFork920,
            };


            XRTableCell[] fork2Cells = new XRTableCell[]
            {
                xrt2Fork100, xrt2Fork110, xrt2Fork120, xrt2Fork130,
                xrt2Fork200, xrt2Fork210, xrt2Fork220,
                xrt2Fork300, xrt2Fork310, xrt2Fork320, xrt2Fork330,
                xrt2Fork400, xrt2Fork410, xrt2Fork420,
                xrt2Fork500, xrt2Fork510, xrt2Fork520, xrt2Fork530,
                xrt2Fork600, xrt2Fork610, xrt2Fork620,
                xrt2Fork700, xrt2Fork710,
                xrt2Fork800, xrt2Fork810, xrt2Fork820,
                xrt2Fork900, xrt2Fork910, xrt2Fork920,
            };
            XRTableCell[] carrCells = new XRTableCell[]
            {
                xrtCarr500,xrtCarr600
            };


            XRTableCell[] funcCells = eleqCells.Concat(liftCells).Concat(travCells).Concat(trav2Cells).Concat(forkCells).Concat(fork2Cells).Concat(carrCells).ToArray();
            // 컨트롤을 담을 리스트 생성
            List<Control> combinedControls = new List<Control>();

            foreach (XRTableCell cell in funcCells)
            {
                cell.Text = ""; // 모든 셀의 텍스트를 빈 문자열로 설정
            }

            // 딕셔너리의 모든 Control[] 배열을 리스트에 추가
            foreach (var controlArray in CS_StaticUnit.dicCtrlFunc.Values)
            {
                combinedControls.AddRange(controlArray);  // 각 Control[] 배열을 리스트에 추가
            }


            // 리스트를 배열로 변환하여 Control[] funcPageCtrls에 할당
            Control[] funcPageCtrls = combinedControls.ToArray();
            
            for (int i = 0; i < funcPageCtrls.Length; i++)
            {
                // 컨트롤이 ComboBox인지 확인
                if (funcPageCtrls[i] is ComboBoxEdit comboBox && comboBox.BackColor == Color.White)
                {
                    funcCells[i].Text = comboBox.Text; // 해당 셀의 텍스트 업데이트
                }
                // 컨트롤이 CheckBox인지 확인
                else if (funcPageCtrls[i] is CheckEdit checkBox)
                {
                    funcCells[i].Text = checkBox.Checked ? checkBox.Text : ""; // Checked 상태에 따라 텍스트 설정
                }
            }



            xrtCarr100.Text = CS_StaticUnit.dtLout.Rows[0][1].ToString();
            xrtCarr110.Text = CS_StaticUnit.dtLout.Rows[0][2].ToString();
            xrtCarr120.Text = CS_StaticUnit.dtLout.Rows[0][3].ToString();
            xrtCarr200.Text = CS_StaticUnit.dtLout.Rows[1][1].ToString();
            xrtCarr210.Text = CS_StaticUnit.dtLout.Rows[1][2].ToString();
            xrtCarr220.Text = CS_StaticUnit.dtLout.Rows[1][3].ToString();
            xrtCarr300.Text = CS_StaticUnit.dtLout.Rows[2][1].ToString();
            xrtCarr310.Text = CS_StaticUnit.dtLout.Rows[2][2].ToString();
            xrtCarr320.Text = CS_StaticUnit.dtLout.Rows[2][3].ToString();
            xrtCarr400.Text = CS_StaticUnit.dtLout.Rows[3][1].ToString();
            xrtCarr410.Text = CS_StaticUnit.dtLout.Rows[3][2].ToString();
            xrtCarr420.Text = CS_StaticUnit.dtLout.Rows[3][3].ToString();

            xrtLift1000.Text = CS_StaticSensor.listLiftSensor;
            xrtTrav1000.Text = CS_StaticSensor.listTrav1Sensor;
            xrtFork1000.Text = CS_StaticSensor.listFork1Sensor;
            xrt2Trav1000.Text = CS_StaticSensor.listTrav2Sensor;
            xrt2Fork1000.Text = CS_StaticSensor.listFork2Sensor;
            xrtCarr1000.Text = CS_StaticSensor.listCarrSensor;

            

        }
    }
}
