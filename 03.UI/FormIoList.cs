using DevExpress.XtraEditors;
using Eplan.MCNS.Lib.UI_CS;
using Eplan.MCNS.Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Eplan.EplAddin.HMX_MCNS._03.UI
{
    public partial class FormIoList : DevExpress.XtraEditors.XtraForm
    {
        CS_DataGridView cs_DataGridView = new CS_DataGridView();

        public FormIoList()
        {
            InitializeComponent();

            ControlFormFunction();



            // GridControl의 데이터 소스 갱신
            gridControl1.DataSource = CS_StaticSensor.sensorCopyIoDt;
            cs_DataGridView.SetIoGridView(gridView1);

            btnSaveIo.MouseClick += (o, e) =>
            {
                // CS_StaticSensor.sensorIoDt의 내용을 지우고
                CS_StaticSensor.sensorIoDt.Clear();

                // copyDt의 수정된 내용을 CS_StaticSensor.sensorIoDt에 복사
                foreach (DataRow row in CS_StaticSensor.sensorCopyIoDt.Rows)
                {
                    CS_StaticSensor.sensorIoDt.ImportRow(row);
                }

                MessageBox.Show("변경 사항이 저장되었습니다.", "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);

            };
        }
        public void ControlFormFunction()
        {
            //마우스 클릭 폼 이동
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

                // 사용자가 "No"를 선택하면 폼 종료 취소
                if (result == DialogResult.No)
                {
                    e.Cancel = true; // 종료 취소
                    return;
                }

                // "Yes"를 선택하면 기본 동작으로 폼이 닫힘
            };



        }
    }
}
