using DevExpress.XtraEditors;
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

namespace Eplan.EplAddin.HMX_MCNS
{
    public partial class FormIoList : DevExpress.XtraEditors.XtraForm
    {
        GridViewManager gvManager = new GridViewManager();

        public FormIoList()
        {
            InitializeComponent();

            ControlFormFunction();



            // GridControl의 데이터 소스 갱신
            gridControl1.DataSource = DataTableUnits.dtSensorCopyIo;
            gvManager.SetIoGridView(gridView1);

            btnSaveIo.MouseClick += (o, e) =>
            {
                // CS_StaticSensor.sensorIoDt의 내용을 지우고
                DataTableUnits.dtSensorIo.Clear();

                // copyDt의 수정된 내용을 CS_StaticSensor.sensorIoDt에 복사
                foreach (DataRow row in DataTableUnits.dtSensorCopyIo.Rows)
                {
                    DataTableUnits.dtSensorIo.ImportRow(row);
                }

                MessageBox.Show("변경 사항이 저장되었습니다.", "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);

            };
        }
        public void ControlFormFunction()
        {
            this.FormClosing += (o, e) =>
            {
                // 종료 확인 메시지 표시
                DialogResult result = MessageBox.Show(
                    "IO 리스트 작성을 종료하시겠습니까?",
                    "IO 리스트 작성하기 종료",
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
