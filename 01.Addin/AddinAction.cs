using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel.MasterData;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Linq;
using System.Runtime.Hosting;
using System.Threading;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Eplan.MCNS.Lib;
using System.Diagnostics;

namespace Eplan.EplAddin.HMX_MCNS
{
    public class AddinAction : IEplAction
    {


        public bool Execute(ActionCallingContext acc)
        {
            Form formConceptSheet = null;
            Form formInitialPage = null;

            // 현재 열려 있는 폼을 확인
            foreach (Form form in Application.OpenForms)
            {
                if (form is FormConceptSheet) formConceptSheet = form;
                if (form is FormInitialPage) formInitialPage = form;
            }

            // FormConceptSheet가 열려 있으면 활성화하고, FormInitialPage도 열려 있으면 활성화
            if (formConceptSheet != null)
            {
                formConceptSheet.WindowState = FormWindowState.Normal;
                formConceptSheet.Activate();

                if (formInitialPage != null)
                {
                    formInitialPage.WindowState = FormWindowState.Normal;
                    formInitialPage.Activate();
                }
                return true;
            }

            // FormConceptSheet가 없으면 FormInitialPage 닫기
            formInitialPage?.Close();

            // FormInitialPage 새로 생성 후 열기
            
            FormUnits.formInitialPage = new FormInitialPage();
            FormUnits.formInitialPage.Show(new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle));

            return true;
        }

        public void GetActionProperties(ref ActionProperties ap)
        {
        }

        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = AddinStatic.actionName;
            return true;
        }
    }
}
