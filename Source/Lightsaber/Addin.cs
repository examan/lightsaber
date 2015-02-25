using NetOffice.OfficeApi.Tools;
using NetOffice.PowerPointApi.Tools;
using NetOffice.Tools;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = NetOffice.PowerPointApi;

namespace Lightsaber
{
    [
        COMAddin("Lightsaber", "Hightlight add-in for PowerPoint", 3),
        ProgId("Lightsaber.Addin"),
        Guid("052E6528-40D1-4A76-BEDE-FA6416F6DA79"),
        RegistryLocation(RegistrySaveLocation.LocalMachine)
    ]
    public partial class Addin : PowerPoint.Tools.COMAddin
    {
        public Addin()
        {
            this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);
            this.OnDisconnection += new OnDisconnectionEventHandler(Addin_OnDisconnection);
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            Console.WriteLine("Addin started in PowerPoint Version {0}", Application.Version);
        }

        private void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {

        }

        protected override void OnError(ErrorMethodKind methodKind, System.Exception exception)
        {
            MessageBox.Show("An error occurend in " + methodKind.ToString(), "Lightsaber");
        }

        [RegisterErrorHandler]
        public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, System.Exception exception)
        {
            MessageBox.Show("An error occurend in " + methodKind.ToString(), "Lightsaber");
        }
    }
}
