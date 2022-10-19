using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TotalMEPProject.UI;

namespace TotalMEPProject.Commands.Login
{
    [Transaction(TransactionMode.Manual)]
    public class CmdLogin : IExternalCommand
    {
        private UIDocument _uiDoc;
        private Document _doc;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            _uiDoc = uiapp.ActiveUIDocument;
            _doc = _uiDoc.Document;

            FrmLogin UI_Login = new FrmLogin();
            UI_Login.ShowDialog();

            //bool isHasInternet = LicenseUtils.CheckForInternetConnection(10000, "http://www.google.com");

            //string errMess = "";
            //bool isValidLicense = LicenseUtils.CheckLicense(isHasInternet, ref errMess);

            return Result.Succeeded;
        }
    }
}