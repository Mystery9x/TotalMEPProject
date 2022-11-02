using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using TotalMEPProject.Request;
using TotalMEPProject.Services;
using TotalMEPProject.UI;
using TotalMEPProject.Ultis;

namespace TotalMEPProject
{
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication
    {
        #region Variable

        public static UIApplication uiApp;
        public static UIControlledApplication cachedUiCtrApp;

        public static WindowHandle hWndRevit = null;
        public static VerticalMEPForm verticalMEPForm = null;

        #endregion Variable

        #region Method

        #region Revit Event

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            // Create ribbon tools
            cachedUiCtrApp = application;
            CreateRibbonButtons(application);

            return Result.Succeeded;
        }

        #endregion Revit Event

        #region Create ribbon

        /// <summary>
        /// Create ribbon tools
        /// </summary>
        /// <param name="app"></param>
        private void CreateRibbonButtons(UIControlledApplication app)
        {
            //Get dll path
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            string iconFolder = GetIconFolder();

            string tabName = "TotalMEP";
            app.CreateRibbonTab(tabName);

            CreateTotalMEPTab(app, tabName, assemblyPath, iconFolder);

            CreatePlumbingTab(app, tabName, assemblyPath, iconFolder);

            CreateFireFightingTab(app, tabName, assemblyPath, iconFolder);

            CreateDuctTab(app, tabName, assemblyPath, iconFolder);

            CreateOpeningTab(app, tabName, assemblyPath, iconFolder);

            CreateModifyTab(app, tabName, assemblyPath, iconFolder);

            CreateLoginTab(app, tabName, assemblyPath, iconFolder);

            bool isHasInternet = LicenseUtils.CheckForInternetConnection();
        }

        /// <summary>
        /// Create login tab
        /// </summary>
        /// <param name="app"></param>
        /// <param name="tabName"></param>
        /// <param name="assemblyPath"></param>
        /// <param name="iconFolder"></param>
        private void CreateLoginTab(UIControlledApplication app,
                                    string tabName,
                                    string assemblyPath,
                                    string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel loginHangerPanel = app.CreateRibbonPanel(tabName, Define.LoginLicense);

            //Create button
            PushButtonData loginData = new PushButtonData("btnLogin", "License Server", assemblyPath, Define.CmdLoginClassName);
            //AddImages(loginData, iconFolder, "Setting.png", "Setting.png");
            loginHangerPanel.AddItem(loginData);
        }

        /// <summary>
        /// Create total mep tab
        /// </summary>
        /// <param name="app"></param>
        /// <param name="tabName"></param>
        /// <param name="assemblyPath"></param>
        /// <param name="iconFolder"></param>
        private void CreateTotalMEPTab(UIControlledApplication app,
                                                    string tabName,
                                                    string assemblyPath,
                                                    string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel totalMEPPanel = app.CreateRibbonPanel(tabName, Define.TotalMEPRibbonTabName);

            //Create button
            PushButtonData pickLineData = new PushButtonData("btnPickLine", "PickLine", assemblyPath, Define.CmdPickLineClassName);
            //AddImages(veticalMEPData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            totalMEPPanel.AddItem(pickLineData);

            //Create button
            PushButtonData veticalMEPData = new PushButtonData("btnVeticalMEP", "Vetical\nMEP", assemblyPath, Define.CmdVerticalMEPClassName);
            //AddImages(veticalMEPData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            totalMEPPanel.AddItem(veticalMEPData);

            //Create button
            PushButtonData fastVerticalData = new PushButtonData("btnFastVertical", "Fast\nVertical", assemblyPath, Define.CmdFastVerticalClassName);
            //AddImages(fastVerticalData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            totalMEPPanel.AddItem(fastVerticalData);

            //Create button
            PushButtonData HolyUpdownData = new PushButtonData("btnHolyUpdown", "Holy\nUpdown", assemblyPath, Define.CmdHolyUpdownClassName);
            //AddImages(HolyUpdownData, iconFolder, "Manual Hanger.png", "Manual Hanger.png");
            totalMEPPanel.AddItem(HolyUpdownData);

            //Create button
            PushButtonData changeElevationData = new PushButtonData("btnChangeElevation", "Change\nElevation", assemblyPath, Define.CmdChangeElevationClassName);
            //AddImages(changeElevationData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            totalMEPPanel.AddItem(changeElevationData);

            //Create button
            PushButtonData MEPConnectionData = new PushButtonData("btnMEPConnection", "MEP\nConnection", assemblyPath, Define.CmdMEPConnectionClassName);
            //AddImages(reinforcedHangerData, iconFolder, "Re-Inforced Hanger.png", "Re-Inforced Hanger.png");
            totalMEPPanel.AddItem(MEPConnectionData);
        }

        /// <summary>
        /// Create plumbing tab
        /// </summary>
        /// <param name="app"></param>
        /// <param name="tabName"></param>
        /// <param name="assemblyPath"></param>
        /// <param name="iconFolder"></param>
        private void CreatePlumbingTab(UIControlledApplication app,
                                                   string tabName,
                                                   string assemblyPath,
                                                   string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel plumbingPanel = app.CreateRibbonPanel(tabName, Define.PlumpingRibbonTabName);

            ////Create button
            PulldownButtonData groupVerticalData = new PulldownButtonData("PulldownGroup1", "Vertical\nConnection");
            PulldownButton groupVertical = plumbingPanel.AddItem(groupVerticalData) as PulldownButton;

            //Create button
            PushButtonData veticalWYEData = new PushButtonData("btnVeticalWYE", "Vetical WYED\nConnection 1", assemblyPath, Define.CmdVerticalWYEConnection1ClassName);
            //AddImages(veticalWYEData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            groupVertical.AddPushButton(veticalWYEData);

            //Create button
            PushButtonData verticalTeeConnectionData = new PushButtonData("btnVerticalTeeConnection", "Vertical Tee\nConnection", assemblyPath, Define.CmdVeticalTeeConnectionClassName);
            //AddImages(verticalTeeConnectionData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            groupVertical.AddPushButton(verticalTeeConnectionData);

            //Create button
            PushButtonData createEndCapData = new PushButtonData("btnCreateEndCap", "Create\nEnd Cap", assemblyPath, Define.CmdCreateEndCapClassName);
            //AddImages(createEndCapData, iconFolder, "Manual Hanger.png", "Manual Hanger.png");
            plumbingPanel.AddItem(createEndCapData);

            //Create button
            PushButtonData extendPipeData = new PushButtonData("btnExtendPipe", "Extend\nPipe", assemblyPath, Define.CmdExtendPipeClassName);
            //AddImages(extendPipeData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            plumbingPanel.AddItem(extendPipeData);

            //Create button
            PushButtonData slopePipeConnectionData = new PushButtonData("btnSlopePipeConnection", "Slope Pipe\nConnection", assemblyPath, Define.CmdSlopePipeConnectionClassName);
            //AddImages(slopePipeConnectionData, iconFolder, "Re-Inforced Hanger.png", "Re-Inforced Hanger.png");
            plumbingPanel.AddItem(slopePipeConnectionData);

            //Create button
            PushButtonData createCouplingData = new PushButtonData("btnCreateCoupling", "Create\nCoupling", assemblyPath, Define.CmdCreateCouplingClassName);
            //AddImages(createCouplingData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            plumbingPanel.AddItem(createCouplingData);

            //Create button
            PushButtonData createNippleData = new PushButtonData("btnCreateNipple", "Create\nNipple", assemblyPath, Define.CmdCreateNippleClassName);
            //AddImages(createNippleData, iconFolder, "Re-Inforced Hanger.png", "Re-Inforced Hanger.png");
            plumbingPanel.AddItem(createNippleData);
        }

        /// <summary>
        /// Create fire fighting tab
        /// </summary>
        /// <param name="app"></param>
        /// <param name="tabName"></param>
        /// <param name="assemblyPath"></param>
        /// <param name="iconFolder"></param>
        private void CreateFireFightingTab(UIControlledApplication app,
                                                  string tabName,
                                                  string assemblyPath,
                                                  string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel fireFightingPanel = app.CreateRibbonPanel(tabName, Define.FireFightingRibbonTabName);

            //Create button
            PushButtonData levelSmartData = new PushButtonData("btnLevelSmart", "2level\nSmart", assemblyPath, Define.Cmd2LevelSmartClassName);
            //AddImages(levelSmartData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            fireFightingPanel.AddItem(levelSmartData);

            ////Create button
            PulldownButtonData groupSprinklerData = new PulldownButtonData("Sprinkler", "Sprinkler\nConnection");
            PulldownButton groupSprinkler = fireFightingPanel.AddItem(groupSprinklerData) as PulldownButton;

            //Create button
            PushButtonData sprinklerUprightData = new PushButtonData("btnSprinklerUpright", "Sprinkler\nUpright", assemblyPath, Define.CmdSprinklerUprightClassName);
            //AddImages(sprinklerUprightData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            groupSprinkler.AddPushButton(sprinklerUprightData);

            //Create button
            PushButtonData sprinklerDownrightData = new PushButtonData("btnSprinklerDownright", "Sprinkler\nDownright", assemblyPath, Define.CmdSprinklerDownrightClassName);
            //AddImages(sprinklerDownrightData, iconFolder, "Manual Hanger.png", "Manual Hanger.png");
            groupSprinkler.AddPushButton(sprinklerDownrightData);

            //Create button
            PushButtonData flexSprinkerData = new PushButtonData("btnFlexSprinker", "Flex Sprinker", assemblyPath, Define.CmdFlexSprinklerClassName);
            //AddImages(flexSprinkerData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            groupSprinkler.AddPushButton(flexSprinkerData);
        }

        /// <summary>
        /// Create duct tab
        /// </summary>
        /// <param name="app"></param>
        /// <param name="tabName"></param>
        /// <param name="assemblyPath"></param>
        /// <param name="iconFolder"></param>
        private void CreateDuctTab(UIControlledApplication app,
                                                  string tabName,
                                                  string assemblyPath,
                                                  string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel ductPanel = app.CreateRibbonPanel(tabName, Define.DuctRibbonTabName);

            ////Create button
            PulldownButtonData groupSplitData = new PulldownButtonData("SplitDuct", "Split\nDuct");
            PulldownButton groupSplit = ductPanel.AddItem(groupSplitData) as PulldownButton;

            //Create button
            PushButtonData splitDuctData = new PushButtonData("btnSplitDuct", "Setting", assemblyPath, Define.CmdSplitDuctClassName);
            //AddImages(splitDuctData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            groupSplit.AddPushButton(splitDuctData);

            //Create button
            PushButtonData deleteUnionData = new PushButtonData("btnDeleteUnion", "Delete\nUnion", assemblyPath, Define.CmdDeleteUnionClassName);
            //AddImages(deleteUnionData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            groupSplit.AddPushButton(deleteUnionData);

            //Create button
            PushButtonData diffuserConnectionData = new PushButtonData("btnDiffuserConnection", "Diffuser\nConnection", assemblyPath, Define.CmdDiffuserConnectionClassName);
            //AddImages(diffuserConnectionData, iconFolder, "Manual Hanger.png", "Manual Hanger.png");
            ductPanel.AddItem(diffuserConnectionData);

            //Create button
            PushButtonData tapFlipData = new PushButtonData("btnTapFlip", "Tap\nFlip", assemblyPath, Define.CmdTapFlipClassName);
            //AddImages(tapFlipData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            ductPanel.AddItem(tapFlipData);

            //Create button
            PushButtonData grillOnDuctData = new PushButtonData("btnGrillOnDuct", "Grill On Duct", assemblyPath, Define.CmdGrillOnDuctClassName);
            //AddImages(grillOnDuctData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            ductPanel.AddItem(grillOnDuctData);
        }

        /// <summary>
        /// Create opening tab
        /// </summary>
        /// <param name="app"></param>
        /// <param name="tabName"></param>
        /// <param name="assemblyPath"></param>
        /// <param name="iconFolder"></param>
        private void CreateOpeningTab(UIControlledApplication app,
                                                 string tabName,
                                                 string assemblyPath,
                                                 string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel openingPanel = app.CreateRibbonPanel(tabName, Define.OpeningRibbonTabName);

            ////Create button
            PulldownButtonData groupOpeningData = new PulldownButtonData("PulldownGroup1", "Opening");
            PulldownButton groupOpening = openingPanel.AddItem(groupOpeningData) as PulldownButton;

            //Create button
            PushButtonData createOpeningData = new PushButtonData("btnCreateOpening", "Create\nOpening", assemblyPath, Define.CmdCreateOpeningClassName);
            //AddImages(createOpeningData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            groupOpening.AddPushButton(createOpeningData);

            PushButtonData deleteAllData = new PushButtonData("btnDeleteAll", "Delete All", assemblyPath, Define.CmdDeleteAllOpeningClassName);
            //AddImages(deleteAllData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            groupOpening.AddPushButton(deleteAllData);

            PushButtonData deleteBySelectionOpeningData = new PushButtonData("btnDeleteSelection", "Delete By Selection", assemblyPath, Define.CmdDeleteBySelectionOpeningClassName);
            //AddImages(deleteAllSleeveData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            groupOpening.AddPushButton(deleteBySelectionOpeningData);

            ////Create button
            PulldownButtonData groupSleeveData = new PulldownButtonData("PulldownGroup2", "Sleeve");
            PulldownButton groupSleeve = openingPanel.AddItem(groupSleeveData) as PulldownButton;

            //Create button
            PushButtonData createSleeveData = new PushButtonData("btnCreateSleeve", "Create\nSleeve", assemblyPath, Define.CmdCreateSleeveClassName);
            //AddImages(createSleeveData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            groupSleeve.AddPushButton(createSleeveData);

            PushButtonData deleteAllSleeveData = new PushButtonData("btnDeleteAll", "Delete All", assemblyPath, Define.CmdDeleteAllSleeveClassName);
            //AddImages(deleteAllSleeveData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            groupSleeve.AddPushButton(deleteAllSleeveData);

            PushButtonData deleteBySelectionData = new PushButtonData("btnDeleteSelection", "Delete By Selection", assemblyPath, Define.CmdDeleteBySelectionSleeveClassName);
            //AddImages(deleteAllSleeveData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            groupSleeve.AddPushButton(deleteBySelectionData);

            ////Create button
            PulldownButtonData groupChangeData = new PulldownButtonData("PulldownGroupChange", "Change \nObject");
            PulldownButton groupChange = openingPanel.AddItem(groupChangeData) as PulldownButton;

            //Create button
            PushButtonData changeOpeningData = new PushButtonData("btnChangeOpening", "Change Opening", assemblyPath, Define.CmdChangeOpeningClassName);
            //AddImages(changeObjectData, iconFolder, "Manual Hanger.png", "Manual Hanger.png");
            groupChange.AddPushButton(changeOpeningData);

            //Create button
            PushButtonData changeSleeveData = new PushButtonData("btnChangeSleeve", "Change Sleeve", assemblyPath, Define.CmdChangeSleeveClassName);
            //AddImages(changeObjectData, iconFolder, "Manual Hanger.png", "Manual Hanger.png");
            groupChange.AddPushButton(changeSleeveData);
        }

        /// <summary>
        /// Create modify tab
        /// </summary>
        /// <param name="app"></param>
        /// <param name="tabName"></param>
        /// <param name="assemblyPath"></param>
        /// <param name="iconFolder"></param>
        private void CreateModifyTab(UIControlledApplication app,
                                               string tabName,
                                               string assemblyPath,
                                               string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel modifyPanel = app.CreateRibbonPanel(tabName, Define.ModifyRibbonTabName);

            //Create button
            PushButtonData fittingRotationData = new PushButtonData("btnFittingRotation", "Fitting\nRotation", assemblyPath, Define.CmdFittingRotationClassName);
            //AddImages(fittingRotationData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            modifyPanel.AddItem(fittingRotationData);

            ////Create button
            PulldownButtonData groupData = new PulldownButtonData("PulldownGroup2", "Tee Tap\nTransfer");
            PulldownButton group = modifyPanel.AddItem(groupData) as PulldownButton;

            //Create button
            PushButtonData tapToTeeData = new PushButtonData("btnTapToTee", "Tap To Tee", assemblyPath, Define.CmdTapToTeeClassName);
            //AddImages(tapToTeeData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            group.AddPushButton(tapToTeeData);

            //Create button
            PushButtonData teeToTapData = new PushButtonData("btnTeeToTap", "Tee To Tap", assemblyPath, Define.CmdTeeToTapClassName);
            //AddImages(teeToTapData, iconFolder, "Manual Hanger.png", "Manual Hanger.png");
            group.AddPushButton(teeToTapData);

            //Create button
            PushButtonData blockCADToFamilyData = new PushButtonData("btnBlockCADToFamily", "Block CAD\nTo Family", assemblyPath, Define.CmdBlockCADToFamilyClassName);
            //AddImages(blockCADToFamilyData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            modifyPanel.AddItem(blockCADToFamilyData);

            //Create button
            PushButtonData smartSelectionData = new PushButtonData("btnSmartSelection", "Smart\nSelection", assemblyPath, Define.CmdSmartSelectionClassName);
            //AddImages(smartSelectionData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            modifyPanel.AddItem(smartSelectionData);

            //Create button
            PushButtonData flipSelectionData = new PushButtonData("btnFlipSelection", "Flip\nSelection", assemblyPath, Define.CmdFlipSelectionClassName);
            //AddImages(flipSelectionData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            modifyPanel.AddItem(flipSelectionData);

            //Create button
            PushButtonData moveFittingData = new PushButtonData("btnMoveFitting", "Move\nFitting", assemblyPath, Define.CmdMoveFittingClassName);
            //AddImages(moveFittingData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            modifyPanel.AddItem(moveFittingData);
        }

        /// <summary>
        /// Add images
        /// </summary>
        /// <param name="buttonData"></param>
        /// <param name="iconFolder"></param>
        /// <param name="largeImage"></param>
        /// <param name="smallImage"></param>
        private void AddImages(ButtonData buttonData,
                               string iconFolder,
                               string largeImage,
                               string smallImage)
        {
            if (!string.IsNullOrEmpty(iconFolder)
                && Directory.Exists(iconFolder))
            {
                string largeImagePath = Path.Combine(iconFolder, largeImage);
                if (File.Exists(largeImagePath))
                    buttonData.LargeImage = new BitmapImage(new Uri(largeImagePath));

                string smallImagePath = Path.Combine(iconFolder, smallImage);
                if (File.Exists(smallImagePath))
                    buttonData.Image = new BitmapImage(new Uri(smallImagePath));
            }
        }

        /// <summary>
        /// Get icon folder
        /// </summary>
        /// <returns></returns>
        private string GetIconFolder()
        {
            string appDir = GetAppFolder();
            string imageDir = Path.Combine(appDir, "Icon");
            return imageDir;
        }

        /// <summary>
        /// Get app folder
        /// </summary>
        /// <returns></returns>
        private string GetAppFolder()
        {
            string location = Assembly.GetExecutingAssembly().Location;
            string dir = Path.GetDirectoryName(location);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            return dir;
        }

        /// <summary>
        /// Get config folder
        /// </summary>
        /// <returns></returns>
        private string GetConfigFolder()
        {
            string appDir = GetAppFolder();
            string configDir = Path.Combine(appDir, "Config");
            if (!Directory.Exists(configDir))
                Directory.CreateDirectory(configDir);

            return configDir;
        }

        #endregion Create ribbon

        #region Show Dialog

        public static bool ShowVerticalMEPForm()
        {
            try
            {
                if (null == hWndRevit)
                {
                    Process process = Process.GetCurrentProcess();

                    IntPtr h = process.MainWindowHandle;
                    hWndRevit = new WindowHandle(h);
                }

                bool isShow = false;

                if (verticalMEPForm == null || verticalMEPForm.IsDisposed)
                {
                    // A new handler to handle request posting by the dialog
                    RequestHandler handler = new RequestHandler();

                    // External Event for the dialog to use (to post requests)
                    ExternalEvent exEvent = ExternalEvent.Create(handler);

                    // We give the objects to the new dialog;
                    // The dialog becomes the owner responsible fore disposing them, eventually.
                    verticalMEPForm = new VerticalMEPForm(exEvent, handler);

                    verticalMEPForm.Show(hWndRevit);
                }
                else
                {
                    isShow = true;
                }

                DisplayService.SetFocus(new HandleRef(null, App.verticalMEPForm.Handle));

                //NativeWin32.SetForegroundWindow(m_MyForm.Handle);

                return true;
            }
            catch (System.Exception ex)
            {
                return false;
            }
        }

        #endregion Show Dialog

        #endregion Method
    }
}