using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using TotalMEPProject.Ultis;

namespace TotalMEPProject
{
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication
    {
        public static UIApplication m_uiApp;
        public static UIControlledApplication _cachedUiCtrApp;
        private ExternalEvent exEvent;

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            // Create ribbon tools
            _cachedUiCtrApp = application;
            CreateRibbonButtons(application);

            return Result.Succeeded;
        }

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

        private void CreateTotalMEPTab(UIControlledApplication app,
                                                    string tabName,
                                                    string assemblyPath,
                                                    string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel totalMEPPanel = app.CreateRibbonPanel(tabName, Define.TotalMEPRibbonTabName);

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

        private void CreatePlumbingTab(UIControlledApplication app,
                                                   string tabName,
                                                   string assemblyPath,
                                                   string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel plumbingPanel = app.CreateRibbonPanel(tabName, Define.PlumpingRibbonTabName);

            //Create button
            PushButtonData veticalWYEData = new PushButtonData("btnVeticalWYE", "Vetical WYED\nConnection 1", assemblyPath, Define.CmdVerticalWYEConnection1ClassName);
            //AddImages(veticalWYEData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            plumbingPanel.AddItem(veticalWYEData);

            //Create button
            PushButtonData verticalTeeConnectionData = new PushButtonData("btnVerticalTeeConnection", "Vertical Tee\nConnection", assemblyPath, Define.CmdVeticalTeeConnectionClassName);
            //AddImages(verticalTeeConnectionData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            plumbingPanel.AddItem(verticalTeeConnectionData);

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

            //Create button
            PushButtonData sprinklerUprightData = new PushButtonData("btnSprinklerUpright", "Sprinkler\nUpright", assemblyPath, Define.CmdSprinklerUprightClassName);
            //AddImages(sprinklerUprightData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            fireFightingPanel.AddItem(sprinklerUprightData);

            //Create button
            PushButtonData sprinklerDownrightData = new PushButtonData("btnSprinklerDownright", "Sprinkler\nDownright", assemblyPath, Define.CmdSprinklerDownrightClassName);
            //AddImages(sprinklerDownrightData, iconFolder, "Manual Hanger.png", "Manual Hanger.png");
            fireFightingPanel.AddItem(sprinklerDownrightData);

            //Create button
            PushButtonData flexSprinkerData = new PushButtonData("btnFlexSprinker", "Flex Sprinker", assemblyPath, Define.CmdFlexSprinklerClassName);
            //AddImages(flexSprinkerData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            fireFightingPanel.AddItem(flexSprinkerData);
        }

        private void CreateDuctTab(UIControlledApplication app,
                                                  string tabName,
                                                  string assemblyPath,
                                                  string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel ductPanel = app.CreateRibbonPanel(tabName, Define.DuctRibbonTabName);

            //Create button
            PushButtonData splitDuctData = new PushButtonData("btnSplitDuct", "Split\nDuct", assemblyPath, Define.CmdSplitDuctClassName);
            //AddImages(splitDuctData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            ductPanel.AddItem(splitDuctData);

            //Create button
            PushButtonData deleteUnionData = new PushButtonData("btnDeleteUnion", "Delete\nUnion", assemblyPath, Define.CmdDeleteUnionClassName);
            //AddImages(deleteUnionData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            ductPanel.AddItem(deleteUnionData);

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

        private void CreateOpeningTab(UIControlledApplication app,
                                                 string tabName,
                                                 string assemblyPath,
                                                 string iconFolder)
        {
            //Create ribbon panel
            RibbonPanel openingPanel = app.CreateRibbonPanel(tabName, Define.OpeningRibbonTabName);

            //Create button
            PushButtonData createOpeningData = new PushButtonData("btnCreateOpening", "Create\nOpening", assemblyPath, Define.CmdCreateOpeningClassName);
            //AddImages(createOpeningData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            openingPanel.AddItem(createOpeningData);

            //Create button
            PushButtonData createSleeveData = new PushButtonData("btnCreateSleeve", "Create\nSleeve", assemblyPath, Define.CmdCreateSleeveClassName);
            //AddImages(createSleeveData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            openingPanel.AddItem(createSleeveData);

            //Create button
            PushButtonData changeObjectData = new PushButtonData("btnChangeObject", "Change\nObject", assemblyPath, Define.CmdChangeObjectClassName);
            //AddImages(changeObjectData, iconFolder, "Manual Hanger.png", "Manual Hanger.png");
            openingPanel.AddItem(changeObjectData);

            ////Create button
            PulldownButtonData group1Data = new PulldownButtonData("PulldownGroup1", "Delete\nOpening");
            PulldownButton group1 = openingPanel.AddItem(group1Data) as PulldownButton;

            PushButtonData deleteAllData = new PushButtonData("btnDeleteAll", "Delete All", assemblyPath, Define.CmdDeleteAllOpeningClassName);
            //AddImages(deleteAllData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            group1.AddPushButton(deleteAllData);

            PushButtonData deleteBySelectionData = new PushButtonData("btnDeleteBySelection", "Delete By Selection", assemblyPath, Define.CmdDeleteBySelectionOpeningClassName);
            //AddImages(deleteBySelectionData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            group1.AddPushButton(deleteBySelectionData);

            ////Create button
            PulldownButtonData group2Data = new PulldownButtonData("PulldownGroup2", "Delete\nSleeve");
            PulldownButton group2 = openingPanel.AddItem(group2Data) as PulldownButton;

            PushButtonData deleteAllSleeveData = new PushButtonData("btnDeleteAll", "Delete All", assemblyPath, Define.CmdDeleteAllSleeveClassName);
            //AddImages(deleteAllSleeveData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            group2.AddPushButton(deleteAllSleeveData);

            PushButtonData deleteBySelectionSleeveData = new PushButtonData("btnDeleteBySelection", "Delete By Selection", assemblyPath, Define.CmdDeleteBySelectionSleeveClassName);
            //AddImages(deleteBySelectionSleeveData, iconFolder, "Auto Hanger.png", "Auto Hanger.png");
            group2.AddPushButton(deleteBySelectionSleeveData);
        }

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

            //Create button
            PushButtonData tapToTeeData = new PushButtonData("btnTapToTee", "Tap To Tee", assemblyPath, Define.CmdTapToTeeClassName);
            //AddImages(tapToTeeData, iconFolder, "LoadFamily.png", "LoadFamily.png");
            modifyPanel.AddItem(tapToTeeData);

            //Create button
            PushButtonData teeToTapData = new PushButtonData("btnTeeToTap", "Tee To Tap", assemblyPath, Define.CmdTeeToTapClassName);
            //AddImages(teeToTapData, iconFolder, "Manual Hanger.png", "Manual Hanger.png");
            modifyPanel.AddItem(teeToTapData);

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
    }
}