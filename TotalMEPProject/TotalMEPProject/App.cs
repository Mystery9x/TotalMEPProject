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
using TotalMEPProject.Ultis.HolyUltis;

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
        public static FastVerticalForm fastVerticalForm = null;
        public static HolyUpDownForm m_HolyUpDownForm = null;

        public static HolySplitUpdown _HolyUpdown = null;

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
            string errMess = string.Empty;
            bool isValidLicense = LicenseUtils.CheckLicense(isHasInternet, ref errMess);

            if (!isValidLicense)
            {
                DisableItemRibbonHanger(app);
                if (errMess != "")
                    IO.ShowError(errMess, "TotalMEP");
            }
        }

        public static void DisableItemRibbonHanger(UIControlledApplication app)
        {
            var ribbonPanels = app.GetRibbonPanels("TotalMEP");
            foreach (var item in ribbonPanels)
            {
                if (item.Name != Define.LoginLicense)
                    item.Enabled = false;
            }
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
            AddImages(loginData, iconFolder, "Setting.png", "Setting.png");
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
            AddImages(pickLineData, iconFolder, "Vertical MEP-01.png", "Vertical MEP-01.png");
            totalMEPPanel.AddItem(pickLineData);

            //Create button
            PushButtonData veticalMEPData = new PushButtonData("btnVeticalMEP", "Vetical\nMEP", assemblyPath, Define.CmdVerticalMEPClassName);
            AddImages(veticalMEPData, iconFolder, "Vertical MEP 2-01-01.png", "Vertical MEP 2-01-01.png");
            totalMEPPanel.AddItem(veticalMEPData);

            //Create button
            PushButtonData fastVerticalData = new PushButtonData("btnFastVertical", "Fast\nVertical", assemblyPath, Define.CmdFastVerticalClassName);
            AddImages(fastVerticalData, iconFolder, "Fast Vertical-02-01.png", "Fast Vertical-02-01.png");
            totalMEPPanel.AddItem(fastVerticalData);

            //Create button
            PushButtonData HolyUpdownData = new PushButtonData("btnHolyUpdown", "Holy\nUpdown", assemblyPath, Define.CmdHolyUpdownClassName);
            AddImages(HolyUpdownData, iconFolder, "Holy Updown-02.png", "Holy Updown-02.png");
            totalMEPPanel.AddItem(HolyUpdownData);

            //Create button
            PushButtonData MEPConnectionData = new PushButtonData("btnMEPConnection", "MEP\nConnection", assemblyPath, Define.CmdMEPConnectionClassName);
            AddImages(MEPConnectionData, iconFolder, "Vertical Connection-01.png", "Vertical Connection-01.png");
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
            AddImagesPullDown(groupVertical, iconFolder, "Vertical Connection-02-01.png", "Vertical Connection-02-01.png");

            //Create button
            PushButtonData veticalWYEData = new PushButtonData("btnVeticalWYE", "Vetical WYED", assemblyPath, Define.CmdVerticalWYEConnection1ClassName);
            AddImages(veticalWYEData, iconFolder, "Vertical WYED.png", "Vertical WYED.png");
            groupVertical.AddPushButton(veticalWYEData);

            //Create button
            PushButtonData verticalTeeConnectionData = new PushButtonData("btnVerticalTeeConnection", "Vertical Tee", assemblyPath, Define.CmdVeticalTeeConnectionClassName);
            AddImages(verticalTeeConnectionData, iconFolder, "Vertical Tee-01.png", "Vertical Tee-01.png");
            groupVertical.AddPushButton(verticalTeeConnectionData);

            //Create button
            PushButtonData createEndCapData = new PushButtonData("btnCreateEndCap", "Create\nEnd Cap", assemblyPath, Define.CmdCreateEndCapClassName);
            AddImages(createEndCapData, iconFolder, "Create End Cap-01.png", "Create End Cap-01.png");
            plumbingPanel.AddItem(createEndCapData);

            //Create button
            PushButtonData extendPipeData = new PushButtonData("btnExtendPipe", "Extend\nPipe", assemblyPath, Define.CmdExtendPipeClassName);
            AddImages(extendPipeData, iconFolder, "Extend Pipe-01.png", "Extend Pipe-01.png");
            plumbingPanel.AddItem(extendPipeData);

            //Create button
            PushButtonData slopePipeConnectionData = new PushButtonData("btnSlopePipeConnection", "Slope Pipe\nConnection", assemblyPath, Define.CmdSlopePipeConnectionClassName);
            AddImages(slopePipeConnectionData, iconFolder, "Slope Pipe Connection-01.png", "Slope Pipe Connection-01.png");
            plumbingPanel.AddItem(slopePipeConnectionData);

            //Create button
            PushButtonData createCouplingData = new PushButtonData("btnCreateCoupling", "Create\nCoupling", assemblyPath, Define.CmdCreateCouplingClassName);
            AddImages(createCouplingData, iconFolder, "Create Coupling-01.png", "Create Coupling-01.png");
            plumbingPanel.AddItem(createCouplingData);

            //Create button
            PushButtonData createNippleData = new PushButtonData("btnCreateNipple", "Create\nNipple", assemblyPath, Define.CmdCreateNippleClassName);
            AddImages(createNippleData, iconFolder, "Create Nipple-01.png", "Create Nipple-01.png");
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
            AddImages(levelSmartData, iconFolder, "2level Smart-01.png", "2level Smart-01.png");
            fireFightingPanel.AddItem(levelSmartData);

            ////Create button
            PulldownButtonData groupSprinklerData = new PulldownButtonData("Sprinkler", "Sprinkler\nConnection");
            PulldownButton groupSprinkler = fireFightingPanel.AddItem(groupSprinklerData) as PulldownButton;
            AddImagesPullDown(groupSprinkler, iconFolder, "Sprinkler Connection-01.png", "Sprinkler Connection-01.png");

            //Create button
            PushButtonData sprinklerUprightData = new PushButtonData("btnSprinklerUpright", "Sprinkler\nUpright", assemblyPath, Define.CmdSprinklerUprightClassName);
            AddImages(sprinklerUprightData, iconFolder, "Sprinkler Upright-01.png", "Sprinkler Upright-01.png");
            groupSprinkler.AddPushButton(sprinklerUprightData);

            //Create button
            PushButtonData sprinklerDownrightData = new PushButtonData("btnSprinklerDownright", "Sprinkler\nDownright", assemblyPath, Define.CmdSprinklerDownrightClassName);
            AddImages(sprinklerDownrightData, iconFolder, "Sprinkler Down Right.png", "Sprinkler Down Right.png");
            groupSprinkler.AddPushButton(sprinklerDownrightData);

            //Create button
            PushButtonData flexSprinkerData = new PushButtonData("btnFlexSprinker", "Flex Sprinker", assemblyPath, Define.CmdFlexSprinklerClassName);
            AddImages(flexSprinkerData, iconFolder, "Flex Sprinkler-01.png", "Flex Sprinkler-01.png");
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
            AddImagesPullDown(groupSplit, iconFolder, "Split Duct-01.png", "Split Duct-01.png");

            //Create button
            PushButtonData splitDuctData = new PushButtonData("btnSplitDuct", "Setting", assemblyPath, Define.CmdSplitDuctClassName);
            AddImages(splitDuctData, iconFolder, "Setting Split-01.png", "Setting Split-01.png");
            groupSplit.AddPushButton(splitDuctData);

            //Create button
            PushButtonData deleteUnionData = new PushButtonData("btnDeleteUnion", "Delete\nUnion", assemblyPath, Define.CmdDeleteUnionClassName);
            AddImages(deleteUnionData, iconFolder, "Delete Split-01.png", "Delete Split-01.png");
            groupSplit.AddPushButton(deleteUnionData);

            //Create button
            PushButtonData diffuserConnectionData = new PushButtonData("btnDiffuserConnection", "Diffuser\nConnection", assemblyPath, Define.CmdDiffuserConnectionClassName);
            AddImages(diffuserConnectionData, iconFolder, "Diffuser Connection-01.png", "Diffuser Connection-01.png");
            ductPanel.AddItem(diffuserConnectionData);

            //Create button
            PushButtonData tapFlipData = new PushButtonData("btnTapFlip", "Tap\nFlip", assemblyPath, Define.CmdTapFlipClassName);
            AddImages(tapFlipData, iconFolder, "Tap Flip-01.png", "Tap Flip-01.png");
            ductPanel.AddItem(tapFlipData);

            //Create button
            PushButtonData grillOnDuctData = new PushButtonData("btnGrillOnDuct", "Grill On Duct", assemblyPath, Define.CmdGrillOnDuctClassName);
            AddImages(grillOnDuctData, iconFolder, "Grill on Duct-01.png", "Grill on Duct-01.png");
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
            AddImagesPullDown(groupOpening, iconFolder, "Opening-01.png", "Opening-01.png");

            //Create button
            PushButtonData createOpeningData = new PushButtonData("btnCreateOpening", "Create\nOpening", assemblyPath, Define.CmdCreateOpeningClassName);
            AddImages(createOpeningData, iconFolder, "Creat Opening-01.png", "Creat Opening-01.png");
            groupOpening.AddPushButton(createOpeningData);

            PushButtonData deleteAllData = new PushButtonData("btnDeleteAll", "Delete All", assemblyPath, Define.CmdDeleteAllOpeningClassName);
            AddImages(deleteAllData, iconFolder, "Delete Opening-01.png", "Delete Opening-01.png");
            groupOpening.AddPushButton(deleteAllData);

            //PushButtonData deleteBySelectionOpeningData = new PushButtonData("btnDeleteSelection", "Delete By Selection", assemblyPath, Define.CmdDeleteBySelectionOpeningClassName);
            //AddImages(deleteBySelectionOpeningData, iconFolder, "Vertical Connection-01.png", "Vertical Connection-01.png");
            //groupOpening.AddPushButton(deleteBySelectionOpeningData);

            ////Create button
            PulldownButtonData groupSleeveData = new PulldownButtonData("PulldownGroup2", "Sleeve");
            PulldownButton groupSleeve = openingPanel.AddItem(groupSleeveData) as PulldownButton;
            AddImagesPullDown(groupSleeve, iconFolder, "Sleeve-01.png", "Sleeve-01.png");

            //Create button
            PushButtonData createSleeveData = new PushButtonData("btnCreateSleeve", "Create\nSleeve", assemblyPath, Define.CmdCreateSleeveClassName);
            AddImages(createSleeveData, iconFolder, "Create Sleeve-01.png", "Create Sleeve-01.png");
            groupSleeve.AddPushButton(createSleeveData);

            PushButtonData deleteAllSleeveData = new PushButtonData("btnDeleteAll", "Delete All", assemblyPath, Define.CmdDeleteAllSleeveClassName);
            AddImages(deleteAllSleeveData, iconFolder, "Delete Sleeve-01.png", "Delete Sleeve-01.png");
            groupSleeve.AddPushButton(deleteAllSleeveData);

            //PushButtonData deleteBySelectionData = new PushButtonData("btnDeleteSelection", "Delete By Selection", assemblyPath, Define.CmdDeleteBySelectionSleeveClassName);
            //AddImages(deleteBySelectionData, iconFolder, "Vertical Connection-01.png", "Vertical Connection-01.png");
            //groupSleeve.AddPushButton(deleteBySelectionData);

            ////Create button
            PulldownButtonData groupChangeData = new PulldownButtonData("PulldownGroupChange", "Change \nObject");
            PulldownButton groupChange = openingPanel.AddItem(groupChangeData) as PulldownButton;
            AddImagesPullDown(groupChange, iconFolder, "Change Object.png", "Change Object.png");
            //Create button
            PushButtonData changeOpeningData = new PushButtonData("btnChangeOpening", "Change Opening", assemblyPath, Define.CmdChangeOpeningClassName);
            AddImages(changeOpeningData, iconFolder, "Change Opening-01.png", "Change Opening-01.png");
            groupChange.AddPushButton(changeOpeningData);

            //Create button
            PushButtonData changeSleeveData = new PushButtonData("btnChangeSleeve", "Change Sleeve", assemblyPath, Define.CmdChangeSleeveClassName);
            AddImages(changeSleeveData, iconFolder, "Change Sleeve-01.png", "Change Sleeve-01.png");
            groupChange.AddPushButton(changeSleeveData);

            //Create button
            PushButtonData settingChangeData = new PushButtonData("btnSettingChange", "Setting Change", assemblyPath, Define.CmdChangeSleeveClassName);
            AddImages(settingChangeData, iconFolder, "Setting change open-01.png", "Setting change open-01.png");
            groupChange.AddPushButton(settingChangeData);
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
            AddImages(fittingRotationData, iconFolder, "Fitting Rotation-01.png", "Fitting Rotation-01.png");
            modifyPanel.AddItem(fittingRotationData);

            ////Create button
            PulldownButtonData groupData = new PulldownButtonData("PulldownGroup2", "Tee Tap\nTransfer");
            PulldownButton group = modifyPanel.AddItem(groupData) as PulldownButton;
            AddImagesPullDown(group, iconFolder, "Tee Tap Transfer-01.png", "Tee Tap Transfer-01.png");

            //Create button
            PushButtonData tapToTeeData = new PushButtonData("btnTapToTee", "Tap To Tee", assemblyPath, Define.CmdTapToTeeClassName);
            AddImages(tapToTeeData, iconFolder, "Tap to Tee-01.png", "Tap to Tee-01.png");
            group.AddPushButton(tapToTeeData);

            //Create button
            PushButtonData teeToTapData = new PushButtonData("btnTeeToTap", "Tee To Tap", assemblyPath, Define.CmdTeeToTapClassName);
            AddImages(teeToTapData, iconFolder, "Tee to Tap.png", "Tee to Tap.png");
            group.AddPushButton(teeToTapData);

            //Create button
            PushButtonData blockCADToFamilyData = new PushButtonData("btnBlockCADToFamily", "Block CAD\nTo Family", assemblyPath, Define.CmdBlockCADToFamilyClassName);
            AddImages(blockCADToFamilyData, iconFolder, "Block CAD To Family-01.png", "Block CAD To Family-01.png");
            modifyPanel.AddItem(blockCADToFamilyData);

            //Create button
            PushButtonData smartSelectionData = new PushButtonData("btnSmartSelection", "Smart\nSelection", assemblyPath, Define.CmdSmartSelectionClassName);
            AddImages(smartSelectionData, iconFolder, "Smart Selection-01.png", "Smart Selection-01.png");
            modifyPanel.AddItem(smartSelectionData);

            ////Create button
            //PushButtonData flipSelectionData = new PushButtonData("btnFlipSelection", "Flip\nSelection", assemblyPath, Define.CmdFlipSelectionClassName);
            //AddImages(flipSelectionData, iconFolder, "Vertical Connection-01.png", "Vertical Connection-01.png");
            //modifyPanel.AddItem(flipSelectionData);

            //Create button
            PushButtonData moveFittingData = new PushButtonData("btnMoveFitting", "Move\nFitting", assemblyPath, Define.CmdMoveFittingClassName);
            AddImages(moveFittingData, iconFolder, "Move Fititng-01.png", "Move Fititng-01.png");
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
        /// Add images
        /// </summary>
        /// <param name="buttonData"></param>
        /// <param name="iconFolder"></param>
        /// <param name="largeImage"></param>
        /// <param name="smallImage"></param>
        private void AddImagesPullDown(PulldownButton buttonData,
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
                    RequestHandler handler = new RequestHandler();

                    ExternalEvent exEvent = ExternalEvent.Create(handler);

                    verticalMEPForm = new VerticalMEPForm(exEvent, handler);

                    verticalMEPForm.Show(hWndRevit);
                }
                else
                {
                    isShow = true;
                }

                DisplayService.SetFocus(new HandleRef(null, App.verticalMEPForm.Handle));

                return true;
            }
            catch (System.Exception ex)
            {
                return false;
            }
        }

        public static bool ShowFastVerticalForm(System.Collections.Generic.Dictionary<string, ElementId> levelIds)
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

                if (fastVerticalForm == null || fastVerticalForm.IsDisposed)
                {
                    RequestHandler handler = new RequestHandler();

                    ExternalEvent exEvent = ExternalEvent.Create(handler);

                    fastVerticalForm = new FastVerticalForm(levelIds, exEvent, handler);

                    fastVerticalForm.Show(hWndRevit);
                }
                else
                {
                    isShow = true;
                }

                DisplayService.SetFocus(new HandleRef(null, App.fastVerticalForm.Handle));

                return true;
            }
            catch (System.Exception ex)
            {
                return false;
            }
        }

        public static bool ShowHolyUpDownForm()
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

                if (m_HolyUpDownForm == null || m_HolyUpDownForm.IsDisposed)
                {
                    RequestHandler handler = new RequestHandler();

                    ExternalEvent exEvent = ExternalEvent.Create(handler);

                    m_HolyUpDownForm = new HolyUpDownForm(exEvent, handler);

                    m_HolyUpDownForm.Show(hWndRevit);
                }
                else
                {
                    isShow = true;
                }

                DisplayService.SetFocus(new HandleRef(null, App.m_HolyUpDownForm.Handle));

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