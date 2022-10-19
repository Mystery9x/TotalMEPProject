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