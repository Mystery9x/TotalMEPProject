using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;

namespace TotalMEPProject.UI
{
    /// <summary>
    /// Interaction logic for ProgessBar.xaml
    /// </summary>
    public partial class ProgessBar : Window
    {
        public double valueElementComplete = 0;
        public double valueConditionComplete = 0;

        public bool isCancel = false;

        public ProgessBar()
        {
            InitializeComponent();
        }

        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, Object value);

        private UpdateProgressBarDelegate updPbSelectElementDelegate = null;
        private UpdateProgressBarDelegate updPbConditionDelegate = null;

        public void IncrementElementProgressBar()
        {
            valueElementComplete++;
            Dispatcher.Invoke(updPbSelectElementDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, valueElementComplete });

            if (valueElementComplete <= progessBarSelectElement.Maximum)
                tbElementComplete.Text = "Element Complete " +
                                         valueElementComplete.ToString() +
                                         " / " +
                                         progessBarSelectElement.Maximum.ToString();
        }

        public void ResetElementProgressBar()
        {
            Dispatcher.Invoke(updPbSelectElementDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, valueElementComplete });

            if (valueElementComplete <= progessBarSelectElement.Maximum)
                tbElementComplete.Text = "Element Complete " +
                                         valueElementComplete.ToString() +
                                         " / " +
                                         progessBarSelectElement.Maximum.ToString();
        }

        public void IncrementConditionProgressBar()
        {
            valueConditionComplete++;
            Dispatcher.Invoke(updPbConditionDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, valueConditionComplete });

            if (valueConditionComplete <= progessBarCondition.Maximum)
                tbConditionComplete.Text = "Condition Complete " +
                                           valueConditionComplete.ToString() +
                                           " / " + progessBarCondition.Maximum.ToString();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            updPbSelectElementDelegate = new UpdateProgressBarDelegate(progessBarSelectElement.SetValue);
            updPbConditionDelegate = new UpdateProgressBarDelegate(progessBarCondition.SetValue);
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            isCancel = true;
            this.Close();
        }

        #region Hide Minimize/Maximize Button

        [DllImport("user32.dll")]
        internal static extern int SetWindowLong(IntPtr hwnd, int index, int value);

        [DllImport("user32.dll")]
        internal static extern int GetWindowLong(IntPtr hwnd, int index);

        internal static void HideMinimizeMaximizeAndCloseButtons(Window window)
        {
            const int GWL_STYLE = -16;
            const int WS_SYSMENU = 0x80000;

            IntPtr hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
            long value = GetWindowLong(hwnd, GWL_STYLE);

            SetWindowLong(hwnd, GWL_STYLE, (int)(value & ~WS_SYSMENU));
        }

        private void Window_SourceInitialized(object sender, EventArgs e)
        {
            HideMinimizeMaximizeAndCloseButtons(this);
        }

        #endregion Hide Minimize/Maximize Button
    }
}