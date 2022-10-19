using System.Windows.Forms;

namespace TotalMEPProject.Ultis
{
    public class IO
    {
        /// <summary>
        /// Show information to user
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        public static void ShowInfor(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Show question for user
        /// </summary>
        /// <param name="content"></param>
        /// <param name="title"></param>
        /// <returns></returns>
        public static DialogResult ShowQuestion(string content, string title = "Question")
        {
            return MessageBox.Show(content, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        /// <summary>
        /// Show warning to user
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        public static void ShowWarning(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Show error to user
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        public static void ShowError(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}