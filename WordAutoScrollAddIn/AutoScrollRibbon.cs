using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Threading;

namespace WordAutoScrollAddIn
{
    public partial class AutoScrollRibbon
    {
        private Thread scrollThread;
        private bool isScrolling = false;
        private int scrollLines = 1; // 默认滚动行数
        private int delayTime = 1000; // 默认停留时间

        private void AutoScrollRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void StartScrollButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!isScrolling)
            {
                isScrolling = true;
                scrollThread = new Thread(new ThreadStart(ScrollDocument));
                scrollThread.Start();
            }
        }

        private void ScrollDocument()
        {
            while (isScrolling)
            {
                Globals.ThisAddIn.Application.ActiveWindow.SmallScroll(Down: scrollLines);
                Thread.Sleep(delayTime);
            }
        }

        private void StopScrollButton_Click(object sender, RibbonControlEventArgs e)
        {
            isScrolling = false;
            if (scrollThread != null && scrollThread.IsAlive)
            {
                scrollThread.Join();
            }
        }

        private void LinesEditBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (int.TryParse(LinesEditBox.Text, out int lines))
            {
                scrollLines = lines;
            }
            else
            {
                MessageBox.Show("请输入有效的行数！");
                LinesEditBox.Text = "1";
            }
        }

        private void DelayEditBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (int.TryParse(DelayEditBox.Text, out int delay))
            {
                delayTime = delay;
            }
            else
            {
                MessageBox.Show("请输入有效的时间！");
                DelayEditBox.Text = "1000";
            }
        }
    }
}
