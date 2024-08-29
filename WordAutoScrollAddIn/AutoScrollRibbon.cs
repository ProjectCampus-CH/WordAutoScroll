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
        private int scrollLines = 1; // Ĭ�Ϲ�������
        private int delayTime = 1000; // Ĭ��ͣ��ʱ��

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
                MessageBox.Show("��������Ч��������");
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
                MessageBox.Show("��������Ч��ʱ�䣡");
                DelayEditBox.Text = "1000";
            }
        }
    }
}
