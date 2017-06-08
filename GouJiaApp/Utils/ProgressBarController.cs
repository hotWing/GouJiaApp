using System.Windows.Forms;
namespace GouJiaApp.Utils
{
    public class ProgressBarController
    {

        //规避掉进度条的动画效果
        public static void setValue(ProgressBar pb, int value)
        {
            if (value == pb.Maximum)
            {
                pb.Value = value;           
                pb.Value = value - 1;      
            }
            else
            {
                pb.Value = value + 1;      
            }
            pb.Value = value;              
        }

    }
}
