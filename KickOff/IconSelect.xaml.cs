using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace KickOff
{
    /// <summary>
    /// Interaction logic for IconSelect.xaml
    /// </summary>
    public partial class IconSelect : Window
    {
        public string _iconSourceFilePath { get; set; }
        //private WrapPanel MainPanel;// = new WrapPanel();
        //ObservableCollection<IconBitMap> ibm = null;
        //Border bb = null;

        public IconSelect(string iconSourceFilePath)
        {
            InitializeComponent();

            _iconSourceFilePath = iconSourceFilePath;

            WindowStyle = WindowStyle.SingleBorderWindow;
            ResizeMode = ResizeMode.CanResize;
            SizeToContent = SizeToContent.Manual; //.WidthAndHeight;
            ////Height = 400;            Width = 150;
            Background = System.Windows.Media.Brushes.AntiqueWhite; // SystemColors.WindowBrush; // Brushes.AntiqueWhite;
            Foreground = System.Windows.SystemColors.WindowTextBrush; // Brushes.DarkBlue;

            //MainPanel = new WrapPanel();
            MainPanel.Margin = new Thickness(2);
            MainPanel.Width = Double.NaN; //auto
            MainPanel.Height = Double.NaN; //auto
            MainPanel.AllowDrop = true;
            ObservableCollection<IconBitMap> ibm=null;
            try
            {
                ibm = ico2bmap.ExtractAllIconBitMapFromFile(_iconSourceFilePath);
                int imgCount = 0;
                foreach (IconBitMap ico in ibm)
                {
                    Image img = new Image() { Name= "img" + imgCount.ToString(), Source = ico.bitmapsource, Width = ico.BitmapSize, Margin = new Thickness(4)};
                    img.MouseDown += Img_MouseDown;

                    DoubleAnimation mo_animation = new DoubleAnimation
                    {
                        From = 1.0,
                        To = 0.1,
                        Duration = new Duration(TimeSpan.FromSeconds(0.25)),
                        AutoReverse = true
                    };
                    var mo_storyboard = new Storyboard
                    {
                        RepeatBehavior = RepeatBehavior.Forever
                    };
                    Storyboard.SetTargetProperty(mo_animation, new PropertyPath("(Opacity)"));
                    mo_storyboard.Children.Add(mo_animation);
                    BeginStoryboard enterBeginStoryboard = new BeginStoryboard
                    {
                        Name = img.Name + "_esb",
                        Storyboard = mo_storyboard
                    };
                    // Set the name of the storyboard so it can be found
                    NameScope.GetNameScope(this).RegisterName(enterBeginStoryboard.Name, enterBeginStoryboard);
                    // create event callbacks
                    var moe = new EventTrigger(MouseEnterEvent);
                    moe.Actions.Add(enterBeginStoryboard);
                    img.Triggers.Add(moe);
                    var mle = new EventTrigger(MouseLeaveEvent);
                    mle.Actions.Add(
                        new StopStoryboard
                        {
                            BeginStoryboardName = enterBeginStoryboard.Name
                        });
                    img.Triggers.Add(mle);

                    img.MouseEnter += Img_MouseEnter;  //+= (s, e) => Mouse.OverrideCursor = Cursors.Hand;
                    img.MouseLeave += Img_MouseLeave;

                    MainPanel.Children.Add(img);

                    imgCount++;
                }
                
            } finally {
                if (ibm != null)
                    ibm.Clear();
            }

            MainPanel.Visibility = Visibility.Visible;

            // create a panel to draw in
            Content = MainPanel;
        }

        private void Img_MouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Arrow;
            //throw new NotImplementedException();
        }

        private void Img_MouseEnter(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;
            //Image i = sender as Image;

            //throw new NotImplementedException();
        }

        private void Img_MouseDown(object sender, MouseButtonEventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
