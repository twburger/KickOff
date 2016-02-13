using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Animation;

namespace KickOff
{
    /// <summary>
    /// Interaction logic for IconSelect.xaml
    /// </summary>
    public partial class IconSelect : Window
    {
        public int idxFile = -1, idxIcon = -1;
        public IconSelect(string[] iconSourceFilePath)
        {
            InitializeComponent();

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
            ObservableCollection<IconBitMap> ibm = null;
            try
            {
                
                for (int fileCount = 0; fileCount < iconSourceFilePath.Length; fileCount++)
                {
                    if (string.Empty != iconSourceFilePath[fileCount]
                        //&& System.IO.File.Exists(iconSourceFilePath[fileCount]) &&
                        //!File.GetAttributes(iconSourceFilePath[fileCount]).HasFlag(FileAttributes.Directory)
                        )
                    {
                        int imgCount = 0;

                        ibm = ico2bmap.ExtractAllIconBitMapFromFile(iconSourceFilePath[fileCount]);
                        
                        foreach (IconBitMap ico in ibm)
                        {
                            IconImage img = new IconImage()
                            {
                                idxFile = fileCount,
                                idxIcon = imgCount,
                                Name = "_img_" + imgCount.ToString() + "_" + fileCount.ToString(),
                                Source = ico.bitmapsource,
                                Width = ico.BitmapSize,
                                Margin = new Thickness(4)
                            };
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

                            // Add it and count it to make sure each has a unique name
                            MainPanel.Children.Add(img); imgCount++;
                        }
                    }
                }
            }
            catch( Exception e )
            {
                lnkio.WriteProgramLog("Get new icon error: " + e.Message);
                //throw new Exception("Get new icon error: " + e.Message);
            }
            finally
            {
                if (ibm != null)
                    ibm.Clear();
            }

            MainPanel.Visibility = Visibility.Visible;

            // create a panel to draw in
            //Content = MainPanel;
        }

        private void Img_MouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Arrow;
            //throw new NotImplementedException();
        }

        private void Img_MouseEnter(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;
            //Image i = sender as IconImage;

            //throw new NotImplementedException();
        }

        private void Img_MouseDown(object sender, MouseButtonEventArgs e)
        {
            IconImage img = sender as IconImage;
            int idx = MainPanel.Children.IndexOf(img);

            this.idxFile = img.idxFile;
            this.idxIcon = img.idxIcon;

            DialogResult = true;
            Close();
        }
    }

    class IconImage : Image
    {
        public int idxFile { get; set; }
        public int idxIcon { get; set; }
    }
}