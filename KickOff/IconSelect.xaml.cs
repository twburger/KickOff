using System;
using System.Collections;
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
    public partial class IconSelect : Window, IDisposable
    {
        public int idxFile = -1, idxIcon = -1;
        ObservableCollection<string> registeredNames = new ObservableCollection<string>();

        public void Dispose()
        {
            try
            {
                this.Triggers.Clear();
                //IEnumerator scs = MainIconPanel.Children.GetEnumerator();
                //scs.Reset();
                //while (scs.MoveNext())
                //{
                //    IconImage s = (IconImage)scs.Current;
                //    if( s.Triggers.Count > 0 )
                //        s.Triggers.Clear();
                //}

                foreach (string regname in registeredNames)
                    NameScope.GetNameScope(this).UnregisterName(regname);

                MainIconPanel.Children.Clear();
            }
            catch (Exception exc)
            {
                lnkio.WriteProgramLog("IconSelect.Dispose() error: " + exc.Message);
            }
            finally
            {
                GC.Collect();
            }
        }
        public IconSelect()
        {
            InitializeComponent();

            //Closed += IconSelect_Closed;
            WindowStyle = WindowStyle.ToolWindow; //.SingleBorderWindow;
            ResizeMode = ResizeMode.CanResize;
            SizeToContent = SizeToContent.Manual; //.WidthAndHeight;
            ////Height = 400;            Width = 150;
            Background = System.Windows.Media.Brushes.LemonChiffon; //.AntiqueWhite; // SystemColors.WindowBrush; // Brushes.AntiqueWhite;
            Foreground = System.Windows.SystemColors.WindowTextBrush; // Brushes.DarkBlue;

            //MainPanel = new WrapPanel();
            MainIconPanel.Margin = new Thickness(2);
            MainIconPanel.Width = Double.NaN; //auto
            MainIconPanel.Height = Double.NaN; //auto
            MainIconPanel.AllowDrop = true;
        }

        public void IconSelectSetIconSourcePaths(string[] iconSourceFilePaths)
        {
            try
            {
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
                    Name = this.Name + "_esb",//img.Name + "_esb",
                    Storyboard = mo_storyboard
                };

                // Set the name of the storyboard so it can be found
                NameScope.GetNameScope(this).RegisterName(enterBeginStoryboard.Name, enterBeginStoryboard);
                registeredNames.Add(enterBeginStoryboard.Name);

                // create event callbacks
                var moe = new EventTrigger(MouseEnterEvent);
                moe.Actions.Add(enterBeginStoryboard);
                var mle = new EventTrigger(MouseLeaveEvent);
                mle.Actions.Add(
                    new StopStoryboard
                    {
                        BeginStoryboardName = enterBeginStoryboard.Name
                    });


                for (int fileCount = 0; fileCount < iconSourceFilePaths.Length; fileCount++)
                {
                    if (string.Empty != iconSourceFilePaths[fileCount])
                    {
                        if (
                        (System.IO.File.Exists(iconSourceFilePaths[fileCount]) ||
                        System.IO.Directory.Exists(iconSourceFilePaths[fileCount]))
                        )
                        {
                            int imgCount = 0;
                            //if (null != ibm) ibm.Clear();
                            ObservableCollection<IconBitMap> iconBitmaps
                                = ico2bmap.ExtractAllIconBitMapFromFile(iconSourceFilePaths[fileCount]);
                            foreach (IconBitMap ico in iconBitmaps)
                            {
                                IconImage img = new IconImage()
                                {
                                    idxFile = fileCount,
                                    idxIcon = imgCount,
                                    Name = "iconselect_img_" + imgCount.ToString() + "_" + fileCount.ToString(),
                                    Source = ico.bitmapsource,
                                    Width = ico.BitmapSize,
                                    Margin = new Thickness(4)
                                };
                                img.MouseDown += Img_MouseDown;

                                /// storyboard
                                /// 

                                img.Triggers.Add(moe);
                                img.Triggers.Add(mle);
                                
                                img.MouseEnter += Img_MouseEnter;  //+= (s, e) => Mouse.OverrideCursor = Cursors.Hand;
                                img.MouseLeave += Img_MouseLeave;

                                // Add it and count it to make sure each has a unique name
                                MainIconPanel.Children.Add(img);

                                imgCount++;
                            }
                            iconBitmaps.Clear();
                        }
                    }
                    //else                    {                        lnkio.WriteProgramLog("Cannot find or read icon source file: " + (string.Empty == iconSourceFilePaths[fileCount] ? "File name is missing" : iconSourceFilePaths[fileCount]));                    }
                }
            }
            catch( Exception e )
            {
                lnkio.WriteProgramLog("IconSelect() error: " + e.Message);
                //throw new Exception("Get new icon error: " + e.Message);
            }
            //finally            {                if (ibm != null)                    ibm.Clear();            }

            MainIconPanel.Visibility = Visibility.Visible;

        // create a panel to draw in
        //Content = MainPanel;

        return;
        }

        /*
        private void IconSelect_Closed(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            try {
                foreach (string regname in registeredNames)
                    NameScope.GetNameScope(this).UnregisterName(regname);
                MainIconPanel.Children.Clear();
            }
            catch (Exception exc )
            {
                lnkio.WriteProgramLog("IconSelect_Closed() error: " + exc.Message);
            }
        }
        */
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
            int idx = MainIconPanel.Children.IndexOf(img);

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