using System.Windows;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media.Animation;
using System;
using System.Diagnostics;
using System.IO;

namespace KickOff
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<Shortcut> shortcutItems;

        //private DockPanel myMainPanel;
        private WrapPanel MainPanel;

        private int IconCounter;


        public MainWindow()
        {
            InitializeComponent();

            IconCounter = 0;

            if (shortcutItems == null)
                shortcutItems = new ObservableCollection<Shortcut>();
            shortcutItems.Clear();


            ResizeMode = ResizeMode.CanResize;
            SizeToContent = SizeToContent.Manual; //.WidthAndHeight;
            Height = 200;
            Width = 400;

            // myMainPanel = new DockPanel();
            //myMainPanel.LastChildFill = false; // last child stretches to fit remaining space

            MainPanel = new WrapPanel();
            MainPanel.Margin = new Thickness(5);
            MainPanel.Width = Double.NaN; //auto
            MainPanel.Height = Double.NaN; //auto

            this.Content = MainPanel; // create a panel to draw in
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="mouseEvent"></param>
        protected override void OnDrop(DragEventArgs mouseEvent)
        {
            // init base code
            base.OnDrop(mouseEvent);

            // If the DataObject contains string data, extract it.
            //if (mouseEvent.Data.GetDataPresent(DataFormats.FileDrop))
            if (mouseEvent.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] FileList = (string[])mouseEvent.Data.GetData(DataFormats.FileDrop, false);

                /// Create the shortcuts
                /// 
                CreateShortCut(FileList);

                // Set Effects to notify the drag source what effect
                // the drag-and-drop operation had.
                // (MOVE the LNK if CTRL or SHFT is pressed; otherwise, copy.)
                if (mouseEvent.KeyStates.HasFlag(DragDropKeyStates.ControlKey) || mouseEvent.KeyStates.HasFlag(DragDropKeyStates.ShiftKey))
                {
                    mouseEvent.Effects = DragDropEffects.Move;

                    // Delete the LNK file

                }
                else
                {
                    mouseEvent.Effects = DragDropEffects.Copy;
                }

                // create the link(s) in main window
                foreach (Shortcut sci in shortcutItems)
                {
                    if ( ! sci.IsRendered ) // not already on display
                    {
                        sci.Name = "_Icon" + IconCounter.ToString(); // used to match image to data

                        sci.Width = sci.Bitmap.bitmap.Width;
                        sci.Source = sci.Bitmap.bitmapsource;
                        sci.Visibility = Visibility.Visible;

                        IconCounter++;

                        DoubleAnimation mo_animation = new DoubleAnimation
                        {
                            From = 1.0,
                            To = 0.1,
                            Duration = new Duration(TimeSpan.FromSeconds(0.5)),
                            AutoReverse = true
                            //RepeatBehavior = RepeatBehavior.Forever
                        };

                        var mo_storyboard = new Storyboard
                        {
                            RepeatBehavior = RepeatBehavior.Forever
                        };

                        Storyboard.SetTargetProperty(mo_animation, new PropertyPath("(Opacity)"));

                        mo_storyboard.Children.Add(mo_animation);

                        BeginStoryboard enterBeginStoryboard = new BeginStoryboard
                        {
                            //Name = sci.Bitmap.iconImage.Name + "_esb",
                            Name = sci.Name + "_esb",
                            Storyboard = mo_storyboard
                        };

                        //NameScope.GetNameScope(this).RegisterName(sci.Bitmap.iconImage.Name + "_esb", enterBeginStoryboard);
                        NameScope.GetNameScope(this).RegisterName(enterBeginStoryboard.Name, enterBeginStoryboard);

                        var moe = new EventTrigger(MouseEnterEvent);
                        moe.Actions.Add(enterBeginStoryboard);
                        //sci.Bitmap.iconImage.Triggers.Add(moe);
                        sci.Triggers.Add(moe);

                        var mle = new EventTrigger(MouseLeaveEvent);
                        mle.Actions.Add(
                            new StopStoryboard
                            {
                                BeginStoryboardName = enterBeginStoryboard.Name
                            });
                        //sci.Bitmap.iconImage.Triggers.Add(mle);
                        sci.Triggers.Add(mle);

                        // Add mouse event handler to image
                        //sci.Bitmap.iconImage.MouseLeftButtonUp += IconImage_MouseLeftButtonUp;
                        sci.MouseLeftButtonUp += IconImage_MouseLeftButtonUp;

                        // load and mark as loaded (first or it is not in the collections's copy)
                        sci.IsRendered = true;
                        //myMainPanel.Children.Add(sci.Bitmap.iconImage);
                        MainPanel.Children.Add(sci);
                    }
                }
            }

            mouseEvent.Handled = true;
        }

        private void IconImage_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                //Image img = (Image)sender;
                Shortcut lnkData = (Shortcut)sender;

                // Look for the data
                //Shortcut lnkData = (Shortcut)shortcutItems.Where(X => X.InternalName == img.Name).FirstOrDefault();

                if (lnkData != null)
                {
                    try {
                        Process p = null;

                        // This is this a reference .appref-ms file or an URL
                        if (lnkData.FileName.ToLower().Contains(".appref-ms"))
                        {
                            //string n = @"rundll32.exe dfshim.dll,ShOpenVerbApplication http://github-windows.s3.amazonaws.com/GitHub.application#GitHub.application,Culture=neutral,PublicKeyToken=317444273a93ac29,processorArchitecture=x86";
                            //rundll32.exe dfshim.dll,ShOpenVerbShortcut D:\Users\User\Desktop\GitHub
                            //Uri a = new Uri(n);
                            //string b = a.ToString();
                            //p = Process.Start(lnkData.TargetLinkPath);
                            p = Process.Start(lnkData.FileName);
                        }
                        else
                        {
                            // Prepare the process to run
                            ProcessStartInfo start = new ProcessStartInfo();
                            // Enter in the command line arguments, everything you would enter after the executable name itself
                            if (File.GetAttributes(lnkData.TargetLinkPath).HasFlag(FileAttributes.Directory)) // a link to a directory
                            {
                                start.Arguments = lnkData.TargetLinkPath;
                                start.FileName = "explorer.exe";
                                start.WorkingDirectory = lnkData.TargetLinkPath;
                            }
                            else {
                                start.Arguments = lnkData.TargetParameters;
                                // Enter the executable to run, including the complete path
                                start.FileName = lnkData.TargetLinkPath;
                                start.WorkingDirectory = lnkData.WorkingDirectory;
                            }
                            // Do you want to show a console window?
                            start.WindowStyle = ProcessWindowStyle.Normal;
                            start.CreateNoWindow = true;
                            start.UseShellExecute = false; // do not open reference using shell or args can not be used
                            p = Process.Start(start);
                        }

                        if (p == null) // could be that it just does not start 2+ instances
                        {
                            // throw new Exception("Program run error: Program did not run");
                        }
                    }
                    catch { throw new Exception("Program run error: Program start failed"); }

                    //int exitCode;
                    //// Run the external process & wait for it to finish
                    //using (Process proc = Process.Start(start))
                    //{
                    //    proc.WaitForExit();

                    //    // Retrieve the app's exit code
                    //    exitCode = proc.ExitCode;
                    //}
                }
                else
                {
                    // bad
                    if (lnkData != null)
                    {
                        if (lnkData.Name != null)
                            throw new Exception("Program run error: Internal data lookup error. The icon has no data or internal name: '" + lnkData.Name + "' is corrupt");
                        else
                            throw new Exception("Program run error: Internal data lookup error. The icon internal name is missing");
                    }
                    else
                        throw new Exception("Program run error: Internal shortcut data lookup error. The data is not found");
                }
            }
            catch
            {
                throw new Exception("Program run error: Unknown internal error. Program data lookup failed.");
            }
            return;
        }

        private object CreateObject(string v)
        {
            throw new NotImplementedException();
        }

        public ObservableCollection<Shortcut> ShortcutItems
        {
            get { return shortcutItems; }
            set { shortcutItems = value; }
        }

        /// <summary>
        /// Create the KickOff link
        /// </summary>
        /// <param name="FileList"></param>
        private void CreateShortCut(string[] FileList)
        {
            foreach (string FilePath in FileList)
            {
                // get the link target
                string WorkDir = string.Empty;
                string TargetParams = string.Empty;
                string scName = string.Empty;
                bool bIsRef = false;
                bool bIsDir = false;

                System.IO.FileAttributes attr = File.GetAttributes(FilePath);
                if ( attr.HasFlag(FileAttributes.Directory) )
                    bIsDir = true;

                if (!File.Exists(FilePath) && ! bIsDir)
                {
                    throw new Exception(String.Format("File: {0} does not exist",FilePath));
                }
                else
                {
                    string target = lnkio.GetShortcutTarget(FilePath, bIsDir, ref WorkDir, ref TargetParams, ref scName, ref bIsRef);
                    if (target.Length > 0)
                    {
                        var items = shortcutItems.Where(x => x.FileName == FilePath);

                        if (items.Count() == 0) // the link imported is not in the existing list so add it
                        {
                            // Is the link target a directory?
                            
                            string IconSourceFile;

                            // if it's a reference ot the link is to a directory use the lnk file's icon as the image
                            if (bIsRef || bIsDir)
                                IconSourceFile = FilePath; // if a reference or a directory use the link itself as the icon source
                            else if(File.GetAttributes(target).HasFlag(FileAttributes.Directory)) // a link to a directory 
                                IconSourceFile = FilePath;
                            else
                                IconSourceFile = target; // 

                            shortcutItems.Add(new Shortcut()
                            {
                                ShortcutName = scName,
                                FileName = FilePath,
                                TargetLinkPath = target,
                                WorkingDirectory = WorkDir,
                                TargetParameters = TargetParams,
                                IsRendered = false,
                                Bitmap = ico2bmap.GetBitmapFromFileIcon(IconSourceFile)
                            });
                        }
                    }
                }
            }
        }

    } //----------- Main Class

    public class Shortcut : Image
    {
        public Shortcut()
        {

        }

        public bool IsReference { get; set; }

        public string ShortcutName { get; set; }

        public string FileName { get; set; }

        public string TargetLinkPath { get; set; }

        public string WorkingDirectory { get; set; }

        public string TargetParameters { get; set; }

        //public System.Windows.Media.Imaging.BitmapSource BitMapIcon { get; set; }

        public IconBitMap Bitmap { get; set; }

        public bool IsRendered { get; set; }
    }


} //------------- Namespace
