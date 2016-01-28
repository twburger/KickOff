// TWB Consulting 2016

using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media.Animation;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Media;

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
        /// Determine if what is being dragged in can be converted into Kickoff shortcuts
        /// </summary>
        /// <param name="eDragEvent"></param>
        protected override void OnDragEnter(DragEventArgs eDragEvent)
        {
            base.OnDragEnter(eDragEvent);

            // Set Effects to notify the drag source what effect
            // the drag-and-drop operation had.
            // (MOVE the LNK if CTRL or SHFT is pressed; otherwise, copy.)
            if (eDragEvent.KeyStates.HasFlag(DragDropKeyStates.ControlKey))
            {
                eDragEvent.Effects = DragDropEffects.Move;
            }
            else
            {
                eDragEvent.Effects = DragDropEffects.Copy;
            }

            eDragEvent.Handled = false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="eDropEvent"></param>
        protected override void OnDrop(DragEventArgs eDropEvent)
        {
            // init base code
            base.OnDrop(eDropEvent);

            // If the DataObject contains string data, extract it.
            if (eDropEvent.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] FileList = (string[])eDropEvent.Data.GetData(DataFormats.FileDrop, false);

                /// Create the shortcuts
                /// 
                CreateShortCut(FileList);

                // create the link(s) in main window
                foreach (Shortcut sci in shortcutItems)
                {
                    if (!sci.IsRendered) // not already on display
                    {
                        sci.Name = "_Icon" + IconCounter.ToString(); // used to match image to data

                        /// Set the image bitmap
                        /// 
                        sci.Width = sci.lnkData.Bitmap.BitmapSize; //.Bitmap.bitmap.Width;
                        sci.Source = sci.lnkData.Bitmap.bitmapsource; //.Bitmap.bitmapsource;
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
                            Name = sci.Name + "_esb",
                            Storyboard = mo_storyboard
                        };

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

                        // Add mouse event handler to run the shortcut
                        sci.MouseLeftButtonUp += IconImage_MouseLeftButtonUp;

                        // Add a popup to display the link name when the mouse is over
                        sci.lnkPopup.IsOpen = false;
                        TextBlock popupText = new TextBlock();
                        popupText.Text = sci.lnkData.Description;
                        popupText.Background = Brushes.AntiqueWhite;
                        popupText.Foreground = Brushes.Black;
                        sci.lnkPopup.Child = popupText;
                        //sci.lnkPopup.PlacementTarget = sci.lnkPopup;  // this crashes everything
                        sci.lnkPopup.Placement = PlacementMode.MousePoint; //.Center;

                        // add handlers for popup
                        sci.MouseEnter += Sci_MouseEnter;
                        sci.MouseLeave += Sci_MouseLeave;

                        // load and mark as loaded (first or it is not in the collections's copy)
                        sci.IsRendered = true;
                        //myMainPanel.Children.Add(sci.Bitmap.iconImage);
                        MainPanel.Children.Add(sci);
                    }
                }
                eDropEvent.Handled = true;

                return;
            }
            else
                eDropEvent.Handled = false;
        }

        private void Sci_MouseLeave(object sender, System.Windows.Input.MouseEventArgs mouseEvent)
        {
            Shortcut sc = (Shortcut)sender;
            sc.lnkPopup.IsOpen = false;

            mouseEvent.Handled = false;
        }

        private void Sci_MouseEnter(object sender, System.Windows.Input.MouseEventArgs mouseEvent)
        {
            Shortcut sc = (Shortcut)sender;
            sc.lnkPopup.IsOpen = true;

            mouseEvent.Handled = false;
        }

        private void IconImage_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs mouseEvent)
        {
            try
            {
                Shortcut scLink = (Shortcut)sender;

                if (scLink != null)
                {
                    try {
                        Process p = null;

                        // This is this a reference .appref-ms file or an URL
                        if (scLink.lnkData.bIsReference)
                        {
                            //string n = @"rundll32.exe dfshim.dll,ShOpenVerbApplication http://github-windows.s3.amazonaws.com/GitHub.application#GitHub.application,Culture=neutral,PublicKeyToken=317444273a93ac29,processorArchitecture=x86";
                            //rundll32.exe dfshim.dll,ShOpenVerbShortcut D:\Users\User\Desktop\GitHub
                            //Uri a = new Uri(n);
                            //string b = a.ToString();

                            p = Process.Start(scLink.lnkData.ShortcutAddress);
                        }
                        else
                        {
                            // Prepare the process to run
                            ProcessStartInfo start = new ProcessStartInfo();
                            // Enter in the command line arguments, everything you would enter after the executable name itself
                            if (scLink.lnkData.bTargetIsDirectory)
                            {

                                start.Arguments = scLink.lnkData.TargetPath;
                                start.FileName = "explorer.exe";
                                start.WorkingDirectory = scLink.lnkData.TargetPath;
                            }
                            else {
                                start.Arguments = scLink.lnkData.Arguments;
                                // Enter the executable to run, including the complete path
                                start.FileName = scLink.lnkData.TargetPath;
                                /// For cases where the starting working directory is like this:  %HOMEDRIVE%%HOMEPATH%
                                start.WorkingDirectory = Environment.ExpandEnvironmentVariables(scLink.lnkData.WorkingDirectory);
                                start.WorkingDirectory = System.IO.Path.GetFullPath(scLink.lnkData.WorkingDirectory);
                            }
                            // Do you want to show a console window?
                            start.WindowStyle = ProcessWindowStyle.Normal;
                            start.CreateNoWindow = true;
                            start.UseShellExecute = false; // do not open reference using shell or args can not be used

                            /// Run the programm
                            /// 
                            try {
                                p = Process.Start(start);
                            } catch
                            {
                                /// if the process fails it seems that the current working directory may be to blame
                                /// For example: The GitBash shortcut uses as command args: "C:\Program Files\Git\git-bash.exe" --cd-to-home
                                /// and the starting working directory of %HOMEDRIVE%%HOMEPATH%
                                start.WorkingDirectory = string.Empty;//@"D:\Users\user";
                                try { p = Process.Start(start); } catch { throw new Exception("Program run error: Program start failed"); }
                            }
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
                    if (scLink != null)
                    {
                        if (scLink.Name != null)
                            throw new Exception("Program run error: Internal data lookup error. The icon has no data or internal name: '" + scLink.Name + "' is corrupt");
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

            mouseEvent.Handled = true;

            return;
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

                LnkData lnk = lnkio.ResolveShortcut(FilePath);
                if (null != lnk)
                {
                    Shortcut sc = new Shortcut
                    {
                        lnkData = lnk,
                        lnkPopup = new Popup(),
                        IsRendered = false
                    };

                    var items = shortcutItems.Where(x => x.lnkData.ShortcutAddress == FilePath);

                    if (items.Count() == 0) // the link imported is not in the existing list so add it
                    {
                        /// Add a popup to display
                        sc.lnkPopup = new Popup();
                        TextBlock popupText = new TextBlock();
                        popupText.Text = sc.lnkData.Description;
                        popupText.Background = Brushes.AntiqueWhite;
                        popupText.Foreground = Brushes.Black;
                        sc.lnkPopup.Child = popupText;
                        sc.lnkPopup.PlacementTarget = sc;
                        sc.lnkPopup.IsOpen = false;

                        /// Add the shortcut to the list
                        shortcutItems.Add(sc);
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

        public LnkData lnkData { get; set; }
        public bool IsRendered { get; set; }
        public Popup lnkPopup { get; set; }
    }

    /*
        public bool IsReference { get; set; }

        public string ShortcutName { get; set; }

        public string FileName { get; set; }

        public string TargetLinkPath { get; set; }

        public string WorkingDirectory { get; set; }

        public string TargetParameters { get; set; }

        //public System.Windows.Media.Imaging.BitmapSource BitMapIcon { get; set; }

        public IconBitMap Bitmap { get; set; }

    }
*/

} //------------- Namespace
