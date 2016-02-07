// TWB Consulting 2016

using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media.Animation;
using System;
using System.Diagnostics;
using System.Windows.Media;
using System.Windows.Input;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Data;
using System.Collections;
//using System.Security;

namespace KickOff
{
    internal static class WindowExtensions
    {
        // from winuser.h
        private const int GWL_STYLE = -16,
                          WS_MAXIMIZEBOX = 0x10000,
                          WS_MINIMIZEBOX = 0x20000;

        [DllImport("user32.dll")]
        extern private static int GetWindowLong(IntPtr hwnd, int index);

        [DllImport("user32.dll")]
        extern private static int SetWindowLong(IntPtr hwnd, int index, int value);

        internal static void HideMinimizeAndMaximizeButtons(this Window window)
        {
            IntPtr hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
            var currentStyle = GetWindowLong(hwnd, GWL_STYLE);

            SetWindowLong(hwnd, GWL_STYLE, (currentStyle & ~WS_MAXIMIZEBOX & ~WS_MINIMIZEBOX));
        }
        internal static void ShowMinimizeAndMaximizeButtons(this Window window)
        {
            IntPtr hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
            var currentStyle = GetWindowLong(hwnd, GWL_STYLE);

            SetWindowLong(hwnd, GWL_STYLE, (currentStyle | WS_MAXIMIZEBOX | WS_MINIMIZEBOX));
        }
    }
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<Shortcut> Shortcuts= new ObservableCollection<Shortcut>();//null;
        //private WrapPanel MainPanel = new WrapPanel();
        //private DockPanel MainPanel = new DockPanel();
        //private Grid MainPanel;
        private UniformGrid MainPanel = new UniformGrid();
        private int IconCounter=0;
        private string ProgramState;

        public MainWindow()
        {
            // Windows logoff or shutdown
            Application.Current.SessionEnding += Current_SessionEnding;

            //Program ending
            Application.Current.Exit += Current_Exit;

            //ObservableCollection<IconBitMap> allico = ico2bmap.ExtractAllIconBitMapFromFile("shell32.dll");

            // allow data to be dragged into app
            AllowDrop = true;

            //WindowStyle = WindowStyle.ToolWindow;
            //WindowStyle = WindowStyle.SingleBorderWindow;//WindowStyle.None;// WindowStyle.SingleBorderWindow;
            //this.WindowStyle = WindowStyle.None; this.
            //AllowsTransparency = true;

            // Turn off the min and max buttons
            WindowStyle = WindowStyle.SingleBorderWindow;
            this.SourceInitialized += (x, y) =>
            {
                this.HideMinimizeAndMaximizeButtons();
            };

            ResizeMode = ResizeMode.CanResize;
            SizeToContent = SizeToContent.Manual; //.WidthAndHeight;
            ////Height = 400;            Width = 150;
            Background = Brushes.AntiqueWhite; // SystemColors.WindowBrush; // Brushes.AntiqueWhite;
            Foreground = SystemColors.WindowTextBrush; // Brushes.DarkBlue;

            //MainPanel = new WrapPanel();
            MainPanel.Margin = new Thickness(2);
            MainPanel.Width = Double.NaN; //auto
            MainPanel.Height = Double.NaN; //auto
            MainPanel.AllowDrop = true;
            MainPanel.Visibility = Visibility.Visible;

            Content = MainPanel; // create a panel to draw in

            // add handlers for mouse leaving the main window
            MouseEnter += MainWindow_MouseEnter;
            MouseLeave += MainWindow_MouseLeave;
            DragEnter += MainWindow_DragEnter;
            Drop += MainWindow_Drop;
            DragOver += MainWindow_DragOver;
            GiveFeedback += MainWindow_GiveFeedback;

            IconCounter = 0;

            //if (Shortcuts == null)                Shortcuts = new ObservableCollection<Shortcut>();            Shortcuts.Clear();

            //DataContext = Shortcuts;
            //ItemsControl itemCntrl = new ItemsControl();
            //itemCntrl.ItemsSource = "{Binding}";

            // add load behavior to change size and position to last saved
            Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // get the program state
            SetProgramState();
        }

        private void Current_SessionEnding(object sender, SessionEndingCancelEventArgs e)
        {
            SaveProgramState();

            //throw new NotImplementedException();
        }

        private void Current_Exit(object sender, ExitEventArgs e)
        {
            SaveProgramState();
        }

        public void SaveProgramState()
        {
            /// Get the current main window size and position, and the 
            /// paths to the shortcuts.
            ProgramState = string.Empty;
            ProgramState += Top.ToString(); //Application.Current.MainWindow.Top.ToString();
            ProgramState += "\n";
            ProgramState += Left.ToString();
            ProgramState += "\n";
            ProgramState += Height.ToString();
            ProgramState += "\n";
            ProgramState += Width.ToString();
            ProgramState += "\n";

            /// get the shortcuts from the panel so a new order
            /// is saved 
            /// 
            //foreach (Shortcut s in Shortcuts)
            IEnumerator scs = MainPanel.Children.GetEnumerator();
            scs.Reset();
            while (scs.MoveNext())
            {
                Shortcut s = (Shortcut) scs.Current;
                ProgramState += s.lnkData.ShortcutAddress + "\n";
            }

            lnkio.WriteProgramState(ProgramState);

            //throw new NotImplementedException();
        }

        public void SetProgramState()
        {
            ProgramState = string.Empty;
            ProgramState = lnkio.ReadProgramState();

            if (string.Empty != ProgramState)
            {
                // Break up on the newline 
                string[] s = ProgramState.Split('\n');

                // first is program name and second is timestamp

                Application.Current.MainWindow.Top = Double.Parse(s[2]);
                Application.Current.MainWindow.Left = Double.Parse(s[3]);

                Application.Current.MainWindow.Height = Double.Parse(s[4]);
                Application.Current.MainWindow.Width = Double.Parse(s[5]);

                // skip the first 3 elements of the array
                IEnumerable<string> items = s.Skip(6);
                s = items.ToArray<string>();

                CreateShortCut(s);

                PlaceShortcutsintoView(s);
            }
        }

        private void MainWindow_GiveFeedback(object sender, GiveFeedbackEventArgs e)
        {
            if (e.Effects == DragDropEffects.None)
            {
                Mouse.SetCursor(Cursors.None);
                e.UseDefaultCursors = false;
            }
            else
            {
                e.UseDefaultCursors = true;
            }

            //throw new NotImplementedException();
        }

        private void MainWindow_DragOver(object sender, DragEventArgs e)
        {
            //string[] dataFormats = eDropEvent.Data.GetFormats(true);

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;

            e.Handled = true;
        }

        private void MainWindow_DragEnter(object sender, DragEventArgs e)
        {
            //string[] dataFormats = e.Data.GetFormats(true);

            //    if (e.Data.GetDataPresent(DataFormats.FileDrop))
            //        e.Effects = DragDropEffects.Copy;
            //    else
            //        e.Effects = DragDropEffects.None;

            //    e.Handled = true;
            e.Handled = false;
        }

        private void MainWindow_MouseEnter(object sender, MouseEventArgs e)
        {
            e.Handled = false;
            //throw new NotImplementedException();
        }

        private void MainWindow_MouseLeave(object sender, MouseEventArgs e)
        {
            e.Handled = false;
            //throw new NotImplementedException();
        }

        private void MainWindow_Drop(object sender, DragEventArgs eDropEvent)
        {
            // init base code
            //base.OnDrop(eDropEvent);  // do not use this code does all needed

            string[] dataFormats = eDropEvent.Data.GetFormats(true);

            // If the DataObject contains string data, extract it.
            if (eDropEvent.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] FileList = (string[])eDropEvent.Data.GetData(DataFormats.FileDrop, false);

                /// Create the shortcuts
                /// 
                CreateShortCut(FileList);

                PlaceShortcutsintoView(FileList);

                eDropEvent.Handled = true;
            }
            else
                eDropEvent.Handled = false;
        }

        private void PlaceShortcutsintoView(string[] FileList)
        {
            // create the link(s) in main window
            foreach (Shortcut sc in Shortcuts)
            {
                if (!sc.IsRendered) // not already on display
                {
                    sc.Name = "_Icon" + IconCounter.ToString(); // used to match image to data

                    /// Set sorting order
                    /// 
                    sc.SortOrder = IconCounter;

                    /// Set the image bitmap
                    /// 
                    sc.Width = sc.lnkData.Bitmap.BitmapSize; //.Bitmap.bitmap.Width;
                    sc.Source = sc.lnkData.Bitmap.bitmapsource; //.Bitmap.bitmapsource;
                    sc.Visibility = Visibility.Visible;

                    IconCounter++;

                    // Create Animations that modofy the shortcut icons
                    DoubleAnimation mo_animation = new DoubleAnimation
                    {
                        From = 1.0,
                        To = 0.1,
                        Duration = new Duration(TimeSpan.FromSeconds(0.75)),
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
                        Name = sc.Name + "_esb",
                        Storyboard = mo_storyboard
                    };

                    // Set the name of the storyboard so it can be found
                    NameScope.GetNameScope(this).RegisterName(enterBeginStoryboard.Name, enterBeginStoryboard);

                    var moe = new EventTrigger(MouseEnterEvent);
                    moe.Actions.Add(enterBeginStoryboard);
                    sc.Triggers.Add(moe);
                    var mle = new EventTrigger(MouseLeaveEvent);
                    mle.Actions.Add(
                        new StopStoryboard
                        {
                            BeginStoryboardName = enterBeginStoryboard.Name
                        });
                    sc.Triggers.Add(mle);

                    // Add a popup to display the link name when the mouse is over
                    TextBlock popupText = new TextBlock();
                    // Description is Comment 
                    popupText.Text =
                    System.IO.Path.GetFileNameWithoutExtension(sc.lnkData.ShortcutAddress);
                    if (popupText.Text != sc.lnkData.Description)
                        popupText.Text += "\n" + sc.lnkData.Description;
                    popupText.Background = Brushes.AntiqueWhite;
                    popupText.Foreground = Brushes.Black;
                    sc.lnkPopup.Child = popupText;
                    sc.lnkPopup.PlacementTarget = sc;
                    sc.lnkPopup.IsOpen = false;
                    sc.lnkPopup.Placement = PlacementMode.MousePoint;//.Center;

                    // add handlers for popup
                    //sc.MouseEnter += SCI_MouseEnter; // turn on popup
                    //sc.MouseLeave += SCI_MouseLeave; // turn off popup
                    sc.MouseLeftButtonUp += SCI_MouseLeftButtonUp; // Add mouse event handler to run the shortcut

                    // right click menu
                    sc.MouseRightButtonUp += Sci_MouseRightButtonUp;
                    //sc.AllowDrop=true;
                    //sc.Drop += Sc_Drop;
                    //sc.MouseRightButtonDown += Sc_MouseRightButtonDown;
                    //sc.DragOver += Sc_DragOver;
                    // load and mark as loaded (first or it is not in the collections's copy)
                    sc.IsRendered = true;

                    MainPanel.Children.Add(sc);
                }
            }

            return;
        }

        private void Sc_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            Shortcut _sc = sender as Shortcut;
            DataObject dragData = new DataObject("Shortcut", _sc.lnkData.ShortcutAddress);
            DragDrop.DoDragDrop(_sc, dragData, DragDropEffects.Move);
        }


        private void Sc_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("Shortcut"))
            {
                e.Effects = DragDropEffects.Move;
                e.Handled = true;
            }
            else
            {
                e.Effects = DragDropEffects.None;
                e.Handled = false;
            }
        }
        private void Sc_Drop(object sender, DragEventArgs e)
        {
            string[] dataFormats = e.Data.GetFormats(true);

            // If the DataObject contains string data, extract it.
            if (e.Data.GetDataPresent("Shortcut"))
            {
                Shortcut sc = (Shortcut)sender;
                Shortcut t = (Shortcut)e.Source;
                if (t != sc)
                {
                    try
                    {
                        // remove the shortcut from the list, the panel and destroy itself
                        if (sc != null && MainPanel.Children.Contains(sc))
                        {
                            // Deleting an object requires anything that may 
                            // reference it like events to be doing this 
                            // need to be cleared

                            int target = MainPanel.Children.IndexOf(t);
                            sc.Triggers.Clear();
                            // must be removed first before inserting it with new index
                            MainPanel.Children.Remove(sc);
                            MainPanel.Children.Insert(target, sc);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    //throw new NotImplementedException();
                }
            }
        }
        private void Sci_MouseRightButtonUp(object sender, MouseButtonEventArgs mssgEvent)
        {
            Shortcut sc = (Shortcut)sender;
            try
            {
                // remove the shortcut from the list, the panel and destroy itself
                if (sc != null && MainPanel.Children.Contains(sc))
                {
                    // Deleting an object requires anything that may 
                    // reference it like events to be doing this 
                    // need to be cleared

                    sc.Triggers.Clear();
                    int target = MainPanel.Children.IndexOf(sc);
                    MainPanel.Children.Remove(sc);  
                    //MainPanel.Children.Insert(target - 2, sc);

                    //Shortcuts.Remove(sc);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        private void SCI_MouseEnter(object sender, MouseEventArgs e)
        {
            ((Shortcut)sender).lnkPopup.IsOpen = true;

            e.Handled = false;
        }
        private void SCI_MouseLeave(object sender, MouseEventArgs e)
        {
            ((Shortcut)sender).lnkPopup.IsOpen = false;
            e.Handled = false;
        }
        private void SCI_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs mouseEvent)
        {
            Shortcut scLink = (Shortcut)sender;
                try
                {
                    if (scLink != null)
                    {
                        try
                        {
                            Process p = null;

                            // This is this a reference .appref-ms file or an URL or a data file
                            if (scLink.lnkData.bTargetIsFile || scLink.lnkData.bIsReference)
                            {
                                //if( scLink.lnkData.bIsReference ) {
                                //string n = @"rundll32.exe dfshim.dll,ShOpenVerbApplication http://github-windows.s3.amazonaws.com/GitHub.application#GitHub.application,Culture=neutral,PublicKeyToken=317444273a93ac29,processorArchitecture=x86";
                                //rundll32.exe dfshim.dll,ShOpenVerbShortcut D:\Users\User\Desktop\GitHub
                                //Uri a = new Uri(n);
                                //string b = a.ToString(); }

                                p = Process.Start(scLink.lnkData.ShortcutAddress);  // just send the link to the OS
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
                                else
                                {
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
                                try
                                {
                                    p = Process.Start(start);
                                }
                                catch
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
            get { return Shortcuts; }
            set { Shortcuts = value; }
        }

        /// <summary>
        /// Create the KickOff link
        /// </summary>
        /// <param name="FileList"></param>
        private void CreateShortCut(string[] FileList)
        {
            foreach (string FilePath in FileList)
            {
                if (null == FilePath || string.Empty == FilePath)
                    continue;

                LnkData lnk = lnkio.ResolveShortcut(FilePath);
                if (null != lnk)
                {
                    if( null == lnk || lnk.ShortcutAddress == string.Empty)
                    {
                        throw new Exception("Shortcut link does not have a path");
                    }

                    Shortcut sc = new Shortcut
                    {
                        lnkData = lnk,
                        lnkPopup = new Popup(),
                        IsRendered = false
                    };

                    var items = Shortcuts.Where(x => x.lnkData.ShortcutAddress == FilePath);

                    if (items.Count() == 0) // the link imported is not in the existing list so add it
                    {
                        /// Add the shortcut to the list
                        Shortcuts.Add(sc);
                    }
                }
            }
        }

    } //----------- Main Class

    public class Shortcut : Image
    {
        public Shortcut()
        {
            Margin = new Thickness(4); // all same// (2, 2, 2, 2); //left,top,right,bottom
        }

        public LnkData lnkData { get; set; }
        public bool IsRendered { get; set; }
        public Popup lnkPopup { get; set; }
        public int SortOrder { get; set; }
    }

} //------------- Namespace


/*

*/
