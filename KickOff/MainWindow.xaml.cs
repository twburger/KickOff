// TWB Consulting 2016

using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media.Animation;
using System.Diagnostics;
using System.Windows.Input;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Collections;
using System.Windows.Media;
using System.Windows.Interop;

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

        [DllImport("user32.dll")]
        private static extern bool EnableMenuItem(IntPtr hMenu,  uint uIDEnableItem, uint uEnable);

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

        // --------------------------------------------------------------------------------------

        [DllImport("user32.dll")]
        extern private static IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        [DllImport("user32.dll")]
        extern private static bool InsertMenu(IntPtr hMenu, Int32 wPosition, Int32 wFlags, Int32 wIDNewItem, string lpNewItem);

        /// Define our Constants we will use
        //private const Int32 WM_SYSCOMMAND = 0x112;
        private const Int32 MF_SEPARATOR = 0x800;
        private const Int32 MF_BYPOSITION = 0x400;
        private const Int32 MF_STRING = 0x0;
        public const Int32 SettingsSysMenuID = 9000;
        public const Int32 AboutSysMenuID = 9001;

        internal static void ModifyMenu(this Window window)
        {
            /// Get the Handle for the Forms System Menu
            IntPtr systemMenuHandle = GetSystemMenu(new System.Windows.Interop.WindowInteropHelper(window).Handle, false);

            /// Create our new System Menu items just before the Close menu item
            InsertMenu(systemMenuHandle, 5, MF_BYPOSITION | MF_SEPARATOR, 0, string.Empty); // <-- Add a menu seperator
            InsertMenu(systemMenuHandle, 6, MF_BYPOSITION, SettingsSysMenuID, "Settings..");
            InsertMenu(systemMenuHandle, 7, MF_BYPOSITION, AboutSysMenuID, "About..");

            // Attach our WndProc handler to this Window
            //HwndSource source = HwndSource.FromHwnd(systemMenuHandle);
            //source.AddHook(new HwndSourceHook(WndProc));
        }
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<Shortcut> Shortcuts = new ObservableCollection<Shortcut>();//null;
        //private WrapPanel MainPanel = new WrapPanel();
        //private DockPanel MainPanel = new DockPanel();
        //private Grid MainPanel;
        private UniformGrid MainPanel = new UniformGrid();
        private int ShortcutCounter = 0;
        private string ProgramState;
        private ContextMenu shortcutCtxMenu = new ContextMenu();

        public MainWindow()
        {
            InitializeComponent();

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

            // Turn off the min and max buttons and add my menu items to the main window context menu
            WindowStyle = WindowStyle.SingleBorderWindow;
            this.SourceInitialized += (x, y) =>
            {
                this.HideMinimizeAndMaximizeButtons();
                this.ModifyMenu();
            };

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
            MainPanel.Visibility = Visibility.Visible;

            // add handlers for mouse leaving the main window
            MouseEnter += MainWindow_MouseEnter;
            MouseLeave += MainWindow_MouseLeave;
            DragEnter += MainWindow_DragEnter;
            DragOver += MainWindow_DragOver;
            GiveFeedback += MainWindow_GiveFeedback;
            Drop += MainWindow_Drop;

            ShortcutCounter = 0;

            /// MENUS
            // Shortcut Right Click context menu
            MenuItem miSC_Delete = new MenuItem();
            miSC_Delete.Width = 120;
            miSC_Delete.Header = "_Delete";
            miSC_Delete.Click += MiSC_Delete_Click;
            shortcutCtxMenu.Items.Add(miSC_Delete);

            // create a panel to draw in
            Content = MainPanel;

            // add load behavior to change size and position to last saved
            Loaded += MainWindow_Loaded;

            //DataContext = this;
        }
        private const Int32 WM_SYSTEMMENU = 0xa4; // 164
        private const Int32 WP_SYSTEMMENU = 0x02; // 2
        private const Int32 WM_SYSCOMMAND = 0x112; // 274 <--------THIS IS the menu Select
        private const Int32 WM_NCRBUTTONDOWN = 0xA4; //164
        private const Int32 WM_NCLBUTTONDOWN = 0xA1; //161
        private const Int32 WM_CONTEXTMENU = 0x7B; //123
        private const Int32 WM_ENTERIDLE = 0x121; //289
        private const Int32 WM_INITMENUPOPUP = 0x0117; // 279
        private const Int32 WM_MENUSELECT = 0x011f; // 287
        private const Int32 MF_MOUSESELECT = 0x00008000; //32768
        private const Int32 MF_SYSMENU = 0x00002000; //8192

        private void debugout(int msg, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                int hiP = High16(lParam); // Flags
                int loP = Low16(lParam); //menu item
                int hiW = High16(wParam); // Flags
                int loW = Low16(wParam); //menu item

                Debug.WriteLine("MSG: {0} wParam low: {1} wParam high: {2} lParam low: {3} lParam high: {4}",
                    msg, loW, hiW, loP, loP);
            }
            catch
            {
                Debug.WriteLine("MSG: {0} wParam: {1} lParam: {2} ", msg, wParam, lParam);
            }

        }
        int GetIntUnchecked(IntPtr value)
        {
            return IntPtr.Size == 8 ? unchecked((int)value.ToInt64()) : value.ToInt32();
        }
        int Low16(IntPtr value)
        {
            return unchecked((short)GetIntUnchecked(value));
        }
        int High16(IntPtr value)
        {
            return unchecked((short)(((uint)GetIntUnchecked(value)) >> 16));
        }
        private IntPtr ProcessCustomMainContext(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (!handled)
            {
                int hiP = High16(wParam); // Flags
                int loP = Low16(wParam); //menu item

                switch ((uint)msg)
                {
                    case WM_SYSCOMMAND:
                        debugout(msg, wParam, lParam);
                        switch (loP)
                        {
                            case WindowExtensions.SettingsSysMenuID:
                                MessageBox.Show("Settingstext", "Caption Settings", MessageBoxButton.OK, MessageBoxImage.Information);
                                break;
                            case WindowExtensions.AboutSysMenuID:
                                MessageBox.Show("About text", "Caption About", MessageBoxButton.OK, MessageBoxImage.Information);
                                break;
                        }
                        break;

                    case WM_MENUSELECT:

                        debugout(msg, wParam, lParam);
                        switch (hiP)
                        {
                            case MF_SYSMENU:
                                debugout(msg, wParam, lParam);
                                break;

                            case MF_MOUSESELECT:
                                debugout(msg, wParam, lParam);
                                break;
                        }
                        break;
                }
            }
            return IntPtr.Zero;
        }

        private void MiSC_Delete_Click(object sender, RoutedEventArgs e)
        {
            MenuItem mi = (MenuItem)sender;
            ContextMenu cm = (ContextMenu)mi.Parent;
            Shortcut sc = cm.PlacementTarget as Shortcut;
            DeleteShortcut(sc);
        }
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // get the program state
            SetProgramState();

            //Add a handler to process custom main context items added
            IntPtr windowhandle = new WindowInteropHelper(this).Handle;
            HwndSource hwndSource = HwndSource.FromHwnd(windowhandle);
            hwndSource.AddHook(new HwndSourceHook(ProcessCustomMainContext));
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
                Shortcut s = (Shortcut)scs.Current;
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
                //Mouse.SetCursor(Cursors.None);
                //e.UseDefaultCursors = false;
                e.UseDefaultCursors = true;
            }
            else
            {
                //Mouse.SetCursor(Cursors.Hand);
                //e.UseDefaultCursors = false;
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

                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    e.Effects = DragDropEffects.Copy;
                else
                    e.Effects = DragDropEffects.None;
                e.Handled = true;
            //e.Handled = false;
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

            //string[] dataFormats = eDropEvent.Data.GetFormats(true);

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
                    sc.Name = "_Icon" + ShortcutCounter.ToString(); // used to match image to data

                    /// Set sorting order
                    /// 
                    sc.SortOrder = ShortcutCounter;

                    /// Set the image bitmap
                    /// 
                    sc.Width = sc.lnkData.icoBitmap.BitmapSize; //.Bitmap.bitmap.Width;
                    sc.Source = sc.lnkData.icoBitmap.bitmapsource; //.Bitmap.bitmapsource;
                    sc.Visibility = Visibility.Visible;

                    ShortcutCounter++;

                    // Create Animations that modify the shortcut icons
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
                    //if (popupText.Text != sc.lnkData.Description)                        popupText.Text += "\n" + sc.lnkData.Description;
                    popupText.Background = System.Windows.Media.Brushes.AntiqueWhite;
                    popupText.Foreground = System.Windows.Media.Brushes.Black;
                    sc.lnkPopup.Child = popupText;
                    sc.lnkPopup.PlacementTarget = sc;
                    sc.lnkPopup.IsOpen = false;
                    sc.lnkPopup.Placement = PlacementMode.MousePoint;//.Center;

                    // add handlers for popup
                    sc.MouseEnter += SCI_MouseEnter; // turn on popup
                    sc.MouseLeave += SCI_MouseLeave; // turn off popup

                    // mouse event handler to run the shortcut
                    sc.MouseLeftButtonUp += SC_MouseLeftButtonUp;

                    // move shortcuts using drag and drop
                    sc.AllowDrop = true;
                    sc.MouseMove += SC_mouseMove;
                    sc.PreviewMouseLeftButtonDown += SC_PreviewMouseLeftButtonDown;
                    sc.DragOver += SC_dragOver;
                    sc.GiveFeedback += SC_GiveFeedback;
                    sc.Drop += SC_drop;

                    // add the context menu
                    sc.ContextMenu = shortcutCtxMenu;

                    // load and mark as loaded (first or it is not in the collections's copy)
                    sc.IsRendered = true;
                    MainPanel.Children.Add(sc);
                }
            }

            return;
        }

        private void SC_GiveFeedback(object sender, GiveFeedbackEventArgs e)
        {
            if (e.Effects == DragDropEffects.None)
            {
                //Mouse.SetCursor(Cursors.None);
                //e.UseDefaultCursors = false;
                e.UseDefaultCursors = true;
            }
            else
            {
                 e.UseDefaultCursors = true;
            }
        }

        private void SC_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            RunShortcut((Shortcut)sender);
        }

        private void SC_drop(object sender, DragEventArgs e)
        {
            //string[] dataFormats = e.Data.GetFormats(true);

            // If the DataObject contains string data, extract it.
            if (e.Data.GetDataPresent(SC_DROP_FORMAT))
            {
                int srcidx = (int)e.Data.GetData(SC_DROP_FORMAT);
                Shortcut src = (Shortcut)MainPanel.Children[srcidx];
                //Shortcut src = (Shortcut)e.Data.GetData(SC_DROP_FORMAT);

                Shortcut trg = (Shortcut)e.OriginalSource;
                //Shortcut trg = (Shortcut)this.InputHitTest(e.GetPosition(this));
                if (trg != src)
                {
                    try
                    {
                        // remove the shortcut from the list, the panel and destroy itself
                        if (src != null && MainPanel.Children.Contains(src))
                        {
                            // Deleting an object requires anything that may 
                            // reference it like events to be doing this 
                            // need to be cleared

                            int targetidx = MainPanel.Children.IndexOf(trg);
                            src.Triggers.Clear();
                            // must be removed first before inserting it with new index
                            MainPanel.Children.Remove(src);
                            MainPanel.Children.Insert(targetidx, src);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    finally
                    {
                        
                    }
                }
            }
            //throw new NotImplementedException();
        }

        private void SC_dragOver(object sender, DragEventArgs e)
        {
            Shortcut trg = (Shortcut)sender;

            if (e.Data.GetDataPresent(SC_DROP_FORMAT) )
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

        private void SC_mouseMove(object sender, MouseEventArgs e)
        {
            // Get the current mouse position
            System.Windows.Point mousePos = e.GetPosition(null);
            Vector diff = startPoint - mousePos;

            if ( e.LeftButton == MouseButtonState.Pressed &&
                (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance))
            {
                // Get the dragged item
                Shortcut sc = sender as Shortcut;
                if( MainPanel.Children.Contains(sc) )
                {
                    // Can just pass its inex or the whole object
                    int sourceidx = MainPanel.Children.IndexOf(sc);
                    //DataObject dragData = new DataObject(SC_DROP_FORMAT, sc);
                    DataObject dragData = new DataObject(SC_DROP_FORMAT, sourceidx);
                    DragDrop.DoDragDrop(sc, dragData, DragDropEffects.Move);
                }
            }
        }

        private System.Windows.Point startPoint;
        private static string SC_DROP_FORMAT = "Shortcut";

        private void SC_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Store the mouse position
            startPoint = e.GetPosition(null);

            e.Handled = true;
        }

        private void DeleteShortcut(Shortcut sc)
        {
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
                    
                    // Delete the shortcut from the internal list
                    Shortcuts.Remove(sc);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
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
        private void RunShortcut(Shortcut scLink)
        {
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

                //mouseEvent.Handled = true;

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

    public class Shortcut : System.Windows.Controls.Image
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
