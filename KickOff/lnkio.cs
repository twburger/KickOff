using System;
using System.IO;
using IWshRuntimeLibrary;

namespace KickOff
{
    public static class lnkio
    {
        public static bool CreateShortcut(LnkData lnkData, string lnkFileName, string pathtoLnk)
        {
            bool bLnkCreated = false;

            WshShell shell = new WshShell();
            try
            {
                string shortcutAddress = System.IO.Path.GetFullPath(pathtoLnk);
                if (null == shortcutAddress || string.Empty == shortcutAddress)
                {
                    object shDesktop = (object)"Desktop";
                    shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop);
                }
                else {
                    shortcutAddress = pathtoLnk;
                }

                shortcutAddress += @"\" + lnkFileName;

                // Do not overwrite
                if (!System.IO.File.Exists(shortcutAddress))
                {
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);

                    shortcut.Arguments = lnkData.Arguments;
                    shortcut.Description = lnkData.Description; // "New shortcut for a Notepad";
                    //shortcut.FullName = string.Empty;  // FullName is read only
                    shortcut.Hotkey = lnkData.Hotkey; //"Ctrl+Shift+N";
                    shortcut.IconLocation = lnkData.IconLocation;
                    shortcut.RelativePath = lnkData.RelativePath;
                    shortcut.TargetPath = lnkData.TargetPath; //Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\notepad.exe";
                    shortcut.WindowStyle = lnkData.WindowStyle; // 1 for default window, 3 for maximize, 7 for minimize

                    shortcut.Save();

                    bLnkCreated = true;
                }
            }
            catch {
                throw new Exception("Write link error");
            }

            return (bLnkCreated);
        }

        public static LnkData ResolveShortcut(string lnkPath)
        {
            return (ResolveShortcut(System.IO.Path.GetFileName(System.IO.Path.GetFullPath(lnkPath)),
                System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(lnkPath))));
        }

        public static LnkData ResolveShortcut(string lnkFileName, string pathtoLnk)
        {
            WshShell shell = new WshShell();
            LnkData lnkData = null;
            try
            {
                string shortcutAddress = null;
                if (pathtoLnk != null)
                    shortcutAddress = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(pathtoLnk));
                if (null == shortcutAddress || string.Empty == shortcutAddress)
                {
                    object shDesktop = (object)"Desktop";
                    shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop);
                }
                else {
                    shortcutAddress = pathtoLnk;
                }

                shortcutAddress += @"\" + lnkFileName;

                bool bTisD = (System.IO.File.GetAttributes(shortcutAddress).HasFlag(FileAttributes.Directory));
                if (bTisD || System.IO.File.Exists(shortcutAddress))
                {
                    // the link itself is a directory
                    if (bTisD)
                    {
                        lnkData = new LnkData
                        {
                            Arguments = string.Empty,
                            Description = System.IO.Path.GetFileNameWithoutExtension(shortcutAddress),
                            FullName = System.IO.Path.GetFileName(shortcutAddress),
                            Hotkey = string.Empty,
                            IconLocation = string.Empty,
                            TargetPath = shortcutAddress,
                            WindowStyle = 1, // 1 for default window, 3 for maximize, 7 for minimize should remap to 1 = ProcessWindowStyle.Normal
                            WorkingDirectory = string.Empty,

                            bIsReference = false,
                            bTargetIsDirectory = true,
                            Bitmap = ico2bmap.GetBitmapFromFileIcon(shortcutAddress),
                            ShortcutAddress = shortcutAddress
                        };
                    }
                    else
                    {
                        bool bIsRef = (".appref-ms" == System.IO.Path.GetExtension(shortcutAddress).ToLower());
                        // if this is a reference use the reference as the target
                        if (bIsRef)
                        {
                            lnkData = new LnkData
                            {
                                Arguments = string.Empty,
                                Description = System.IO.Path.GetFileNameWithoutExtension(shortcutAddress),
                                FullName = System.IO.Path.GetFileName(shortcutAddress),
                                Hotkey = string.Empty,
                                IconLocation = string.Empty,
                                TargetPath = shortcutAddress,
                                WindowStyle = 1, // 1 for default window, 3 for maximize, 7 for minimize should remap to 1 = ProcessWindowStyle.Normal
                                WorkingDirectory = string.Empty,

                                bIsReference = bIsRef,
                                bTargetIsDirectory = false,
                                Bitmap = ico2bmap.GetBitmapFromFileIcon(shortcutAddress),
                                ShortcutAddress = shortcutAddress
                            };
                        }
                        else
                        {
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
                            bTisD = (System.IO.File.GetAttributes(shortcut.TargetPath).HasFlag(FileAttributes.Directory));
                            lnkData = new LnkData
                            {
                                Arguments = shortcut.Arguments,
                                Description = shortcut.Description, // A Comment
                                FullName = shortcut.FullName,  // shortcut FullName is read only
                                Hotkey = shortcut.Hotkey, //"Ctrl+Shift+N";
                                IconLocation = shortcut.IconLocation,
                                //RelativePath = shortcut.RelativePath, // The link relative path is set only
                                TargetPath = shortcut.TargetPath,
                                WindowStyle = shortcut.WindowStyle, // 1 for default window, 3 for maximize, 7 for minimize
                                WorkingDirectory = shortcut.WorkingDirectory,

                                bIsReference = bIsRef,
                                bTargetIsDirectory = bTisD,
                                Bitmap = ico2bmap.GetBitmapFromFileIcon(bTisD ? shortcutAddress : shortcut.TargetPath),
                                ShortcutAddress = shortcutAddress
                            };

                        }
                    }
                }
            }
            
            catch
            {
                lnkData = null;
            }

            return (lnkData);
        }

    } //lnkio

    public class LnkData
    {
        public LnkData() { }

        public string Arguments { get; set; }
        public string Description { get; set; }       // "New shortcut for a Notepad"
        public string FullName { get; set; }           // FullName is read only in the link
        public string Hotkey { get; set; }             //"Ctrl+Shift+N";
        public string IconLocation { get; set; }
        public string RelativePath { get; set; }        // write only in link
        public string TargetPath { get; set; }         //Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\notepad.exe";
        public int WindowStyle { get; set; }           // 1 for default window, 3 for maximize, 7 for minimize
        public string WorkingDirectory { get; set; }

        public bool bIsReference { get; set; }
        public bool bTargetIsDirectory { get; set; }
        public IconBitMap Bitmap { get; set; }
        public string ShortcutAddress { get; set; }
    }

}  // namespace

/*
        public static string GetShortcutTarget(string FilePath, bool bIsDir, ref string WorkingDirectory, 
            ref string TargetParameters, ref string ShortcutName, ref bool IsRef)
        {
            string LinkTargetPath = string.Empty;

            try
            {
                if (bIsDir || System.IO.File.Exists(FilePath))  // never assume thie file exists
                {
                    
                    WorkingDirectory = string.Empty;
                    TargetParameters = string.Empty;
                    ShortcutName = string.Empty;
                    IsRef = false;

                    // Files named {name}.appref-ms are Application References. 
                    // Contains text URL for the application, culture, 
                    //processor architecture and key used to sign the application
                    // These files are simply text and do not have the .lnk format
                    string ext = System.IO.Path.GetExtension(FilePath).ToLower();

                    // if not a link or reference it's a file or directory
                    if ( ext != ".lnk" && ext != ".appref-ms")
                    {
                        /// 
                        /// Is this an application or a directory? 
                        /// If so use the path as the target
                        /// 
                        IsRef = false;
                        ShortcutName = System.IO.Path.GetFileName(FilePath);
                        if (string.IsNullOrEmpty(ShortcutName))
                            ShortcutName = System.IO.Path.GetDirectoryName(FilePath);
                        LinkTargetPath = FilePath;
                        WorkingDirectory = System.IO.Path.GetDirectoryName(FilePath);

                        //throw new Exception(
                        //    string.Format(
                        //        "Can not process {0}. Supplied file must " +
                        //        "be a .LNK or a 'Click Once' application reference .appref-ms file"
                        //        , FilePath));
                    }
                    else
                    {
                        IsRef = (ext == ".appref-ms");

                        if (IsRef)
                        {
                            //@"rundll32.exe dfshim.dll,ShOpenVerbApplication http://github-windows.s3.amazonaws.com/GitHub.application#GitHub.application, Culture=neutral, PublicKeyToken=317444273a93ac29, processorArchitecture=x86";
                            //@"rundll32.exe dfshim.dll,ShOpenVerbShortcut D:\Users\User\Desktop\GitHub.appref-ms";

                            //TargetParameters = System.IO.File.ReadAllText(FilePath, Encoding.Default);
                            //LinkTargetPath = "rundll32.exe dfshim.dll,ShOpenVerbApplication "; // + TargetParameters;
                            //link = "rundll32.exe dfshim.dll,ShOpenVerbShortcut";
                            //link = file;
                            //TargetParameters = file;

                            LinkTargetPath = FilePath;
                            ShortcutName = System.IO.Path.GetFileName(FilePath);
                        }
                        else
                        {
                            LinkTargetPath = GetLnkTarget(FilePath, ref WorkingDirectory, ref TargetParameters, ref ShortcutName);
                        }
                    }
                }
                else
                {
                    throw new Exception("Supplied file does not exist");
                }

                return LinkTargetPath;
            }
            catch
            {
                return string.Empty;
            }
        }
*/
/*
        //public static string GetLnkTarget(string lnkPath, ref string workDir, 
        //    ref string targetParams, ref string scName)
        //{
        //    // needs reference to Windows\system32\shell32.dll

        //    //var shl = new Shell32.Shell(); 
        //    //lnkPath = System.IO.Path.GetFullPath(lnkPath);
        //    //var dir = shl.NameSpace(System.IO.Path.GetDirectoryName(lnkPath));
        //    //var itm = dir.Items().Item(System.IO.Path.GetFileName(lnkPath));
        //    //var lnk = (Shell32.ShellLinkObject)itm.GetLink;
        //    //return lnk.Target.Path;


        //    // Needs Interop.IWshRuntimeLibrary

        //    IWshShell shell = new IWshRuntimeLibrary.WshShell();

        //    // Create a link
        //    //IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(lnkPath);

        //    //read the link
        //    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(lnkPath);

        //    // read a link
        //    string pathOnly = System.IO.Path.GetDirectoryName(lnkPath);
        //    string filenameOnly = System.IO.Path.GetFileName(lnkPath);

        //    Shell shell = new Shell();
        //    Folder folder = shell.NameSpace(pathOnly);
        //    FolderItem folderItem = folder.ParseName(filenameOnly);
        //    if (folderItem != null)
        //    {
        //        Shell32.ShellLinkObject link = (Shell32.ShellLinkObject)folderItem.GetLink;
        //        return link.Path;
        //    }
        //    // the description is not returned when the shortcut target is a path
        //    if (string.IsNullOrEmpty(shortcut.Description))
        //        scName = System.IO.Path.GetFileName(shortcut.TargetPath);
        //    else
        //        scName = shortcut.Description;
        //    workDir = shortcut.WorkingDirectory;
        //    targetParams = shortcut.Arguments;
        //    shortcut.
        //    return shortcut.TargetPath;
        //}
*/

/*
FileStream fileStream = System.IO.File.Open(file, FileMode.Open, FileAccess.Read);
using (System.IO.BinaryReader fileReader = new BinaryReader(fileStream))
{
    fileStream.Seek(0x14, SeekOrigin.Begin);     // Seek to flags
    uint flags = fileReader.ReadUInt32();        // Read flags
    if ((flags & 1) == 1)
    {                      // Bit 1 set means we have to
                           // skip the shell item ID list
        fileStream.Seek(0x4c, SeekOrigin.Begin); // Seek to the end of the header
        uint offset = fileReader.ReadUInt16();   // Read the length of the Shell item ID list
        fileStream.Seek(offset, SeekOrigin.Current); // Seek past it (to the file locator info)
    }

    long fileInfoStartsAt = fileStream.Position; // Store the offset where the file info
                                                 // structure begins
    uint totalStructLength = fileReader.ReadUInt32(); // read the length of the whole struct
    fileStream.Seek(0xc, SeekOrigin.Current); // seek to offset to base pathname
    uint fileOffset = fileReader.ReadUInt32(); // read offset to base pathname
                                               // the offset is from the beginning of the file info struct (fileInfoStartsAt)
    fileStream.Seek((fileInfoStartsAt + fileOffset), SeekOrigin.Begin); // Seek to beginning of
                                                                        // base pathname (target)
    long pathLength = (totalStructLength + fileInfoStartsAt) - fileStream.Position - 2; // read
                                                                                        // the base pathname. I don't need the 2 terminating nulls.
    char[] linkTarget = fileReader.ReadChars((int)pathLength); // should be unicode safe
    link = new string(linkTarget);

    int begin = link.IndexOf("\0\0");
    if (begin > -1)
    {
        int end = link.IndexOf("\\\\", begin + 2) + 2;
        end = link.IndexOf('\0', end) + 1;

        string firstPart = link.Substring(0, begin);
        string secondPart = link.Substring(end);

        return firstPart + secondPart;
    }
    else {
        return link;
    }
*/
