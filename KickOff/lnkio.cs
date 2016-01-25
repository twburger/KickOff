using System;
using System.Text;
using System.IO;

namespace KickOff
{
    public static class lnkio
    {
        public static string GetShortcutTarget(string FilePath, bool bIsDir, ref string WorkingDirectory, ref string TargetParameters, ref string ShortcutName, ref bool IsRef)
        {
            string LinkTargetPath = string.Empty;

            try
            {
                if (bIsDir || File.Exists(FilePath))  // never assume thie file exists
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

        public static string GetLnkTarget(string lnkPath, ref string workDir, 
            ref string targetParams, ref string scName)
        {
            // needs reference to Windows\system32\shell32.dll

            //var shl = new Shell32.Shell(); 
            //lnkPath = System.IO.Path.GetFullPath(lnkPath);
            //var dir = shl.NameSpace(System.IO.Path.GetDirectoryName(lnkPath));
            //var itm = dir.Items().Item(System.IO.Path.GetFileName(lnkPath));
            //var lnk = (Shell32.ShellLinkObject)itm.GetLink;
            //return lnk.Target.Path;


            // Needs Interop.IWshRuntimeLibrary

            IWshRuntimeLibrary.IWshShell shell = new IWshRuntimeLibrary.WshShell();

            IWshRuntimeLibrary.IWshShortcut shortcut =
                (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(lnkPath);

            scName = shortcut.Description;
            workDir = shortcut.WorkingDirectory;
            targetParams = shortcut.Arguments;
            return shortcut.TargetPath;
        }
    }

}


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
