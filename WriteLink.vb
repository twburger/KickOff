  ' Make sure variables are declared.
  option explicit

  ' Routine to create "mylink.lnk" on the Windows desktop.
  sub CreateShortCut()
    dim objShell, strDesktopPath, objLink
    set objShell = CreateObject("WScript.Shell")
    strDesktopPath = objShell.SpecialFolders("Desktop")
    set objLink = objShell.CreateShortcut(strDesktopPath & "\mylink.lnk")
    objLink.Arguments = "c:\windows\tips.txt"
    objLink.Description = "Shortcut to Notepad.exe"
    objLink.TargetPath = "c:\windows\notepad.exe"
    objLink.WindowStyle = 1
    objLink.WorkingDirectory = "c:\windows"
    objLink.Save
  end sub

  ' Program starts running here.
  call CreateShortCut()