#Include, Gdip.ahk

if (!pToken := Gdip_Startup()) {
    MsgBox, 48, Gdiplus Error!, Gdiplus failed to start. Please ensure you have Gdiplus on your system.
    ExitApp
}
OnExit, Quit
return

/*  f1 hotkey
 *  Paste a screenshot of the active window into an open Word document.
 *  Shapes.AddPicture Method - http://msdn.microsoft.com/en-us/library/ff940072.aspx
 */
f1::
if (!WinExist("ahk_class OpusApp")) {
    MsgBox, Could not find a Word window!
    return
}
oWord := ComObjActive("Word.Application")
oWord.ActiveDocument.Shapes.AddPicture(TempFile:=CaptureWindow(), false, true)
FileDelete, % TempFile
return

/*  CaptureWindow(Window, Ext)
 *  Creates a screenshot of a window.
 *
 *  Parameters:
 *      Window      - The WinTitle of the window to capture. (Default: the active window)
 *      Ext         - The type of image to create. Supported extensions are: BMP, DIB, RLE, JPG,
 *                    JPEG, JPE, JFIF, GIF, TIF, TIFF, PNG. (Default: PNG)
 *  Returns:
 *      Path        - The path of the file created.
 */
CaptureWindow(Window:="A", Ext:="png") {
    Gdip_SaveBitmapToFile(Gdip_BitmapFromHWND(WinExist(Window)), Path := A_ScriptDir "\Temp" A_Now "." Ext)
    return Path
}

Quit:
Gdip_Shutdown(pToken)
ExitApp