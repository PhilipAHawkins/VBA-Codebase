' using shell commands within VBA
' including renaming a file extension, such as from zip to txt
Sub renamethis()

Dim objShell as Object
Dim objFolder as Object
Dim objFile as Object

set objShell = CreateObject("Shell.Application")
set objFolder = objShell.namespace("\\PATH\Folder")
For Each objFile in objFolder.Items
    If objFile.Name = "rename_this.txt" Then
        objFile  =Replace(objFile, "txt", "zip")
    end If
next objFile
end Sub

' unzip image files from a zip file
' Thanks, Ron de Bruin

Sub unzip2()

Dim FSO as Object
Dim oApp as Object
dim fname as variant
dim fnamefldr as variant
dim defpath as string
dim strDate as string
dim fnamezip as variant
dim fnameinsubfolder as Object

fname = Application.GetOpenFilename(filefilter:="Zip Files (*.zip), *.zip",_
                                    MultiSelect:=False)

if fname = false then
    ' do nothing
else
    ' set root
    defpath = Application.DefaultFilepath
    if right(defpath,1) <> "\" then
        defpath = defpath & "\"
    end if

    ' create temp unzip folder
    strDate = Format(Now, " dd-mm-yy h-mm-ss")
    fnamefldr = defpath & "MyUnzipFolder" & strDate & "\"
    MkDir fnamefldr

    'extract files with shell
    Set oApp = CreateObject("shell.Application")
        For Each fnamezip in oApp.namespace(fname).Items
            ' zip files can have folders as well as files so we must look in folders inside the zip
            if fnamezip.Type <> "File folder" Then 
                if lcase(fnamezip)like lcase ("*.png") then
                    oApp.namespace(fnamefldr).CopyHere _
                        oApp.namespace(fname).Items.Item(CStr(fnamezip))
                end if
            else
                'TODO
                ' for each fnameinsubfolder in oapp.namespace(fname).items(fnamezip).subfolders
                    'if lcase(fnamezip) like lcase("*.png") then
                     '                   oApp.namespace(fnamefldr).CopyHere _
                      '  oApp.namespace(fname).Items.Item(CStr(fnamezip))
                '   end if
            end if
        next
    
    debug.print("files located in: " & fnamefldr)

    on error resume next
    set FSO = CreateObject("scripting.filesystemobject")
    FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
end if
end sub
