' This all goes in a module.

Public Type extfinder
	FullPath As String
	File As String
	FolderPath As String
	Extension As String
End Type

 
Function extfind() as extfinder

Dim fd As FileDialog

	Set fd = Application.FileDialog(msoFileDialogFilePicker)
	With fd
		.Title = "Pick your file"
		.AllowMultiSelect = False
		If .Show = True Then
			extfind.FullPath = .SelectedItems(1) 'full path of folders and file and extension
			extfind.File = Dir(extfind.FullPath) ' file and extension
			extfind.FolderPath = Left$(extfind.FullPath, (Len(extfind.FullPath)) - Len(extfind.File)) ' just the folder path
			extfind.Extension =  Right$(extfind.FullPath, (Len(extfind.FullPath)) - InStrRev(extfind.FullPath, ".")) ' just the file's extension
		end if
	End With
End Function

Sub attempttype()
' sample: getting all attributes of our new Type when 
Dim my_attrs As extfinder
	my_attrs = extfind() ' returns all extfind attrs of the selected file
	my_attrs.Extension ' returns that attr.
End Sub

Sub getFilePath()
' sample: get one attribute when you only want the one.
Dim my_path as String
	my_path = extfind.FullPath
End Sub