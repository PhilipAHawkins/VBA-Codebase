Private Sub regedit_value_storage()
    Dim my_regedit As String, LastOpen As String, Msg As String
'   Create, get, and save setting from registry
'   Apparently any old Str or Int can go in these things.
'   "New entry" is what you want to change if you want a new setting,
'   and the argument you set after that is the setting that appears under "Data" in regedit.
'   Saves to 1. open Regedit, 2. Computer\HKEY_CURRENT_USER\SOFTWARE\
'   VB and VBA Program Settings\my_reg_entries\my_folder
    my_regedit = GetSetting("my_reg_entries", "my_folder", "New entry", "I set this just now")
'   final argument must be omitted if you just want to access the entry, not edit it.

'   Display the information
    Msg = "The registry edit is " & my_regedit & "."
    MsgBox Msg
 
'   store it
    SaveSetting "my_reg_entries", "my_folder", "New entry", my_regedit
' any computer that runs this sub will now have my_regedit saved in its registry.
End Sub

Sub get_a_reg_entry()
    my_regedit = GetSetting("my_reg_entries", "my_folder", "New entry")
    MsgBox my_regedit
End Sub