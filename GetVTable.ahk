#Requires AutoHotkey v2.0

; ===============================================================================================================================
; GetVTable(Interface, FilePattern, Recurse := false, ToolTipNo := false)
; Function:       Gets information about the specified interface's virtual table from a header file.
;                 Visual Studio may need to installed in order to access the header files.
;                 The user is encouraged to verify the results.
; Parameters:     - Interface - (Required) The name (not case-sensitive) of an interface, e.g. 'IShellItem'.
;                 - FilePattern - (Required) The name of a file, directory, or file pattern to search for the interface.
;                   If FilePattern is a directory, the function will search for header (.h) files in that directory.
;                 - Recurse - (Optional) A boolean value that indicates whether to recurse through the subfolders or not.
;                   If omitted, it defaults to false.
;                 - ToolTipNo - (Optional) An integer between 0 and 20 to indicate which tooltip, if any, will show the path
;                   being searched. If omitted, it defaults to false (0), i.e. don't show a tooltip.
; Return values:  - If the interface is not found, the return value is an empty array.
;                 - If the interface is found the return value is an array whose elements are object literals with the
;                   following property names:
;                     - Index - The position of the method within the original interface definition.
;                       This value is the Index parameter in the ComCall function.
;                     - Name - The name of the interface method.
;                     - LineNo - The line number where the method is defined in the file.
;                       Note: This is the line that contains the 'STDMETHOD' or 'STDMETHODCALLTYPE' strings.
;                     - Interface - The name of the interface that defines the method. For interfaces that don't have a
;                       'InterfaceNameVtbl' definition, this is the name of the interface that the Interface input parameter
;                       inherits from.
;                     - IID - The identifier string of the interface that defines the method.
;                   The array object will also have the following own property:
;                     - Path - The path of the file that contains the interface definition.
;                 - If FilePattern doesn't match any files, an error is thrown.
; Global vars:    None
; Depenencies:    None
; Requirements:   AHK v2.0
; Tested with:    AHK v2.0.0 (U32/U64)
; Tested on:      Win 10 Pro (x64)
; Written by:     iPhilip
; Forum link:     https://www.autohotkey.com/boards/viewtopic.php?f=83&t=130455
; References:     https://www.autohotkey.com/docs/v2/lib/ComCall.htm
; ===============================================================================================================================

GetVTable(Interface, FilePattern, Recurse := false, ToolTipNo := false) {
   static IUnknown := []
   
   ; Initialize the IUnknown array for interfaces where 'InterfaceNameVtbl' is not defined (InterfaceType = 2).
   ; The IUnknown interface is not defined in the header files associated with those interfaces.
   
   if !IUnknown.Length
      for Method in ['QueryInterface', 'AddRef', 'Release']
         IUnknown.Push({Index: A_Index - 1, Name: Method, LineNo: 'N/A', Interface: 'IUnknown'})
   
   VTable := []
   static GUID := '[[:xdigit:]]{8}-[[:xdigit:]]{4}-[[:xdigit:]]{4}-[[:xdigit:]]{4}-[[:xdigit:]]{12}'
   NeedleRegEx1 := 'iS)\r?\n +MIDL_INTERFACE\("(' GUID ')"\)\r?\n +(' Interface ')( : public \w+)?\r?\n'
   NeedleRegEx2 := 'iS)\r\ninterface DX_DECLARE_INTERFACE\("(' GUID ')"\) (' Interface ')  : public (\w+)\r\n'
   Loop Files DirExist(FilePattern) ? FilePattern '\*.h' : FilePattern, Recurse ? 'FR' : 'F' {
      if ToolTipNo
         ToolTip A_LoopFileFullPath, , , ToolTipNo
      Text := FileRead(A_LoopFileFullPath)
      if ((FoundPos := RegExMatch(Text, NeedleRegEx1, &Match)) && (InterfaceType := 1))
      || ((FoundPos := RegExMatch(Text, NeedleRegEx2, &Match)) && (InterfaceType := 2)) {
         IID := '{' Match[1] '}', Interface := Match[2]
         if InterfaceType = 1 {
            NeedleRegEx1 := 'sS)(.+?)STDMETHODCALLTYPE( \*|\* )(\w+) ?\)\(.*?\r?\n'
            NeedleRegEx2 := 'S)\r?\n +END_INTERFACE\r?\n *} .*' Interface 'Vtbl;\r?\n'
            FoundPos := RegExMatch(Text, 'S)' Interface 'Vtbl', &Match, FoundPos)
            Index := -1
         } else {
            VTable.Push((Match[3] = 'IUnknown' ? IUnknown : GetVTable(Match[3], A_LoopFileFullPath, false, ToolTipNo))*)
            NeedleRegEx1 := 'sS)(.+?)STDMETHOD_?\((\w+, )?(\w+)\)\(\r\n'
            NeedleRegEx2 := 'S)\r\n}; // interface ' Interface '\r\n'
            Index := VTable.Length - 1
         }
         while (FoundPos := RegExMatch(Text, NeedleRegEx1, &Match, FoundPos + Match.Len)) && !RegExMatch(Match[1], NeedleRegEx2) {
            LineNo := StrReplace(SubStr(Text, 1, FoundPos + Match.Len), '`n', , , &Count) && Count
            VTable.Push({Index: Index + A_Index, Name: Match[3], LineNo: LineNo, Interface: Interface, IID: IID})
         }
         VTable.Path := A_LoopFileFullPath
         break
      }
   }
   if !IsSet(FoundPos)
      throw ValueError('FilePattern does not match any files.', -1, FilePattern)
   if ToolTipNo
      ToolTip , , , ToolTipNo
   return VTable
}

; ---------------
