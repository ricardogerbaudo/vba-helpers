Function GetWorkbookPath(Optional wb As Workbook)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' Purpose:  Returns a workbook's physical path, even when they are saved in
    '           synced OneDrive Personal, OneDrive Business or Microsoft Teams folders.
    '           If no value is provided for wb, it's set to ThisWorkbook object instead.
    ' Author:   Ricardo Gerbaudo
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    If InStr(1, wb.Path, "https://") = 0 Then
        GetWorkbookPath = wb.Path
        Exit Function
    Else
        Const HKEY_CURRENT_USER = &H80000001
        Dim objRegistryProvider As Object
        Dim strRegistryPath As String
        Dim arrSubKeys()
        Dim strSubKey As Variant
        Dim strUrlNamespace As String
        Dim strMountPoint As String
        Dim strLocalPath As String
        Dim strRemainderPath As String
        Dim strLibraryType As String
    
        Set objRegistryProvider = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    
        strRegistryPath = "SOFTWARE\SyncEngines\Providers\OneDrive"
        objRegistryProvider.EnumKey HKEY_CURRENT_USER, strRegistryPath, arrSubKeys
        
        For Each strSubKey In arrSubKeys
            objRegistryProvider.GetStringValue HKEY_CURRENT_USER, strRegistryPath & "\" & strSubKey & "\", "UrlNamespace", strUrlNamespace
            If InStr(1, wb.Path, strUrlNamespace) <> 0 Or InStr(1, strUrlNamespace, wb.Path) <> 0 Then
                objRegistryProvider.GetStringValue HKEY_CURRENT_USER, strRegistryPath & "\" & strSubKey & "\", "MountPoint", strMountPoint
                objRegistryProvider.GetStringValue HKEY_CURRENT_USER, strRegistryPath & "\" & strSubKey & "\", "LibraryType", strLibraryType
                
                If InStr(1, wb.Path, strUrlNamespace) <> 0 Then
                    strRemainderPath = Replace(wb.Path, strUrlNamespace, vbNullString)
                Else
                    GetWorkbookPath = strMountPoint
                    Exit Function
                End If
                
                'If Non-Business OneDrive, skips the GUID part of the URL to match with physical path
                If InStr(1, strUrlNamespace, "https://d.docs.live.net") <> 0 Then
                    If InStr(2, strRemainderPath, "/") = 0 Then
                        strRemainderPath = vbNullString
                    Else
                        strRemainderPath = Mid(strRemainderPath, InStr(2, strRemainderPath, "/"))
                    End If
                End If
                
                'If OneDrive Business, adds extra slash at the start of string to match the pattern
                strRemainderPath = IIf(InStr(1, strUrlNamespace, "sharepoint.com") <> 0, "/", vbNullString) & strRemainderPath
                strRemainderPath = Replace(strRemainderPath, "/", "\")
                
                strLocalPath = ""
                
                If (InStr(1, strRemainderPath, "\")) <> 0 Then
                    strLocalPath = Mid(strRemainderPath, InStr(1, strRemainderPath, "\"))
                ElseIf strRemainderPath <> "" Then
                    strLocalPath = "\" & strRemainderPath
                End If
                
                strLocalPath = strMountPoint & strLocalPath
                
                If Dir(strLocalPath & "\" & wb.Name) = "" Then
                    strLocalPath = Mid(strRemainderPath, InStr(2, strRemainderPath, "\"))
                    strLocalPath = strMountPoint & strLocalPath
                    If Dir(strLocalPath & "\" & wb.Name) <> "" Then
                        GetWorkbookPath = strLocalPath
                        Exit Function
                    End If
                Else
                    GetWorkbookPath = strLocalPath
                    Exit Function
                End If
            End If
        Next
    End If
        
End Function
