Public Site


site = InputBox("Website To Ping:")


WScript.Echo "Site Is: " & Ping( site )

Function Ping( myHostName )


    ' Standard housekeeping
    Dim colPingResults, objPingResult, strQuery

    ' Define the WMI query
    strQuery = "SELECT * FROM Win32_PingStatus WHERE Address = '" & myHostName & "'"

    ' Run the WMI query
    Set colPingResults = GetObject("winmgmts://./root/cimv2").ExecQuery( strQuery )

    ' Translate the query results to either True or False
    For Each objPingResult In colPingResults
        If Not IsObject( objPingResult ) Then
            Ping = "Offline"
        ElseIf objPingResult.StatusCode = 0 Then
            Ping = "Online"
        Else
            Ping = "Online"
        End If
    Next

    Set colPingResults = Nothing
End Function