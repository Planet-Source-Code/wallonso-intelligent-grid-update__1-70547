Attribute VB_Name = "Mod_DataBase"
Option Explicit

Public Function ConnectDataBase(sDBPath As String, Optional sPassWord As String) As ADODB.Connection
Dim cnADO As New ADODB.Connection
Dim strconnect As String
    Set cnADO = New ADODB.Connection
    If Len(sPassWord) Then
        'cnADO.Provider = "Microsoft.Jet.OLEDB.4.0;"
        'strconnect = "Datasource=" & sDBPath & ";Password=" & sPassWord & ";"
    strconnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & sDBPath & ";" & _
                   "Jet Oledb:Database Password=" & sPassWord & ";" '& _
                   '"Persist Security Info=False;"
    Else
        'cnADO.Provider = "Microsoft.Jet.OLEDB.4.0;"
        strconnect = sDBPath
            strconnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sDBPath & ";"                    '& _
                   '"Jet Oledb:Database Password=" & sPassWord & ";" & _
                   '"Persist Security Info=False;"
    End If
    
    cnADO.ConnectionString = strconnect ' sDBPath ' m_sDatabaseName
    cnADO.Open
    Set ConnectDataBase = cnADO
End Function

Public Function CloseDatabase(cnAD As ADODB.Connection)
    If Not cnAD Is Nothing Then
        cnAD.Close
    End If
End Function


