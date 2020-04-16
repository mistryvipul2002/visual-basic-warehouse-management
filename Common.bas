Attribute VB_Name = "Common"
Public sConnection As New ADODB.Connection

Public Sub initializeConnection()
    If Not sConnection.State = adStateOpen Then
        sConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\WareHouseDB.mdb;Persist Security Info=False;"
    End If
End Sub

Public Sub closeConnection()
    If Not sConnection.State = adStateClosed Then
        sConnection.Close
    End If
End Sub

Public Sub set_rec_getData(ByRef sRecordset As ADODB.recordSet, ByVal sSQL As String)
If Not sRecordset.State = adStateClosed Then
    sRecordset.Close
End If
With sRecordset
    .CursorLocation = adUseClient
    .Open sSQL, sConnection, adOpenKeyset, adLockOptimistic
End With
End Sub

Public Function search2(ByRef recordSet As ADODB.recordSet, ByVal field As String, ByVal query As String) As Boolean
    recordSet.MoveFirst
    recordSet.Find field & " LIKE '" & query & "'", , adSearchForward
    
    If recordSet.EOF Then
        search2 = False
    Else
        search2 = True
    End If
End Function

Public Function Search(ByVal tableName As String, ByVal field As String, ByVal query As String) As Boolean
    Dim recordSetProduct As New ADODB.recordSet
    Call set_rec_getData(recordSetProduct, "select * from " & tableName)
    Search = search2(recordSetProduct, field, query)
    recordSetProduct.Close
End Function
