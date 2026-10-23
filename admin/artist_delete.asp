<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../config/security.asp"-->
<%
' ============================================================================
'  admin/artist_delete.asp - Elimina un artista y sus vinculos (transaccion)
' ============================================================================
RequireAdmin()

Dim iId, oConn, bOk

iId = ToLong(Request.Form("id"))
If iId <= 0 Then iId = ToLong(Request.QueryString("id"))

If iId <= 0 Then
    Response.Redirect GetAppBaseUrl() & "admin/artists.asp?msg=error"
    Response.End
End If

On Error Resume Next
Set oConn = OpenConnection()
If Err.Number <> 0 Then
    Response.Redirect GetAppBaseUrl() & "admin/artists.asp?msg=error"
    Response.End
End If

bOk = True
oConn.BeginTrans

bOk = bOk And (ExecNonQueryId(oConn, "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = ?", iId) = 0)
bOk = bOk And (ExecNonQueryId(oConn, "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?", iId) = 0)

If bOk Then
    oConn.CommitTrans
Else
    oConn.RollbackTrans
End If

CloseConn oConn
On Error GoTo 0

If bOk Then
    Response.Redirect GetAppBaseUrl() & "admin/artists.asp?msg=deleted"
Else
    Response.Redirect GetAppBaseUrl() & "admin/artists.asp?msg=error"
End If
Response.End
%>
