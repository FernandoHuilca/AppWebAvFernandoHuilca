<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../config/security.asp"-->
<%
' ============================================================================
'  admin/song_delete.asp - Elimina una cancion y sus vinculos (transaccion)
'  Solo POST. Solo rol ADMIN. Todos los borrados van por parametros ADO.
' ============================================================================
RequireAdmin()

Dim iId, oConn, bOk

iId = ToLong(Request.Form("id"))
If iId <= 0 Then iId = ToLong(Request.QueryString("id"))

If iId <= 0 Then
    Response.Redirect GetAppBaseUrl() & "admin/songs.asp?msg=error"
    Response.End
End If

On Error Resume Next
Set oConn = OpenConnection()
If Err.Number <> 0 Then
    Response.Redirect GetAppBaseUrl() & "admin/songs.asp?msg=error"
    Response.End
End If

bOk = True
oConn.BeginTrans

bOk = bOk And (ExecNonQueryId(oConn, "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", iId) = 0)
bOk = bOk And (ExecNonQueryId(oConn, "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", iId) = 0)
bOk = bOk And (ExecNonQueryId(oConn, "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?", iId) = 0)
bOk = bOk And (ExecNonQueryId(oConn, "DELETE FROM TBL_SONG WHERE ID_SONG = ?", iId) = 0)

If bOk Then
    oConn.CommitTrans
Else
    oConn.RollbackTrans
End If

CloseConn oConn
On Error GoTo 0

If bOk Then
    Response.Redirect GetAppBaseUrl() & "admin/songs.asp?msg=deleted"
Else
    Response.Redirect GetAppBaseUrl() & "admin/songs.asp?msg=error"
End If
Response.End
%>
