<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="config/dbConnection.asp"-->
<!--#include file="config/security.asp"-->
<%
' ============================================================================
'  song_plays.asp  ->  POST ?id={ID_SONG}
'  Incrementa el contador PLAYS de una cancion (llamado por fetch en JS).
'  Salvaguardas: solo sesiones activas; id entero via parametro ADO.
' ============================================================================

Dim oConn, oCmd, sOut, iSongId

Response.ContentType = "text/plain"
sOut = "0"

If IsLoggedIn() Then
    iSongId = ToLong(Request.QueryString("id"))
    If iSongId > 0 Then
        On Error Resume Next
        Err.Clear
        Set oConn = OpenConnection()
        If Err.Number = 0 Then
            Set oCmd = NewCommand(oConn, _
                "UPDATE TBL_SONG SET PLAYS = CASE WHEN PLAYS IS NULL THEN 1 ELSE PLAYS + 1 END WHERE ID_SONG = ?")
            AddIntParam oCmd, "@id", iSongId
            oCmd.Execute
            If Err.Number = 0 Then sOut = "1"
            CloseConn oConn
        End If
        On Error GoTo 0
    End If
End If

Response.Write sOut
Response.End
%>
