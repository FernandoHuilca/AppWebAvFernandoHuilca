<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../config/security.asp"-->
<%
' ============================================================================
'  admin/songs.asp - Gestion de Canciones (CRUD: listar)
'  Solo para rol ADMIN.
' ============================================================================
RequireAdmin()

Dim sPageTitle, sActiveNav
sPageTitle = "PoliMusic - Gestion de Canciones"
sActiveNav = "songs"

Dim oConn, oRs
Dim vSongs, vCount
Dim dArt
Dim lId, sName, sPath, lPlays, sArtText, sFileChip, sConfirm
Dim i, sMsg, sMsgClass
Dim sTmp1, sTmp2, lTmp

sMsg = AdminMsg()
sMsgClass = "alert-ok"
If LCase(Trim("" & Request.QueryString("msg"))) = "error" Then sMsgClass = "alert-error"

On Error Resume Next
Set oConn = OpenConnection()
If Err.Number <> 0 Then
    ShowErrorPage "No hay conexi&oacute;n con SQL Server. Revisa config/dbConnection.asp."
End If
On Error GoTo 0

Set oRs = OpenRs(oConn, "SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG ORDER BY ID_SONG")
If oRs Is Nothing Then
    CloseConn oConn
    ShowErrorPage "No se pudo leer la tabla de canciones."
End If
If Not oRs.EOF Then vSongs = oRs.GetRows()
CloseRs oRs
Set oRs = Nothing

' Artistas por cancion
Set dArt = Server.CreateObject("Scripting.Dictionary")
Set oRs = OpenRs(oConn, "SELECT P.ID_SONG, A.ARTIST_NAME, A.ARTIST_LASTNAME FROM TBL_PLAYER P " & _
                        "INNER JOIN TBL_ARTIST A ON A.ID_ARTIST = P.ID_ARTIST ORDER BY P.ID_SONG, A.ARTIST_NAME")
If Not (oRs Is Nothing) Then
    Do While Not oRs.EOF
        lTmp = NzLong(oRs("ID_SONG"))
        sTmp1 = Trim(Nz(oRs("ARTIST_NAME")))
        sTmp2 = Trim(Nz(oRs("ARTIST_LASTNAME")))
        If Len(sTmp2) > 0 Then sTmp1 = sTmp1 & " " & sTmp2
        If dArt.Exists(CStr(lTmp)) Then
            dArt(CStr(lTmp)) = dArt(CStr(lTmp)) & ", " & sTmp1
        Else
            dArt.Add CStr(lTmp), sTmp1
        End If
        oRs.MoveNext
    Loop
    CloseRs oRs
End If
vCount = 0
If IsArray(vSongs) Then vCount = UBound(vSongs, 2) + 1
CloseConn oConn
%>
<!--#include file="../includes/header.asp"-->

<header class="page-head">
  <h1>Canci&oacute;n &rarr; <span class="muted">Listado</span></h1>
  <a class="btn btn-primary" href="song_form.asp">+ Nueva canci&oacute;n</a>
  <a class="btn btn-ghost" href="<%= H(GetAppBaseUrl()) %>index.asp">Volver al inicio</a>
</header>

<% If Len(sMsg) > 0 Then %>
<div class="alert <%= sMsgClass %>"><%= sMsg %></div>
<% End If %>

<div class="admin-tools">
  <span class="chip"><%= vCount %> canci&oacute;n<% If vCount <> 1 Then %>es<% End If %></span>
  <span class="muted">El audio debe existir en la carpeta <b>SongFiles</b> del proyecto.</span>
</div>

<div class="table-wrap">
  <table class="tbl">
    <thead>
      <tr>
        <th>ID</th>
        <th class="hide-md">Archivo MP3</th>
        <th>Canci&oacute;n</th>
        <th>Artistas</th>
        <th class="num">Reproducc.</th>
        <th class="text-right">Acciones</th>
      </tr>
    </thead>
    <tbody>
<%
If vCount = 0 Then
%>
      <tr><td colspan="6" class="muted" style="text-align:center;padding:26px;">No hay canciones registradas todav&iacute;a. Crea la primera con el bot&oacute;n &laquo;Nueva canci&oacute;n&raquo;.</td></tr>
<%
Else
    For i = 0 To vCount - 1
        lId = NzLong(vSongs(0, i))
        sName = Nz(vSongs(1, i))
        sPath = Nz(vSongs(2, i))
        lPlays = NzLong(vSongs(3, i))
        If dArt.Exists(CStr(lId)) Then sArtText = dArt(CStr(lId)) Else sArtText = ""
        If ResourceFileExists(MapResourceUrl(sPath)) Then
            sFileChip = "<span class='ticket-chip file-ok'>OK</span>"
        Else
            sFileChip = "<span class='ticket-chip file-missing'>sin archivo</span>"
        End If
        sConfirm = "&iquest;Eliminar definitivamente la canci&oacute;n &quot;" & H(sName) & "&quot;? " & _
                   "Tambi&eacute;n se quitar&aacute; de artistas, &aacute;lbumes y playlists."
%>
      <tr>
        <td class="id-cell num"><%= lId %></td>
        <td class="hide-md path-cell"><%= H(sPath) %> <%= sFileChip %></td>
        <td><b><%= H(sName) %></b></td>
        <td><% If Len(sArtText) > 0 Then %><%= H(sArtText) %><% Else %><span class="muted">&mdash;</span><% End If %></td>
        <td class="num"><%= PlaysText(lPlays) %></td>
        <td class="text-right">
          <span class="cell-actions">
            <a class="btn btn-sm btn-outline" href="song_form.asp?id=<%= lId %>">Editar</a>
            <form method="post" action="song_delete.asp" data-confirm="<%= sConfirm %>" style="display:inline;">
              <input type="hidden" name="id" value="<%= lId %>">
              <button class="btn btn-sm btn-danger" type="submit">Eliminar</button>
            </form>
          </span>
        </td>
      </tr>
<%
    Next
End If
%>
    </tbody>
  </table>
</div>

<script src="<%= H(GetAppBaseUrl()) %>js/admin.js" defer></script>
<!--#include file="../includes/footer.asp"-->
