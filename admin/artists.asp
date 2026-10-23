<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../config/security.asp"-->
<%
' ============================================================================
'  admin/artists.asp - Gestion de Artistas (listado)
'  Solo rol ADMIN.
' ============================================================================
RequireAdmin()

Dim sPageTitle, sActiveNav
sPageTitle = "PoliMusic - Gestion de Artistas"
sActiveNav = "artists"

Dim oConn, oRs
Dim vArt, vCount
Dim dCnt
Dim lId, sName, sLast, sFull, sImgUrl, sImgChip, lCnt, sConfirm
Dim i, sMsg, sMsgClass

sMsg = AdminMsg()
sMsgClass = "alert-ok"
If LCase(Trim("" & Request.QueryString("msg"))) = "error" Then sMsgClass = "alert-error"

On Error Resume Next
Set oConn = OpenConnection()
If Err.Number <> 0 Then
    ShowErrorPage "No hay conexi&oacute;n con SQL Server. Revisa config/dbConnection.asp."
End If
On Error GoTo 0

Set oRs = OpenRs(oConn, "SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST ORDER BY ARTIST_NAME, ARTIST_LASTNAME")
If oRs Is Nothing Then
    CloseConn oConn
    ShowErrorPage "No se pudo leer la tabla de artistas."
End If
vArt = Empty
If Not oRs.EOF Then vArt = oRs.GetRows()
CloseRs oRs

' cuantas canciones tiene cada artista
Set dCnt = Server.CreateObject("Scripting.Dictionary")
Set oRs = OpenRs(oConn, "SELECT ID_ARTIST, COUNT(*) AS CANT FROM TBL_PLAYER GROUP BY ID_ARTIST")
If Not (oRs Is Nothing) Then
    Do While Not oRs.EOF
        If Not dCnt.Exists(CStr(NzLong(oRs("ID_ARTIST")))) Then
            dCnt.Add CStr(NzLong(oRs("ID_ARTIST"))), NzLong(oRs("CANT"))
        End If
        oRs.MoveNext
    Loop
    CloseRs oRs
End If
CloseConn oConn

vCount = 0
If IsArray(vArt) Then vCount = UBound(vArt, 2) + 1
%>
<!--#include file="../includes/header.asp"-->

<header class="page-head">
  <h1>Artistas &rarr; <span class="muted">Listado</span></h1>
  <a class="btn btn-primary" href="artist_form.asp">+ Nuevo artista</a>
  <a class="btn btn-ghost" href="<%= H(GetAppBaseUrl()) %>index.asp">Volver al inicio</a>
</header>

<% If Len(sMsg) > 0 Then %>
<div class="alert <%= sMsgClass %>"><%= sMsg %></div>
<% End If %>

<div class="admin-tools">
  <span class="chip"><%= vCount %> artista<% If vCount <> 1 Then %>s<% End If %></span>
  <span class="muted">Las fotos deben existir en la carpeta <b>ArtistImages</b> del proyecto.</span>
</div>

<div class="table-wrap">
  <table class="tbl">
    <thead>
      <tr>
        <th>ID</th>
        <th>Foto</th>
        <th>Nombre completo</th>
        <th class="hide-md">Ruta de foto</th>
        <th class="num">Canciones</th>
        <th class="text-right">Acciones</th>
      </tr>
    </thead>
    <tbody>
<%
If vCount = 0 Then
%>
      <tr><td colspan="6" class="muted" style="text-align:center;padding:26px;">No hay artistas registrados todav&iacute;a.</td></tr>
<%
Else
    For i = 0 To vCount - 1
        lId = NzLong(vArt(0, i))
        sName = Nz(vArt(1, i))
        sLast = Nz(vArt(2, i))
        sFull = sName
        If Len(Trim(sLast)) > 0 Then sFull = sFull & " " & Trim(sLast)
        sImgUrl = AssetUrl(vArt(3, i))
        If ResourceFileExists(MapResourceUrl(Nz(vArt(3, i)))) Then
            sImgChip = ""
        Else
            sImgChip = "<span class='ticket-chip file-missing'>sin archivo</span>"
        End If
        If dCnt.Exists(CStr(lId)) Then lCnt = dCnt(CStr(lId)) Else lCnt = 0
        sConfirm = "&iquest;Eliminar definitivamente al artista &quot;" & H(sFull) & "&quot;? " & _
                   "Se quitar&aacute;n sus v&iacute;nculos con canciones."
%>
      <tr>
        <td class="id-cell num"><%= lId %></td>
        <td>
          <img class="thumb" src="<%= H(sImgUrl) %>" alt="Foto de <%= H(sFull) %>" style="border-radius:50%;"
               onerror="this.onerror=null;this.src='<%= H(GetAppBaseUrl()) %>assets/placeholder.svg';">
        </td>
        <td><b><%= H(sFull) %></b></td>
        <td class="hide-md path-cell"><%= H(Nz(vArt(3, i))) %> <%= sImgChip %></td>
        <td class="num"><%= lCnt %></td>
        <td class="text-right">
          <span class="cell-actions">
            <a class="btn btn-sm btn-outline" href="artist_form.asp?id=<%= lId %>">Editar</a>
            <form method="post" action="artist_delete.asp" data-confirm="<%= sConfirm %>" style="display:inline;">
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
