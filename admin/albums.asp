<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../config/security.asp"-->
<%
' ============================================================================
'  admin/albums.asp - Gestion de Albumes (listado)
'  Solo rol ADMIN.
' ============================================================================
RequireAdmin()

Dim sPageTitle, sActiveNav
sPageTitle = "PoliMusic - Gestion de Albumes"
sActiveNav = "albums"

Dim oConn, oRs
Dim vAlb, vCount
Dim dCnt
Dim lId, sTitle, iTmpYear, sYear, sImgUrl, sImgChip, lCnt, sConfirm
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

Set oRs = OpenRs(oConn, "SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM ORDER BY ID_ALBUM")
If oRs Is Nothing Then
    CloseConn oConn
    ShowErrorPage "No se pudo leer la tabla de albumes."
End If
vAlb = Empty
If Not oRs.EOF Then vAlb = oRs.GetRows()
CloseRs oRs

' cuantas canciones tiene cada album
Set dCnt = Server.CreateObject("Scripting.Dictionary")
Set oRs = OpenRs(oConn, "SELECT ID_ALBUM, COUNT(*) AS CANT FROM TBL_TRACKLIST GROUP BY ID_ALBUM")
If Not (oRs Is Nothing) Then
    Do While Not oRs.EOF
        If Not dCnt.Exists(CStr(NzLong(oRs("ID_ALBUM")))) Then
            dCnt.Add CStr(NzLong(oRs("ID_ALBUM"))), NzLong(oRs("CANT"))
        End If
        oRs.MoveNext
    Loop
    CloseRs oRs
End If
CloseConn oConn

vCount = 0
If IsArray(vAlb) Then vCount = UBound(vAlb, 2) + 1
%>
<!--#include file="../includes/header.asp"-->

<header class="page-head">
  <h1>&Aacute;lbumes &rarr; <span class="muted">Listado</span></h1>
  <a class="btn btn-primary" href="album_form.asp">+ Nuevo &aacute;lbum</a>
  <a class="btn btn-ghost" href="<%= H(GetAppBaseUrl()) %>index.asp">Volver al inicio</a>
</header>

<% If Len(sMsg) > 0 Then %>
<div class="alert <%= sMsgClass %>"><%= sMsg %></div>
<% End If %>

<div class="admin-tools">
  <span class="chip"><%= vCount %> &aacute;lb<% If vCount = 1 Then %>um<% Else %>umes<% End If %></span>
  <span class="muted">Las portadas deben existir en la carpeta <b>AlbumImages</b> del proyecto.</span>
</div>

<div class="table-wrap">
  <table class="tbl">
    <thead>
      <tr>
        <th>ID</th>
        <th>Portada</th>
        <th>T&iacute;tulo</th>
        <th>A&ntilde;o</th>
        <th class="hide-md">Ruta de portada</th>
        <th class="num">Canciones</th>
        <th class="text-right">Acciones</th>
      </tr>
    </thead>
    <tbody>
<%
If vCount = 0 Then
%>
      <tr><td colspan="7" class="muted" style="text-align:center;padding:26px;">No hay &aacute;lbumes registrados todav&iacute;a.</td></tr>
<%
Else
    For i = 0 To vCount - 1
        lId = NzLong(vAlb(0, i))
        sTitle = Nz(vAlb(1, i))
        iTmpYear = NzLong(vAlb(2, i))
        sYear = ""
        If Len(Trim(Nz(vAlb(2, i)))) > 0 And NzLong(vAlb(2, i)) > 0 Then sYear = CStr(NzLong(vAlb(2, i)))
        sImgUrl = AssetUrl(vAlb(3, i))
        If ResourceFileExists(MapResourceUrl(Nz(vAlb(3, i)))) Then
            sImgChip = ""
        Else
            sImgChip = "<span class='ticket-chip file-missing'>sin archivo</span>"
        End If
        If dCnt.Exists(CStr(lId)) Then lCnt = dCnt(CStr(lId)) Else lCnt = 0
        sConfirm = "&iquest;Eliminar definitivamente el &aacute;lbum &quot;" & H(sTitle) & "&quot;? " & _
                   "Se quitar&aacute;n sus v&iacute;nculos del tracklist."
%>
      <tr>
        <td class="id-cell num"><%= lId %></td>
        <td>
          <img class="thumb" src="<%= H(sImgUrl) %>" alt="Portada de <%= H(sTitle) %>"
               onerror="this.onerror=null;this.src='<%= H(GetAppBaseUrl()) %>assets/placeholder.svg';">
        </td>
        <td><b><%= H(sTitle) %></b></td>
        <td class="num"><% If Len(sYear) > 0 Then %><%= sYear %><% Else %><span class="muted">&mdash;</span><% End If %></td>
        <td class="hide-md path-cell"><%= H(Nz(vAlb(3, i))) %> <%= sImgChip %></td>
        <td class="num"><%= lCnt %></td>
        <td class="text-right">
          <span class="cell-actions">
            <a class="btn btn-sm btn-outline" href="album_form.asp?id=<%= lId %>">Editar</a>
            <form method="post" action="album_delete.asp" data-confirm="<%= sConfirm %>" style="display:inline;">
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
