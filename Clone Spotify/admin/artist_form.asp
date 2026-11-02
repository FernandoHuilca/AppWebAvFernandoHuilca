<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../config/security.asp"-->
<%
' ============================================================================
'  admin/artist_form.asp - Crear / Editar Artista (CRUD)
' ============================================================================
RequireAdmin()

Dim sPageTitle, sActiveNav, sSubmitLabel, sBase
Dim bPost, iId, sName, sLast, sImage, sError
Dim oConn, oCmd, oRs, bOk
Dim sImageOpts, oFs, oFolder, oFile

sBase = GetAppBaseUrl()
sPageTitle = "PoliMusic - Nuevo Artista"
sActiveNav = "artists"
sError = ""
bPost = (UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST")
iId = 0
sName = ""
sLast = ""
sImage = ""
sImageOpts = ""

On Error Resume Next
Set oConn = OpenConnection()
If Err.Number <> 0 Then
    ShowErrorPage "No hay conexi&oacute;n con SQL Server. Revisa config/dbConnection.asp."
End If
On Error GoTo 0

If bPost Then
    iId = ToLong(Request.Form("id"))
    sName = DbText(Request.Form("artist_name"), 100)
    sLast = DbText(Request.Form("artist_lastname"), 100)
    sImage = DbText(Request.Form("artist_image"), 255)

    If Len(sName) = 0 Then
        sError = "El nombre del artista es obligatorio (maximo 100 caracteres)."
    Else
        On Error Resume Next
        bOk = True
        If iId > 0 Then
            Set oCmd = NewCommand(oConn, _
                "UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = ?")
            AddVarCharParam oCmd, "@nombre", sName, 100
            AddVarCharParam oCmd, "@apellido", sLast, 100
            AddVarCharParam oCmd, "@foto", sImage, 255
            AddIntParam oCmd, "@id", iId
            oCmd.Execute
            If Err.Number <> 0 Then bOk = False
        Else
            Set oCmd = NewCommand(oConn, _
                "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (?, ?, ?)")
            AddVarCharParam oCmd, "@nombre", sName, 100
            AddVarCharParam oCmd, "@apellido", sLast, 100
            AddVarCharParam oCmd, "@foto", sImage, 255
            oCmd.Execute
            If Err.Number <> 0 Then bOk = False
        End If
        CloseConn oConn
        On Error GoTo 0

        If bOk Then
            Response.Redirect GetAppBaseUrl() & "admin/artists.asp?msg=saved"
            Response.End
        Else
            sError = "No se pudo guardar el artista. Revisa los datos e intenta de nuevo."
        End If
    End If
End If

If Not bPost Then
    iId = ToLong(Request.QueryString("id"))
    If iId > 0 Then
        sPageTitle = "PoliMusic - Editar Artista"

        Set oRs = OpenRsId(oConn, "SELECT ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", iId)
        If oRs Is Nothing Then
            CloseConn oConn
            ShowErrorPage "No se pudo leer el artista solicitado."
        End If
        If oRs.EOF Then
            CloseRs oRs
            CloseConn oConn
            ShowErrorPage "El artista solicitado no existe en la base de datos."
        End If
        sName = Trim(Nz(oRs("ARTIST_NAME")))
        sLast = Trim(Nz(oRs("ARTIST_LASTNAME")))
        sImage = Trim(Nz(oRs("ARTIST_IMAGE")))
        CloseRs oRs
    End If
End If
CloseConn oConn

' Datalist con imagenes disponibles en /ArtistImages
On Error Resume Next
Set oFs = Server.CreateObject("Scripting.FileSystemObject")
Set oFolder = oFs.GetFolder(GetAppPhysicalRoot() & "ArtistImages")
On Error GoTo 0
If Not (oFolder Is Nothing) Then
    For Each oFile In oFolder.Files
        If LCase(Right(oFile.Name, 4)) = ".jpg" Or LCase(Right(oFile.Name, 4)) = ".png" Or LCase(Right(oFile.Name, 5)) = ".jpeg" Then
            sImageOpts = sImageOpts & "<option value=""~/ArtistImages/" & H(oFile.Name) & """>" & H(oFile.Name) & "</option>"
        End If
    Next
End If
If Len(sImage) > 0 Then
    sImageOpts = "<option value=""" & H(sImage) & """>" & H(sImage) & "</option>" & sImageOpts
End If

If iId > 0 Then
    sSubmitLabel = "Guardar cambios"
Else
    sSubmitLabel = "Crear artista"
End If
%>
<!--#include file="../includes/header.asp"-->

<header class="page-head">
  <h1><% If iId > 0 Then %>Editar artista<% Else %>Nuevo artista<% End If %></h1>
  <a class="btn btn-ghost" href="artists.asp">&larr; Volver a Artistas</a>
</header>

<% If Len(sError) > 0 Then %>
<div class="alert alert-error"><%= sError %></div>
<% End If %>

<form class="form-card" method="post" action="artist_form.asp">
  <h2 class="form-title"><% If iId > 0 Then %>Formulario de edici&oacute;n<% Else %>Formulario de registro<% End If %></h2>
  <input type="hidden" name="id" value="<%= iId %>">
  <div class="form-grid">
    <label class="field">
      <span>Nombre (ARTIST_NAME) *</span>
      <input type="text" name="artist_name" maxlength="100" required value="<%= H(sName) %>" placeholder="ej. Juan">
    </label>

    <label class="field">
      <span>Apellido (ARTIST_LASTNAME)</span>
      <input type="text" name="artist_lastname" maxlength="100" value="<%= H(sLast) %>" placeholder="ej. Gabriel">
      <small class="hint">Campo opcional.</small>
    </label>

    <label class="field">
      <span>Foto (ARTIST_IMAGE)</span>
      <input type="text" name="artist_image" maxlength="255" list="imgList" value="<%= H(sImage) %>" placeholder="~/ArtistImages/artist1.jpg">
      <datalist id="imgList"><%= sImageOpts %></datalist>
      <small class="hint">Elige una imagen de la carpeta <b>ArtistImages</b> (si no existe a&uacute;n, se mostrar&aacute; un placeholder).</small>
    </label>
  </div>

  <div class="form-actions">
    <button class="btn btn-primary" type="submit"><%= sSubmitLabel %></button>
    <a class="btn btn-ghost" href="artists.asp">Cancelar</a>
  </div>
  <p class="form-note">Los valores viajan por <b>par&aacute;metros ADO</b> (protecci&oacute;n contra inyecci&oacute;n SQL).</p>
</form>

<script src="<%= H(GetAppBaseUrl()) %>js/admin.js" defer></script>
<!--#include file="../includes/footer.asp"-->
