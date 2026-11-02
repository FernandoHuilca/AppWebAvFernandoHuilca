<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../config/security.asp"-->
<%
' ============================================================================
'  admin/album_form.asp - Crear / Editar Album (CRUD)
' ============================================================================
RequireAdmin()

Dim sPageTitle, sActiveNav, sSubmitLabel, sBase
Dim bPost, iId, sTitle, sYear, sImage, sError
Dim oConn, oCmd, oRs, bOk
Dim sImageOpts, oFs, oFolder, oFile

sBase = GetAppBaseUrl()
sPageTitle = "PoliMusic - Nuevo Album"
sActiveNav = "albums"
sError = ""
bPost = (UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST")
iId = 0
sTitle = ""
sYear = ""
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
    sTitle = DbText(Request.Form("title_album"), 50)
    sYear = DbText(Request.Form("release_year"), 4)
    sImage = DbText(Request.Form("album_image"), 255)

    If Len(sTitle) = 0 Then
        sError = "El titulo del album es obligatorio (maximo 50 caracteres)."
    ElseIf Len(sYear) > 0 And Not IsNumeric(sYear) Then
        sError = "El ano debe ser un numero de 4 digitos."
    ElseIf Len(sYear) > 0 And (CLng(sYear) < 1000 Or CLng(sYear) > 9999) Then
        sError = "El ano debe ser un numero de 4 digitos (ejemplo: 1980)."
    Else
        On Error Resume Next
        bOk = True
        If iId > 0 Then
            Set oCmd = NewCommand(oConn, _
                "UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?")
            AddVarCharParam oCmd, "@titulo", sTitle, 50
            AddNullableIntParam oCmd, "@anio", sYear
            AddVarCharParam oCmd, "@imagen", sImage, 255
            AddIntParam oCmd, "@id", iId
            oCmd.Execute
            If Err.Number <> 0 Then bOk = False
        Else
            Set oCmd = NewCommand(oConn, _
                "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (?, ?, ?)")
            AddVarCharParam oCmd, "@titulo", sTitle, 50
            AddNullableIntParam oCmd, "@anio", sYear
            AddVarCharParam oCmd, "@imagen", sImage, 255
            oCmd.Execute
            If Err.Number <> 0 Then bOk = False
        End If
        CloseConn oConn
        On Error GoTo 0

        If bOk Then
            Response.Redirect GetAppBaseUrl() & "admin/albums.asp?msg=saved"
            Response.End
        Else
            sError = "No se pudo guardar el album. Revisa los datos e intenta de nuevo."
        End If
    End If
End If

If Not bPost Then
    iId = ToLong(Request.QueryString("id"))
    If iId > 0 Then
        sPageTitle = "PoliMusic - Editar Album"

        Set oRs = OpenRsId(oConn, "SELECT TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", iId)
        If oRs Is Nothing Then
            CloseConn oConn
            ShowErrorPage "No se pudo leer el album solicitado."
        End If
        If oRs.EOF Then
            CloseRs oRs
            CloseConn oConn
            ShowErrorPage "El album solicitado no existe en la base de datos."
        End If
        sTitle = Trim(Nz(oRs("TITLE_ALBUM")))
        If IsNull(oRs("RELEASE_YEAR")) Then
            sYear = ""
        Else
            sYear = CStr(NzLong(oRs("RELEASE_YEAR")))
        End If
        sImage = Trim(Nz(oRs("ALBUM_IMAGE_PATH")))
        CloseRs oRs
    End If
End If
CloseConn oConn

' Datalist con imagenes disponibles en /AlbumImages
On Error Resume Next
Set oFs = Server.CreateObject("Scripting.FileSystemObject")
Set oFolder = oFs.GetFolder(GetAppPhysicalRoot() & "AlbumImages")
On Error GoTo 0
If Not (oFolder Is Nothing) Then
    For Each oFile In oFolder.Files
        If LCase(Right(oFile.Name, 4)) = ".jpg" Or LCase(Right(oFile.Name, 4)) = ".png" Or LCase(Right(oFile.Name, 5)) = ".jpeg" Then
            sImageOpts = sImageOpts & "<option value=""~/albumImages/" & H(oFile.Name) & """>" & H(oFile.Name) & "</option>"
        End If
    Next
End If
If Len(sImage) > 0 Then
    sImageOpts = "<option value=""" & H(sImage) & """>" & H(sImage) & "</option>" & sImageOpts
End If

If iId > 0 Then
    sSubmitLabel = "Guardar cambios"
Else
    sSubmitLabel = "Crear album"
End If
%>
<!--#include file="../includes/header.asp"-->

<header class="page-head">
  <h1><% If iId > 0 Then %>Editar &aacute;lbum<% Else %>Nuevo &aacute;lbum<% End If %></h1>
  <a class="btn btn-ghost" href="albums.asp">&larr; Volver a &Aacute;lbumes</a>
</header>

<% If Len(sError) > 0 Then %>
<div class="alert alert-error"><%= sError %></div>
<% End If %>

<form class="form-card" method="post" action="album_form.asp">
  <h2 class="form-title"><% If iId > 0 Then %>Formulario de edici&oacute;n<% Else %>Formulario de registro<% End If %></h2>
  <input type="hidden" name="id" value="<%= iId %>">
  <div class="form-grid">
    <label class="field">
      <span>T&iacute;tulo (TITLE_ALBUM) *</span>
      <input type="text" name="title_album" maxlength="50" required value="<%= H(sTitle) %>" placeholder="ej. Rock Album">
    </label>

    <label class="field">
      <span>A&ntilde;o (RELEASE_YEAR)</span>
      <input type="number" name="release_year" min="1000" max="9999" step="1" value="<%= H(sYear) %>" placeholder="ej. 1980">
      <small class="hint">Campo opcional. Deja vac&iacute;o para guardar NULL.</small>
    </label>

    <label class="field">
      <span>Portada (ALBUM_IMAGE_PATH)</span>
      <input type="text" name="album_image" maxlength="255" list="imgList" value="<%= H(sImage) %>" placeholder="~/albumImages/album1.png">
      <datalist id="imgList"><%= sImageOpts %></datalist>
      <small class="hint">Elige una imagen de la carpeta <b>AlbumImages</b> (si no existe a&uacute;n se mostrar&aacute; placeholder).</small>
    </label>
  </div>

  <div class="form-actions">
    <button class="btn btn-primary" type="submit"><%= sSubmitLabel %></button>
    <a class="btn btn-ghost" href="albums.asp">Cancelar</a>
  </div>
  <p class="form-note">Los valores viajan por <b>par&aacute;metros ADO</b> (protecci&oacute;n contra inyecci&oacute;n SQL).</p>
</form>

<script src="<%= H(GetAppBaseUrl()) %>js/admin.js" defer></script>
<!--#include file="../includes/footer.asp"-->
