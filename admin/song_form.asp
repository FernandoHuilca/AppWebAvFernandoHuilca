<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="../config/dbConnection.asp"-->
<!--#include file="../config/security.asp"-->
<%
' ============================================================================
'  admin/song_form.asp - Crear / Editar Cancion (CRUD)
'   GET  ?id=1  -> carga la cancion 1 para editar
'   GET  ?id=   -> formulario en blanco (nueva cancion)
'   POST        -> guarda (INSERT o UPDATE) y asigna artistas (TBL_PLAYER)
'  Todo con parametros ADO y transaccion para integridad.
' ============================================================================
RequireAdmin()

Dim sPageTitle, sActiveNav, sSubmitLabel, sBase
Dim bPost, iId, sName, sPath, sPlays, sError
Dim oConn, oCmd, oRs, sSql, bOk
Dim dSel, sArtistParam, aSelIds, iSel, lTmp, sTmp
Dim vArt, iArtCount, iArtRow, lArtId, sArtName, sArtImg, sChecked, sArtistsHtml
Dim sAudioOpts, oFs, oFolder, oFile

sBase = GetAppBaseUrl()
sPageTitle = "PoliMusic - Nueva Cancion"
sActiveNav = "songs"
sError = ""
bPost = (UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST")
iId = 0
sName = ""
sPath = ""
sPlays = ""
sArtistsHtml = ""
sAudioOpts = ""

Set dSel = Server.CreateObject("Scripting.Dictionary")

On Error Resume Next
Set oConn = OpenConnection()
If Err.Number <> 0 Then
    ShowErrorPage "No hay conexi&oacute;n con SQL Server. Revisa config/dbConnection.asp."
End If
On Error GoTo 0

If bPost Then
    iId = ToLong(Request.Form("id"))
    sName = DbText(Request.Form("song_name"), 50)
    sPath = DbText(Request.Form("song_path"), 255)
    sPlays = DbText(Request.Form("plays"), 10)
    sArtistParam = Trim("" & Request.Form("artIds"))

    If Len(sName) = 0 Then
        sError = "El nombre de la cancion es obligatorio (maximo 50 caracteres)."
    ElseIf Len(sPath) = 0 Then
        sError = "La ruta del archivo de audio es obligatoria."
    ElseIf Len(sPlays) > 0 And Not IsNumeric(sPlays) Then
        sError = "El campo reproducciones debe ser un numero entero."
    ElseIf Len(sPlays) > 0 And (CDbl(sPlays) < 0 Or CDbl(sPlays) > 999999999) Then
        sError = "El campo reproducciones debe estar entre 0 y 999999999."
    Else
        On Error Resume Next
        bOk = True
        oConn.BeginTrans
        If Err.Number <> 0 Then bOk = False

        If bOk And iId > 0 Then
            Set oCmd = NewCommand(oConn, _
                "UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?, PLAYS = ? WHERE ID_SONG = ?")
            AddVarCharParam oCmd, "@nombre", sName, 50
            AddVarCharParam oCmd, "@ruta", sPath, 255
            AddNullableIntParam oCmd, "@plays", sPlays
            AddIntParam oCmd, "@id", iId
            oCmd.Execute
            If Err.Number <> 0 Then bOk = False
        ElseIf bOk And iId <= 0 Then
            sSql = "SET NOCOUNT ON; INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, ?); " & _
                   "SELECT CAST(SCOPE_IDENTITY() AS INT) AS NEW_ID;"
            Set oCmd = NewCommand(oConn, sSql)
            AddVarCharParam oCmd, "@nombre", sName, 50
            AddVarCharParam oCmd, "@ruta", sPath, 255
            AddNullableIntParam oCmd, "@plays", sPlays
            Set oRs = oCmd.Execute
            If Err.Number <> 0 Then
                bOk = False
            ElseIf Not (oRs Is Nothing) And Not oRs.EOF Then
                ' obtenemos el id autogenerado por la BD
                iId = CLng(oRs("NEW_ID"))
            Else
                bOk = False
            End If
            CloseRs oRs
        End If

        ' Reasignar artistas de la cancion (borrar e insertar seleccionados)
        If bOk Then
            If ExecNonQueryId(oConn, "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", iId) <> 0 Then bOk = False
        End If

        If bOk And Len(sArtistParam) > 0 Then
            aSelIds = Split(sArtistParam, ",")
            For iSel = 0 To UBound(aSelIds)
                lTmp = ToLong(Trim("" & aSelIds(iSel)))
                If lTmp > 0 Then
                    If ExecNonQuery2Id(oConn, "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)", iId, lTmp) <> 0 Then
                        bOk = False
                        Exit For
                    End If
                End If
            Next
        End If

        If bOk Then
            oConn.CommitTrans
        Else
            oConn.RollbackTrans
        End If
        On Error GoTo 0

        If bOk Then
            CloseConn oConn
            Response.Redirect GetAppBaseUrl() & "admin/songs.asp?msg=saved"
            Response.End
        End If

        ' La conexion sigue abierta: se reutiliza para volver a llenar el formulario
        sError = "No se pudieron guardar los cambios. Revisa que los artistas existan y los datos sean v&aacute;lidos."
    End If
End If

If Not bPost Then
    iId = ToLong(Request.QueryString("id"))
    If iId > 0 Then
        sPageTitle = "PoliMusic - Editar Cancion"

        Set oRs = OpenRsId(oConn, "SELECT SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG WHERE ID_SONG = ?", iId)
        If oRs Is Nothing Then
            CloseConn oConn
            ShowErrorPage "No se pudo leer la canci&oacute;n solicitada."
        End If
        If oRs.EOF Then
            CloseRs oRs
            CloseConn oConn
            ShowErrorPage "La canci&oacute;n solicitada no existe en la base de datos."
        End If
        sName = Trim(Nz(oRs("SONG_NAME")))
        sPath = Trim(Nz(oRs("SONG_PATH")))
        If IsNull(oRs("PLAYS")) Then
            sPlays = ""
        Else
            sPlays = PlaysText(oRs("PLAYS"))
        End If
        CloseRs oRs

        Set oRs = OpenRsId(oConn, "SELECT ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = ?", iId)
        If Not (oRs Is Nothing) Then
            Do While Not oRs.EOF
                If Not dSel.Exists(CStr(NzLong(oRs("ID_ARTIST")))) Then
                    dSel.Add CStr(NzLong(oRs("ID_ARTIST"))), True
                End If
                oRs.MoveNext
            Loop
            CloseRs oRs
        End If
    Else
        sPageTitle = "PoliMusic - Nueva Cancion"
    End If
Else
    ' re-marcar los artistas seleccionados tras un error de validacion
    aSelIds = Split(sArtistParam, ",")
    For iSel = 0 To UBound(aSelIds)
        lTmp = ToLong(Trim("" & aSelIds(iSel)))
        If lTmp > 0 Then
            If Not dSel.Exists(CStr(lTmp)) Then dSel.Add CStr(lTmp), True
        End If
    Next
End If

' Lista de todos los artistas (checkboxes)
Set oRs = OpenRs(oConn, "SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST ORDER BY ARTIST_NAME, ARTIST_LASTNAME")
If oRs Is Nothing Then
    CloseConn oConn
    ShowErrorPage "No se pudo leer la tabla de artistas."
End If
vArt = Empty
If Not oRs.EOF Then vArt = oRs.GetRows()
CloseRs oRs
iArtCount = 0
If IsArray(vArt) Then iArtCount = UBound(vArt, 2) + 1
CloseConn oConn

sArtistsHtml = ""
For iArtRow = 0 To iArtCount - 1
    lArtId = NzLong(vArt(0, iArtRow))
    sArtName = Trim(Nz(vArt(1, iArtRow)))
    sTmp = Trim(Nz(vArt(2, iArtRow)))
    If Len(sTmp) > 0 Then sArtName = sArtName & " " & sTmp
    sArtImg = AssetUrl(vArt(3, iArtRow))
    sChecked = ""
    If dSel.Exists(CStr(lArtId)) Then sChecked = " checked"
    sArtistsHtml = sArtistsHtml & _
        "<label class='artist-check'>" & _
        "<input type='checkbox' name='artIds' value='" & lArtId & "'" & sChecked & ">" & _
        "<img src='" & H(sArtImg) & "' alt='' onerror='this.onerror=null;this.src=&quot;" & H(PlaceholderUrl()) & "&quot;;'>" & _
        "<span>" & H(sArtName) & "</span></label>"
Next

' Datalist con los mp3 disponibles en /SongFiles
On Error Resume Next
Set oFs = Server.CreateObject("Scripting.FileSystemObject")
Set oFolder = oFs.GetFolder(GetAppPhysicalRoot() & "SongFiles")
On Error GoTo 0
If Not (oFolder Is Nothing) Then
    For Each oFile In oFolder.Files
        If LCase(Right(oFile.Name, 4)) = ".mp3" Then
            sAudioOpts = sAudioOpts & "<option value=""~/songFiles/" & H(oFile.Name) & """>" & H(oFile.Name) & "</option>"
        End If
    Next
End If
If Len(sPath) > 0 Then
    sAudioOpts = "<option value=""" & H(sPath) & """>" & H(sPath) & "</option>" & sAudioOpts
End If

If iId > 0 Then
    sSubmitLabel = "Guardar cambios"
Else
    sSubmitLabel = "Crear canci&oacute;n"
End If
%>
<!--#include file="../includes/header.asp"-->

<header class="page-head">
  <h1><% If iId > 0 Then %>Editar canci&oacute;n<% Else %>Nueva canci&oacute;n<% End If %></h1>
  <a class="btn btn-ghost" href="songs.asp">&larr; Volver a Canciones</a>
</header>

<% If Len(sError) > 0 Then %>
<div class="alert alert-error"><%= sError %></div>
<% End If %>

<form class="form-card" method="post" action="song_form.asp">
  <h2 class="form-title"><% If iId > 0 Then %>Formulario de edici&oacute;n<% Else %>Formulario de registro<% End If %></h2>
  <input type="hidden" name="id" value="<%= iId %>">
  <div class="form-grid">
    <label class="field">
      <span>Nombre de la canci&oacute;n (SONG_NAME) *</span>
      <input type="text" name="song_name" maxlength="50" required value="<%= H(sName) %>" placeholder="ej. Adventure">
    </label>

    <label class="field">
      <span>Archivo de audio (SONG_PATH) *</span>
      <input type="text" name="song_path" maxlength="255" required list="audioList" value="<%= H(sPath) %>" placeholder="~/songFiles/bensound-adventure.mp3">
      <datalist id="audioList"><%= sAudioOpts %></datalist>
      <small class="hint">Elige un mp3 de la carpeta <b>SongFiles</b> escribi&eacute;ndolo o seleccion&aacute;ndolo. Las rutas se guardan como &ldquo;~/songFiles/archivo.mp3&rdquo;.</small>
    </label>

    <label class="field">
      <span>Reproducciones (PLAYS)</span>
      <input type="number" name="plays" min="0" max="999999999" step="1" value="<%= H(sPlays) %>" placeholder="0">
      <small class="hint">Deja el campo vac&iacute;o para guardar NULL.</small>
    </label>

    <div class="field">
      <span>Artistas (TBL_PLAYER)</span>
      <% If iArtCount > 0 Then %>
        <div class="artist-checks"><%= sArtistsHtml %></div>
      <% Else %>
        <span class="no-artists">No hay artistas registrados todav&iacute;a. Crea artistas en la secci&oacute;n correspondiente.</span>
      <% End If %>
      <small class="hint">Puedes marcar varios. Los cambios se aplican al guardar.</small>
    </div>
  </div>

  <div class="form-actions">
    <button class="btn btn-primary" type="submit"><%= sSubmitLabel %></button>
    <a class="btn btn-ghost" href="songs.asp">Cancelar</a>
  </div>
  <p class="form-note">Toda la operaci&oacute;n se guarda en una transacci&oacute;n y los valores viajan por <b>par&aacute;metros ADO</b> (protecci&oacute;n contra inyecci&oacute;n SQL).</p>
</form>

<script src="<%= H(GetAppBaseUrl()) %>js/admin.js" defer></script>
<!--#include file="../includes/footer.asp"-->
