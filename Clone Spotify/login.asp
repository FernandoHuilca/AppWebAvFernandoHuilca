<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="config/dbConnection.asp"-->
<!--#include file="config/security.asp"-->
<%
' ============================================================================
'  login.asp - Pantalla de acceso (roles: 0=Usuario, 1=Administrador)
'  Valida contra TBL_USER con parametros ADO (protege contra SQL Injection).
'  Acepta contrasenas planas o su hash SHA-1 en mayusculas (HASHBYTES).
' ============================================================================

Dim sError, sInfo, sUsername, sPassword
Dim oConn, oCmd, oRs
Dim lIdUser, iUserType, sUserPhoto

' Si ya hay sesion, ir directo al dashboard
If IsLoggedIn() Then
    Response.Redirect GetAppBaseUrl() & "index.asp"
    Response.End
End If

sError = ""
sInfo = ""
sUsername = ""

Select Case LCase(Trim("" & Request.QueryString("msg")))
    Case "logout"
        sInfo = "Sesi&oacute;n finalizada correctamente. &iexcl;Hasta pronto!"
    Case "expired"
        sInfo = "Tu sesi&oacute;n expir&oacute;. Vuelve a iniciar sesi&oacute;n."
End Select

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    sUsername = DbText(Request.Form("username"), 50)
    sPassword = DbText(Request.Form("password"), 50)

    If Len(sUsername) = 0 Or Len(sPassword) = 0 Then
        sError = "Debes escribir usuario y contrase&ntilde;a."
    Else
        On Error Resume Next
        Err.Clear
        Set oConn = OpenConnection()
        If Err.Number = 0 Then
            Set oCmd = NewCommand(oConn, _
                "SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, USER_PHOTO " & _
                "FROM TBL_USER " & _
                "WHERE USERNAME = ? " & _
                "  AND (PASSWORD = ? OR PASSWORD = CONVERT(VARCHAR(40), HASHBYTES('SHA1', CONVERT(VARCHAR(50), ?)), 2))")
            AddVarCharParam oCmd, "@usuario", sUsername, 50
            AddVarCharParam oCmd, "@clave", sPassword, 50
            AddVarCharParam oCmd, "@claveHash", sPassword, 50
            Set oRs = oCmd.Execute
            If Err.Number <> 0 Then
                sError = "No hay conexi&oacute;n con la base de datos. Verifica SQL Server y config/dbConnection.asp."
            ElseIf oRs.EOF Then
                sError = "Usuario o contrase&ntilde;a incorrectos."
            Else
                lIdUser = CLng(oRs("ID_USER"))
                iUserType = CLng(oRs("USER_TYPE"))
                sUserPhoto = AssetUrl(oRs("USER_PHOTO"))
                Session.TimeOut = 60
                Session("UserID") = CStr(lIdUser)
                Session("Username") = CStr(oRs("USERNAME"))
                Session("Email") = CStr(oRs("EMAIL"))
                Session("UserType") = iUserType
                Session("UserPhotoUrl") = sUserPhoto
                CloseRs oRs
                CloseConn oConn
                Response.Redirect GetAppBaseUrl() & "index.asp"
                Response.End
            End If
            CloseRs oRs
            CloseConn oConn
        Else
            sError = "No hay conexi&oacute;n con la base de datos. Verifica SQL Server y config/dbConnection.asp."
        End If
        On Error GoTo 0
    End If
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PoliMusic - Iniciar sesi&oacute;n</title>
<link rel="icon" type="image/svg+xml" href="<%= H(GetAppBaseUrl()) %>assets/favicon.svg">
<link rel="stylesheet" href="<%= H(GetAppBaseUrl()) %>css/login.css">
</head>
<body class="login-body">
<main class="login-split">
  <!-- Panel de marca con estetica pastel -->
  <section class="brand-panel" aria-hidden="true">
    <div class="brand-blob blob-1"></div>
    <div class="brand-blob blob-2"></div>
    <div class="brand-blob blob-3"></div>
    <div class="brand-content">
      <div class="brand-logo">
        <svg width="72" height="72" viewBox="0 0 64 64">
          <circle cx="32" cy="32" r="28" fill="#332E45"/>
          <circle cx="32" cy="32" r="16" fill="none" stroke="#CDB9F5" stroke-width="2"/>
          <circle cx="32" cy="32" r="6" fill="#BFE8D8"/>
          <text x="44" y="50" font-size="20" font-family="Arial, sans-serif" fill="#332E45">&#9834;</text>
        </svg>
      </div>
      <h1 class="brand-title">Poli<span>Music</span></h1>
      <p class="brand-tagline">Tu m&uacute;sica, tus playlists, tus reglas.</p>
      <ul class="brand-features">
        <li>&#127911; Reproductor con lista completa</li>
        <li>&#127908; Gesti&oacute;n de canciones, artistas y &aacute;lbumes</li>
        <li>&#127744; Interfaz pastel moderna y ligera</li>
      </ul>
    </div>
  </section>

  <!-- Formulario -->
  <section class="form-panel">
    <form class="login-card" method="post" action="login.asp" autocomplete="on">
      <div class="login-card-head">
        <span class="brand-mini">Poli<b>Music</b></span>
        <h2>Bienvenido</h2>
        <p>Ingresa con tu cuenta para escuchar tu m&uacute;sica.</p>
      </div>

      <% If Len(sInfo) > 0 Then %>
      <div class="alert alert-ok" role="status"><%= sInfo %></div>
      <% End If %>
      <% If Len(sError) > 0 Then %>
      <div class="alert alert-error" role="alert"><%= sError %></div>
      <% End If %>

      <label class="field">
        <span>Usuario</span>
        <input type="text" name="username" maxlength="50" required autofocus
               placeholder="ej. userNormal" value="<%= H(sUsername) %>" autocomplete="username">
      </label>

      <label class="field">
        <span>Contrase&ntilde;a</span>
        <input type="password" name="password" maxlength="50" required
               placeholder="&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull;" autocomplete="current-password">
      </label>

      <button type="submit" class="btn btn-primary btn-block">Entrar</button>

      <details class="demo-creds">
        <summary>Cuentas de prueba</summary>
        <ul>
          <li><code>userAdmin</code> / <code>12345678</code> &mdash; rol ADMIN</li>
          <li><code>userNormal</code> / <code>12345678</code> &mdash; rol USER</li>
          <li><code>username1</code> / <code>12345678</code> &mdash; clave guardada como SHA-1</li>
        </ul>
      </details>
    </form>
  </section>
</main>
</body>
</html>
