<%
' ============================================================================
'  includes/header.asp
'  Incluye el <head>, la sidebar y abre <main>.
'  La pagina que lo incluye DEBE definir antes:
'     Dim sPageTitle, sActiveNav   ("home" | "songs" | "artists" | "albums")
' ============================================================================
Dim hdrBase, hdrAvatar, hdrUser
hdrBase = GetAppBaseUrl()
hdrUser = "" & Session("Username")
hdrAvatar = "" & Session("UserPhotoUrl")
If Len(Trim(hdrAvatar)) = 0 Then hdrAvatar = hdrBase & "assets/placeholder.svg"
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><%= H(sPageTitle) %></title>
<link rel="icon" type="image/svg+xml" href="<%= H(hdrBase) %>assets/favicon.svg">
<link rel="stylesheet" href="<%= H(hdrBase) %>css/styles.css">
</head>
<body>
<a class="skip-link" href="#main">Saltar al contenido</a>
<div class="app-shell">

  <aside class="sidebar" aria-label="Men&uacute; principal">
    <a class="brand" href="<%= H(hdrBase) %>index.asp">
      <span class="brand-disc" aria-hidden="true">
        <svg width="34" height="34" viewBox="0 0 34 34"><circle cx="17" cy="17" r="15" fill="#332E45"/><circle cx="17" cy="17" r="9" fill="none" stroke="#8B79CC" stroke-width="1.6"/><circle cx="17" cy="17" r="4" fill="#F7F5FC"/></svg>
      </span>
      <span class="brand-name">Poli<span class="brand-accent">Music</span></span>
    </a>

    <nav class="side-nav" aria-label="Secciones">
      <span class="side-label">Men&uacute;</span>
      <a class="side-link<% If sActiveNav = "home" Then Response.Write " is-active" %>" href="<%= H(hdrBase) %>index.asp">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 11.5 12 5l8 6.5V20h-5.5v-5h-5v5H4z"/></svg>
        <span>Inicio</span>
      </a>
      <a class="side-link" href="<%= H(hdrBase) %>index.asp#playlists">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><path d="M4 6h13M4 10h9M4 14h7"/><circle cx="17" cy="15" r="2.6"/><path d="M19.6 15V8.4l2 .7"/></svg>
        <span>Mis playlists</span>
      </a>

      <% If IsAdmin() Then %>
      <span class="side-label">Administraci&oacute;n</span>
      <a class="side-link<% If sActiveNav = "songs" Then Response.Write " is-active" %>" href="<%= H(hdrBase) %>admin/songs.asp">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><path d="M9.5 17.5V7l9-2v10.3"/><circle cx="7" cy="17.5" r="2.5"/><circle cx="16" cy="15.3" r="2.5"/></svg>
        <span>Canciones</span>
      </a>
      <a class="side-link<% If sActiveNav = "artists" Then Response.Write " is-active" %>" href="<%= H(hdrBase) %>admin/artists.asp">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><circle cx="9" cy="8.5" r="3.2"/><path d="M3.5 19c.7-3 2.9-4.6 5.5-4.6s4.8 1.6 5.5 4.6"/><circle cx="16.8" cy="9.5" r="2.5"/><path d="M15.8 18.8c.4-2 1.9-3.2 3.7-3.2.9 0 1.7.3 2.4.9"/></svg>
        <span>Artistas</span>
      </a>
      <a class="side-link<% If sActiveNav = "albums" Then Response.Write " is-active" %>" href="<%= H(hdrBase) %>admin/albums.asp">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><circle cx="12" cy="12" r="8.2"/><circle cx="12" cy="12" r="2.2"/><path d="M12 3.8v4" stroke-linecap="round"/></svg>
        <span>&Aacute;lbumes</span>
      </a>
      <% End If %>
    </nav>

    <div class="side-user">
      <img class="avatar" src="<%= H(hdrAvatar) %>" alt="Foto de <%= H(hdrUser) %>"
           onerror="this.onerror=null;this.src='<%= H(hdrBase) %>assets/placeholder.svg';">
      <div class="side-user-info">
        <span class="side-username"><%= H(hdrUser) %></span>
        <span class="chip<% If IsAdmin() Then Response.Write " chip-admin" %>"><%= RoleLabel() %></span>
      </div>
      <a class="btn btn-logout" href="<%= H(hdrBase) %>logout.asp" title="Cerrar sesi&oacute;n">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 5H6v14h8"/><path d="M11 12h9M17 8.5 20.5 12 17 15.5"/></svg>
        <span>Salir</span>
      </a>
    </div>
  </aside>

  <main class="main-content" id="main">
