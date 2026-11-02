<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="config/dbConnection.asp"-->
<!--#include file="config/security.asp"-->
<%
' ============================================================================
'  index.asp - Dashboard (Listado de Canciones + reproductor + playlists)
'  Requiere sesion activa. Los rutas de imagenes/audio se leen de la BD y se
'  resuelven a las carpetas reales de este proyecto (SongFiles, UserImages...).
' ============================================================================
RequireLogin()

Dim sPageTitle, sActiveNav
sPageTitle = "PoliMusic - Listado de Canciones"
sActiveNav = "home"

Dim oConn, oCmd, oRs, sSql
Dim iUserId, iPlaylist
Dim vPl, iPlCount, iPlRow
Dim dArt, dAlb, dPlFilter
Dim sPlaylistsHtml, bHasFilter, sFilterTitle
Dim vSongs, iSongs, iRow
Dim lTmp, sTmpKey, sTmp1, sTmp2
Dim lPlId, sPlTitle, sPlType, sPlPhoto, lPlCount, sPlActive
Dim lId, sName, sAudioUrl, lPlays, sArtText
Dim sAlbumTitle, sAlbumImg, sCoverUrl, iIdx, aAlbumParts
Dim sSongsJson

iUserId = ToLong(Session("UserID"))
iPlaylist = ToLong(Request.QueryString("pl"))
bHasFilter = False
sFilterTitle = ""
sSongsJson = "[]"
Set dArt = Server.CreateObject("Scripting.Dictionary")
Set dAlb = Server.CreateObject("Scripting.Dictionary")
Set dPlFilter = Server.CreateObject("Scripting.Dictionary")

On Error Resume Next
Set oConn = OpenConnection()
If Err.Number <> 0 Then
    ShowErrorPage "No hay conexi&oacute;n con SQL Server. Revisa que el servicio este iniciado y la cadena de conexion en <b>config/dbConnection.asp</b>."
End If

' ---------------------------------------------------------------- PLAYLISTS
sSql = "SELECT P.ID_PLAYLIST, P.TITLE, P.TYPE, P.PHOTO, " & _
       "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = P.ID_PLAYLIST) AS SONG_COUNT " & _
       "FROM TBL_PLAYLIST P WHERE P.ID_USER = ? ORDER BY P.TITLE, P.ID_PLAYLIST"
Set oRs = Nothing
Set oCmd = NewCommand(oConn, sSql)
If Err.Number = 0 Then
    AddIntParam oCmd, "idUser", iUserId
    Set oRs = oCmd.Execute
End If

If Not (oRs Is Nothing) Then
    If Not oRs.EOF Then vPl = oRs.GetRows()
    CloseRs oRs
End If
iPlCount = 0
If IsArray(vPl) Then iPlCount = UBound(vPl, 2) + 1

sPlaylistsHtml = ""
For iPlRow = 0 To iPlCount - 1
    lPlId = NzLong(vPl(0, iPlRow))
    sPlTitle = Nz(vPl(1, iPlRow))
    sPlType = Nz(vPl(2, iPlRow))
    sPlPhoto = AssetUrl(vPl(3, iPlRow))
    lPlCount = NzLong(vPl(4, iPlRow))
    sPlActive = ""
    If lPlId = iPlaylist Then sPlActive = " is-active"
    dPlFilter.Add CStr(lPlId), sPlTitle
    If ToLong(sPlType) = 1 Then
        sPlType = "Privada"
    Else
        sPlType = "P&uacute;blica"
    End If
    sPlaylistsHtml = sPlaylistsHtml & _
        "<a class='playlist-card" & sPlActive & "' href='" & H(GetAppBaseUrl()) & "index.asp?pl=" & lPlId & "'>" & _
        "<img class='playlist-photo' src='" & H(sPlPhoto) & "' alt='Portada de la playlist " & H(sPlTitle) & "' loading='lazy'" & _
        " onerror='this.onerror=null;this.src=&quot;" & H(PlaceholderUrl()) & "&quot;;'>" & _
        "<span class='playlist-name'>" & H(sPlTitle) & "</span>" & _
        "<span class='playlist-meta'><span class='chip chip-mini'>" & sPlType & "</span>" & _
        "<span>&#9654; " & lPlCount & " canciones</span></span></a>"
Next

' Validar el filtro solicitado por querystring
If iPlaylist > 0 Then
    If dPlFilter.Exists(CStr(iPlaylist)) Then
        bHasFilter = True
        sFilterTitle = dPlFilter(CStr(iPlaylist))
    End If
End If

' -------------------------------------------------------------- CANCIONES
Set oRs = Nothing
If bHasFilter Then
    sSql = "SELECT S.ID_SONG, S.SONG_NAME, S.SONG_PATH, S.PLAYS " & _
           "FROM TBL_SONG S INNER JOIN TBL_PLAYLIST_SONG PS ON PS.ID_SONG = S.ID_SONG " & _
           "AND PS.ID_PLAYLIST = ? ORDER BY S.SONG_NAME, S.ID_SONG"
    Set oCmd = NewCommand(oConn, sSql)
    If Err.Number = 0 Then
        AddIntParam oCmd, "idPlaylist", iPlaylist
        Set oRs = oCmd.Execute
    End If
Else
    Set oRs = OpenRs(oConn, "SELECT S.ID_SONG, S.SONG_NAME, S.SONG_PATH, S.PLAYS " & _
                            "FROM TBL_SONG S ORDER BY S.SONG_NAME, S.ID_SONG")
End If
vSongs = Empty
If Not (oRs Is Nothing) Then
    If Not oRs.EOF Then vSongs = oRs.GetRows()
    CloseRs oRs
End If
iSongs = 0
If IsArray(vSongs) Then iSongs = UBound(vSongs, 2) + 1

' ---------------------------------------------------- ARTISTAS POR CANCION
Set oRs = OpenRs(oConn, "SELECT P.ID_SONG, A.ARTIST_NAME, A.ARTIST_LASTNAME FROM TBL_PLAYER P " & _
                        "INNER JOIN TBL_ARTIST A ON A.ID_ARTIST = P.ID_ARTIST ORDER BY P.ID_SONG, A.ARTIST_NAME")
If Not (oRs Is Nothing) Then
    Do While Not oRs.EOF
        lTmp = NzLong(oRs("ID_SONG"))
        sTmp1 = Trim(Nz(oRs("ARTIST_NAME")))
        sTmp2 = Trim(Nz(oRs("ARTIST_LASTNAME")))
        If Len(sTmp2) > 0 Then sTmp1 = sTmp1 & " " & sTmp2
        sTmpKey = CStr(lTmp)
        If dArt.Exists(sTmpKey) Then
            dArt(sTmpKey) = dArt(sTmpKey) & ", " & sTmp1
        Else
            dArt.Add sTmpKey, sTmp1
        End If
        oRs.MoveNext
    Loop
    CloseRs oRs
End If

' ------------------------------------------------------- ALBUM POR CANCION
Set oRs = OpenRs(oConn, "SELECT T.ID_SONG, A.TITLE_ALBUM, A.ALBUM_IMAGE_PATH, A.RELEASE_YEAR " & _
                        "FROM TBL_TRACKLIST T INNER JOIN TBL_ALBUM A ON A.ID_ALBUM = T.ID_ALBUM " & _
                        "ORDER BY T.ID_SONG, A.ID_ALBUM")
If Not (oRs Is Nothing) Then
    Do While Not oRs.EOF
        lTmp = NzLong(oRs("ID_SONG"))
        If dAlb.Exists(CStr(lTmp)) Then
            ' primer album (ORDER BY ID_ALBUM)
        Else
            dAlb.Add CStr(lTmp), Nz(oRs("TITLE_ALBUM")) & Chr(1) & Nz(oRs("ALBUM_IMAGE_PATH")) & Chr(1) & CStr(NzLong(oRs("RELEASE_YEAR")))
        End If
        oRs.MoveNext
    Loop
    CloseRs oRs
End If

CloseConn oConn
On Error GoTo 0

' -------------------------------------------------------- JSON DEL PLAYER
sSongsJson = "["
For iRow = 0 To iSongs - 1
    lId = NzLong(vSongs(0, iRow))
    sName = Nz(vSongs(1, iRow))
    sAudioUrl = AudioUrlOrEmpty(Nz(vSongs(2, iRow)))
    lPlays = NzLong(vSongs(3, iRow))
    sTmpKey = CStr(lId)
    If dArt.Exists(sTmpKey) Then sArtText = dArt(sTmpKey) Else sArtText = ""
    If dAlb.Exists(sTmpKey) Then
        aAlbumParts = Split(dAlb(sTmpKey), Chr(1))
        sAlbumTitle = aAlbumParts(0)
        sAlbumImg = aAlbumParts(1)
    Else
        sAlbumTitle = ""
        sAlbumImg = ""
    End If
    sCoverUrl = AssetUrl(sAlbumImg)
    If iRow > 0 Then sSongsJson = sSongsJson & ","
    sSongsJson = sSongsJson & "{""id"":" & lId & _
        ",""t"":""" & JsonEncode(sName) & """" & _
        ",""u"":""" & JsonEncode(sAudioUrl) & """" & _
        ",""a"":""" & JsonEncode(sArtText) & """" & _
        ",""al"":""" & JsonEncode(sAlbumTitle) & """" & _
        ",""c"":""" & JsonEncode(sCoverUrl) & """" & _
        ",""p"":" & lPlays & "}"
Next
If iSongs > 0 Then sSongsJson = sSongsJson & "]"
%>
<!--#include file="includes/header.asp"-->

<header class="welcome-banner">
  <img class="welcome-avatar" src="<%= H(Session("UserPhotoUrl")) %>" alt=""
       onerror="this.onerror=null;this.src='<%= H(hdrBase) %>assets/placeholder.svg';">
  <div class="welcome-text">
    <span class="chip<% If IsAdmin() Then Response.Write " chip-admin" %>"><%= RoleLabel() %></span>
    <h1>&iexcl;Hola, <%= H(Session("Username")) %>!</h1>
    <p>Bienvenido a tu m&uacute;sica. Sigue bajando hasta el <a href="#songs">Listado de Canciones</a>.</p>
  </div>
  <div class="welcome-actions">
    <a class="btn btn-logout" href="<%= H(hdrBase) %>logout.asp">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 5H6v14h8"/><path d="M11 12h9M17 8.5 20.5 12 17 15.5"/></svg>
      <span>Cerrar sesi&oacute;n</span>
    </a>
  </div>
</header>

<% If IsAdmin() Then %>
<section class="admin-actions" aria-label="Panel de administraci&oacute;n">
  <div class="admin-actions-head">
    <h2>Panel de administraci&oacute;n</h2>
    <span class="chip chip-admin">Modo ADMIN</span>
  </div>
  <div class="admin-actions-grid">
    <a class="admin-action action-songs" href="<%= H(hdrBase) %>admin/songs.asp">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><path d="M9.5 17.5V7l9-2v10.3"/><circle cx="7" cy="17.5" r="2.5"/><circle cx="16" cy="15.3" r="2.5"/></svg>
      <span class="admin-action-title">Canciones</span>
      <span class="admin-action-sub">Crear, editar y eliminar</span>
    </a>
    <a class="admin-action action-artists" href="<%= H(hdrBase) %>admin/artists.asp">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><circle cx="9" cy="8.5" r="3.2"/><path d="M3.5 19c.7-3 2.9-4.6 5.5-4.6s4.8 1.6 5.5 4.6"/><circle cx="16.8" cy="9.5" r="2.5"/><path d="M15.8 18.8c.4-2 1.9-3.2 3.7-3.2.9 0 1.7.3 2.4.9"/></svg>
      <span class="admin-action-title">Artistas</span>
      <span class="admin-action-sub">Fichas e im&aacute;genes</span>
    </a>
    <a class="admin-action action-albums" href="<%= H(hdrBase) %>admin/albums.asp">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><circle cx="12" cy="12" r="8.2"/><circle cx="12" cy="12" r="2.2"/><path d="M12 3.8v4" stroke-linecap="round"/></svg>
      <span class="admin-action-title">&Aacute;lbumes</span>
      <span class="admin-action-sub">Tracklists y portadas</span>
    </a>
  </div>
</section>
<% End If %>

<% If iPlCount > 0 Then %>
<section class="playlists-section" id="playlists" aria-label="Tus playlists">
  <div class="playlists-head">
    <h2 class="section-title">Tus playlists</h2>
    <% If bHasFilter Then %>
    <a class="chip chip-link" href="<%= H(hdrBase) %>index.asp">&#10003; Ver todas las canciones</a>
    <% End If %>
  </div>
  <div class="playlists-row"><%= sPlaylistsHtml %></div>
</section>
<% End If %>

<section class="songs-section" id="songs" aria-label="Listado de canciones">
  <div class="songs-head">
    <h2 class="section-title">
      <% If bHasFilter Then %>
      Listado de Canciones de &laquo;<%= H(sFilterTitle) %>&raquo;
      <% Else %>
      Listado de Canciones
      <% End If %>
      <span class="chip"><%= iSongs %><% If iSongs = 1 Then %> canci&oacute;n<% Else %>&nbsp;canciones<% End If %></span>
    </h2>
    <input class="search-box" type="search" id="songSearch" placeholder="Buscar canci&oacute;n, artista o &aacute;lbum&hellip;" aria-label="Buscar canciones">
  </div>

  <% If iSongs > 0 Then %>
  <div class="songs-grid">
  <%
  For iRow = 0 To iSongs - 1
      lId = NzLong(vSongs(0, iRow))
      sName = Nz(vSongs(1, iRow))
      sAudioUrl = AudioUrlOrEmpty(Nz(vSongs(2, iRow)))
      lPlays = NzLong(vSongs(3, iRow))
      sTmpKey = CStr(lId)
      If dArt.Exists(sTmpKey) Then sArtText = dArt(sTmpKey) Else sArtText = ""
      If dAlb.Exists(sTmpKey) Then
          aAlbumParts = Split(dAlb(sTmpKey), Chr(1))
          sAlbumTitle = aAlbumParts(0)
          sAlbumImg = aAlbumParts(1)
      Else
          sAlbumTitle = ""
          sAlbumImg = ""
      End If
      sCoverUrl = AssetUrl(sAlbumImg)
      iIdx = iRow
  %>
    <article class="song-card" data-index="<%= iIdx %>" role="button" tabindex="0" aria-label="Reproducir <%= H(sName) %>">
      <div class="card-art" style="--tint:<%= PastelTint(lId) %>;">
        <img src="<%= H(sCoverUrl) %>" alt="Portada de <%= H(sName) %>" loading="lazy"
             onerror="this.onerror=null;this.src='<%= H(hdrBase) %>assets/placeholder.svg';">
        <span class="card-playbtn" aria-hidden="true">
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M9 6.5v11l9-5.5z"/></svg>
        </span>
      </div>
      <h3 class="card-title"><%= H(sName) %></h3>
      <p class="card-artist"><% If Len(sArtText) > 0 Then %><%= H(sArtText) %><% Else %>Artista desconocido<% End If %></p>
      <div class="card-meta">
        <span class="chip chip-mini"><% If Len(sAlbumTitle) > 0 Then %><%= H(sAlbumTitle) %><% Else %>Sencillo<% End If %></span>
        <span class="card-plays" title="Reproducciones">&#9654; <span data-plays="<%= lId %>"><%= PlaysText(lPlays) %></span></span>
      </div>
    </article>
  <%
  Next
  %>
  </div>
  <% Else %>
  <div class="empty-state">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" aria-hidden="true"><path d="M9.5 17.5V7l9-2v10.3"/><circle cx="7" cy="17.5" r="2.5"/><circle cx="16" cy="15.3" r="2.5"/></svg>
    <p>No hay canciones que mostrar en esta vista.</p>
  </div>
  <% End If %>
</section>

<div class="player-bar" id="playerBar" aria-label="Reproductor de m&uacute;sica">
  <div class="player-info">
    <img class="player-art" id="playerArt" src="<%= H(hdrBase) %>assets/placeholder.svg" alt="Portada de la canci&oacute;n">
    <div class="player-meta">
      <span class="player-title" id="playerTitle">Listo para reproducir</span>
      <span class="player-artist" id="playerArtist">Selecciona una canci&oacute;n</span>
    </div>
  </div>
  <div class="player-center">
    <div class="player-controls">
      <button id="btnShuffle" class="pbtn" type="button" title="Aleatorio" aria-pressed="false">
        <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M10.59 9.17 5.41 4 4 5.41l5.17 5.17 1.42-1.41zM14.5 4l2.04 2.04L4 18.59 5.41 20 17.96 7.46 20 9.5V4h-5.5zm.33 9.41-1.41 1.41 3.13 3.13L14.5 20H20v-5.5l-2.04 2.04-3.13-3.13z"/></svg>
      </button>
      <button id="btnPrev" class="pbtn" type="button" title="Anterior">
        <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M6 6h2v12H6zm3.5 6 8.5 6V6z"/></svg>
      </button>
      <button id="btnPlay" class="pbtn pbtn-main" type="button" title="Reproducir / Pausar" aria-pressed="false">
        <svg id="icoPlay" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M9 6.5v11l9-5.5z"/></svg>
        <svg id="icoPause" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true" style="display:none"><path d="M8.5 6.5h3v11h-3zM12.5 6.5h3v11h-3z"/></svg>
      </button>
      <button id="btnNext" class="pbtn" type="button" title="Siguiente">
        <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M6 18l8.5-6L6 6zM16 6h2v12h-2z"/></svg>
      </button>
      <button id="btnRepeat" class="pbtn" type="button" title="Repetir" aria-pressed="false">
        <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M7 7h10v2l4-4-4-4v2H5v6h2V7zm10 10H7v-2l-4 4 4 4v-2h12v-6h-2v2z"/></svg>
      </button>
    </div>
    <div class="player-seek">
      <span class="ptime" id="timeNow">0:00</span>
      <input class="range" id="seekBar" type="range" min="0" max="1000" value="0" step="1" aria-label="Progreso de la canci&oacute;n">
      <span class="ptime" id="timeLeft">0:00</span>
    </div>
  </div>
  <div class="player-side">
    <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M3 9v6h4l5 5V4L7 9H3zm13.5 3a4.5 4.5 0 0 0-2.5-4v8a4.5 4.5 0 0 0 2.5-4z"/></svg>
    <input class="range range-vol" id="volBar" type="range" min="0" max="100" value="100" step="1" aria-label="Volumen">
  </div>
</div>
<audio id="audioPlayer" preload="metadata"></audio>

<script>
window.POLIMUSIC_CTX = {
  base: "<%= JsonEncode(GetAppBaseUrl()) %>",
  songs: <%= sSongsJson %>
};
</script>
<script src="<%= H(GetAppBaseUrl()) %>js/player.js" defer></script>
<!--#include file="includes/footer.asp"-->
