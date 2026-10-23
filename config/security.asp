<%
' ============================================================================
'  config/security.asp
'  Sesion y roles + utilidades de seguridad (XSS, rutas de recursos, JSON).
'  NOTA: este archivo NO lleva "Option Explicit" porque dbConnection.asp
'  (incluido siempre antes) ya lo declara.
' ============================================================================

' --- Cabeceras para contenido dinamico siempre fresco y UTF-8 ---
Response.CodePage = 65001
Response.Charset = "utf-8"
Response.Expires = -1
Response.CacheControl = "no-cache"

' ============================================================================
'  SESION / ROLES
' ============================================================================

Function IsLoggedIn()
    IsLoggedIn = (Len(Trim("" & Session("UserID"))) > 0)
End Function

Function IsAdmin()
    IsAdmin = False
    If IsLoggedIn() Then
        If "" & Session("UserType") = "1" Then IsAdmin = True
    End If
End Function

Function RoleLabel()
    If IsAdmin() Then
        RoleLabel = "Administrador"
    Else
        RoleLabel = "Usuario"
    End If
End Function

' Redirige al login si la sesion no existe.
Sub RequireLogin()
    If Not IsLoggedIn() Then
        Response.Redirect GetAppBaseUrl() & "login.asp"
        Response.End
    End If
End Sub

' Solo rol admin (USER_TYPE = 1). Si no, al dashboard sin acciones de admin.
Sub RequireAdmin()
    If Not IsLoggedIn() Then
        Response.Redirect GetAppBaseUrl() & "login.asp"
        Response.End
    End If
    If Not IsAdmin() Then
        Response.Redirect GetAppBaseUrl() & "index.asp"
        Response.End
    End If
End Sub

' ============================================================================
'  CODIFICACION (evita XSS al imprimir en HTML y JS)
' ============================================================================

Function H(sValue)          ' para cualquier salida en HTML
    H = Server.HTMLEncode("" & sValue)
End Function

Function JsonEncode(sValue) ' texto seguro para meter en un string JSON/JS
    Dim sVal
    sVal = "" & sValue
    sVal = Replace(sVal, "\", "\\")                 ' 1) escapar diagonales
    sVal = Replace(sVal, Chr(34), "\" & Chr(34))    ' 2) escapar comillas dobles
    sVal = Replace(sVal, vbCrLf, "\n")
    sVal = Replace(sVal, vbLf, "\n")
    sVal = Replace(sVal, vbCr, "")
    sVal = Replace(sVal, vbTab, "\t")
    sVal = Replace(sVal, "</", "<\/")               ' 3) evita cerrar <script>
    JsonEncode = sVal
End Function

' ============================================================================
'  URL BASE DE LA APLICACION
'  Funciona en la raiz del sitio, en cualquier sub-carpeta de wwwroot, en un
'  directorio virtual o como sub-aplicacion IIS sin tocar configuracion.
'  Metodo: relacionar la URL del script pedido (PATH_INFO) con su ruta fisica
'  (PATH_TRANSLATED) y subir por carpetas hasta encontrar config\security.asp.
'  (Antes dependia solo de APPL_MD_PATH: al mover la carpeta dentro de wwwroot
'  corre como parte del sitio Default, APPL_MD_PATH queda en ...\ROOT y la
'  base salia "/" -> css/js/images/redirects a 404. Este metodo no depende
'  de si la carpeta es o no una Aplicacion de IIS.)
' ============================================================================

Dim gAppBaseUrl, gAppPhysRoot    ' cache por peticion (visible en toda la pagina)
gAppBaseUrl = ""
gAppPhysRoot = ""

' Carpeta fisica raiz del proyecto (la que contiene config\security.asp),
' deducida desde la ruta fisica del script en ejecucion. Termina en "\".
Function AppRootFromScript()
    AppRootFromScript = ""
    Dim oFs, sPhys, sDir, lSlash, iDocs
    On Error Resume Next
    sPhys = "" & Request.ServerVariables("PATH_TRANSLATED")
    If Len(sPhys) = 0 Then Exit Function
    If Right(sPhys, 1) = "\" Then
        sDir = sPhys                                    ' fisica = carpeta (doc por defecto)
    Else
        lSlash = InStrRev(sPhys, "\")
        If lSlash = 0 Then Exit Function
        sDir = Left(sPhys, lSlash)                      ' carpeta del script
    End If
    Set oFs = Server.CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then Exit Function
    iDocs = 0
    Do While iDocs < 15                                 ' tope de profundidad
        If oFs.FolderExists(sDir & "config") Then
            If oFs.FileExists(sDir & "config\security.asp") Then
                AppRootFromScript = sDir
                Err.Clear
                Exit Function
            End If
        End If
        Err.Clear
        iDocs = iDocs + 1
        If Right(sDir, 1) = "\" Then sDir = Left(sDir, Len(sDir) - 1)
        If Len(sDir) <= 3 Then Exit Function            ' llego a la raiz de la unidad
        lSlash = InStrRev(sDir, "\")
        If lSlash < 4 Then Exit Function
        sDir = Left(sDir, lSlash)
    Loop
    Err.Clear
End Function

' Base URL real de la app: "/AppWebAvFernandoHuilca/" o "/" si corre en la raiz.
Function ComputeAppBaseUrl()
    ComputeAppBaseUrl = ""
    Dim sUrl, sPhys, sRoot, sRel, aUrl, uU, sLast, sFile, kCnt, r, i, sOut
    On Error Resume Next
    sUrl  = "" & Request.ServerVariables("PATH_INFO")        ' /AppWebAvFernandoHuilca/login.asp
    sPhys = "" & Request.ServerVariables("PATH_TRANSLATED")  ' C:\...\AppWebAvFernandoHuilca\login.asp
    If Len(sUrl) = 0 Or Len(sPhys) = 0 Then Exit Function

    sUrl = Replace(sUrl, "\", "/")
    Do While Len(sUrl) > 1 And Right(sUrl, 1) = "/"
        sUrl = Left(sUrl, Len(sUrl) - 1)
    Loop
    If sUrl = "/" Then                        ' peticion a la raiz del sitio
        ComputeAppBaseUrl = "/"
        Exit Function
    End If

    sRoot = AppRootFromScript()
    If Len(sRoot) = 0 Then Exit Function

    aUrl = Split(sUrl, "/")                   ' ["","AppWebAvFernandoHuilca",...]
    uU = UBound(aUrl)                         ' = numero de segmentos de URL
    If uU < 1 Then Exit Function
    sRel = Mid(sPhys, Len(sRoot) + 1)         ' "login.asp" | "admin\artists.asp" | ""
    sLast = LCase(aUrl(uU))

    ' numero de carpetas fisicas entre la raiz de la app y el script
    kCnt = 0
    If Right(sPhys, 1) = "\" Then
        ' PATH_TRANSLATED es una carpeta (la URL entro por documento por defecto)
        If Len(sRel) > 0 Then
            If Right(sRel, 1) = "\" Then sRel = Left(sRel, Len(sRel) - 1)
            If Len(sRel) > 0 Then kCnt = UBound(Split(sRel, "\")) + 1
        End If
        r = uU - kCnt                                          ' la URL no trae archivo
    Else
        sFile = LCase(Mid(sPhys, InStrRev(sPhys, "\") + 1))
        If Len(sRel) > 0 Then kCnt = UBound(Split(sRel, "\"))  ' el ultimo segmento es el archivo
        If sLast = sFile Then
            r = uU - kCnt - 1                                  ' la URL incluye el archivo
        Else
            r = uU - kCnt                                      ' la URL no incluye el archivo
        End If
    End If
    If r < 0 Then Exit Function

    ' base URL = los primeros r+1 segmentos, siempre con "/" final
    sOut = ""
    For i = 0 To r
        sOut = sOut & aUrl(i) & "/"
    Next
    ComputeAppBaseUrl = sOut
End Function

Function GetAppBaseUrl()
    If Len(gAppBaseUrl) = 0 Then gAppBaseUrl = ComputeAppBaseUrl()
    If Len(gAppBaseUrl) = 0 Then          ' respaldo: configuracion de IIS
        Dim sMd, sUrl, iPos
        On Error Resume Next
        sMd = "" & Request.ServerVariables("APPL_MD_PATH")
        On Error GoTo 0
        sUrl = ""
        If Len(sMd) > 0 Then
            iPos = InStr(1, sMd, "/ROOT", vbTextCompare)
            If iPos > 0 Then sUrl = Mid(sMd, iPos + 5)
        End If
        If Len(sUrl) = 0 Then sUrl = "/"
        If Right(sUrl, 1) <> "/" Then sUrl = sUrl & "/"
        gAppBaseUrl = sUrl
    End If
    GetAppBaseUrl = gAppBaseUrl
End Function

Function GetAppPhysicalRoot()
    ' Ruta fisica del root de la aplicacion, terminando en "\"
    If Len(gAppPhysRoot) = 0 Then
        gAppPhysRoot = AppRootFromScript()
        If Len(gAppPhysRoot) = 0 Then
            On Error Resume Next
            gAppPhysRoot = "" & Request.ServerVariables("APPL_PHYSICAL_PATH")
            On Error GoTo 0
        End If
        If Len(gAppPhysRoot) = 0 Then
            On Error Resume Next
            gAppPhysRoot = Server.MapPath(GetAppBaseUrl())
            On Error GoTo 0
        End If
        If Len(gAppPhysRoot) > 0 And Right(gAppPhysRoot, 1) <> "\" Then
            gAppPhysRoot = gAppPhysRoot & "\"
        End If
    End If
    GetAppPhysicalRoot = gAppPhysRoot
End Function

' ============================================================================
'  MAPEO DE RUTAS DE LA BD  ("~/songFiles/x.mp3" -> "SongFiles/x.mp3")
'  Carpetas reales de este proyecto: SongFiles, UserImages, PlaylistImages.
'  AlbumImages / ArtistImages se detectaran automaticamente cuando existan.
' ============================================================================

Function MapResourceUrl(sDbPath)
    Dim sVal, sFolder, sFile, iPos, sMapped
    MapResourceUrl = ""
    If IsNull(sDbPath) Then Exit Function
    sVal = Trim("" & sDbPath)
    If Len(sVal) = 0 Then Exit Function

    ' URLs absolutas (http/https) y rutas ya absolutas del sitio: tal cual
    If InStr(sVal, "://") > 0 Or Left(sVal, 1) = "/" Then
        MapResourceUrl = sVal
        Exit Function
    End If

    If Left(sVal, 2) = "~/" Then sVal = Mid(sVal, 3)
    sVal = Replace(sVal, "\", "/")

    iPos = InStr(sVal, "/")
    If iPos > 0 Then
        sFolder = LCase(Left(sVal, iPos - 1))
        sFile = Mid(sVal, iPos + 1)
    Else
        sFolder = ""
        sFile = sVal
    End If

    Select Case sFolder
        Case "songfiles"
            sMapped = "SongFiles/"
        Case "userimages"
            sMapped = "UserImages/"
        Case "playlistimages"
            sMapped = "PlaylistImages/"
        Case "albumimages"
            sMapped = "AlbumImages/"
        Case "artistimages"
            sMapped = "ArtistImages/"
        Case Else
            sMapped = ""
    End Select

    If Len(sMapped) = 0 Then sMapped = sFolder & "/"
    If Len(sFile) > 0 Then sMapped = sMapped & sFile
    MapResourceUrl = GetAppBaseUrl() & sMapped
End Function

' True si el archivo fisico existe dentro de la aplicacion
Function ResourceFileExists(sUrl)
    Dim sRel, sPhys, oFs, sBase
    ResourceFileExists = False
    On Error Resume Next
    If Len("" & sUrl) = 0 Then Exit Function
    If InStr(sUrl, "://") > 0 Then
        ResourceFileExists = True        ' no verificamos URLs externas
        Exit Function
    End If
    ' Solo URLs propias (empiezan con la base de la aplicacion)
    sBase = GetAppBaseUrl()
    If InStr(1, sUrl, sBase, vbTextCompare) <> 1 Then Exit Function
    sRel = Mid(sUrl, Len(sBase) + 1)
    sRel = Replace(sRel, "%20", " ")
    sPhys = GetAppPhysicalRoot() & Replace(sRel, "/", "\")
    Set oFs = Server.CreateObject("Scripting.FileSystemObject")
    ResourceFileExists = oFs.FileExists(sPhys)
    Set oFs = Nothing
    On Error GoTo 0
End Function

Function PlaceholderUrl()
    PlaceholderUrl = GetAppBaseUrl() & "assets/placeholder.svg"
End Function

' Devuelve la URL del recurso o "" si el fichero no existe (usado para audio)
Function AudioUrlOrEmpty(sDbPath)
    Dim sUrl
    AudioUrlOrEmpty = ""
    sUrl = MapResourceUrl(sDbPath)
    If Len(sUrl) = 0 Then Exit Function
    If ResourceFileExists(sUrl) Then AudioUrlOrEmpty = sUrl
End Function

' Devuelve la URL del recurso o un placeholder agradable si no existe
Function AssetUrl(sDbPath)
    Dim sUrl
    AssetUrl = PlaceholderUrl()
    sUrl = MapResourceUrl(sDbPath)
    If Len(sUrl) = 0 Then Exit Function
    If ResourceFileExists(sUrl) Then AssetUrl = sUrl
End Function

' --- Tinte pastel deterministico para las portadas sin imagen ---
Function PastelTint(iSeed)
    Dim iMod
    iMod = ((ToLong(iSeed) Mod 5) + 5) Mod 5
    Select Case iMod
        Case 0
            PastelTint = "var(--pastel-1)"
        Case 1
            PastelTint = "var(--pastel-2)"
        Case 2
            PastelTint = "var(--pastel-3)"
        Case 3
            PastelTint = "var(--pastel-4)"
        Case Else
            PastelTint = "var(--pastel-5)"
    End Select
End Function

Function PlaysText(vValue)
    If IsNull(vValue) Or Len("" & vValue) = 0 Then
        PlaysText = "0"
    ElseIf IsNumeric(vValue) Then
        PlaysText = CStr(CLng(vValue))
    Else
        PlaysText = "0"
    End If
End Function

' ============================================================================
'  PAGINA DE ERROR GENERICA
' ============================================================================

Function Nz(sValue)      ' Null-safe a texto
    If IsNull(sValue) Then
        Nz = ""
    Else
        Nz = "" & sValue
    End If
End Function

Function NzLong(vValue)  ' Null-safe a entero
    If IsNull(vValue) Then
        NzLong = 0
    Else
        NzLong = ToLong(vValue)
    End If
End Function

Sub ShowErrorPage(sMsg)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PoliMusic - Error</title>
<link rel="stylesheet" href="<%= H(GetAppBaseUrl()) %>css/styles.css">
</head>
<body class="error-page">
  <main class="error-card">
    <div class="error-ico" aria-hidden="true">&#9888;&#65039;</div>
    <h1>Ups, algo sali&oacute; mal</h1>
    <p><%= sMsg %></p>
    <div class="error-actions">
      <a class="btn btn-primary" href="<%= H(GetAppBaseUrl()) %>index.asp">Volver al inicio</a>
      <a class="btn btn-ghost" href="<%= H(GetAppBaseUrl()) %>logout.asp">Cerrar sesi&oacute;n</a>
    </div>
  </main>
</body>
</html>
<%
    Response.End
End Sub

' Texto de alerta a partir de ?msg=... en las listas del administrador
Function AdminMsg()
    Select Case LCase(Trim("" & Request.QueryString("msg")))
        Case "created"
            AdminMsg = "Registro creado correctamente."
        Case "saved"
            AdminMsg = "Cambios guardados correctamente."
        Case "deleted"
            AdminMsg = "Registro eliminado correctamente."
        Case "error"
            AdminMsg = "No se pudo completar la operaci&oacute;n. Revisa que los datos no rompan reglas (duplicados, tama&ntilde;os, etc.)."
        Case Else
            AdminMsg = ""
    End Select
End Function
%>
