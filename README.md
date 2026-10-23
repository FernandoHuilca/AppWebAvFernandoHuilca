# PoliMusic — Clon de Spotify en ASP Clásico (VBScript)

Aplicación web tipo streaming desarrollada **100 % en ASP Clásico (VBScript)** con
Vanilla JavaScript, CSS3 puro y SQL Server (base `BDD_PoliMusic`).
Sin ASP.NET, sin C#, sin frameworks.

---

## 1. Estructura del proyecto

```
AppWebAvFernandoHuilca/
├─ login.asp                  Pantalla de login (valida rol USER / ADMIN)
├─ logout.asp                 Destruye la sesión y regresa al login
├─ index.asp                  Dashboard: bienvenida, playlists, Listado de
│                             Canciones, reproductor y accesos admin
├─ song_plays.asp             Endpoint POST: incrementa PLAYS (fetch de JS)
├─ web.config                 Documento por defecto (login.asp) para IIS
│
├─ config/
│   ├─ dbConnection.example.asp  Plantilla SIN contraseñas (va al repo)
│   ├─ dbConnection.asp          Cadena con tu contraseña real (local,
│   │                            protegido por .gitignore)
│   └─ security.asp           Sesiones, roles, anti-XSS, mapeo de rutas ~/...
│
├─ includes/
│   ├─ header.asp             Layout compartido (sidebar, avatar, logout)
│   └─ footer.asp             Cierre del layout + zona de toast
│
├─ admin/                     CRUD (solo rol ADMIN)
│   ├─ songs.asp              Listado de canciones
│   ├─ song_form.asp          Crear / editar canción (+ artistas en TBL_PLAYER)
│   ├─ song_delete.asp        Eliminar canción (transacción)
│   ├─ artists.asp / artist_form.asp / artist_delete.asp
│   └─ albums.asp  / album_form.asp  / album_delete.asp
│
├─ css/  (styles.css, login.css)
├─ js/   (player.js, admin.js)
├─ bd/BDD_PoliMusic.sql       Script de base de datos
├─ assets/                    favicon.svg y placeholder.svg (fallback visual)
│
├─ SongFiles/                 Audios .mp3  (mapea "~/songFiles/…" de la BD)
├─ UserImages/                Fotos de usuario (mapea "~/userImages/…")
├─ PlaylistImages/            Portadas de playlists (mapea "~/playlistImages/…")
├─ AlbumImages/               Portadas de álbumes (mapea "~/albumImages/…")
└─ ArtistImages/              Fotos de artistas (mapea "~/ArtistImages/…")
```

> Si una imagen/álbum/artista aún no tiene archivo (p. ej. `AlbumImages`,
> `ArtistImages` todavía no existen), la app dibuja automáticamente un
> **placeholder pastel** (`assets/placeholder.svg`). Al crear la carpeta y copiar
> los archivos (mis nombres que indica la BD), las imágenes aparecen solas.
>
> ℹ️ Las carpetas de media (`SongFiles/`, `UserImages/`, `PlaylistImages/`…)
> van **vacías** en el repo (`.gitignore` las excluye; solo llevan `.gitkeep`):
> los audios y fotos viven únicamente en tu servidor IIS.

## 2. Base de datos

1. Abre `bd/BDD_PoliMusic.sql` en **SSMS / Azure Data Studio** (o `sqlcmd`)
   sobre `localhost` y ejecútalo completo.
2. Crea el login/usuario `usr_polimusic` con la contraseña que prefieras
   (el script usa el placeholder `PON_TU_CONTRASENA_AQUI`; cámbialo antes de
   ejecutarlo, no publishes tu contraseña real) y todas las tablas del esquema
   original, **sin tablas adicionales**.

Credenciales de prueba (semilla del script):

| Usuario     | Contraseña | Rol  |
|-------------|------------|------|
| `userAdmin` | `12345678` | ADMIN (USER_TYPE = 1) |
| `userNormal`| `12345678` | USER  (USER_TYPE = 0) |
| `username1` | `12345678` | USER  (contraseña guardada como SHA-1) |

El login soporta **ambas formas**: contraseñas planas y contraseñas almacenadas
como hash `SHA1` en mayúsculas, comparadas con
`CONVERT(VARCHAR(40), HASHBYTES('SHA1', …), 2)` del propio SQL Server.

## 3. Montaje en IIS (Windows)

1. **Activar IIS + ASP clásico** (Panel de control → Turn Windows features on or off):
   - Internet Information Services → World Wide Web Services →
     Application Development Features → **ASP**  ✅
   - (Opcional) Common HTTP Features → Static Content, Default Document ✅
2. **Crear el sitio/aplicación** apuntando a esta carpeta:
   - IIS Manager → *Add Website* (o *Convert to Application* dentro de Default Web Site)
   - Physical path: `…\AppWebAvFernandoHuilca`
   - Application Pool: **No Managed Code** (Classic, v4.0azu o sin código)
   - Autenticación: **Anonymous** habilitada (la sesión se maneja con cookies ASP)
3. **Permisos**: que el Application Pool Identity tenga lectura/escritura de la
   carpeta (IIS_IUSRS / AppPool identity) para que FSO pueda listar `SongFiles`.
4. Seguridad integrada de SQL: la app usa el login SQL `usr_polimusic`;
   asegúrate de que el servidor permita **SQL Server Authentication**
   (propiedades → Security → Mixed Mode) y reinicia el servicio si lo cambias.
5. Abrir `http://localhost:<puerto>/login.asp`.

Si instalas la app como **sub-aplicación** (ej. `http://localhost/Polimusic/`)
todo funciona igual: la URL base y las rutas de recursos se calculan
dinámicamente desde la propia URL del script (`GetAppBaseUrl()` en
`config/security.asp`), sin importar si corre en la raíz, como sub-carpeta o
como aplicación. Las páginas de `admin/` necesitan que el montaje esté como
**Aplicación IIS** (paso 2) o tener activado *ASP → Parent Paths*.

## 4. Cadena de conexión

1. Copia `config/dbConnection.example.asp` como `config/dbConnection.asp`
   (el archivo real con tu contraseña nunca se sube a GitHub: está en
   `.gitignore`) y edita **una sola línea**:

```vb
Const DB_CONN_STRING = "Provider=MSOLEDBSQL;Server=localhost;Database=BDD_PoliMusic;Uid=usr_polimusic;Pwd=TU_PASSWORD;"
```

Alternativas listadas en el mismo archivo (descomentar una):

- `Provider=MSOLEDBSQL;…` (Microsoft OLE DB Driver for SQL Server moderno)
- `Driver={ODBC Driver 17 for SQL Server};…`
- Con instancia con nombre: `Server=localhost\SQLEXPRESS;…`

## 5. Seguridad implementada

- ** Anti inyección SQL**: *ninguna* consulta concatena valores del usuario.
  Todo entra por `ADODB.Command` + `CreateParameter` con tipo y longitud
  (`adVarChar`/`adInteger`). Los ids del querystring se convierten con
  `ToLong()` antes de parametrizar.
- **Anti XSS**: todo valor de la BD se imprime con `Server.HTMLEncode`
  (helper `H()`), incluido el JSON del reproductor (`JsonEncode` escapa
  comillas, backslashes, saltos y `</`).
- **Sesiones**: `Session("UserID"/"Username"/"UserType"/"UserPhotoUrl")`.
  - `RequireLogin()` en todas las vistas privadas.
  - `RequireAdmin()` en todo `admin/` (rol `USER_TYPE = 1`).
  - `logout.asp` ejecuta `Session.Contents.RemoveAll()` + `Session.Abandon()`.
- **Borrados en transacción**: primero las tablas hijas
  (`TBL_PLAYER`, `TBL_TRACKLIST`, `TBL_PLAYLIST_SONG`) y luego el padre;
  si algo falla, `RollbackTrans`.
- **JSON del player**: `u` va vacío si el mp3 no existe (FSO lo verifica) y
  `player.js` salta a la siguiente canción disponible con un toast.

## 6. Funcionalidades

- **login.asp**: validación contra BD, mensajes de error, cuentas demo, banner
  de "sesión finalizada" al salir.
- **Dashboard (`index.asp`)**:
  - Bienvenida dinámica usando la variable de sesión.
  - Botón **Cerrar sesión** en banner y sidebar.
  - **Listado de Canciones** en grilla con portada real de la BD
    (o placeholder), artistas (TBL_PLAYER→TBL_ARTIST), álbum (TBL_TRACKLIST→TBL_ALBUM).
  - **Playlists** del usuario (TBL_PLAYLIST) con filtro por canción
    (`index.asp?pl=1`) usando TBL_PLAYLIST_SONG.
  - **Reproductor HTML5**: play/pause, anterior/siguiente, **aleatorio**,
    **repetir**, barra de progreso, volumen, auto-avance al terminar,
    contador de **PLAYS** incrementado vía `song_plays.asp`.
  - Buscador en vivo (nombre/artista/álbum).
- **Panel admin** (solo ADMIN):
  - Botones de acceso rápido (Canciones / Artistas / Álbumes) + sidebar.
  - CRUD completo de las tres entidades con validación de tamaños/tipos,
    transacciones y mensajes de éxito/error.
  - En el formulario de canción puedes **marcar artistas** (checkboxes) que se
    persisten en `TBL_PLAYER`.

## 7. Solución de problemas

| Síntoma | Causa probable | Solución |
|---|---|---|
| `No hay conexión con SQL Server` | Servicio detenido o provider ausente | Iniciar SQL Server (MSSQLSERVER) y/o cambiar `DB_CONN_STRING` a `MSOLEDBSQL` u ODBC 17 |
| Login falla con error 18456 | Autenticación Mixed Mode desactivada | Propiedades del servidor → Security → "SQL Server and Windows Authentication mode" |
| HTTP 500 en `.asp` | ASP no habilitado | Instalar la característica **ASP** en Windows Features |
| Imágenes genéricas | Falta la carpeta `AlbumImages` / `ArtistImages` con esos nombres | Crear carpeta y copiar los archivos con los nombres de la BD |
| Audio absent | El mp3 no está en `SongFiles` | Copiar p. ej. `bensound-dreams.mp3` a `SongFiles/` (el reproductor salta a la siguiente disponible) |
| CSS/JS 404 como sub-app | Se movió la aplicación | Verificar que las carpetas `css/js/assets` estén dentro del root de la aplicación IIS |

## 8. Notas técnicas

- Todas las páginas ASP están en **UTF-8 sin BOM** con
  `<%@ Language="VBScript" CodePage="65001" %>`; los textos VBScript usan sólo
  ASCII y los acentos se escriben como entidades HTML (`&oacute;`) para que no
  dependan del code page del servidor.
- `server.MapPath`/FSO comprueban la existencia física de cada recurso para
  servir placeholder o vacío cuando falta el archivo real.
- El layout es responsive (sidebar colapsa a barra superior y el player se
  compacta en móvil).
