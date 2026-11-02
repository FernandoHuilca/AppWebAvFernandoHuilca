<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="utf-8" />
    <title>Gesti&oacute;n de Canciones | PoliMusic</title>
    <link rel="stylesheet" href="Css/PoliMusic.css" />
</head>
<body>
    <!-- Header con el usuario autenticado y el link para cerrar sesion -->
    <div class="header">
        <div class="brand">Poli<span class="accent">Music</span></div>
        <div class="user-box">
            <span class="user-chip">Usuario autenticado: <strong><asp:Label ID="lblUserName" runat="server"></asp:Label></strong></span>
            <asp:HyperLink ID="lnkLogout" runat="server" NavigateUrl="~/Logout.aspx" CssClass="btn-logout">Cerrar sesi&oacute;n</asp:HyperLink>
        </div>
    </div>

    <div class="container">
        <h1 class="page-title">Gesti&oacute;n de Canciones</h1>
        <p class="subtitle">Administra tu cat&aacute;logo musical: agrega, escucha, edita y elimina canciones.</p>

        <form id="formSong" runat="server">
            <!-- Panel: agregar cancion -->
            <div class="panel">
                <h2>Nueva canci&oacute;n</h2>
                <div class="form-row">
                    <label class="field-label" for="txtSongName">Nombre de la canci&oacute;n *</label>
                    <asp:TextBox ID="txtSongName" runat="server" CssClass="text-input" placeholder="Ej.: Adventure"></asp:TextBox>
                </div>
                <div class="form-row">
                    <label class="field-label" for="fupSong">Archivo MP3 *</label>
                    <asp:FileUpload ID="fupSong" runat="server" CssClass="upload-input" />
                </div>
                <div class="form-row">
                    <asp:Button ID="btnUpload" runat="server" Text="Subir canci&oacute;n" CssClass="btn-primary" OnClick="BtnUpload_Click" />
                </div>
                <asp:Label ID="lblMessage" runat="server"></asp:Label>
            </div>

            <!-- Panel: catalogo -->
            <div class="panel">
                <h2>Cat&aacute;logo de canciones</h2>
                <asp:GridView ID="gdvSongs" runat="server"
                    AutoGenerateColumns="false"
                    DataKeyNames="ID_SONG"
                    GridLines="None"
                    CssClass="grid"
                    EmptyDataText="No hay canciones registradas todav&iacute;a."
                    OnRowCommand="gdvSongs_RowCommand"
                    OnRowEditing="gdvSongs_RowEditing"
                    OnRowCancelingEdit="gdvSongs_RowCancelingEdit"
                    OnRowUpdating="gdvSongs_RowUpdating"
                    OnRowDeleting="gdvSongs_RowDeleting"
                    OnRowDataBound="gdvSongs_RowDataBound">
                    <Columns>
                        <asp:BoundField DataField="ID_SONG" HeaderText="#" ReadOnly="true">
                            <ItemStyle Width="50px" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Canci&oacute;n">
                            <ItemTemplate>
                                <span class="song-title"><%# Eval("SONG_NAME") %></span><br />
                                <span class="song-path">
                                    <asp:HyperLink ID="lnkSongFile" runat="server" NavigateUrl='<%# Eval("SONG_PATH") %>' Text='<%# Eval("SONG_PATH") %>' Target="_blank" />
                                </span>
                            </ItemTemplate>
                            <EditItemTemplate>
                                <asp:TextBox ID="txtEditName" runat="server" Text='<%# Bind("SONG_NAME") %>' CssClass="text-input"></asp:TextBox>
                            </EditItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Reproducciones">
                            <ItemTemplate>
                                <span class="plays-chip"><%# Eval("PLAYS") %> &#9654;</span>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Escuchar">
                            <ItemTemplate>
                                <asp:LinkButton ID="lnkPlay" runat="server" CssClass="btn-play" CommandName="Play" CommandArgument='<%# Eval("ID_SONG") %>' Text="&#9654; Reproducir" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:CommandField HeaderText="Acciones" ShowEditButton="true" ShowDeleteButton="true" EditText="Editar" DeleteText="Eliminar" ControlStyle-CssClass="action-link" />
                    </Columns>
                </asp:GridView>

                <!-- Reproductor -->
                <asp:Panel ID="pnlPlayer" runat="server" Visible="false" CssClass="player">
                    <span class="now-playing">&#9679; Reproduciendo ahora</span>
                    <asp:Label ID="lblPlaying" runat="server" CssClass="now-playing-title"></asp:Label>
                    <audio id="audioPlayer" runat="server" controls></audio>
                </asp:Panel>
            </div>
        </form>
    </div>
</body>
</html>
