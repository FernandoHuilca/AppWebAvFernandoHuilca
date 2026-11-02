using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;
        string strSongFolder = ConfigurationManager.AppSettings["songFolderPath"];
        string strSongUrlFolder = "~/songFiles/";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session[Util.Constants.USER] == null)
            {
                Response.Redirect("Login.aspx", true);
                return;
            }
            Entities.User authenticatedUser = (Entities.User)Session[Util.Constants.USER];
            lblUserName.Text = authenticatedUser.UserName;
            if (!Page.IsPostBack)
            {
                lblMessage.Text = string.Empty;
                lblMessage.CssClass = string.Empty;
                pnlPlayer.Visible = false;
                BindSongsGrid();
            }
        }

        private void BindSongsGrid()
        {
            DataTable dtSongs = new BLL.Song(strConnString).Read();
            gdvSongs.DataSource = dtSongs;
            gdvSongs.DataBind();
        }

        private void ShowMessage(string message, bool isSuccess)
        {
            lblMessage.Text = message;
            lblMessage.CssClass = isSuccess ? "message message-success" : "message message-error";
        }

        protected void BtnUpload_Click(object sender, EventArgs e)
        {
            Entities.Song song = new Entities.Song();
            song.SongName = txtSongName.Text.Trim();
            song.Plays = 0;

            string messageValidations = new BLL.Song(strConnString).ValidationsRequired(song);
            if (messageValidations != string.Empty)
            {
                ShowMessage(messageValidations, false);
                return;
            }
            if (!fupSong.HasFile)
            {
                ShowMessage("Select a song file first!!", false);
                return;
            }

            // The same base name is used for the file on the server and for the path saved in the database.
            string newFileName = song.SongName.Replace(" ", string.Empty) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            messageValidations = new FileManagement().SaveSongOnServer(fupSong, strSongFolder, newFileName);
            if (messageValidations != string.Empty)
            {
                ShowMessage(messageValidations, false);
                return;
            }
            song.SongPath = strSongUrlFolder + newFileName + Path.GetExtension(fupSong.FileName).ToLower();

            int rows = new BLL.Song(strConnString).Save(song);
            if (rows > 0)
            {
                ShowMessage("Song added to the catalog successfully.", true);
                txtSongName.Text = string.Empty;
                BindSongsGrid();
            }
            else
            {
                ShowMessage("The song could not be saved on the database.", false);
            }
        }

        protected void gdvSongs_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName == "Play")
            {
                int idSong = Convert.ToInt32(e.CommandArgument);
                Entities.Song song = new BLL.Song(strConnString).ReadByID(idSong);
                if (song != null && song.SongPath != null && song.SongPath != string.Empty)
                {
                    new BLL.Song(strConnString).IncreasePlays(idSong);
                    audioPlayer.Attributes["src"] = ResolveUrl(song.SongPath);
                    audioPlayer.Attributes["autoplay"] = "autoplay";
                    lblPlaying.Text = song.SongName;
                    pnlPlayer.Visible = true;
                    ShowMessage(string.Empty, true);
                    BindSongsGrid();
                }
                else
                {
                    ShowMessage("The song file was not found.", false);
                }
            }
        }

        protected void gdvSongs_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                // Ask for confirmation before deleting a song.
                foreach (Control control in e.Row.Cells[4].Controls)
                {
                    LinkButton lnkCommand = control as LinkButton;
                    if (lnkCommand != null && lnkCommand.CommandName == "Delete")
                    {
                        lnkCommand.OnClientClick = "return confirm('\u00BFSeguro que deseas eliminar esta canci\u00F3n?');";
                    }
                }
            }
        }

        protected void gdvSongs_RowEditing(object sender, GridViewEditEventArgs e)
        {
            gdvSongs.EditIndex = e.NewEditIndex;
            BindSongsGrid();
        }

        protected void gdvSongs_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            gdvSongs.EditIndex = -1;
            BindSongsGrid();
        }

        protected void gdvSongs_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            TextBox txtEditName = (TextBox)gdvSongs.Rows[e.RowIndex].FindControl("txtEditName");
            int idSong = Convert.ToInt32(gdvSongs.DataKeys[e.RowIndex].Value);

            if (txtEditName == null || txtEditName.Text.Trim() == string.Empty)
            {
                ShowMessage("Song name can't be empty.", false);
                return;
            }

            Entities.Song song = new BLL.Song(strConnString).ReadByID(idSong);
            if (song == null)
            {
                ShowMessage("The song does not exist anymore.", false);
                gdvSongs.EditIndex = -1;
                BindSongsGrid();
                return;
            }
            song.SongName = txtEditName.Text.Trim();

            int rows = new BLL.Song(strConnString).Update(song);
            gdvSongs.EditIndex = -1;
            if (rows > 0)
            {
                ShowMessage("Song updated successfully.", true);
            }
            else
            {
                ShowMessage("The song could not be updated.", false);
            }
            BindSongsGrid();
        }

        protected void gdvSongs_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            int idSong = Convert.ToInt32(gdvSongs.DataKeys[e.RowIndex].Value);
            Entities.Song song = new BLL.Song(strConnString).ReadByID(idSong);

            int rows = new BLL.Song(strConnString).Delete(idSong);
            if (rows > 0)
            {
                // Remove the mp3 file attached to the song, if it exists on the server.
                if (song != null && song.SongPath != null && song.SongPath != string.Empty)
                {
                    try
                    {
                        string strServerPath = Server.MapPath(song.SongPath);
                        if (File.Exists(strServerPath))
                        {
                            File.Delete(strServerPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        ShowMessage("Song deleted, but the mp3 file could not be removed: " + ex.Message, false);
                    }
                }
                ShowMessage("Song deleted from the catalog successfully.", true);
            }
            else
            {
                ShowMessage("The song could not be deleted.", false);
            }
            BindSongsGrid();
        }
    }

}
