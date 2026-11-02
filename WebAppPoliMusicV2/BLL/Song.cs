using System;
using System.Collections.Generic;
using System.Data;
using System.Web;

namespace WebAppPoliMusicV2.BLL
{
    public class Song
    {
        private string strConnectionString;
        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }
        public int Save(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Save(song);
        }
        public DataTable Read()
        {
            return new DAL.Song(strConnectionString).Read();
        }
        public Entities.Song ReadByID(int idSong)
        {
            return new DAL.Song(strConnectionString).ReadByID(idSong);
        }
        public int Update(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Update(song);
        }
        public int Delete(int idSong)
        {
            return new DAL.Song(strConnectionString).Delete(idSong);
        }
        public int IncreasePlays(int idSong)
        {
            return new DAL.Song(strConnectionString).IncreasePlays(idSong);
        }
        public string ValidationsRequired(Entities.Song song)
        {
            if (song.SongName == null || song.SongName.Trim() == string.Empty)
            {
                return "Song name is required!";
            }
            if (song.SongPath == null || song.SongPath.Trim() == string.Empty)
            {
                return "Song file is required!";
            }
            return string.Empty;
        }
    }

}
