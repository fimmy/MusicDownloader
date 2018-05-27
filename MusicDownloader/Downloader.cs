using Aspose.Cells;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;

namespace MusicDownloader
{
    public class Downloader
    {
        private Logger log;
        private string dir;
        private string id;
        private Worksheet worksheet;
        public Downloader(string id, Worksheet ws)
        {
            log = LogManager.GetCurrentClassLogger();
            worksheet = ws;
            this.id = id;
        }
        /// <summary>
        /// 请求获取json
        /// </summary>
        /// <param name="requestUrl"></param>
        /// <returns></returns>
        public string GetJSON(string requestUrl)
        {
            string result;
            try
            {
                HttpWebRequest httpWebRequest = WebRequest.Create(requestUrl) as HttpWebRequest;
                CookieContainer cookieContainer = new CookieContainer();
                httpWebRequest.CookieContainer = cookieContainer;
                httpWebRequest.AllowAutoRedirect = true;
                httpWebRequest.Method = "GET";
                httpWebRequest.ContentType = "application/x-www-form-urlencoded";
                httpWebRequest.Headers.Add("Authorization", "Basic YWRtaW46YWRtaW4=");
                HttpWebResponse httpWebResponse = httpWebRequest.GetResponse() as HttpWebResponse;
                Stream responseStream = httpWebResponse.GetResponseStream();
                StreamReader streamReader = new StreamReader(responseStream, Encoding.UTF8);
                string text = streamReader.ReadToEnd();
                string empty = string.Empty;
                result = text;
            }
            catch (Exception ex)
            {
                log.Error(ex);
                result = "";
            }
            return result;
        }
        public string GetMP3URL(string jsontext)
        {
            JObject jobject = JObject.Parse(jsontext);
            JArray jarray = JArray.Parse(jobject["data"].ToString());
            JObject jobject2 = JObject.Parse(jarray[0].ToString());
            return jobject2["url"].ToString();
        }
        public void DownloadPlaylist()
        {
            try
            {
                //歌曲列表
                var songlist = new List<Song>();
                var songidlist = new List<string>();
                //歌单信息
                var playliststr = GetJSON("http://localhost:3000/playlist/detail?id=" + id.Trim());
                JObject playlist = (JObject)JsonConvert.DeserializeObject(playliststr);
                //文件夹名设置为歌单名
                this.dir = playlist["playlist"]["name"].ToString();
                JArray songs = JArray.Parse(playlist["playlist"]["tracks"].ToString());

                for (int i = 0; i < songs.Count; i++)
                {
                    JObject song = JObject.Parse(songs[i].ToString());
                    var art = song["ar"].Select(ih => ih["name"].ToString()).ToArray();
                    songlist.Add(new Song
                    {
                        id = int.Parse(song["id"].ToString()),
                        title = song["name"].ToString(),
                        authors = art,
                        source = "网易"
                    });
                    songidlist.Add(song["id"].ToString());

                }
                //获取每一首歌的下载链接
                var ids = string.Join(',', songidlist);
                var songurls = GetJSON("http://localhost:3000/music/url?id=" + ids);
                JObject songurlsObj = (JObject)JsonConvert.DeserializeObject(songurls);
                JArray urls = JArray.Parse(songurlsObj["data"].ToString());
                for (int i = 0; i < urls.Count; i++)
                {
                    var url = urls[i]["url"].ToString();
                    var sid = int.Parse(urls[i]["id"].ToString());
                    var br = int.Parse(urls[i]["br"].ToString());
                    var t = songlist.FirstOrDefault(s => s.id == sid);
                    if (t is Song)
                    {
                        t.url = url;
                        t.br = br;
                    }
                }
                //若没有下载链接则到qq音乐搜索
                var urlnull = songlist.Where(s => string.IsNullOrEmpty(s.url)).ToList();
                foreach (var item in urlnull)
                {
                    try
                    {
                        var song = GetSongFromQQNew(item.title, item.authors[0] ?? "");
                        if (!string.IsNullOrEmpty(song.url))
                        {
                            item.oldauthor = item.author;
                            item.oldtitle = item.title;
                            item.title = song.title;
                            item.authors = new string[] { song.author };
                            item.url = song.url;
                            item.source = song.source;
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }

                }
                var cell = worksheet.Cells;
                cell[0, 0].PutValue("演唱者");
                cell[0, 1].PutValue("标题");
                cell[0, 2].PutValue("源演唱者");
                cell[0, 3].PutValue("源标题");
                cell[0, 4].PutValue("来源");
                cell[0, 5].PutValue("url");
                cell[0, 6].PutValue("下载结果");
                cell[0, 7].PutValue("文件名");
                cell[0, 8].PutValue("码率");
                int cellindex = 1;
                //开始下载
                foreach (var item in songlist)
                {
                    try
                    {
                        log.Info($"[{songlist.IndexOf(item)}/{songlist.Count}]开始下载{item.author} - {item.title}");
                        DownloadFile(item.url, $"{item.author} - {item.title}.mp3");
                        log.Info($"{item.author} - {item.title} 下载完成");
                        cell[cellindex, 6].PutValue("成功");
                    }
                    catch (Exception ex)
                    {
                        cell[cellindex, 6].PutValue("失败");
                        log.Info($"{item.author} - {item.title} 下载失败");
                        cell[cellindex, 9].PutValue(ex.ToString());
                    }
                    cell[cellindex, 0].PutValue(item.author);
                    cell[cellindex, 1].PutValue(item.title);
                    cell[cellindex, 2].PutValue(item.oldauthor);
                    cell[cellindex, 3].PutValue(item.oldtitle);
                    cell[cellindex, 4].PutValue(item.source);
                    cell[cellindex, 5].PutValue(item.url);
                    cell[cellindex, 7].PutValue($"{item.author} - {item.title}");
                    cell[cellindex, 8].PutValue(item.br);
                    cellindex++;
                }
                worksheet.Name = dir;
            }
            catch (Exception ex2)
            {
                log.Error(ex2);
            }

        }
        public string GetQQMusic_vkey()
        {
            string json = GetJSON("http://base.music.qq.com/fcgi-bin/fcg_musicexpress.fcg?json=3&guid=5150825362&format=json");
            JObject jobject = JObject.Parse(json);
            return jobject["key"].ToString();
        }
        public Song GetSongFromQQNew(string title, string author)
        {
            var song = new Song();
            try
            {
                var titleauthor = $"{title} {author}".Trim();
                var uri = $"https://c.y.qq.com/splcloud/fcgi-bin/smartbox_new.fcg?is_xml=0&format=json&key={titleauthor}&g_tk=5381&loginUin=0&hostUin=0&inCharset=utf8&outCharset=utf-8&notice=0&platform=yqq&needNewCode=0";
                var json = GetJSON(uri);
                JObject jobject = JObject.Parse(json);
                JObject jobject2 = JObject.Parse(jobject["data"].ToString());
                JObject jobject3 = JObject.Parse(jobject2["song"].ToString());
                var count = int.Parse(jobject3["count"].ToString());
                if (count > 0)
                {
                    JArray jarray = JArray.Parse(jobject3["itemlist"].ToString());
                    if (jarray.Count > 0)
                    {
                        JObject jobject4 = JObject.Parse(jarray[0].ToString());
                        var mid = jobject4["mid"].ToString();
                        //根据mid获取歌曲的mediaid
                        var mediamiduri = $"https://c.y.qq.com/v8/fcg-bin/fcg_play_single_song.fcg?songmid={mid}&tpl=yqq_song_detail&format=json&g_tk=5381&loginUin=0&hostUin=0&format=jsonp&inCharset=utf8&outCharset=utf-8&notice=0&platform=yqq&needNewCode=0";
                        var songdownloadurl = GetJSON(mediamiduri);
                        var songjson = JObject.Parse(songdownloadurl)["data"][0];
                        var realmid = songjson["file"]["media_mid"].ToString();
                        var singer = string.Join("、", songjson["singer"].Select(i => i["name"].ToString()).ToArray());
                        var songname = songjson["title"].ToString();

                        var qqurl = string.Concat(new string[]
                        {
                        "http://dl.stream.qqmusic.qq.com/M800",
                        realmid,
                        ".mp3?vkey=",
                        GetQQMusic_vkey(),
                        "&guid=5150825362&fromtag=1"
                        });

                        song.url = qqurl;
                        song.authors = new string[] { singer };
                        song.title = songname;
                        song.source = "QQ音乐";
                    }
                }

                return song;
            }
            catch (Exception)
            {
                if (!string.IsNullOrEmpty(author))
                {
                    song = GetSongFromQQNew(title, "");

                }
                else
                {
                    song = GetQQMusicSearch(title);
                }
                return song;
            }

        }
        public Song GetQQMusicSearch(string title)
        {

            var song = new Song();
            try
            {
                string json = this.GetJSON("http://s.music.qq.com/fcgi-bin/music_search_new_platform?t=0&n=1&aggr=1&cr=1&loginUin=0&format=json&inCharset=GB2312&outCharset=utf-8&notice=0&platform=jqminiframe.json&needNewCode=0&p=1&catZhida=0&remoteplace=sizer.newclient.next_song&w=" + title.Trim());
                JObject jobject = JObject.Parse(json);
                JObject jobject2 = JObject.Parse(jobject["data"].ToString());
                JObject jobject3 = JObject.Parse(jobject2["song"].ToString());
                JArray jarray = JArray.Parse(jobject3["list"].ToString());
                JObject jobject4 = JObject.Parse(jarray[0].ToString());
                var mid = jobject4["f"].ToString().Split('|')[20].ToString();
                var mediamiduri = $"https://c.y.qq.com/v8/fcg-bin/fcg_play_single_song.fcg?songmid={mid}&tpl=yqq_song_detail&format=json&g_tk=5381&loginUin=0&hostUin=0&format=jsonp&inCharset=utf8&outCharset=utf-8&notice=0&platform=yqq&needNewCode=0";
                var songdownloadurl = GetJSON(mediamiduri);
                var songjson = JObject.Parse(songdownloadurl)["data"][0];
                var realmid = songjson["file"]["media_mid"].ToString();
                var singer = string.Join("、", songjson["singer"].Select(i => i["name"].ToString()).ToArray());
                var songname = songjson["title"].ToString();

                var qqurl = string.Concat(new string[]
                {
                          "http://dl.stream.qqmusic.qq.com/M800",
                        realmid,
                        ".mp3?vkey=",
                        GetQQMusic_vkey(),
                        "&guid=5150825362&fromtag=1"
                });
                song.url = qqurl;
                song.authors = new string[] { singer };
                song.title = songname;
                song.source = "QQ音乐";
                return song;
            }
            catch (Exception)
            {
                return song;
            }

        }
        public void DownloadFile(string URL, string filename)
        {
            if (!Directory.Exists(this.dir))
            {
                Directory.CreateDirectory(this.dir);
            }
            filename = filename.Replace("/", "、");
            string invalid = new string(Path.GetInvalidFileNameChars());
            foreach (char c in invalid)
            {
                filename = filename.Replace(c.ToString(), "");
            }
            bool flag = File.Exists(this.dir + "/" + filename);
            if (!flag)
            {
                try
                {
                    if (string.IsNullOrEmpty(URL))
                    {
                        throw new Exception("url empty");
                    }
                    WebClient myWebClient = new WebClient();

                    myWebClient.DownloadFile(URL, this.dir + "/" + filename);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

    }
}
