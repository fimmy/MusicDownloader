using Aspose.Cells;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace MusicDownloader
{
    class Program
    {
        static void Main(string[] args)
        {
            var log = NLog.LogManager.GetCurrentClassLogger();
            log.Info("开始");
            //var dl = new Downloader("", null);
            //var song = dl.GetQQMusicSearch("无忘花");
            //var song = dl.GetQQMusicSearch("无忘花");
            //dl.DownloadFile(song.url, $"{song.author} - {song.title}.mp3");
            var array = new string[]
            {
                "83169609",
                "34279751",
                "2139305008",
                "981484303",
                "21870141",
                "50415859",
                "915946943",
                "576742030"
            };

            var workbook = new Workbook();
            workbook.Worksheets.RemoveAt(0);
            foreach (var item in array)
            {
                var wsindex = workbook.Worksheets.Add();
                var downloader = new Downloader(item, workbook.Worksheets[wsindex]);
                downloader.DownloadPlaylist();
            }
            workbook.Save("result.xlsx");
            log.Info("结束");
            Console.Read();
        }

    }

}
