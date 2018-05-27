using System;
using System.Collections.Generic;
using System.Text;

namespace MusicDownloader
{
    public class Song
    {
        public string author
        {
            get
            {
                return string.Join("、", authors);
            }
        }
        /// <summary>
        /// 演唱者
        /// </summary>
        public string[] authors { get; set; }
        /// <summary>
        /// 标题
        /// </summary>
        public string title { get; set; }
        /// <summary>
        /// id
        /// </summary>
        public int id { get; set; }
        /// <summary>
        /// 下载地址
        /// </summary>
        public string url { get; set; }
        /// <summary>
        /// 码率
        /// </summary>
        public int br { get; set; }
        /// <summary>
        /// 来源
        /// </summary>
        public string source { get; set; }
        public string oldauthor { get; set; }
        public string oldtitle { get; set; }
    }
}
