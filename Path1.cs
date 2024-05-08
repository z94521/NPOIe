using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIe
{
    internal class Path1
    {
       private  string readPath_;
        private string writePath_; 
        public string ReadPath
        {
            get { return @"C:\Users\Administrator\Desktop\" + readPath_ + ".xls"; }
            set { readPath_ = value; }
        }
        public string WritePath
        {
            get { return @"C:\Users\Administrator\Desktop\" + writePath_ + ".xls"; }
            set { readPath_ = value; }
        }
        public Path1(string readpath,string writpath)
        {
            this.readPath_ = readpath;
            this.writePath_ = writpath;
        }
        public Path1  ()
        {

        }
        public string Writefilename(string fileName)
        {
         return   this.writePath_ = fileName;
        }
        public string Readfilename(string fileName)
        {
            return (this.readPath_=fileName);
        }
        public byte[] bytes = new byte[102400];
        
    }
}
