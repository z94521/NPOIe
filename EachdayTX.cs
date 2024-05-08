using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIe
{
    public class EachdayTX
    {
        public string Sta { get; set; }
        public string Date { get; set; }
        public int TXnum { get; set; }
        public static void A1()
        {
            Path1 path1 = new Path1();
            path1.ReadPath = "n1";
            path1.WritePath = "Max";
            FileStream read = new FileStream(path1.ReadPath, FileMode.Open, FileAccess.Read);
            FileStream write = new FileStream(path1.WritePath, FileMode.OpenOrCreate, FileAccess.Write);

            int bytesread;
            //while ((bytesread = read.Read(path1.bytes, 0, path1.bytes.Length)) != 0)
            //{
            //    //    string st = Encoding.UTF8.GetString(data_.bytes, 0, bytesread);
            //    // Console.WriteLine(st);
            //    // data_.bytes = Encoding.UTF8.GetBytes(st);
            //    write.Write(path1.bytes, 0, path1.bytes.Length);
            //}
            read.Close();
            write.Close();
            Console.WriteLine("OK!");
        }
        public static byte[] A2()
        {
            Path1 path1 = new Path1();
            path1.Readfilename("n1");
            int bytesread;
            FileStream read = new FileStream(path1.ReadPath, FileMode.Open, FileAccess.Read);
            bytesread = read.Read(path1.bytes, 0, path1.bytes.Length);

            read.Close();
            return path1.bytes;
        }
        public static void A3()
        {
            Path1 path1 = new Path1();
            path1.Writefilename("Max");
            FileStream write = new FileStream(path1.WritePath, FileMode.OpenOrCreate, FileAccess.Write);

            Byte[] bytes = new byte[102400];
            bytes = A2();
            write.Write(bytes, 0, bytes.Length);
            write.Close();
        }

    }

}
