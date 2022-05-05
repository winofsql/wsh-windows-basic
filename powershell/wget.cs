using System;
using System.IO;
using System.Net;
using System.Web;
using System.Windows.Forms;

public class Program
{
    public static void Main(string[] args)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;	

        string[] param = Environment.GetCommandLineArgs();
        if (args.Length != 0)
        {
            Console.WriteLine(string.Format("第一引数 : {0}", args[0]));
        }
        else
        {
            MessageBox.Show("ダウンロードする URL を引数に指定して下さい");
            Environment.Exit(0);
        }

        string localFileName = Path.GetFileName(args[0]);
        Console.WriteLine(string.Format("ファイル名 : {0}", localFileName));

        using( WebClient wc = new WebClient() ) {
            wc.DownloadFile( args[0], localFileName );
        }

        // *******************************************
        // -ReferencedAssemblies の複数テスト用
        // *******************************************
        string percent_encoding = HttpUtility.UrlEncode(args[0]);
        Console.WriteLine( percent_encoding );

        MessageBox.Show("処理が終了しました");

    }
}
