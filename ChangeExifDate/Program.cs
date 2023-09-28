using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace ChangeExifDate
{
    internal class Program
    {
        static void Main(string[] args)
        {

            //続行確認
            Console.WriteLine("以下のフォルダのjpegの撮影日時を書き換えますよ？");
            Console.WriteLine("フォルダ名＝" + Environment.CurrentDirectory + "\\in");
            Console.Write("y=続行、n=終了：");

            //yかnが入力されるまで繰り返し
            while (true)
            {

                //入力文字
                string yn = Console.ReadLine();

                if (yn == "y")
                {

                    //yなら続行
                    break;
                }
                else if (yn == "n")
                {
                    //nならリターン
                    Console.Write("やらないなら何かキーを押してたら終了しますよ。：");
                    Console.ReadLine();
                    return;
                }
            }
            string inPath = Path.Combine(Environment.CurrentDirectory, "in");
            //フォルダ内のエクセルファイル取得
            string[] fileList = Directory.GetFiles(inPath, "*.jpg");


            //何もなかったら終了
            if (fileList.Length == 0)
            {
                Console.WriteLine("jpegファイルが無い、やりなおし");
                Console.ReadLine();
                return;
            }

            //DBファイル存在確認
            if (!File.Exists(Path.Combine(Environment.CurrentDirectory, "fec.db")))
            {
                Console.WriteLine("DBファイルが無い！やりなおし：");
                Console.ReadLine();
                return;
            }

            //CSVファイル存在確認
            if (!File.Exists("data.csv"))
            {
                Console.WriteLine("CSVファイルが無い！やりなおし：");
                Console.ReadLine();
                return;
            }

            List<DataCsv> dataList = new List<DataCsv>();

            //CSVの変換リスト読んで分解
            using(StreamReader sr = new StreamReader("data.csv"))
            {
                //終わりまで
                while (!sr.EndOfStream)
                {

                    string line = sr.ReadLine();

                    //先頭行は不要
                    if (line == "file,no,parts")
                        continue;

                    string[] lines = line.Split(',');

                    //３列じゃなければ変なやつ
                    if (lines.Length != 3)
                    {
                        Console.WriteLine("変な行がありました。" + line);
                    }
                    else
                    {
                        dataList.Add(new DataCsv { File = lines[0], SerialNo = lines[1], Parts = lines[2] });
                    }
                }
            }

            //出力フォルダ
            string outPath = Path.Combine(Environment.CurrentDirectory, "out");

            //無ければ作成
            if (!File.Exists(outPath))
            {
                Directory.CreateDirectory(outPath);
            }

            //DB接続文字列
            string connect = "Data Source=" + Path.Combine(Environment.CurrentDirectory, "fec.db");
            int cnt = 0;

            try
            {

                //DBアクセス
                using (SqliteConnection con = new SqliteConnection())
                using (SqliteCommand cmd = con.CreateCommand())
                {

                    //オープン
                    con.ConnectionString = connect;
                    Console.WriteLine("constr" + con.ConnectionString);
                    con.Open();

                    //変換リスト毎に処理
                    foreach (DataCsv data in dataList)
                    {

                        Console.WriteLine("ファイル名＝{0}、管理番号＝{1}、部位＝{2}", data.File, data.SerialNo, data.Parts);

                        //変換リストのファイルがあるかな
                        if (!File.Exists(Path.Combine(inPath, data.File)))
                        {
                            Console.WriteLine("ファイル無かったです。失敗。");
                            Console.WriteLine();
                            continue;
                        }

                        DateTime startTs = default(DateTime);
                        DateTime endTs = default(DateTime);
                        DateTime midTs = default(DateTime);

                        string no;
                        string insPoint;
                        string lineNum;
                        string action;
                        bool isSuccess = false;

                        //SQL生成
                        cmd.CommandText = getSql(data.SerialNo, data.Parts);

                        //SQL実行
                        using (SqliteDataReader reader = cmd.ExecuteReader())
                        {
                            //先頭のみ読込（複数ないはずだから）
                            isSuccess = reader.Read();

                            //タイムスタンプあるかな
                            if (isSuccess)
                            {
                                no = reader[0].ToString();
                                insPoint = reader[3].ToString();
                                lineNum = reader[4].ToString();
                                action = reader[5].ToString();

                                Console.WriteLine("{0}、{1}、{2}、{3}、{4}、{5}", reader[0], reader[1], reader[2], reader[3], reader[4], reader[5]);

                                //タイムスタンプ開始終了取得
                                startTs = DateTime.ParseExact(reader[1].ToString(), "yyyy-MM-dd hh:mm:ss", null);
                                endTs = DateTime.ParseExact(reader[2].ToString(), "yyyy-MM-dd hh:mm:ss", null);
                            }
                        }

                        //無かったら残念
                        if (!isSuccess)
                        {
                            Console.WriteLine("タイムスタンプが無かったです。失敗。");
                            Console.WriteLine();
                            continue;
                        }

                        //タイムスタンプ開始終了の中間時刻を算出
                        TimeSpan ts = endTs - startTs;
                        midTs = startTs.AddSeconds(ts.TotalSeconds / 2);

                        //画像ファイルを開く
                        using (Bitmap bmp = new Bitmap(Path.Combine(inPath, data.File)))
                        {

                            //画像ファイルから撮影日時を取得
                            foreach (var item in bmp.PropertyItems)
                            {
                                //データの型を判断
                                if (item.Id == 0x9003 && item.Type == 2)
                                {
                                    string date = midTs.ToString("yyyy:MM:dd hh:mm:ss\0");

                                    Console.WriteLine("変更前＝{0}、変更後＝{1}", Encoding.ASCII.GetString(item.Value), date);

                                    //撮影日時を変更設定
                                    item.Value = Encoding.ASCII.GetBytes(date);
                                    item.Len = item.Value.Length;
                                    bmp.SetPropertyItem(item);
                                    break;
                                }
                            }

                            //保存
                            bmp.Save(Path.Combine(outPath, data.File));
                            Console.WriteLine("成功");
                            Console.WriteLine();
                            cnt++;
                        }


                    }

                    con.Close();

                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("何かでエラーになりました");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine();

                if(ex.InnerException != null)
                {
                    Console.WriteLine(ex.InnerException.Message);
                    Console.WriteLine(ex.InnerException.StackTrace);
                    if(ex.InnerException.InnerException != null)
                    {
                        Console.WriteLine(ex.InnerException.InnerException.Message);
                        Console.WriteLine(ex.InnerException.InnerException.StackTrace);
                    }
                }
            }
            Console.WriteLine("出力完了　合計：" + cnt.ToString() + "件");
            Console.WriteLine();
            Console.Write("キーを押すと終了します。：");
            Console.ReadLine();
        }



        private static string getSql(string paramNo, string parts)
        {

            string sqlBase = "select " +
                "FEC_FacilitySpec.SerialNumber, " +
                "FEC_DeviceLog.TimeStampStart," +
                "FEC_DeviceLog.TimeStampEnd," +
                "FEC_InspectionDetail.InspectionPoint," +
                "FEC_DeviceLog.LineNumber," +
                "FEC_DeviceLog.Action " +
                "from FEC_DeviceLog left outer join FEC_FacilitySpec on FEC_DeviceLog.FacilitySpecId = FEC_FacilitySpec.FacilitySpecId " +
                "left outer join FEC_DamageLog on FEC_DeviceLog.DamageLogId = FEC_DamageLog.DamageLogId and FEC_DeviceLog.LineNumber = FEC_DamageLog.LineNumber " +
                "left outer join FEC_Inspection on FEC_DeviceLog.InspectionId = FEC_Inspection.InspectionId " +
                "left outer join FEC_InspectionDetail on FEC_Inspection.InspectionId = FEC_InspectionDetail.InspectionId and FEC_DeviceLog.LineNumber = FEC_InspectionDetail.LineNumber " +
                "where " +
                "FEC_FacilitySpec.SerialNumber like \"{0}%\" and " +
                "FEC_InspectionDetail.InspectionPoint = \"{1}\"";

            return string.Format(sqlBase,paramNo, parts);

        }
    }

    public class DataCsv
    {
        public string File { get; set; }
        public string SerialNo { get; set; }
        public string Parts { get; set; }
    }


}
