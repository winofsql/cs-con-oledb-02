using System;
using System.Data.OleDb;
using System.IO;
using System.Net;

namespace cs_con_oledb_02
{
    class Program
    {
        static void Main(string[] args)
        {
            string target_file = @"\app\workspace\販売管理.accdb";
            string link_file = @"\app\workspace\商品マスタ.xlsx";

            // string target_file = @"\app\workspace\販売管理.xlsx";
            // string link_file = @"\app\workspace\販売管理.accdb";

            OleDbConnection myConAccess;
            OleDbCommand myCommand;

            // *************************************
            // System.Data.OleDb
            // *************************************
            myConAccess = new OleDbConnection();
            myConAccess.ConnectionString =
                $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={target_file};";
                // @$"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={target_file};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""";

            // 接続を開く
            try
            {
                myConAccess.Open();

                string myQuery;

                myQuery =
                    $"select * into [Excel 12.0 xml;DATABASE={link_file}].商品マスタ from 商品マスタ";
                    // $"select * into [;DATABASE={link_file}].商品マスタ2 from 商品マスタ";
                    // $"drop table [;DATABASE={link_file}].商品マスタ2";

                using (myCommand = new OleDbCommand())
                {
                    try
                    {
                        // コマンドオブジェクトに接続をセット
                        myCommand.Connection = myConAccess;
                        myCommand.CommandText = myQuery;
                        myCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                myConAccess.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("接続エラーです:" + ex.Message);
            }

            Console
                .WriteLine("処理が終了しました : Enter キーを入力してください");
            Console.ReadLine();
        }
    }
}
