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
            string target_accdb = @"\app\workspace\販売管理.accdb";
            string export_xlsx = @"\app\workspace\販売管理.xlsx";

            OleDbConnection myConAccess;
            OleDbCommand myCommand;
            OleDbDataReader myReader;

            // *************************************
            // System.Data.OleDb
            // *************************************
            myConAccess = new OleDbConnection();
            myConAccess.ConnectionString =
                $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={target_accdb};";

            // 接続を開く
            try
            {
                myConAccess.Open();

                string myQuery;

                myQuery =
                    $"select * into [Excel 12.0 xml;DATABASE={export_xlsx}].商品マスタ from 商品マスタ";

                using (myCommand = new OleDbCommand())
                {
                    // *********************
                    // 接続
                    // *********************
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
