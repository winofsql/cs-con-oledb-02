using System;
using System.Data.OleDb;

namespace cs_con_oledb_02
{
    class Program
    {
        static void Main(string[] args)
        {
            OleDbConnection myConAccess;
            OleDbCommand myCommand;
            OleDbDataReader myReader;

            // *************************************
            // System.Data.OleDb
            // *************************************
            myConAccess = new OleDbConnection();
            myConAccess.ConnectionString =
                string
                    .Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};",
                    @"\app\workspace\販売管理.accdb");

            // 接続を開く
            try
            {
                myConAccess.Open();

                string myQuery;

                myQuery = @"select * into [Excel 12.0 xml;DATABASE=\app\workspace\販売管理.xlsx].商品マスタ from 商品マスタ";

                using (myCommand = new OleDbCommand()) {

                    // *********************
                    // 接続
                    // *********************
                    try {
                        // コマンドオブジェクトに接続をセット
                        myCommand.Connection = myConAccess;
                        myCommand.CommandText = myQuery;
                        myCommand.ExecuteNonQuery();

                    }
                    catch (Exception ex) {
                        Console.WriteLine(ex.Message);
                        
                    }
                }

                myConAccess.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("接続エラーです:" + ex.Message);
            }

            Console.WriteLine("処理が終了しました : Enter キーを入力してください");
            Console.ReadLine();

        }
    }
}
