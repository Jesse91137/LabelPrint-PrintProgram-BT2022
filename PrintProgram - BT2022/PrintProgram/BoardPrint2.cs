using Seagull.BarTender.Print;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Telerik.WinControls.UI;

namespace PrintProgram
{
    /// <summary>
    /// 提供標籤列印相關功能的靜態類別。
    /// </summary>
    static class BoardPrint2
    {
        /// <summary>
        /// 列印指定路徑的標籤檔，並將序號清單依序填入標籤欄位，於訊息清單顯示印表機名稱。
        /// </summary>
        /// <param name="LabelPath">標籤檔案的完整路徑。</param>
        /// <param name="SNList">要填入標籤的序號清單。</param>
        /// <param name="List_Msg">用於顯示印表機名稱的 RadListControl 控制項。</param>
        /// <returns>列印成功回傳 true，失敗回傳 false。</returns>
        [STAThread]
        public static bool PrintLabel(string LabelPath, List<string> SNList, RadListControl List_Msg)
        {
            try
            {

                // 建立 BarTender 列印引擎物件
                Engine engine = null;
                // 建立標籤格式文件物件
                LabelFormatDocument btFormat = null;
                // 啟動 BarTender 列印引擎
                engine = new Engine();
                engine.Start(); // 啟動列印引擎
                // 開啟指定路徑的標籤格式文件
                // 標籤檔案（通常是 .btw 格式）
                // 在 BarTender Designer 軟體中，開啟你的 .btw 標籤檔案。
                btFormat = engine.Documents.Open(LabelPath); // 開啟標籤檔案

                // 依序設定標籤欄位的值
                for (int i = 0; i < SNList.Count; i++)
                {
                    string SN_Name = "SN" + (i + 1).ToString();// 動態生成欄位名稱，如 SN1, SN2, ...
                    btFormat.SubStrings[SN_Name].Value = SNList[i].ToString().ToUpper(); //標籤檔中所設定的欄位名稱 。
                }
                btFormat.PrintSetup.IdenticalCopiesOfLabel = int.Parse("1"); //列印標籤數

                // 取得目前設定的印表機名稱，並顯示在 List_Msg 控制項
                string printerName = btFormat.PrintSetup.PrinterName;
                List_Msg.Items.Add($"使用的印表機：{printerName}");

                /*在 BarTender Designer 軟體中，開啟你的 .btw 標籤檔案。
                 * 點選「檔案」→「列印」→「選擇印表機」。
                 */
                // 在 BarTender Designer 軟體中，開啟你的 .btw 標籤檔案。
                btFormat.Print(); // 執行列印
                engine.Stop(); // 停止 BarTender 引擎

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message); // 顯示錯誤訊息
                return false;
            }
        }
    }
}
