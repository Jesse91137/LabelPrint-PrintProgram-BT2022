using Seagull.BarTender.Print;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Telerik.WinControls.UI;

namespace PrintProgram
{
    /// <summary>
    /// 提供 OEM 標籤列印相關功能的靜態類別。
    /// </summary>
    static class OemPrint
    {
        /// <summary>
        /// 列印標籤。
        /// </summary>
        /// <param name="LabelPath">標籤檔案的完整路徑。</param>
        /// <param name="WO">工單號。</param>
        /// <param name="QTY">數量。</param>
        /// <param name="Weight">重量。</param>
        /// <param name="SNList">序號清單。</param>
        /// <param name="List_Msg">用於顯示印表機名稱的 RadListControl 控制項。</param>
        /// <returns>列印成功回傳 true，失敗回傳 false。</returns>
        [STAThread]
        public static bool PrintLabel(string LabelPath, string WO, string QTY, string Weight, List<string> SNList, RadListControl List_Msg)
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
                btFormat = engine.Documents.Open(LabelPath); // 開啟標籤檔案

                // 設定標籤欄位值
                btFormat.SubStrings["WO"].Value = WO;// 設定標籤上的工單號碼
                //btFormat.SubStrings["PC"].Value = PC;
                btFormat.SubStrings["QTY"].Value = QTY;// 設定標籤上的數量
                btFormat.SubStrings["Weight"].Value = Weight;

                // 設定序號欄位值
                for (int i = 0; i < SNList.Count; i++)
                {
                    string SN_Name = "SN" + (i + 1).ToString();// 動態生成欄位名稱，如 SN1, SN2, ...
                    btFormat.SubStrings[SN_Name].Value = SNList[i].ToString().ToUpper(); // 標籤檔中所設定的欄位名稱
                }
                btFormat.PrintSetup.IdenticalCopiesOfLabel = int.Parse("1"); // 列印標籤數量

                // 取得目前設定的印表機名稱，並顯示在 List_Msg 控制項
                string printerName = btFormat.PrintSetup.PrinterName;
                List_Msg.Items.Add($"使用的印表機：{printerName}");

                /*在 BarTender Designer 軟體中，開啟你的 .btw 標籤檔案。
                * 點選「檔案」→「列印」→「選擇印表機」。
                */
                // 在 BarTender Designer 軟體中，開啟你的 .btw 標籤檔案。
                btFormat.Print(); // 執行列印
                engine.Stop(); // 停止列印引擎

                return true; // 列印成功
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message); // 顯示錯誤訊息
                return false; // 列印失敗
            }
        }
    }
}
