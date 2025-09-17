using Seagull.BarTender.Print;
using System;
using Telerik.WinControls.UI;

namespace PrintProgram
{
    static class AccPrint
    {
        /// <summary>
        /// 列印標籤。
        /// </summary>
        /// <param name="LabelPath">標籤檔案的完整路徑。</param>
        /// <param name="WO">工單號碼。</param>
        /// <param name="PN">料號。</param>
        /// <param name="QTY">數量。</param>
        /// <param name="No_Number">流水號。</param>
        /// <param name="List_Msg">用於顯示印表機名稱的 RadListControl 控制項。</param>
        /// <returns>若列印成功則回傳 true，否則回傳 false。</returns>
        [STAThread]
        public static bool PrintLabel(string LabelPath, string WO, string PN, string QTY, string No_Number, RadListControl List_Msg)
        {
            try
            {
                // 建立 BarTender 列印引擎物件
                Engine engine = null;
                // 建立標籤格式文件物件
                LabelFormatDocument btFormat = null;

                // 啟動 BarTender 列印引擎
                engine = new Engine();
                engine.Start();
                // 開啟指定路徑的標籤格式文件
                // 標籤檔案（通常是 .btw 格式）
                btFormat = engine.Documents.Open(LabelPath);

                // 設定標籤上的工單號碼
                btFormat.SubStrings["WO"].Value = WO;
                // 設定標籤上的料號
                btFormat.SubStrings["PN"].Value = PN;
                // 設定標籤上的數量
                btFormat.SubStrings["QTY"].Value = QTY;
                // 設定標籤上的流水號
                btFormat.SubStrings["No_Number"].Value = No_Number;

                // 設定列印標籤的份數為 1
                btFormat.PrintSetup.IdenticalCopiesOfLabel = int.Parse("1");

                // 取得目前設定的印表機名稱，並顯示在 List_Msg 控制項
                string printerName = btFormat.PrintSetup.PrinterName;
                List_Msg.Items.Add($"使用的印表機：{printerName}");

                /*在 BarTender Designer 軟體中，開啟你的 .btw 標籤檔案。
                 * 點選「檔案」→「列印」→「選擇印表機」。
                 */
                // 執行列印
                btFormat.Print();// （BarTender 會依照標籤檔案設定的印表機執行）
                // 停止 BarTender 列印引擎
                engine.Stop();// 停止 BarTender 列印引擎，釋放資源。
                // 列印成功回傳 true
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
