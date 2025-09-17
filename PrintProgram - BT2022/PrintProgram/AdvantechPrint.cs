using Seagull.BarTender.Print;
using System;
using System.Windows.Forms;
using Telerik.WinControls.UI;

namespace PrintProgram
{
    /// <summary>
    /// AdvantechPrint 類別提供標籤列印相關功能。
    /// </summary>
    static class AdvantechPrint
    {
        /// <summary>
        /// 列印標籤。
        /// </summary>
        /// <param name="MAC">要列印的 MAC 位址。</param>
        /// <param name="LabelPath">標籤檔案的路徑。</param>
        /// <param name="List_Msg">用於顯示印表機名稱的 RadListControl 控制項。</param>
        /// <returns>列印成功回傳 true，失敗回傳 false。</returns>
        [STAThread]
        public static bool PrintLabel(string MAC, string LabelPath, RadListControl List_Msg)
        {
            //partNumber, choiceRdb, printlabe, ecn, biosVer, bioscs,radTxt_SN.Text, radTxt_Mac.Text
            try
            {
                // 建立 Engine 物件，並初始化為 null
                Engine engine = null;
                // 建立 LabelFormatDocument 物件，並初始化為 null
                LabelFormatDocument btFormat = null;
                // 實例化 Engine，啟動 BarTender 列印引擎
                engine = new Engine();
                engine.Start();

                // 開啟指定路徑的標籤格式文件
                // 標籤檔案（通常是 .btw 格式）
                btFormat = engine.Documents.Open(LabelPath);
                // 設定標籤上的 MAC 子字串為大寫
                btFormat.SubStrings["MAC"].Value = MAC.ToUpper();

                // 設定要列印的標籤份數為 1
                btFormat.PrintSetup.IdenticalCopiesOfLabel = int.Parse("1"); //列印標籤數

                // 取得目前設定的印表機名稱，並顯示在 List_Msg 控制項
                string printerName = btFormat.PrintSetup.PrinterName;
                List_Msg.Items.Add($"使用的印表機：{printerName}");

                /*
                 * 在 BarTender Designer 軟體中，開啟你的 .btw 標籤檔案。
                 * 點選「檔案」→「列印」→「選擇印表機」。
                */
                // 在 BarTender Designer 軟體中，開啟你的 .btw 標籤檔案。
                btFormat.Print();// 執行列印
                // 停止 BarTender 列印引擎
                engine.Stop();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
    }
}
