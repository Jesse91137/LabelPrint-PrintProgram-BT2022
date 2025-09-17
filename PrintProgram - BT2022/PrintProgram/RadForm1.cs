using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls.UI;
using Seagull.BarTender.Print;
using Newtonsoft.Json;
using System.IO;
using System.Net;
using System.Threading;
using System.Diagnostics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace PrintProgram
{

    /// <summary>
    /// RadForm1 類別為主視窗，負責處理標籤列印、FTP 檔案操作、資料庫查詢、Excel 讀取等功能。
    /// </summary>
    public partial class RadForm1 : Telerik.WinControls.UI.RadForm
    {
        #region Fields
        /// <summary>
        /// 應用程式名稱常數。
        /// </summary>
        private const string appName = "PrintProgram";
        /// <summary>
        /// 資料來源常數字串。
        /// </summary>
        private const string dataSourced = "Data Sourced";

        /// <summary>
        /// BarTender 列印引擎物件。
        /// </summary>
        private Engine engine = null; // The BarTender Print Engine
        /// <summary>
        /// 目前開啟的 BarTender 標籤格式文件。
        /// </summary>
        private LabelFormatDocument format = null; // The currently open Format
        /// <summary>
        /// 指示視窗是否正在關閉，避免縮圖載入時發生例外。
        /// </summary>
        private bool isClosing = false; // Set to true if we are closing. This helps discontinue thumbnail loading.

        /// <summary>
        /// 目前資料夾下所有標籤檔案名稱清單。
        /// </summary>
        private string[] browsingFormats; // The list of filenames in the current folder
        /// <summary>
        /// 以標籤檔名為索引的 ListViewItem 雜湊表，用於追蹤已載入圖片的標籤。
        /// </summary>
        Hashtable listItems; // A hash table containing ListViewItems and indexed by format name.
                             // It keeps track of what formats have had their image loaded.

        /// <summary>
        /// 標籤縮圖產生佇列，存放待產生縮圖的標籤索引。
        /// </summary>
        Queue<int> generationQueue; // A queue containing indexes into browsingFormats
                                    // to facilitate the generation of thumbnails

        /// <summary>
        /// 標籤瀏覽器中最上方可見的索引。
        /// </summary>
        int topIndex; // The top visible index in the lstLabelBrowser
        /// <summary>
        /// 標籤瀏覽器中目前選取的索引。
        /// </summary>
        int selectedIndex; // The selected index in the lstLabelBrowser
        #endregion

        #region 宣告變數
        /// <summary>
        /// SN：序號欄位，儲存目前操作的序號字串。
        /// Item：項次欄位，儲存目前操作的項次字串。
        /// WO：工單欄位，儲存目前操作的工單字串。
        /// PN：料號欄位，儲存目前操作的料號字串。
        /// PC：電源線編號欄位，儲存目前操作的電源線編號字串。
        /// Qty：數量欄位，儲存目前操作的數量字串。
        /// VER：BIOS版本欄位，儲存目前操作的BIOS版本字串。
        /// No_Number：無編號欄位，儲存目前操作的無編號字串。
        /// ENGSR：工程SR欄位，儲存目前操作的工程SR字串。
        /// </summary>
        string SN = string.Empty, Item = string.Empty, WO = string.Empty, PN = string.Empty, PC = string.Empty, Qty = string.Empty, VER = string.Empty, No_Number = string.Empty, ENGSR = string.Empty;

        /// <summary>
        /// version_old：舊版本號，儲存目前程式的舊版本字串。
        /// version_new：新版本號，儲存目前程式的最新版本字串。
        /// </summary>
        string version_old = "", version_new = "";

        /// <summary>
        /// 全域資料集物件，儲存查詢結果或暫存資料。
        /// </summary>
        DataSet dataSet = new DataSet();
        /// <summary>
        /// Advantech Excel 資料表，儲存安勤標籤資料。
        /// </summary>
        DataTable advantechDt;
        /// <summary>
        /// SQL 檔案大小，儲存目前 SQL 檔案的大小（用於進度計算）。
        /// </summary>
        public long SQLfile_size;
        /// <summary>
        /// 標籤預覽圖片的儲存路徑，預設為啟動目錄下的 exp.jpg。
        /// </summary>
        private string _bmp_path = Application.StartupPath + @"\exp.jpg";
        /// <summary>
        /// 目前選擇的 BarTender 標籤檔案路徑。
        /// </summary>
        private string _btw_path = "";
        /// <summary>
        /// 目前選擇的印表機名稱。
        /// </summary>
        string _PrinterName = "";
        /// <summary>
        /// SQL 檔案總大小（用於上傳進度計算）。
        /// </summary>
        public long Sum_Of_SQLfile_size = 0;
        /// <summary>
        /// 標示 FTP 上傳或下載結果是否成功。
        /// </summary>
        public bool TempUploadResult = false;

        /// <summary>
        /// 標籤檔案路徑：Accpath 為配件箱、Boardpath 為主板箱、Boardpath2~4 為主板箱不同版本、Systempath 為系統箱、Systempath2 為系統箱第二版、PhilEPC 為 PHIL_EPC 標籤、小張標籤、AdvantechLabel 為安勤標籤。
        /// </summary>
        public string Accpath, Boardpath, Boardpath2, Boardpath3, Boardpath4, Systempath, Systempath2 = string.Empty, PhilEPC, AdvantechLabel;
        #endregion

        #region FTP使用
        /// <summary>
        /// FTP 下載工廠名稱，儲存目前 FTP 下載工廠的字串。
        /// </summary>
        public string ftpdlfactory;
        /// <summary>
        /// FTP 伺服器位址，儲存目前 FTP 伺服器的 IP 或網域名稱。
        /// </summary>
        public string ftpServer;
        /// <summary>
        /// FTP 使用者名稱，儲存目前 FTP 登入帳號。
        /// </summary>
        public string ftpuser;
        /// <summary>
        /// FTP 密碼，儲存目前 FTP 登入密碼。
        /// </summary>
        public string ftppassword;
        /// <summary>
        /// FTP 檔案路徑、上傳檔案名稱、下載檔案名稱、檔案名稱、ini 設定檔路徑。
        /// </summary>
        public string ftpfilepath, ftpPutFile, ftpGetFile, filename, ini_filepath;
        /// <summary>
        /// 標籤檔案下載目錄，預設為啟動目錄下的 Btw_Folder。
        /// </summary>
        string download_Path = System.Windows.Forms.Application.StartupPath + "\\" + "Btw_Folder";
        /// <summary>
        /// 下載檔案名稱、上傳檔案名稱。
        /// </summary>
        public string DLfilename, UPfilename;
        #endregion

        /// <summary>
        /// 設定檔名稱，預設為 Setup.ini。
        /// </summary>
        public string Cfgname = "Setup.ini";

        /// <summary>
        /// 全域 Mutex 物件，用於判斷程式是否重複執行。
        /// </summary>
        static Mutex m;
        /// <summary>
        /// INI 設定檔操作物件，提供讀寫 INI 檔案功能。
        /// </summary>
        SetupIniIP ini = new SetupIniIP();

        /// <summary>
        /// 建立 RadForm1 類別的新執行個體 <see cref="RadForm1"/> class.
        /// </summary>
        public RadForm1()
        {
            InitializeComponent();
            // 讓 List_Oem_Msg 文字自動換行顯示
            List_Oem_Msg.AutoSizeItems = true;
        }

        /// <summary>
        /// RadForm1_Load 事件處理函式。
        /// 當主視窗載入時，執行下列初始化步驟：
        /// 1. 取得 FTP 連線資訊。
        /// 2. 檢查程式是否重複執行，若已執行則顯示提示並關閉視窗。
        /// 3. 讀取舊版與新版程式版本號，若有新版本則提示並啟動自動更新。
        /// 4. 初始化各項 UI 控制項的啟用狀態。
        /// 5. 依照 INI 設定檔載入各種標籤檔案路徑。
        /// 6. 若 AdvantechMAC 設定為 True，則載入安勤標籤 Excel 資料並切換頁籤；否則僅顯示標準頁籤。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">包含事件資料的 EventArgs。</param>
        private void RadForm1_Load(object sender, EventArgs e)
        {
            Getftp("Print");
            //判斷程式是否重複執行
            if (IsMyMutex("PrintProgram64"))
            {
                MessageBox.Show("程式正在執行中!!");
                Dispose();//關閉
            }
            #region 系統更新
            version_old = ini.IniReadValue("Version", "version", Cfgname);
            version_new = selectVerSQL_new("PrintProgram64");//BT2022版
            //lbl_ver.Text = "VER:V" + version_old;
            //判斷版本號
            int v_old = Convert.ToInt32(version_old.Replace(".", ""));
            int v_new = Convert.ToInt32(version_new.Replace(".", ""));
            if (v_old < v_new)
            {
                MessageBox.Show("有新版本更新VER: V" + version_new);
                autoupdate();
            }
            #endregion

            // 以下程式碼為初始化各項 UI 控制項的啟用狀態，並根據 INI 設定檔載入各種標籤檔案路徑
            txt_Pn.Enabled = false;         // 料號欄位預設不可編輯
            txt_Qty.Enabled = false;        // 數量欄位預設不可編輯
            txt_Pc.Enabled = false;         // 電源線編號欄位預設不可編輯
            txt_Bios_Ver.Enabled = false;   // BIOS版本欄位預設不可編輯
            txt_Sn.Enabled = false;         // 序號欄位預設不可編輯
            txt_No_Number.Enabled = false;  // 無編號欄位預設不可編輯
            // 依照 INI 設定檔載入各種標籤檔案路徑
            Accpath = download_Path + "\\" + ini.IniReadValue("Option", "Acc_Carton_Label_Name", Cfgname);                // 配件箱標籤路徑
            Boardpath = download_Path + "\\" + ini.IniReadValue("Option", "Board_Carton_Label_Name", Cfgname);            // 主板箱標籤路徑
            Boardpath2 = download_Path + "\\" + ini.IniReadValue("Option", "Board_Carton_Label2_Name ", Cfgname);         // 主板箱第二版標籤路徑
            Boardpath3 = download_Path + "\\" + ini.IniReadValue("Option", "Board_Carton_Label3_Name ", Cfgname);         // 主板箱第三版標籤路徑
            Boardpath4 = download_Path + "\\" + ini.IniReadValue("Option", "Board_Carton_Label4_Name ", Cfgname);         // 主板箱第四版標籤路徑
            Systempath = download_Path + "\\" + ini.IniReadValue("Option", "System_Carton_Label_Name", Cfgname);          // 系統箱標籤路徑
            Systempath2 = download_Path + "\\" + ini.IniReadValue("Option", "System_Carton_Label2_Name", Cfgname);        // 系統箱第二版標籤路徑
            PhilEPC = download_Path + "\\" + ini.IniReadValue("Option", "Phil_EPC_WHL_Small_Name", Cfgname);              // PHIL_EPC小張標籤路徑
            AdvantechLabel = download_Path + "\\" + ini.IniReadValue("Option", "AdvantechLabelName ", Cfgname);           // 安勤標籤路徑
            string Advantech = ini.IniReadValue("Option", "AdvantechMAC", Cfgname);                                       // 取得安勤MAC設定
            string AdvantechPath = ini.IniReadValue("Option", "AdvantechMACPath", Cfgname);                               // 取得安勤Excel路徑
            //version_old = ini.IniReadValue("Version", "version", filename);
            //version_new = selectVerSQL_new("E-SOP");

            // 依照 ini 設定載入版本資訊，並根據 AdvantechMAC 設定切換頁籤與載入安勤 Excel 資料
            radLabelElement1.Text = version_new;
            if (Advantech == "True")
            {
                advantechDt = new DataTable();
                // 隱藏工單設定、標準品、OEM、OEM上傳、預設頁籤
                this.Wo_Set_Page.Parent = null;
                this.Standard_Page.Parent = null;
                this.Oem_Page.Parent = null;
                this.Oem_Page_UP.Parent = null;
                this.radPageViewPage1.Parent = null;
                // 載入安勤 Excel 資料表
                advantechDt = LoadExcelAsDataTable(AdvantechPath);
            }
            else
            {
                // 隱藏安勤標籤頁籤
                this.radPageViewPage2.Parent = null;
            }
        }

        /// <summary>
        /// Timer1_Tick 事件處理函式。
        /// 每次計時器觸發時，更新 System_date_ID 控制項的文字為目前的日期與時間。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">包含事件資料的 EventArgs。</param>
        private void Timer1_Tick(object sender, EventArgs e)
        {
            System_date_ID.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }


        #region 標準品外箱Label設定
        #region 序號輸入
        /*
         * 驗證 SN 列表與數量是否吻合；根據選擇的列印製程 (Acc/System/Board) 組裝 SNList，
         * 處理分頁（System 超過 6、Board 超過 26、Board 大量分組）並呼叫相應的 PrintLabel 函式；更新 UI 與 List_Msg；處理列印結果狀態
         */
        /// <summary>
        /// Btn_Print_Click 方法為「列印」按鈕的事件處理函式。
        /// 根據目前選擇的列印模式（Acc、System、Board），將輸入的工單、料號、序號等資訊組合後，呼叫對應的列印方法進行標籤列印。
        /// 支援多頁列印（如 System 超過 6 組、Board 超過 26 組時分頁），並於列印完成或失敗時更新 UI 狀態。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">包含事件資料的 EventArgs。</param>
        private void Btn_Print_Click(object sender, EventArgs e)
        {
            try
            {
                // 檢查序號與數量輸入狀態，並提示使用者
                List_Msg.Items.Clear();
                // 若序號列表為空且非配件箱模式，提示尚未輸入序號
                if (SN_GV.Rows.Count < 1)
                {
                    if (!rdb_Acc.IsChecked)
                    {
                        List_Msg.Items.Add("尚未輸入序號.....");
                        // 將焦點移至序號輸入欄位
                        txt_Sn.Focus();
                        return;
                    }
                }
                // 若序號數量與輸入數量不符且非配件箱模式，提示數量不符
                if (SN_GV.Rows.Count != int.Parse(txt_Qty.Text))
                {
                    // 檢查序號數量與輸入數量是否一致，若不符則提示並將焦點移至序號輸入欄位
                    if (!rdb_Acc.IsChecked)
                    {
                        List_Msg.Items.Add("數量與輸入序號數量不符");
                        // 將焦點移至序號輸入欄位
                        txt_Sn.Focus();
                        return;
                    }
                }

                // 將列印按鈕設為不可用，清空序號欄位並設為不可用
                btn_Print.Enabled = false;
                txt_Sn.Text = string.Empty;
                txt_Sn.Enabled = false;
                // 建立序號清單與項次清單
                List<string> SNList = new List<string>();
                List<int> NOList = new List<int>();

                // 列印檔案路徑與列印結果旗標
                string printlabe = string.Empty;
                //列印結果旗標，預設為 false
                bool printresult = false;

                #region 欄位轉換成值
                // 以下程式碼負責將 UI 輸入欄位的值轉換成對應的變數，並處理空值與預設值
                WO = txt_Wo.Text.Trim(); // 工單
                PN = txt_Pn.Text.Trim(); // 料號
                if (!string.IsNullOrWhiteSpace(txt_Pc.Text.Trim()))
                {
                    PC = txt_Pc.Text.Trim().Replace(",", " "); // 電源線編號，將逗號換成空格
                }
                else
                {
                    PC = "NA"; // 若未輸入則預設為 NA
                }
                if (string.IsNullOrWhiteSpace(txt_Qty.Text))
                {
                    List_Msg.Items.Add("尚未輸入數量"); // 提示未輸入數量
                    txt_Qty.Focus();
                    return;
                }
                else
                {
                    Qty = txt_Qty.Text.Trim(); // 數量
                }
                if (!string.IsNullOrWhiteSpace(txt_Bios_Ver.Text.Trim()))
                {
                    VER = txt_Bios_Ver.Text.Trim(); // BIOS 版本
                }
                else
                {
                    VER = "NA"; // 若未輸入則預設為 NA
                    txt_Bios_Ver.Text = "NA";
                }
                ENGSR = txt_engSR.Text.Trim(); // 工程SR
                if (!string.IsNullOrWhiteSpace(txt_No_Number.Text.Trim()))
                {
                    No_Number = "(" + txt_No_Number.Text.Trim().ToUpper() + ")"; // 無編號，轉大寫並加括號
                }
                else
                {
                    No_Number = string.Empty; // 若未輸入則為空字串
                }
                #endregion

                // 以下程式碼負責處理配件箱列印流程
                List_Msg.Items.Add("列印中.....");
                if (rdb_Acc.IsChecked)
                {
                    printlabe = Accpath;
                    // 呼叫 AccPrint.PrintLabel 方法進行配件箱標籤列印
                    // 傳入標籤路徑、工單、料號、數量、無編號等參數
                    if (AccPrint.PrintLabel(printlabe, WO, PN, Qty, No_Number, List_Msg) == true)
                    {
                        printresult = true; // 列印成功
                    }
                    else
                    {
                        printresult = false; // 列印失敗
                    }
                }
                else if (rdb_System.IsChecked)
                {

                    printlabe = Systempath;
                    //20210818改system列印超過6組印第二頁
                    #region MyRegion
                    // 以下程式碼負責處理 System 列印超過 6 組時的分頁邏輯
                    if (int.Parse(txt_Qty.Text) > 6)
                    {
                        // 先將前 6 筆序號加入 SNList
                        for (int i = 0; i < 6; i++)
                        {
                            SNList.Add(SN_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }
                        // 呼叫 SystemPrint 列印第一頁（前 6 筆序號）
                        if (SystemPrint.PrintLabel(printlabe, WO, PN, PC, VER, Qty, No_Number, ENGSR, SNList, List_Msg) == true)
                        {
                            SNList.Clear();
                            // 將第 7 筆到最後一筆序號加入 SNList
                            for (int i = 6; i < SN_GV.Rows.Count; i++)
                            {
                                SNList.Add(SN_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                            }
                            // 呼叫 SystemPrint2 列印第二頁（剩餘序號）
                            if (SystemPrint2.PrintLabel(Systempath2, SNList) == true)
                            {
                                printresult = true; // 列印成功
                            }
                            else
                            {
                                printresult = false; // 列印失敗
                            }
                        }
                        else
                        {
                            printresult = false; // 列印失敗
                        }
                    }
                    #endregion
                    else
                    {
                        // 這段程式碼將 SN_GV 中所有序號欄位的值加入 SNList，然後呼叫 SystemPrint.PrintLabel 方法進行列印，根據回傳結果設定 printresult 狀態。
                        for (int i = 0; i < SN_GV.Rows.Count; i++)
                        {
                            SNList.Add(SN_GV.Rows[i].Cells["SN"].Value.ToString().Trim()); // 將每一列的序號加入序號清單
                        }
                        if (SystemPrint.PrintLabel(printlabe, WO, PN, PC, VER, Qty, No_Number, ENGSR, SNList, List_Msg) == true)
                        {
                            printresult = true; // 列印成功
                        }
                        else
                        {
                            printresult = false; // 列印失敗
                        }

                    }
                }
                else
                {
                    // 這一行設定主板箱標籤的檔案路徑，供後續列印使用
                    printlabe = Boardpath;
                    int number = int.Parse(txt_Qty.Text);     // 輸入的數字
                    int groupSize = 36; // 每組的大小

                    if (number <= 26)
                    {
                        // 小於等於 26 的情況，直接列印所有序號
                        for (int i = 0; i < SN_GV.Rows.Count; i++)
                        {
                            // 將每一列的序號加入序號清單
                            SNList.Add(SN_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }
                        // 呼叫 BoardPrint.PrintLabel 方法進行主板箱標籤列印
                        if (BoardPrint.PrintLabel(printlabe, WO, PN, VER, Qty, No_Number, ENGSR, SNList, List_Msg) == true)
                        {
                            printresult = true; // 列印成功
                        }
                        else
                        {
                            printresult = false; // 列印失敗
                        }
                    }
                    else
                    {
                        //大於26情況, 先印1~26
                        // 以下程式碼為主板箱標籤列印流程，當數量大於 26 時，先列印前 26 筆序號，
                        // 之後每 36 筆為一組分批列印，直到所有序號都處理完畢。
                        // 1. 先將前 26 筆序號加入 SNList，呼叫 BoardPrint.PrintLabel 列印第一組
                        // 2. 清空 SNList，設定 startNumber 為 27，endNumber 為 26 + groupSize (36)
                        // 3. 使用 while 迴圈，當剩餘數量大於 endNumber 時，
                        //    - 依序將序號與項次加入 SNList、NOList
                        //    - 呼叫 BoardPrint4.PrintLabel 列印該組
                        //    - 若列印失敗則跳出迴圈
                        //    - 更新 startNumber、endNumber，group++
                        for (int i = 0; i < 26; i++)
                        {
                            SNList.Add(SN_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }
                        // 呼叫 BoardPrint.PrintLabel 方法列印前 26 筆序號
                        if (BoardPrint.PrintLabel(printlabe, WO, PN, VER, Qty, No_Number, ENGSR, SNList, List_Msg) == true)
                        {
                            //groupSize=36 ,每36為一組
                            SNList.Clear();
                            int startNumber = 27; // 第一組的起始數字
                            int endNumber = 26 + groupSize; // 第一組的結束數字
                            int group = 2; // 組數

                            // 以下 while 迴圈用於主板箱標籤分組列印，每組最多 36 筆序號
                            // 1. 只要剩餘數量大於 endNumber，就持續分組
                            // 2. 每組將序號與項次加入 NOList、SNList
                            // 3. 呼叫 BoardPrint4.PrintLabel 進行分組列印
                            // 4. 若列印失敗則跳出迴圈
                            // 5. 更新 startNumber、endNumber，進入下一組
                            while (number > endNumber)
                            {
                                // 清空本組的序號與項次清單
                                NOList.Clear();
                                SNList.Clear();

                                // 依序將本組的序號與項次加入清單
                                for (int i = startNumber; i <= endNumber; i++)
                                {
                                    NOList.Add(i); // 加入項次編號
                                    SNList.Add(SN_GV.Rows[i - 1].Cells["SN"].Value.ToString().Trim()); // 加入序號
                                }

                                // 呼叫 BoardPrint4.PrintLabel 方法進行主板箱分組列印，根據回傳結果設定 printresult 狀態
                                if (BoardPrint4.PrintLabel(Boardpath4, SNList, NOList, List_Msg) == true)
                                {
                                    printresult = true;
                                }
                                else
                                {
                                    printresult = false;
                                    break; // 停止處理後面的組數
                                }
                                // 更新下一組的起始與結束編號
                                startNumber = endNumber + 1;
                                endNumber = startNumber + groupSize - 1;
                                group++;
                            }

                            if (endNumber >= number)
                            {
                                // 數字在指定的範圍內 ,不足36一組
                                // 處理該組的數字
                                NOList.Clear();
                                SNList.Clear();

                                // 以下 for 迴圈用於將主板箱標籤分組，每組最多 36 筆序號，若不足則補空字串
                                for (int i = startNumber; i <= endNumber; i++)
                                {
                                    NOList.Add(i); // 加入項次編號
                                    // 檢查序號列表是否有資料，若有則加入序號，否則補空字串
                                    if (SN_GV.Rows.Count >= i && SN_GV.Rows[i - 1].Cells["SN"].Value != null)
                                    {
                                        SNList.Add(SN_GV.Rows[i - 1].Cells["SN"].Value.ToString().Trim()); // 加入序號
                                    }
                                    else
                                    {
                                        SNList.Add(string.Empty); // 補空字串
                                    }
                                }

                                // 呼叫 BoardPrint4.PrintLabel 方法進行主板箱分組列印，根據回傳結果設定 printresult 狀態
                                if (BoardPrint4.PrintLabel(Boardpath4, SNList, NOList, List_Msg) == true)
                                {
                                    printresult = true; // 列印成功
                                }
                                else
                                {
                                    printresult = false; // 列印失敗
                                }
                            }
                        }
                    }
                }

                // 根據列印結果更新 UI 狀態
                if (printresult == true)
                {
                    btn_Print.Enabled = true;
                    if (!rdb_Acc.IsChecked) //更新Print_Carton_Number_Table目前Carton編號
                    {
                        SN_GV.Rows.Clear();         // 清空序號列表
                        List_Msg.Items.Clear();     // 清空訊息顯示區
                        List_Msg.Items.Add("列印完成....."); // 顯示列印完成訊息
                        txt_Sn.Enabled = true;      // 允許序號欄位輸入
                    }
                    else
                    {
                        // 配件箱/小箱列印不做 Carton_Number 更新，只更新 UI
                        // Acc/bcc列印不做Carton_Number更新
                        SN_GV.Rows.Clear();         // 清空序號列表
                        List_Msg.Items.Clear();     // 清空訊息顯示區
                        List_Msg.Items.Add("列印完成....."); // 顯示列印完成訊息
                    }
                }
                else
                {
                    btn_Print.Enabled = true;       // 列印失敗時恢復按鈕可用
                    List_Msg.Items.Add("列印失敗....."); // 顯示列印失敗訊息
                }
            }
            catch (Exception ex)
            {
                List_Msg.Items.Add("612行\r\n" + ex.Message);
                btn_Print.Enabled = true;
            }
        }

        /*
         * 清空 txt_Wo, txt_Pn, txt_Pc, txt_Qty, txt_Bios_Ver, txt_No_Number, txt_Sn, txt_reprint_Sn，重置 rcb_Re_Print，清空 SN_GV 與 List_Msg
         */
        /// <summary>
        /// 標準品外箱Label設定-Btn_Clear_Click 方法為「清除」按鈕的事件處理函式。
        /// 當使用者點擊此按鈕時，會清空所有標準品外箱列印相關的輸入欄位、
        /// 重設重列印序號欄位的啟用狀態、取消重列印核取方塊、
        /// 並清空序號列表及訊息顯示區。
        /// 此方法主要用於重設標準品外箱列印頁面的 UI 狀態，方便使用者重新操作。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為「清除」按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Clear_Click(object sender, EventArgs e)
        {
            // 清空標準品外箱列印頁面的所有輸入欄位與狀態
            txt_Wo.Text = string.Empty;           // 工單欄位設為空
            txt_Pn.Text = string.Empty;           // 料號欄位設為空
            txt_Pc.Text = string.Empty;           // 電源線編號欄位設為空
            txt_Qty.Text = string.Empty;          // 數量欄位設為空
            txt_Bios_Ver.Text = string.Empty;     // BIOS版本欄位設為空
            txt_No_Number.Text = string.Empty;    // 無編號欄位設為空
            txt_Sn.Text = string.Empty;           // 序號欄位設為空
            txt_reprint_Sn.Text = string.Empty;   // 重列印序號欄位設為空
            txt_reprint_Sn.Enabled = false;       // 重列印序號欄位設為不可編輯
            rcb_Re_Print.Checked = false;         // 取消重列印核取方塊
            SN_GV.Rows.Clear();                   // 清空序號列表
            List_Msg.Items.Clear();               // 清空訊息顯示區
        }

        /*
         * 外箱列印流程（不以 SN_GV 為主）；檢查 Qty、組裝欄位值；針對配件箱或系統箱呼叫 AccPrint.PrintLabel / SystemPrint.Out_PrintLabel；更新按鈕狀態與 List_Msg
         */
        /// <summary>
        /// Btn_Out_Print_Click 方法為「外箱列印」按鈕的事件處理函式。
        /// 根據目前選擇的列印模式（配件箱/系統箱），將輸入的工單、料號、序號等資訊組合後，呼叫對應的列印方法進行標籤列印。
        /// 列印完成或失敗時，會更新 UI 狀態並顯示訊息。
        /// </summary>
        /// <param name="sender">事件來源物件（通常為按鈕）。</param>
        /// <param name="e">事件參數（EventArgs）。</param>
        private void Btn_Out_Print_Click(object sender, EventArgs e)
        {
            try
            {
                // 清空訊息列表，並將列印相關按鈕設為不可用
                List_Msg.Items.Clear();
                btn_Print.Enabled = false;
                btn_Out_Print.Enabled = false;
                txt_Sn.Text = string.Empty;
                txt_Sn.Enabled = false;

                // 欄位值轉換
                string printlabe = string.Empty;
                bool printresult = false;
                #region 欄位轉換成值
                // 取得工單、料號、電源線編號、數量、BIOS版本、無編號等欄位值，並處理空值與預設值
                WO = txt_Wo.Text.Trim(); // 工單
                PN = txt_Pn.Text.Trim(); // 料號
                if (!string.IsNullOrWhiteSpace(txt_Pc.Text.Trim()))
                {
                    PC = txt_Pc.Text.Trim().Replace(",", " "); // 電源線編號，將逗號換成空格
                }
                else
                {
                    PC = "NA"; // 若未輸入則預設為 NA
                }
                if (string.IsNullOrWhiteSpace(txt_Qty.Text))
                {
                    List_Msg.Items.Add("尚未輸入數量"); // 提示未輸入數量
                    txt_Qty.Focus();
                    return;
                }
                else
                {
                    Qty = txt_Qty.Text.Trim(); // 數量
                }

                VER = txt_Bios_Ver.Text.Trim(); // BIOS 版本
                if (!string.IsNullOrWhiteSpace(txt_No_Number.Text.Trim()))
                {
                    No_Number = "(" + txt_No_Number.Text.Trim().ToUpper() + ")"; // 無編號，轉大寫並加括號
                }
                else
                {
                    No_Number = string.Empty; // 若未輸入則為空字串
                }
                #endregion

                List_Msg.Items.Add("列印中.....");
                // 判斷列印模式並呼叫對應的列印方法
                if (rdb_Acc.IsChecked)
                {
                    printlabe = Accpath;
                    // 呼叫配件箱列印方法
                    if (AccPrint.PrintLabel(printlabe, WO, PN, Qty, No_Number, List_Msg) == true)
                    {
                        printresult = true;
                    }
                    else
                    {
                        printresult = false;
                    }
                }
                else if (rdb_System.IsChecked)
                {
                    printlabe = Systempath;
                    // 呼叫系統箱列印方法
                    if (SystemPrint.Out_PrintLabel(printlabe, WO, PN, PC, VER, Qty, No_Number, ENGSR, List_Msg) == true)
                    {
                        printresult = true;
                    }
                    else
                    {
                        printresult = false;
                    }
                }
                else
                {
                    List_Msg.Items.Add("請確認列印製程.....");
                }

                // 根據列印結果更新 UI 狀態
                if (printresult == true)
                {
                    btn_Print.Enabled = true;
                    btn_Out_Print.Enabled = true;
                    List_Msg.Items.Add("列印完成.....");
                }
                else
                {
                    btn_Out_Print.Enabled = true;
                    btn_Print.Enabled = true;
                    List_Msg.Items.Add("列印失敗.....");
                }
            }
            catch (Exception ex)
            {
                // 顯示錯誤訊息
                List_Oem_Msg.Items.Add("1326行\r\n" + ex.Message);
                // 發生例外時確保按鈕可用
                btn_Print.Enabled = true;
            }
        }

        /*
         * 從資料表 Print_SN_Bind_Carton_Table 刪除對應 SN 與 Record_Time；刪除失敗時在 List_Msg 顯示錯誤
         */
        /// <summary>
        /// 處理標準品外箱序號列表刪除事件。
        /// 當使用者於 SN_GV 刪除序號列時，
        /// 會根據目前選取的序號與時間，
        /// 刪除資料庫中對應的 [Print_SN_Bind_Carton_Table] 資料。
        /// 若刪除失敗則顯示錯誤訊息。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為序號列表控制項。</param>
        /// <param name="e">包含刪除事件資料的 GridViewRowCancelEventArgs。</param>
        private void SN_GV_UserDeletingRow(object sender, GridViewRowCancelEventArgs e)
        {
            // 取得目前選取的序號與時間，組成刪除 SQL，刪除 [Print_SN_Bind_Carton_Table] 資料表中對應的資料
            string DelSN = SN_GV.CurrentRow.Cells["SN"].Value.ToString();
            string DelTime = Convert.ToDateTime(SN_GV.CurrentRow.Cells["Time"].Value).ToString("yyyy-MM-dd HH:mm:ss");
            string DelSql = " Delete  [Print_SN_Bind_Carton_Table] where Sn = '" + DelSN + "' and Record_Time = '" + DelTime + "'";
            if (db.Exsql(DelSql) == true)
            {
                // 刪除成功不需額外處理
            }
            else
            {
                // 若刪除失敗則顯示錯誤訊息
                List_Msg.Items.Add("資料庫刪除錯誤");
            }
        }

        /*
         * 按 Enter 時將 SN 加入 SN_GV（檢查重複、數量上限），寫入資料庫 Print_SN_Bind_Carton_Table，若已達 qty 則詢問確認後呼叫 Btn_Print_Click 自動列印。
         */
        /// <summary>
        /// 處理標準品外箱序號輸入欄位的 KeyPress 事件。
        /// 當使用者在 txt_Sn 輸入序號並按下 Enter 鍵時，執行以下流程：
        /// 1. 檢查是否已輸入數量，若未輸入則提示並回復焦點。
        /// 2. 檢查序號是否重複，若重複則提示並終止流程。
        /// 3. 若序號未重複且未超過數量上限，則新增序號至序號列表，並寫入資料庫。
        /// 4. 若序號數量已達設定數量，則自動觸發列印流程。
        /// 5. 若序號欄位為空則提示錯誤。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Sn 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void Txt_Sn_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 處理標準品外箱序號輸入欄位的 KeyPress 事件
            // 當使用者在 txt_Sn 輸入序號並按下 Enter 鍵時，執行以下流程：
            // 1. 檢查是否已輸入數量，若未輸入則提示並回復焦點。
            // 2. 檢查序號是否重複，若重複則提示並終止流程。
            // 3. 若序號未重複且未超過數量上限，則新增序號至序號列表，並寫入資料庫。
            // 4. 若序號數量已達設定數量，則自動觸發列印流程。
            // 5. 若序號欄位為空則提示錯誤。
            bool AddSN = true;
            string Bind_SN = string.Empty;
            if (string.IsNullOrWhiteSpace(txt_Qty.Text))
            {
                List_Msg.Items.Add("請先輸入數量 ");
                txt_Qty.Focus();
                return;
            }

            // 按下 Enter 鍵時執行
            if (Convert.ToInt32(e.KeyChar) == 13)
            {
                // 檢查序號是否重複
                if (!string.IsNullOrWhiteSpace(txt_Sn.Text.Trim().ToUpper()))
                {
                    for (int i = 0; i < SN_GV.Rows.Count; i++)
                    {
                        // 逐筆檢查目前序號列表中是否已有相同序號
                        if (txt_Sn.Text.Trim().ToUpper() == SN_GV.Rows[i].Cells["SN"].Value.ToString().ToUpper())
                        {
                            List_Msg.Items.Add(txt_Sn.Text.Trim().ToUpper() + "序號重複 ");
                            AddSN = false;
                            break;
                        }
                        else
                        {
                            AddSN = true;
                        }
                    }

                    // 若序號可新增且未超過數量上限
                    // 以下程式碼片段為標準品外箱序號輸入流程，包含重複檢查、數量上限判斷、資料庫寫入與自動列印防呆確認
                    if (AddSN == true)
                    {
                        // 若目前序號列表為空，則項次設為 "1"，並清空訊息顯示區，綁定序號為目前輸入序號
                        if (SN_GV.Rows.Count == 0)
                        {
                            Item = "1";
                            List_Msg.Items.Clear();
                            Bind_SN = txt_Sn.Text.Trim().ToUpper();
                        }
                        else
                        {
                            // 否則項次為目前序號數量加一，綁定序號為第一筆序號
                            Item = (SN_GV.Rows.Count + 1).ToString();
                            Bind_SN = SN_GV.Rows[0].Cells["SN"].Value.ToString().ToUpper();
                            // 若超過數量上限則顯示錯誤訊息並結束
                            if (int.Parse(Item) > int.Parse(txt_Qty.Text.Trim()))
                            {
                                List_Msg.Items.Add(txt_Sn.Text.Trim().ToUpper() + "EEROR!! 數量滿箱");
                                return;
                            }
                        }
                        // 新增序號至序號列表
                        SN = txt_Sn.Text.Trim().ToUpper();
                        SN_GV.Rows.Add(new object[] { Item, SN, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") });
                        #region 資料庫寫入SN資訊
                        // 新增序號綁定資料至資料庫
                        string InsSql = " INSERT INTO [Print_SN_Bind_Carton_Table] (Record_Time,Work_Order,Part_No,Power_Code,Quantity,Bios_Ver,SN,Bind_SN,No_Number) VALUES("
                                                                   + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                                                                   + "'" + txt_Wo.Text.Trim() + "',"
                                                                   + "'" + txt_Pn.Text.Trim() + "',"
                                                                   + "'" + txt_Pc.Text.Trim() + "',"
                                                                   + "'" + txt_Qty.Text.Trim() + "',"
                                                                   + "'" + txt_Bios_Ver.Text.Trim().ToUpper() + "',"
                                                                   + "'" + txt_Sn.Text.Trim().ToUpper() + "',"
                                                                   + "'" + Bind_SN + "',"
                                                                   + "'" + txt_No_Number.Text.Trim().ToUpper() + "')";
                        // 執行 SQL 寫入，若成功則清空序號輸入欄位
                        if (db.Exsql(InsSql) == true)
                        {
                            txt_Sn.Text = string.Empty;
                            // 若序號數量已達設定值則自動觸發列印流程（防呆確認）
                            if (Item == txt_Qty.Text.Trim())
                            {
                                // 新增防呆：彈出確認視窗 20250912 By Jesse
                                var result = MessageBox.Show("是否確認要列印？", "列印確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                                if (result == DialogResult.OK)
                                {
                                    // 標準品外箱Label設定-列印
                                    Btn_Print_Click(sender, e);
                                }
                                // 按取消則不執行
                            }
                        }
                        else
                        {
                            List_Msg.Items.Add("資料庫寫入錯誤");
                        }
                        #endregion
                    }
                }
                else
                {
                    List_Msg.Items.Add("請檢查序號 ");
                }
            }
        }
        #endregion

        #region 工單資訊
        /*
         * 按 Enter 時，檢查是否選擇列印製程，根據選擇啟動 FTP 下載套版（若本機不存在），清理/開啟相關欄位
         * ，呼叫外部 API 查工單並填入 txt_Pn, txt_Pc, txt_Bios_Ver 等
         */
        /// <summary>
        /// 處理工單輸入欄位的 KeyPress 事件。
        /// 當使用者在 txt_Wo 輸入工單並按下 Enter 鍵時，依據所選列印製程（主板箱/系統箱/配件箱）
        /// 進行套版下載、欄位清空與工單相關資訊查詢，並自動填入料號、電源線、BIOS版本等欄位。
        /// 若查無工單資訊則顯示錯誤訊息。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Wo 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void Txt_Wo_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                // 按下 Enter 鍵時執行
                if (Convert.ToInt32(e.KeyChar) == 13)
                {
                    // 以下區塊為工單輸入後依據選擇的列印製程自動下載對應的套版檔案
                    if (rdb_Board.IsChecked || rdb_System.IsChecked || rdb_Acc.IsChecked)
                    {
                        // 若選擇主板箱列印
                        if (rdb_Board.IsChecked)
                        {
                            // 檢查主板箱套版檔案是否存在，若不存在則啟動FTP下載
                            if (!File.Exists(Boardpath))
                            {
                                List_Msg.Items.Add("列印套版下載中......");
                                DLfilename = ini.IniReadValue("Option", "Board_Carton_Label_Name", Cfgname);
                                this.FTP_Dl_Btw_thread.WorkerSupportsCancellation = true; //允許中斷
                                this.FTP_Dl_Btw_thread.RunWorkerAsync(); //呼叫背景程式
                                txt_Sn.ReadOnly = false;
                            }
                        }
                        // 若選擇系統箱列印
                        else if (rdb_System.IsChecked)
                        {
                            // 檢查系統箱套版檔案是否存在，若不存在則啟動FTP下載
                            if (!File.Exists(Systempath))
                            {
                                DLfilename = ini.IniReadValue("Option", "System_Carton_Label_Name", Cfgname);
                                this.FTP_Dl_Btw_thread.WorkerSupportsCancellation = true; //允許中斷
                                this.FTP_Dl_Btw_thread.RunWorkerAsync(); //呼叫背景程式
                            }
                        }
                        // 若選擇配件箱列印
                        else
                        {
                            // 檢查配件箱套版檔案是否存在，若不存在則啟動FTP下載
                            if (!File.Exists(Accpath))
                            {
                                DLfilename = ini.IniReadValue("Option", "Acc_Carton_Label_Name", Cfgname);
                                this.FTP_Dl_Btw_thread.WorkerSupportsCancellation = true; //允許中斷
                                this.FTP_Dl_Btw_thread.RunWorkerAsync(); //呼叫背景程式
                            }
                        }
                    }
                    else
                    {
                        // 未選擇任何列印製程時，顯示錯誤訊息並將焦點移至選擇區
                        List_Msg.Items.Add("ERROR!!請先選擇列印製程");
                        radGroupBox7.Focus();
                        return;
                    }
                    #region 清空欄位
                    txt_Sn.ReadOnly = true;//工單輸入後將序號欄位設為唯讀，待套版下載完成後再開啟
                    txt_Pn.Text = string.Empty;//清空料號欄位
                    txt_Pc.Text = string.Empty;//清空電源線欄位
                    txt_Qty.Text = string.Empty;//清空數量欄位
                    txt_Bios_Ver.Text = string.Empty; //清空BIOS版本欄位
                    txt_No_Number.Text = string.Empty;//清空無編號欄位
                    List_Msg.Items.Clear();
                    #endregion

                    // 檢查工單欄位是否有輸入
                    if (!string.IsNullOrWhiteSpace(txt_Wo.Text.Trim()))
                    {
                        // 檢查工單長度是否為 10 碼
                        if (Convert.ToInt32(txt_Wo.Text.Trim().Length) == 10)
                        {
                            #region 開啟欄位
                            txt_Pn.Enabled = true;//開啟料號欄位
                            txt_Qty.Enabled = true;//開啟數量欄位
                            txt_Pc.Enabled = true;//開啟電源線欄位
                            txt_Bios_Ver.Enabled = true;//開啟BIOS版本欄位
                            txt_Sn.Enabled = true;//開啟序號欄位
                            btn_Print.Enabled = true;//開啟列印按鈕
                            btn_Out_Print.Enabled = true;//開啟列印按鈕
                            txt_No_Number.Enabled = true;//開啟無編號欄位
                            #endregion

                            string WipInfo = string.Empty;//宣告一個字串用來接Wip資訊
                            string[] relatedwip;//宣告一個陣列用來接相關工單資訊

                            //改為透過安勤工單查昶亨工單;在透過昶亨工單查詢必要資訊
                            //[0] : wono ; [1] : engsr
                            relatedwip = Auto_Route.WipInfoByRelatedWoNo(txt_Wo.Text.Trim()).Split('&');

                            // 根據使用者選擇的列印製程，呼叫不同的工單查詢方法
                            if (rdb_Board.IsChecked)
                            {
                                // 若選擇主板箱，則呼叫 WipBoard_Eve 方法查詢主板工單資訊
                                WipInfo = Auto_Route.WipBoard_Eve(relatedwip[0]);
                            }
                            else if (rdb_System.IsChecked)
                            {
                                // 若選擇系統箱，則呼叫 WipSystem_Eve 方法查詢系統工單資訊
                                WipInfo = Auto_Route.WipSystem_Eve(relatedwip[0]);
                            }

                            // 反序列化工單資訊，取得 SFIS 物件
                            bool amesFg = true, sfisFg = true;
                            SFIS descJsonStu = JsonConvert.DeserializeObject<SFIS>(WipInfo.ToString());//反序列化

                            if (descJsonStu == null)
                            {
                                amesFg = false;
                                List_Msg.Items.Add("AMES_2.0查無工單資訊");
                                sfisFg = AvSFIS_Method();
                                if (!amesFg && !sfisFg)
                                {
                                    List_Msg.Items.Add("舊系統查無工單資訊");
                                    return;
                                }
                                else if (sfisFg)
                                {
                                    foreach (DataRow row in dataSet.Tables[0].Rows)
                                    {
                                        txt_Pn.Text = row["ProductID"].ToString().Trim();
                                        txt_Pn.Focus();
                                        string[] rows = row["ProductID_MF"].ToString().Trim().Split('-');
                                        ENGSR = rows[0];
                                        if (!string.IsNullOrWhiteSpace(ENGSR))
                                        {
                                            txt_engSR.Text = ENGSR;
                                        }
                                        else
                                        {
                                            ENGSR = txt_engSR.Text;
                                        }

                                        txt_Bios_Ver.Text = row["biosVer"].ToString().Trim();
                                        string sqlPC = "select pc from Print_AvSFIS_PowerCode_Table where mid = '" + txt_Wo.Text.Trim() + "' ";

                                        if (db.reDs(sqlPC).Tables[0].Rows.Count > 0)
                                        {
                                            txt_Pc.Text = db.reDs(sqlPC).Tables[0].Rows[0]["pc"].ToString().Trim();
                                        }
                                        else
                                        {
                                            txt_Pc.Text = "NA";
                                        }
                                    }
                                }
                            }
                            if (amesFg)
                            {
                                #region //新版查詢用
                                if (!string.IsNullOrWhiteSpace(descJsonStu.biosVer))
                                {
                                    txt_Bios_Ver.Text = descJsonStu.biosVer;
                                    if (string.IsNullOrWhiteSpace(descJsonStu.biosVer))
                                    {
                                        txt_Bios_Ver.Text = "NA";
                                    }
                                }
                                if (!string.IsNullOrWhiteSpace(descJsonStu.powercord))
                                {
                                    txt_Pc.Text = descJsonStu.powercord;
                                }
                                string[] WinInfoList = Auto_Route.Wip_To_Unint(txt_Wo.Text).Split('&');//取得要做的MAC
                                DataTable test = new DataTable();
                                test = Auto_Route.PowerCord(txt_Wo.Text.Trim());
                                if (test != null && test.Rows.Count > 0)
                                {
                                    if (test.Rows.Count == 1)
                                    {
                                        if (int.Parse(test.Rows[0]["realsendQty"].ToString()) > 0)
                                        {
                                            txt_Pc.Text = test.Rows[0]["materialNo"].ToString();
                                        }
                                    }
                                    else
                                    {
                                        for (int i = 0; i < test.Rows.Count; i++)
                                        {
                                            if (int.Parse(test.Rows[i]["realsendQty"].ToString()) > 0)
                                            {
                                                txt_Pc.Text += test.Rows[i]["materialNo"].ToString() + ",";
                                            }
                                        }
                                    }
                                }
                                if (!string.IsNullOrWhiteSpace(relatedwip[1]))
                                {
                                    string[] rows = relatedwip[1].Trim().Split('-');
                                    ENGSR = rows[0];
                                    txt_engSR.Text = ENGSR;
                                }
                                else
                                {
                                    ENGSR = txt_engSR.Text;
                                }
                                //查機種名稱
                                string WipPn = Auto_Route.WipSystem(txt_Wo.Text.Trim());
                                AvSFIS JsonStu = JsonConvert.DeserializeObject<AvSFIS>(WipPn.ToString());//反序列化
                                if (!string.IsNullOrWhiteSpace(JsonStu.itemNo))
                                {
                                    txt_Pn.Text = JsonStu.itemNo;
                                    txt_Pn.Focus();
                                }
                                else
                                {
                                    List_Msg.Items.Add("ERROR!!查無機種");
                                }
                                #endregion
                            }
                        }
                        else
                        {
                            List_Msg.Items.Add("ERROR!!請確認工單長度");
                        }
                    }
                    else
                    {
                        List_Msg.Items.Add("ERROR!!請輸入工單 ");
                    }
                }
            }
            catch (Exception Ex)
            {
                List_Msg.Items.Add("1135行\r\n" + Ex.Message);
            }

        }

        /*
         * 離開時開啟 txt_Sn（設 ReadOnly = false）；（Designer 曾有註解說可在此自動下載套版）
         */
        /// <summary>
        /// Txt_Pn_Leave 事件處理函式。
        /// 當料號欄位 (txt_Pn) 離開焦點時觸發。
        /// 1. 檢查料號欄位是否有輸入內容。
        /// 2. 若有輸入，則將序號欄位 (txt_Sn) 設為可編輯 (ReadOnly = false)。
        /// 3. 此方法目前未自動下載套版檔案，僅開啟序號欄位，
        ///    若需自動下載可參考註解區塊的 SQL 查詢與 FTP 下載流程。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Pn 控制項。</param>
        /// <param name="e">事件參數，包含離開事件資訊。</param>
        private void Txt_Pn_Leave(object sender, EventArgs e)
        {
            string sqlBTWCmd = "";
            //搜尋列印檔案
            if (!string.IsNullOrWhiteSpace(txt_Pn.Text.Trim()))
            {
                //List_Msg.Items.Add("列印套版下載中......");
                //sqlBTWCmd = "SELECT * FROM [Print_Carton_Table] where Model = '" + txt_Pn.Text.Trim() + "' ";
                //DataSet dsBTW = db.reDs(sqlBTWCmd);
                //if (dsBTW.Tables[0].Rows.Count > 0)
                //{
                //    DLfilename = dsBTW.Tables[0].Rows[0]["Filename"].ToString();
                //    this.FTP_Dl_Btw_thread.WorkerSupportsCancellation = true; //允許中斷
                //    this.FTP_Dl_Btw_thread.RunWorkerAsync(); //呼叫背景程式
                //    txt_Sn.ReadOnly = false;
                //}
                //else
                //{
                //    List_Msg.Items.Add("查無" + txt_Pn.Text.Trim() + "列印套版");
                //}
                txt_Sn.ReadOnly = false;
            }
        }

        /*
         * 只允許數字與 Backspace，輸入後啟用 txt_Sn、btn_Print、btn_Out_Print。
         */
        /// <summary>
        /// 處理數量輸入欄位的 KeyPress 事件。
        /// 僅允許使用者輸入數字（0-9）與 Backspace，其他按鍵將被禁止。
        /// 當輸入非數字或非 Backspace 時，會自動啟用序號輸入欄位 (txt_Sn) 及列印按鈕 (btn_Print, btn_Out_Print)。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Qty 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void Txt_Qty_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允許輸入數字（0-9）和 Backspace，其餘按鍵禁止，並啟用序號輸入欄位與列印按鈕
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true; // 禁止非數字或非 Backspace 的按鍵
                btn_Out_Print.Enabled = true; // 啟用外箱列印按鈕
                txt_Sn.Enabled = true;        // 啟用序號輸入欄位
                btn_Print.Enabled = true;     // 啟用列印按鈕
            }
        }

        /*
         * 按 Enter 時檢查有無選擇列印製程，查 Print_SN_Bind_Carton_Table 找出所有與 Bind_SN 相同的 SN
         * ，填回 txt_Wo, txt_Pn, txt_Pc, txt_Qty, txt_Bios_Ver, txt_No_Number 並把 SN 填入 SN_GV；若數量達到則提示並呼叫 Btn_Print_Click
         */
        ///<summary>
        /// 處理標準品外箱重列印序號輸入欄位的 KeyPress 事件。
        /// 當使用者在 <c>txt_reprint_Sn</c> 輸入序號並按下 Enter 鍵時，
        /// 會依序執行以下流程：
        /// 1. 檢查是否已選擇列印製程（主板箱/系統箱/配件箱），若未選擇則顯示錯誤訊息並回復焦點。
        /// 2. 清空所有標準品外箱列印相關欄位與序號列表，準備載入重列印資料。
        /// 3. 若輸入序號不為空，則查詢資料庫 <c>Print_SN_Bind_Carton_Table</c> 取得所有綁定序號的資料，
        ///    並自動填入工單、料號、電源線、數量、BIOS版本、無編號等欄位，
        ///    並逐筆填入序號列表 <c>SN_GV</c>。
        /// 4. 若序號數量已達設定值則自動觸發列印流程（呼叫 <c>Btn_Print_Click</c>）。
        /// 5. 若查無資料則顯示錯誤訊息，並將焦點移回輸入欄位。
        /// 6. 若未輸入序號則顯示錯誤訊息並回復焦點。
        /// 7. 捕捉所有例外並以訊息視窗顯示錯誤內容。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 <c>txt_reprint_Sn</c> 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void Txt_reprint_Sn_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 處理標準品外箱重列印序號輸入欄位的 KeyPress 事件
            // 當使用者在 txt_reprint_Sn 輸入序號並按下 Enter 鍵時，執行重列印流程
            try
            {
                if (Convert.ToInt32(e.KeyChar) == 13)
                {
                    // 檢查是否已選擇列印製程（主板箱/系統箱/配件箱）
                    if (rdb_Board.IsChecked || rdb_System.IsChecked || rdb_Acc.IsChecked)
                    {
                        // 已選擇列印製程，繼續執行
                    }
                    else
                    {
                        // 未選擇列印製程，顯示錯誤訊息並回復焦點
                        List_Msg.Items.Add("ERROR!!請先選擇列印製程");
                        radGroupBox7.Focus();
                        return;
                    }
                    #region 清空欄位
                    // 清空所有標準品外箱列印相關欄位與序號列表
                    txt_Sn.ReadOnly = true;
                    txt_Pn.Text = string.Empty;
                    txt_Pc.Text = string.Empty;
                    txt_Qty.Text = string.Empty;
                    txt_Bios_Ver.Text = string.Empty;
                    txt_No_Number.Text = string.Empty;
                    txt_No_Number.Enabled = true;
                    List_Msg.Items.Clear();
                    SN_GV.Rows.Clear();
                    #endregion

                    // 檢查是否有輸入第一筆 SN
                    if (!string.IsNullOrWhiteSpace(txt_reprint_Sn.Text.Trim()))
                    {
                        // 查詢資料庫取得所有綁定序號的資料
                        string sqlCmd = "SELECT * FROM [Print_SN_Bind_Carton_Table] where Bind_SN = '" + txt_reprint_Sn.Text.Trim() + "' ";
                        DataSet ds = db.reDs(sqlCmd);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            // 將查詢到的工單、料號、電源線、數量、BIOS版本、無編號填入欄位
                            txt_Wo.Text = ds.Tables[0].Rows[0]["Work_Order"].ToString();
                            txt_Pn.Text = ds.Tables[0].Rows[0]["Part_No"].ToString();
                            txt_Pc.Text = ds.Tables[0].Rows[0]["Power_Code"].ToString();
                            txt_Qty.Text = ds.Tables[0].Rows[0]["Quantity"].ToString();
                            txt_Bios_Ver.Text = ds.Tables[0].Rows[0]["Bios_Ver"].ToString();
                            txt_No_Number.Text = ds.Tables[0].Rows[0]["No_Number"].ToString();

                            // 逐筆填入序號列表
                            // 逐筆填入序號列表
                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                // 若目前序號列表為空，則項次設為 "1"
                                if (SN_GV.Rows.Count == 0)
                                {
                                    Item = "1";
                                }
                                else
                                {
                                    // 否則項次為目前序號數量加一
                                    Item = (SN_GV.Rows.Count + 1).ToString();
                                }
                                // 新增一列，包含項次、序號、記錄時間
                                SN_GV.Rows.Add(new object[] { Item, ds.Tables[0].Rows[i]["SN"].ToString(), ds.Tables[0].Rows[i]["Record_Time"].ToString() });
                            }
                            // 若序號數量已達設定值則自動觸發列印流程
                            if (Item == txt_Qty.Text.Trim())
                            {
                                // 新增防呆：彈出確認視窗 20250912 By Jesse
                                var result = MessageBox.Show("是否確認要列印？", "列印確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                                if (result == DialogResult.OK)
                                {
                                    // 標準品外箱Label設定-列印
                                    Btn_Print_Click(sender, e);
                                }
                                // 按取消則不執行
                            }
                        }
                        else
                        {
                            // 查無資料時顯示錯誤訊息
                            List_Msg.Items.Add("ERROR!!查無第一筆SN資料");
                        }
                    }
                    else
                    {
                        // 未輸入第一筆 SN 時顯示錯誤訊息並回復焦點
                        List_Msg.Items.Add("ERROR!!請輸入第一筆SN ");
                        txt_reprint_Sn.Focus();
                    }
                }
            }
            catch (Exception Ex)
            {
                // 發生例外時顯示錯誤訊息
                MessageBox.Show(Ex.ToString());
            }
        }

        /*
         * 切換 txt_reprint_Sn.Enabled = !rcb_Re_Print.Checked（勾選則禁止輸入 reprint SN）。
         */
        /// <summary>
        /// 處理「重新列印(Re-Print)」核取方塊的 Click 事件，切換重列印序號輸入欄位的可用狀態。
        /// </summary>
        /// <param name="sender">事件來源物件 (通常為 rcb_Re_Print 控制項)。</param>
        /// <param name="e">事件參數 (EventArgs)。</param>
        /// <remarks>
        /// 當 rcb_Re_Print 被勾選時，表示要進行重列印模式，應禁止輸入第一筆序號 (txt_reprint_Sn) 以避免誤輸入；
        /// 取消勾選則恢復輸入功能。此方法僅變更 UI 狀態，不會進行任何資料庫或列印操作。
        /// </remarks>
        private void Rcb_Re_Print_Click(object sender, EventArgs e)
        {
            // 直接將 txt_reprint_Sn.Enabled 設為 rcb_Re_Print.Checked 的反向值，
            // 簡潔且清楚地反映「勾選時禁止輸入」的行為。
            txt_reprint_Sn.Enabled = !rcb_Re_Print.Checked;
        }
        #endregion
        #endregion

        #region Click 事件
        /// <summary>
        /// 處理「搜尋檔案」按鈕點擊事件。
        /// 開啟檔案選擇對話框，選擇 BarTender 標籤檔案（*.btw），
        /// 並產生標籤預覽圖片顯示於 UP_PictureBox。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Search_File_Click(object sender, EventArgs e)
        {

            // 清空圖片與訊息
            UP_PictureBox.Image = null;
            List2_Msg.Items.Clear();
            // 建立檔案選擇對話框
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//該值確定是否可以選擇多個檔案
            dialog.Title = "請選擇列印檔案";
            dialog.Filter = "列印檔案(*.btw)|*.btw";
            // 若選擇檔案
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txt_Btw_Path.Text = dialog.FileName;
                _btw_path = dialog.FileName;
                filename = dialog.SafeFileName;
                //PrintBar(true);
                UP_PictureBox.Image = null;
                // 使用 BarTender Engine 產生預覽圖
                try
                {
                    using (Engine btEngine = new Engine(true))
                    {
                        LabelFormatDocument labelFormat = null;
                        try
                        {
                            labelFormat = btEngine.Documents.Open(_btw_path);

                        }
                        catch (Seagull.BarTender.Print.PrintEngineException ex)
                        {

                            // 判斷是否為版本不相容
                            if (ex.Message.Contains("較新版本") || ex.Message.Contains("cannot be opened") || ex.Message.Contains("newer version"))
                            {
                                MessageBox.Show("標籤檔案版本不相容，請用2022版 BarTender 版儲存。", "版本錯誤");
                                return;
                            }
                            else
                            {
                                MessageBox.Show($"BarTender 錯誤：{ex.Message}", "操作提示");
                                return;
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException ex)
                        {
                            if (ex.Message.Contains("較新版本") || ex.Message.Contains("cannot be opened") || ex.Message.Contains("newer version"))
                            {
                                MessageBox.Show("標籤檔案版本不相容，請用2022版 BarTender 版儲存。", "版本錯誤");
                                return;
                            }
                            else
                            {
                                MessageBox.Show($"BarTender COM 錯誤：{ex.Message}", "操作提示");
                                return;
                            }
                        }
                        if (labelFormat != null)
                        {
                            Seagull.BarTender.Print.Messages m;
                            labelFormat.ExportPrintPreviewToFile(Application.StartupPath, @"\exp.bmp", ImageType.JPEG, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(300, 300), System.Drawing.Color.White, OverwriteOptions.Overwrite, true, true, out m);
                            labelFormat.ExportImageToFile(_bmp_path, ImageType.JPEG, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(300, 300), OverwriteOptions.Overwrite);

                            Image image = Image.FromFile(_bmp_path);
                            Bitmap NmpImage = new Bitmap(image);
                            UP_PictureBox.Image = NmpImage;
                            image.Dispose();
                        }
                        else
                        {
                            MessageBox.Show("生成圖片錯誤", "操作提示");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"發生未預期錯誤：{ex.Message}", "操作提示");
                }
            }
        }

        /// <summary>
        /// 處理「OEM 套版搜尋」按鈕的 Click 事件，根據使用者輸入的工單號 (txt_Oem_Wo)
        /// 從資料表 Print_Carton_Table 查詢對應的列印套版設定，並將結果更新到 UI 控制項，
        /// 若查詢成功則啟動背景執行緒下載套版檔案 (FTP_Dl_Btw_thread)。
        /// </summary>
        /// <param name="sender">事件發送者，通常為按鈕物件 (object)。</param>
        /// <param name="e">事件參數 (EventArgs)。</param>
        /// <remarks>
        /// - 會先清除 List_Oem_Msg 與 txt_Oem_Pn，並在 List_Oem_Msg 顯示「列印套版下載中......」提示。<br/>
        /// - 若 txt_Oem_Wo 為空，則不做查詢動作。<br/>
        /// - 查詢結果若有資料：<br/>
        ///     * 將資料表中的 Pn 與 Filename 帶入到 txt_Oem_Pn 與 DLfilename。<br/>
        ///     * 根據 Qty_Set、Weight_Set、Sn_Set、Mac_Set、Bios_Set 的值啟用或停用相對應的輸入欄位。<br/>
        ///     * 若 Mac_Set 為 "Yes"，會將焦點設定到 txt_Oem_Mac。<br/>
        ///     * 啟動 FTP_Dl_Btw_thread 進行非同步下載（同時允許取消）。<br/>
        /// - 若查無資料會在 List_Oem_Msg 顯示查無訊息。<br/>
        /// - 例外處理：方法內捕捉所有例外並保留現有行為；若需要，可在 catch 區段增加日誌或錯誤回報機制。<br/>
        /// </remarks>
        private void Btn_Oem_Search_Click(object sender, EventArgs e)
        {
            try
            {
                // 2. 清除 List_Oem_Msg 與相關顯示 (txt_Oem_Pn)，顯示「列印套版下載中......」提示文字。
                List_Oem_Msg.Items.Clear();
                //DL_PictureBox.Image = null;
                txt_Oem_Pn.Text = string.Empty;

                List_Oem_Msg.Items.Add("列印套版下載中......");
                string sqlCmd = "";

                // 1. 驗證 txt_Oem_Wo 是否為空。若為空則結束動作（目前實作不顯示額外錯誤）。
                if (!string.IsNullOrWhiteSpace(txt_Oem_Wo.Text.Trim()))
                {
                    // 3. 執行 SQL 查詢 Print_Carton_Table（依 time desc），取得第一筆資料（若有）。
                    // 4. 若查詢到資料：
                    //    4.1  將 Pn 與 Filename 分別寫入 txt_Oem_Pn 及 DLfilename。
                    //    4.2  依資料欄位 Qty_Set、Weight_Set、Sn_Set、Mac_Set、Bios_Set 的值設定對應 txt_Oem_* 控制項的 Enabled 屬性。
                    //    4.3  若 Mac_Set 為 Yes，將焦點移到 txt_Oem_Mac。
                    //    4.4  啟動 FTP_Dl_Btw_thread 的背景工作 (允許中斷並非同步下載套版檔案)。
                    // 5. 若查無資料：在 List_Oem_Msg 顯示查無訊息。

                    // 依據 SQL 查詢結果設定 OEM/ODM 列印頁面欄位與狀態
                    sqlCmd = "SELECT * FROM [Print_Carton_Table] where Wo = '" + txt_Oem_Wo.Text.Trim() + "' order by time desc ";
                    DataSet ds = db.reDs(sqlCmd);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        // 設定料號欄位
                        txt_Oem_Pn.Text = ds.Tables[0].Rows[0]["Pn"].ToString();
                        // 設定下載檔案名稱
                        DLfilename = ds.Tables[0].Rows[0]["Filename"].ToString();
                        // 根據資料庫設定啟用/停用數量欄位
                        if (ds.Tables[0].Rows[0]["Qty_Set"].ToString() == "Yes")
                        {
                            txt_Oem_Qty.Enabled = true;
                        }
                        else
                        {
                            txt_Oem_Qty.Enabled = false;
                        }
                        // 根據資料庫設定啟用/停用重量欄位
                        if (ds.Tables[0].Rows[0]["Weight_Set"].ToString() == "Yes")
                        {
                            txt_Oem_Weight.Enabled = true;
                        }
                        else
                        {
                            txt_Oem_Weight.Enabled = false;
                        }
                        // 根據資料庫設定啟用/停用序號欄位
                        if (ds.Tables[0].Rows[0]["Sn_Set"].ToString() == "Yes")
                        {
                            txt_Oem_Sn.Enabled = true;
                        }
                        else
                        {
                            txt_Oem_Sn.Enabled = false;
                        }
                        // 根據資料庫設定啟用/停用 MAC 欄位，並設定焦點
                        if (ds.Tables[0].Rows[0]["Mac_Set"].ToString() == "Yes")
                        {
                            txt_Oem_Mac.Enabled = true;
                            txt_Oem_Mac.Focus();
                        }
                        else
                        {
                            txt_Oem_Mac.Enabled = false;
                        }
                        // 根據資料庫設定啟用/停用 BIOS 欄位
                        if (ds.Tables[0].Rows[0]["Bios_Set"].ToString() == "Yes")
                        {
                            txt_Oem_Bios.Enabled = true;
                        }
                        else
                        {
                            txt_Oem_Bios.Enabled = false;
                        }
                        // 啟動 FTP 下載套版檔案的背景執行緒
                        this.FTP_Dl_Btw_thread.WorkerSupportsCancellation = true; //允許中斷
                        this.FTP_Dl_Btw_thread.RunWorkerAsync(); //呼叫背景程式
                    }
                    else
                    {
                        // 查無資料時顯示提示訊息
                        List_Oem_Msg.Items.Add("查無" + txt_Oem_Wo.Text.Trim() + "列印套版");
                    }
                }
            }
            catch (Exception ex)
            {
                List_Oem_Msg.Items.Add("796行\r\n" + ex.Message);
                // 6. try/catch：保留原有行為（捕捉例外但不拋出）；可擴充為寫入日誌或顯示詳細錯誤訊息。
                // 7. 注意：此方法會直接存取 UI 控制項並啟動 BackgroundWorker，呼叫端層級需注意多執行緒安全及 UI 更新時的 invoke。
                // 保留現有行為：不顯示例外詳細內容。
                // 若需要更完整的錯誤處理，可在此加入日誌或用戶提示：
                // e.g. Logger.Error(ex); List_Oem_Msg.Items.Add("查詢發生錯誤，請聯絡系統管理員");
            }
        }

        /// <summary>
        /// 處理 OEM/ODM 列印按鈕事件。
        /// 此方法會根據 txt_Oem_Bios 控制項是否啟用，分為兩種主要列印流程：
        /// 1. OEM-STD_Bios 模式 (txt_Oem_Bios.Enabled == true)
        ///    - 若數量大於 26，先取前 26 筆序號由 OemPrint2 列印，再將剩餘序號交由 BoardPrint2 列印。
        ///    - 若數量小於等於 26，直接以 OemPrint2 列印全部序號。
        /// 2. 標準列印以外模式
        ///    - 以 BarTender Engine 開啟下載目錄下的標籤檔 (download_Path + "\\" + DLfilename)。
        ///    - 設定標籤的 SubStrings（QTY、Weight、mac_n、sn_n、SN1、SN2 以及 SN1..SNn）。
        ///    - 若有輸入 MAC，會將 IdenticalCopiesOfLabel 設為 2（列印兩份）。
        ///    - 呼叫 btFormat.Print() 執行列印，並於最後停止 Engine。
        ///
        /// 方法結束時會清除 Sn_Oem_GV、清空輸入欄位並更新 List_Oem_Msg 狀態文字。
        /// </summary>
        /// <param name="sender">事件來源物件（通常為按鈕）。</param>
        /// <param name="e">事件參數（EventArgs）。</param>
        /// <remarks>
        /// - 呼叫此方法會變更 UI 控制項的 Enabled/Focus 與 List_Oem_Msg 顯示內容，需在 UI 執行緒上呼叫。
        /// - 若發生例外，方法會在 catch 區段處理並將 btn_Oem_Print 恢復可用，避免例外導致 UI 無法回應。
        /// - 本方法不會回傳列印結果給呼叫端，僅透過介面顯示列印狀態。
        /// - 若需要更細緻的錯誤處理或記錄，可在 catch 區段加入日誌記錄或顯示詳細錯誤資訊。
        /// </remarks>
        /// <exception cref="Exception">方法會捕捉所有例外並在 catch 區段處理，不會向上
        private void Btn_Oem_Print_Click(object sender, EventArgs e)
        {
            try
            {
                // 訊息列表與序號集合
                List<string> SNList = new List<string>();

                #region OEM-STD_Bios

                // 判斷 OEM BIOS 欄位是否啟用；啟用則使用 OEM-STD_Bios 的列印流程
                if (txt_Oem_Bios.Enabled == true)
                {
                    // 組合列印檔案路徑（下載資料夾 + DLfilename）
                    string printlabe = download_Path + "\\" + DLfilename;

                    // 若數量 > 26，需分頁處理：先列印前 26 筆
                    if (int.Parse(txt_Oem_Qty.Text) > 26)
                    {
                        // 取得前 26 筆序號並加入 SNList
                        for (int i = 0; i < 26; i++)
                        {
                            SNList.Add(Sn_Oem_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }

                        // 使用 OemPrint2 列印前 26 筆序號；成功則繼續列印剩餘序號
                        if (OemPrint2.PrintLabel(printlabe, txt_Oem_Wo.Text, txt_Oem_Qty.Text.Trim(), txt_Oem_Bios.Text.Trim(), SNList, List_Msg) == true)
                        {
                            // 清除 SNList，準備收集剩餘序號供下一次列印使用
                            SNList.Clear();

                            // 收集第 27 筆到結尾的序號
                            for (int i = 26; i < Sn_Oem_GV.Rows.Count; i++)
                            {
                                SNList.Add(Sn_Oem_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                            }

                            // 使用 BoardPrint2 列印剩餘序號（可為不同的標籤樣式或列印邏輯）
                            BoardPrint2.PrintLabel(Boardpath2, SNList, List_Msg);
                        }
                    }
                    else // 若數量 <= 26，直接將所有序號一次列印
                    {
                        // 走訪所有已輸入的序號並加入 SNList
                        for (int i = 0; i < Sn_Oem_GV.Rows.Count; i++)
                        {
                            SNList.Add(Sn_Oem_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }

                        // 直接呼叫 OemPrint2 列印全部序號
                        OemPrint2.PrintLabel(printlabe, txt_Oem_Wo.Text.Trim(), txt_Oem_Qty.Text, txt_Oem_Bios.Text.Trim(), SNList, List_Msg);
                        // if (OemPrint2.PrintLabel(printlabe, WO, SNList) == true) // 原始備用呼叫（註解保留），未改動程式邏輯
                    }
                }

                #endregion

                #region 標準列印以外

                // 若不是 OEM-STD_Bios 模式，進行標準 BarTender 列印流程
                else
                {
                    // 清除顯示訊息並進行前置驗證
                    List_Oem_Msg.Items.Clear();

                    if (!string.IsNullOrWhiteSpace(txt_Oem_Qty.Text) && Sn_Oem_GV.Rows.Count < 1)
                    {
                        List_Oem_Msg.Items.Add("尚未輸入序號.....");
                        txt_Oem_Sn.Focus();
                        return;
                    }
                    if (!string.IsNullOrWhiteSpace(txt_Oem_Qty.Text) && Sn_Oem_GV.Rows.Count != int.Parse(txt_Oem_Qty.Text))
                    {
                        List_Oem_Msg.Items.Add("數量與輸入序號數量不符");
                        txt_Oem_Sn.Focus();
                        return;
                    }

                    // 設定列印參數、UI 狀態
                    int printCount = 1;
                    btn_Oem_Print.Enabled = false;
                    txt_Oem_Sn.Text = string.Empty;
                    txt_Oem_Sn.Enabled = false;

                    List_Oem_Msg.Items.Add("列印中.....");

                    // 使用 BarTender Engine 進行列印
                    // 建立 BarTender 列印引擎與標籤文件物件
                    Engine engine = null;
                    // 建立標籤格式文件物件
                    LabelFormatDocument btFormat = null;
                    // 啟動 BarTender 列印引擎
                    engine = new Engine();
                    engine.Start();

                    // 開啟指定路徑的標籤格式文件
                    // 標籤檔案（通常是 .btw 格式）
                    btFormat = engine.Documents.Open(download_Path + "\\" + DLfilename);

                    // 設定標籤欄位值
                    #region 設定標籤欄位值
                    // 若使用者有填寫數量欄位，將 QTY 子字串設置
                    if (!string.IsNullOrWhiteSpace(txt_Oem_Qty.Text.Trim()))
                    {
                        btFormat.SubStrings["QTY"].Value = txt_Oem_Qty.Text.Trim();
                    }

                    // 若使用者有填寫重量欄位，將 Weight 子字串設置
                    if (!string.IsNullOrWhiteSpace(txt_Oem_Weight.Text.Trim()))
                    {
                        btFormat.SubStrings["Weight"].Value = txt_Oem_Weight.Text.Trim();
                    }

                    // 若有 MAC 欄位，則設定雙份列印與相關子字串（mac_n, sn_n, SN1, SN2）
                    if (!string.IsNullOrWhiteSpace(txt_Oem_Mac.Text.Trim()))
                    {
                        printCount = 2;// 列印標籤數
                        btFormat.SubStrings["mac_n"].Value = txt_Oem_Mac.Text.Trim();
                        btFormat.SubStrings["sn_n"].Value = Sn_Oem_GV.Rows[0].Cells["SN"].Value.ToString().Trim();

                        // 標籤檔內預期的欄位名稱：SN1, SN2, ...
                        btFormat.SubStrings["SN1"].Value = Sn_Oem_GV.Rows[0].Cells["SN"].Value.ToString().Trim();
                        btFormat.SubStrings["SN2"].Value = txt_Oem_Mac.Text.Trim();
                    }

                    // 若有多個序號，將對應的 SN 子字串逐一設置（SN1..SNn）
                    if (Sn_Oem_GV.Rows.Count > 0 && txt_Oem_Qty.Enabled == true)
                    {
                        for (int i = 0; i < Sn_Oem_GV.Rows.Count; i++)
                        {
                            string SN_Name = "SN" + (i + 1).ToString();
                            btFormat.SubStrings[SN_Name].Value = Sn_Oem_GV.Rows[i].Cells["SN"].Value.ToString().Trim(); // 標籤檔中所設定的欄位名稱
                        }
                    }
                    #endregion

                    // 設定要列印的份數，並送出列印
                    btFormat.PrintSetup.IdenticalCopiesOfLabel = printCount; // 列印標籤數

                    // 取得目前設定的印表機名稱，並顯示在 List_Msg 控制項
                    string printerName = btFormat.PrintSetup.PrinterName;
                    List_Msg.Items.Add($"使用的印表機：{printerName}");

                    /*在 BarTender Designer 軟體中，開啟你的 .btw 標籤檔案。
                     * 點選「檔案」→「列印」→「選擇印表機」。
                     */
                    // 在 BarTender Designer 軟體中，開啟你的 .btw 標籤檔案。
                    // 執行列印
                    btFormat.Print();

                    // 停止 Engine
                    engine.Stop(); // 停止 BarTender 引擎
                }

                #endregion

                // 列印完成後清理 UI 與狀態
                Sn_Oem_GV.Rows.Clear();
                List_Oem_Msg.Items.Clear();
                List_Oem_Msg.Items.Add("列印完成.....");
                txt_Oem_Sn.Enabled = true;
                txt_Oem_Mac.Text = string.Empty;
                if (txt_Oem_Mac.Enabled == true)
                {
                    txt_Oem_Mac.Focus();
                }

            }
            catch (Exception ex)
            {
                List_Oem_Msg.Items.Add("999行\r\n" + ex.Message);
                // 發生例外時確保按鈕可用，並視需要可在此處加入日誌或顯示更詳細錯誤資訊
                btn_Oem_Print.Enabled = true;
            }
        }

        // 計畫 (Pseudocode) - 詳細步驟說明：
        // 1. 當使用者點擊 re-print 的 checkbox (rcb_Re_Print) 時，觸發 Click 事件。
        // 2. 檢查 rcb_Re_Print 是否被勾選：
        //    - 若已勾選：將輸入欄位 txt_reprint_Sn 設為不可用 (Enabled = false)，以避免使用者輸入。
        //    - 若未勾選：將 txt_reprint_Sn 設為可用 (Enabled = true)。
        // 3. 不做其他業務邏輯或資料存取，只負責 UI 的狀態切換。
        // 4. 程式碼採用簡潔寫法，直接將 Enabled 屬性與 rcb_Re_Print.Checked 狀態對應，避免重複判斷。

        #region 工單外箱Label設定
        // 計畫 (Pseudocode) - 詳細步驟：
        // 1. 清空 List_Wo_Set_Msg 顯示訊息。
        // 2. 驗證 txt_Wo_Set 是否為空：若空則顯示錯誤並回復焦點，結束方法。
        // 3. 判斷三種列印模式 rdb_Standard / rdb_Oem_Of / rdb_Oem_On，並設定 Print_Type 字串；若都未選，提示使用者並回復焦點，結束方法。
        // 4. 組裝 SQL INSERT 語句：將 Record_Time (目前時間)、Work_Order (txt_Wo_Set)、Print_Type 寫入 Print_Wo_Setting_Table。
        // 5. 呼叫 db.Exsql(InsSql) 執行寫入：
        //    - 若成功：在 List_Wo_Set_Msg 顯示設定完成並清空 txt_Wo_Set。
        //    - 若失敗：在 List_Wo_Set_Msg 顯示「資料庫寫入錯誤」。
        // 6. 捕捉可能的例外（由呼叫端或框架處理），本方法不拋出例外到上層以維持 UI 穩定性。

        /// <summary>
        /// 工單外箱Labely設定-處理「工單設定」按鈕的 Click 事件，將使用者所選的列印模式與工單寫入資料庫中的 Print_Wo_Setting_Table。
        /// </summary>
        /// <param name="sender">事件來源物件 (通常為按鈕)。</param>
        /// <param name="e">事件參數資訊。</param>
        /// <remarks>
        /// - 若 txt_Wo_Set 為空，會顯示提示並將焦點移回 txt_Wo_Set。<br/>
        /// - 支援三種列印模式：Standard、Oem_Off_Line、Oem_On_Line；若未選擇任何模式則提示使用者。<br/>
        /// - 成功寫入資料庫後會在 List_Wo_Set_Msg 顯示完成訊息並清空輸入欄位；失敗時顯示錯誤訊息。<br/>
        /// - 本方法不做進一步的錯誤回報或日誌，若需要可在 catch 區段補充日誌功能。
        /// </remarks>
        private void Btn_Wo_Set_Click(object sender, EventArgs e)
        {
            List_Wo_Set_Msg.Items.Clear();
            string msg = string.Empty, Print_Type = string.Empty;

            // 1) 檢查工單是否輸入
            if (string.IsNullOrWhiteSpace(txt_Wo_Set.Text))
            {
                msg += "尚未輸入工單" + "\r\n";
                MessageBox.Show(msg);
                txt_Wo_Set.Focus();
                return;
            }

            // 2) 決定列印模式
            if (rdb_Standard.IsChecked)
            {
                Print_Type = "Standard";
            }
            else if (rdb_Oem_Of.IsChecked)
            {
                Print_Type = "Oem_Off_Line";
            }
            else if (rdb_Oem_On.IsChecked)
            {
                Print_Type = "Oem_On_Line";
            }
            else
            {
                List_Wo_Set_Msg.Items.Add("請設定列印模式");
                txt_Wo_Set.Focus();
                return;
            }

            // 3) 組裝 SQL 並寫入資料庫
            string InsSql = " INSERT INTO [Print_Wo_Setting_Table] (Record_Time,Work_Order,Print_Type) VALUES("
                               + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                               + "'" + txt_Wo_Set.Text.Trim() + "',"
                               + "'" + Print_Type + "')";
            try
            {
                if (db.Exsql(InsSql) == true)
                {
                    List_Wo_Set_Msg.Items.Add(txt_Wo_Set.Text + " 設定完成");
                    txt_Wo_Set.Text = string.Empty;
                }
                else
                {
                    List_Wo_Set_Msg.Items.Add("資料庫寫入錯誤");
                }
            }
            catch (Exception ex)
            {
                List_Oem_Msg.Items.Add("1104行\r\n" + ex.Message);
                // 避免拋出例外造成 UI 異常，僅保留簡短回報
            }
        }

        /// <summary>
        /// 工單外箱Labely設定-處理「工單設定搜尋」按鈕的 Click 事件。
        /// </summary>
        /// <remarks>
        /// 此方法會：
        /// - 驗證使用者是否輸入查詢的工單號（txt_Wo_Serach）。若未輸入則顯示錯誤並結束。
        /// - 以輸入的工單號查詢資料表 Print_Wo_Setting_Table，取得最新一筆的 Work_Order 與 Print_Type。
        /// - 根據查詢結果，切換畫面中的列印頁籤與對應的欄位：
        ///   * 若 Print_Type 為 "Standard"：將工單帶入標準列印頁面 (Standard_Page)。
        ///   * 若 Print_Type 為 "Oem_On_Line"：將工單帶入 OEM 上線列印頁面 (Oem_Page)。
        ///   * 其他值視為 OFF Line，僅顯示訊息。
        /// - 若查無資料，會在 List_Wo_Search_Msg 顯示查無設定訊息。
        ///
        /// 注意：
        /// - 本方法直接存取 UI 控制項並呼叫資料庫存取物件 db.reDs，假設該物件已正確初始化。
        /// - 本方法對於資料庫查詢結果沒有做更多防護性檢查（例如 null 檢查），呼叫端應確保 db.reDs 行為一致。
        /// </remarks>
        /// <param name="sender">事件來源 (object)，通常為按鈕物件。</param>
        /// <param name="e">事件參數 (EventArgs)。</param>
        private void Btn_Wo_Serch_Click(object sender, EventArgs e)
        {
            List_Wo_Search_Msg.Items.Clear();
            string msg = string.Empty, Print_Type = string.Empty, Work_Order = string.Empty;
            //To do ini upload ftp
            if (string.IsNullOrWhiteSpace(txt_Wo_Serach.Text))
            {
                msg = msg + "尚未輸入工單" + "\r\n";
                MessageBox.Show(msg);
                txt_Wo_Serach.Focus();
                return;
            }
            string sqlCmd = "SELECT [Work_Order],[Print_Type] FROM [Print_Wo_Setting_Table] where [Work_Order] ='" + txt_Wo_Serach.Text.Trim() + "' order by Record_Time desc ";
            DataSet ds = db.reDs(sqlCmd);
            if (ds.Tables[0].Rows.Count != 0)
            {
                Print_Type = ds.Tables[0].Rows[0]["Print_Type"].ToString().Trim();
                Work_Order = ds.Tables[0].Rows[0]["Work_Order"].ToString().Trim();

                if (Print_Type == "Standard")
                {
                    List_Wo_Search_Msg.Items.Add("此工單為標準品列印");
                    txt_Wo.Text = txt_Wo_Serach.Text;
                    Print_Page.SelectedPage = Standard_Page;
                    txt_Wo_Serach.Text = string.Empty;
                }
                else if (Print_Type == "Oem_On_Line")
                {
                    List_Wo_Search_Msg.Items.Add("此工單為OEM/ODM ON Line列印");
                    txt_Oem_Wo.Text = txt_Wo_Serach.Text;
                    Print_Page.SelectedPage = Oem_Page;
                    txt_Wo_Serach.Text = string.Empty;
                }
                else
                {
                    List_Wo_Search_Msg.Items.Add("此工單為OFF Line");
                }
            }
            else
            {
                List_Wo_Search_Msg.Items.Add("查無工單:" + txt_Wo_Serach.Text + "設定");
            }


        }

        #endregion

        /// <summary>
        /// OEM外箱Label設定-清除 OEM/ODM 列印頁面所有輸入欄位與狀態。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為「清除」按鈕。</param>
        /// <param name="e">事件參數。</param>
        /// <remarks>
        /// 此方法會將 OEM 工單、料號、數量、重量、重列印序號等欄位清空，
        /// 並重設相關控制項的啟用狀態，清除序號列表與訊息顯示。
        /// 適用於使用者需重設 OEM/ODM 列印頁面時呼叫。
        /// </remarks>
        private void Btn_Oem_Clear_Click(object sender, EventArgs e)
        {
            List_Oem_Msg.Items.Clear();               // 清空訊息顯示區

            // 清除 OEM/ODM 列印頁面所有輸入欄位與狀態
            txt_Oem_Wo.Text = string.Empty;           // 清空工單欄位
            txt_Oem_Pn.Text = string.Empty;           // 清空料號欄位
            txt_Oem_Qty.Text = string.Empty;          // 清空數量欄位
            txt_Oem_Weight.Text = string.Empty;       // 清空重量欄位
            txt_reprint_Oem_Sn.Text = string.Empty;   // 清空重列印序號欄位
            rcb_Oem_Re_Print.Checked = false;         // 取消重列印核取方塊
            txt_Oem_Qty.Enabled = false;              // 數量欄位設為不可編輯
            txt_Oem_Weight.Enabled = false;           // 重量欄位設為不可編輯
            txt_reprint_Oem_Sn.Enabled = false;       // 重列印序號欄位設為不可編輯
            Sn_Oem_GV.Rows.Clear();                   // 清空序號列表


        }

        /// <summary>
        /// OEM外箱Label列印上傳-Btn_Up_Clear_Click 方法為「清除」按鈕的事件處理函式。
        /// 當使用者點擊此按鈕時，會清空上傳工單欄位 (txt_Up_WO)、標籤料號欄位 (txt_Up_Pn)、
        /// 套版檔案路徑欄位 (txt_Btw_Path)，並將預覽圖片 (UP_PictureBox) 設為空，
        /// 同時清除訊息列表 (List2_Msg) 的所有內容。
        /// 此方法主要用於重設上傳套版相關的 UI 狀態，方便使用者重新操作。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為「清除」按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Up_Clear_Click(object sender, EventArgs e)
        {
            // 清空上傳工單欄位、標籤料號欄位、套版檔案路徑欄位
            txt_Up_WO.Text = string.Empty;      // 工單欄位設為空字串
            txt_Up_Pn.Text = string.Empty;      // 標籤料號欄位設為空字串
            txt_Btw_Path.Text = string.Empty;   // 套版檔案路徑欄位設為空字串
            UP_PictureBox.Image = null;         // 預覽圖片設為空
            List2_Msg.Items.Clear();            // 清空訊息列表
        }

        #endregion

        /// <summary>
        /// 判斷輸入值是否在指定範圍內。
        /// 支援 long 與 string 型別的比較。
        /// </summary>
        /// <param name="input">要判斷的輸入值，可為 long 或 string。</param>
        /// <param name="min">範圍下限，可為 long 或 string。</param>
        /// <param name="max">範圍上限，可為 long 或 string。</param>
        /// <returns>若 input 在 min 與 max 範圍內則回傳 true，否則回傳 false。</returns>
        /// <exception cref="ArgumentException">若輸入型別不支援則拋出例外。</exception>
        public static bool IsInRange(object input, object min, object max)
        {
            // 判斷 input 是否為 long 型別，且 min、max 也為 long 型別
            if (input is long inputNumber && min is long minNumber && max is long maxNumber)
            {
                // 檢查 inputNumber 是否在 minNumber 與 maxNumber 範圍內
                return inputNumber >= minNumber && inputNumber <= maxNumber;
            }
            // 判斷 input 是否為 string 型別，且 min、max 也為 string 型別
            else if (input is string inputString && min is string minString && max is string maxString)
            {
                // 使用字串比較，判斷 inputString 是否在 minString 與 maxString 範圍內
                return string.Compare(inputString, minString) >= 0 && string.Compare(inputString, maxString) <= 0;
            }
            else
            {
                // 若型別不支援則丟出例外
                throw new ArgumentException("輸入的數據類型不支持。");
            }
        }

        /// <summary>
        /// 處理 OEM/ODM 序號輸入欄位的 KeyPress 事件。
        /// 當使用者在 txt_Oem_Sn 輸入序號並按下 Enter 鍵時，執行以下流程：
        /// 1. 檢查是否已輸入數量，若未輸入則提示並回復焦點。
        /// 2. 若啟用 MAC 欄位，則查詢工單對應的序號區間，判斷輸入序號是否在合法範圍內。
        /// 3. 檢查序號是否重複，若重複則提示並終止流程。
        /// 4. 若序號未重複且未超過數量上限，則新增序號至序號列表，並寫入資料庫。
        /// 5. 若序號數量已達設定數量或 MAC 欄位啟用，則自動觸發列印流程。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Oem_Sn 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void Txt_Oem_Sn_KeyPress(object sender, KeyPressEventArgs e)
        {
            bool AddSN = true;
            string Bind_SN = string.Empty;
            // 檢查是否已輸入數量
            if (string.IsNullOrWhiteSpace(txt_Oem_Qty.Text) && txt_Oem_Qty.Enabled == true)
            {
                List_Oem_Msg.Items.Add("請先輸入數量 ");
                txt_Oem_Qty.Focus();
                return;
            }

            // 按下 Enter 鍵時執行
            if (Convert.ToInt32(e.KeyChar) == 13)
            {
                #region IGT-OEM
                // 若啟用 MAC 欄位，檢查序號區間
                if (txt_Oem_Mac.Enabled)
                {
                    // 查詢對應的 Eversun 工單
                    string sqlCmd = "select top(1)Eversun_WoNo from Print_CustomWONO_Table where Avalue_WoNo = '" + txt_Oem_Wo.Text.Trim() + "' order by sno desc";
                    DataSet ds = db.reDs(sqlCmd);
                    string eversun_wono = ds.Tables[0].Rows[0][0].ToString();
                    string WipInfo = Auto_Route.WipbarcodeOther(eversun_wono.Trim());

                    EversunWoNo descJsonStu = JsonConvert.DeserializeObject<EversunWoNo>(WipInfo.ToString());//反序列化

                    string inputString = txt_Oem_Sn.Text.Trim();
                    bool isInRange = false;
                    if (long.TryParse(inputString, out long inputNumber))
                    {
                        // 判斷 long 型別輸入是否在範圍內
                        isInRange = IsInRange(inputNumber, descJsonStu.startNO, descJsonStu.endNO);
                    }
                    else
                    {
                        // 判斷 string 型別輸入是否在範圍內
                        isInRange = IsInRange(inputString, descJsonStu.startNO, descJsonStu.endNO);
                    }

                    // 若不在合法範圍則提示錯誤
                    if (isInRange)
                    {
                        AddSN = true;
                    }
                    else
                    {
                        List_Oem_Msg.Items.Add(txt_Oem_Sn.Text.Trim().ToUpper() + "序號區間錯誤!! ");
                        AddSN = false;
                        return;
                    }
                }
                #endregion
                // 檢查序號是否重複
                if (!string.IsNullOrWhiteSpace(txt_Oem_Sn.Text.Trim().ToUpper()))
                {
                    for (int i = 0; i < Sn_Oem_GV.Rows.Count; i++)
                    {
                        if (txt_Oem_Sn.Text.Trim().ToUpper() == Sn_Oem_GV.Rows[i].Cells["SN"].Value.ToString().ToUpper())
                        {
                            List_Oem_Msg.Items.Add(txt_Oem_Sn.Text.Trim().ToUpper() + "序號重複 ");
                            AddSN = false;
                            break;
                        }
                        else
                        {
                            AddSN = true;
                        }
                    }

                    // 若序號可新增且未超過數量上限
                    if (AddSN == true)
                    {
                        if (Sn_Oem_GV.Rows.Count == 0)
                        {
                            Item = "1";
                            List_Oem_Msg.Items.Clear();
                            Bind_SN = txt_Oem_Sn.Text.Trim().ToUpper();
                        }
                        else
                        {
                            Item = (Sn_Oem_GV.Rows.Count + 1).ToString();
                            Bind_SN = Sn_Oem_GV.Rows[0].Cells["SN"].Value.ToString().ToUpper();
                            if (int.Parse(Item) > int.Parse(txt_Oem_Qty.Text.Trim()))
                            {
                                List_Oem_Msg.Items.Add(txt_Oem_Sn.Text.Trim().ToUpper() + "EEROR!! 數量滿箱");
                                return;
                            }
                        }

                        SN = txt_Oem_Sn.Text.Trim().ToUpper();

                        Sn_Oem_GV.Rows.Add(new object[] { Item, SN, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") });

                        #region 資料庫寫入SN資訊
                        // 新增序號綁定資料至資料庫
                        string InsSql = " INSERT INTO [Print_SN_Bind_Carton_Table] " +
                                        "(Record_Time,Work_Order,Quantity,SN,Bind_SN,Weight,MAC) VALUES("
                                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                                    + "'" + txt_Oem_Wo.Text.Trim() + "',"
                                    + "'" + txt_Oem_Qty.Text.Trim() + "',"
                                    + "'" + txt_Oem_Sn.Text.Trim().ToUpper() + "',"
                                    + "'" + Bind_SN + "',"
                                    + "'" + txt_Oem_Weight.Text.Trim().ToUpper() + "',"
                                    + "'" + txt_Oem_Mac.Text.Trim().ToUpper() + "')";
                        if (db.Exsql(InsSql) == true)
                        {
                            txt_Oem_Sn.Text = string.Empty;
                            // 若序號數量已達設定數量或 MAC 欄位啟用則自動列印
                            if (Item == txt_Oem_Qty.Text.Trim() || txt_Oem_Mac.Enabled == true)
                            {
                                // 新增防呆：彈出確認視窗 20250912 By Jesse
                                var result = MessageBox.Show("是否確認要列印？", "列印確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                                if (result == DialogResult.OK)
                                {
                                    Btn_Oem_Print_Click(sender, e);
                                }
                                // 按取消則不執行

                            }
                        }
                        else
                        {
                            List_Oem_Msg.Items.Add("資料庫寫入錯誤");
                        }
                        #endregion
                    }
                }
                else
                {
                    List_Oem_Msg.Items.Add("請檢查序號 ");
                }
            }
        }

        /// <summary>
        /// 處理 OEM/ODM 序號列表刪除事件。
        /// 當使用者於 Sn_Oem_GV 刪除序號列時，
        /// 會根據目前選取的序號與時間，
        /// 刪除資料庫中對應的 [Print_SN_Bind_Carton_Table] 資料。
        /// 若刪除失敗則顯示錯誤訊息。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為序號列表控制項。</param>
        /// <param name="e">包含刪除事件資料的 GridViewRowCancelEventArgs。</param>
        private void Sn_Oem_GV_UserDeletingRow(object sender, GridViewRowCancelEventArgs e)
        {
            // 取得目前選取的序號與時間，組成刪除 SQL，刪除 [Print_SN_Bind_Carton_Table] 資料表中對應的資料
            string DelSN = Sn_Oem_GV.CurrentRow.Cells["SN"].Value.ToString();
            string DelTime = Convert.ToDateTime(Sn_Oem_GV.CurrentRow.Cells["Time"].Value).ToString("yyyy-MM-dd HH:mm:ss");
            string DelSql = " Delete  [Print_SN_Bind_Carton_Table] where Sn = '" + DelSN + "' and Record_Time = '" + DelTime + "'";
            if (db.Exsql(DelSql) == true)
            {
                // 刪除成功不需額外處理
            }
            else
            {
                // 若刪除失敗則顯示錯誤訊息
                List_Msg.Items.Add("資料庫刪除錯誤");
            }
        }

        /// <summary>
        /// 處理 OEM/ODM 重列印核取方塊的 Click 事件。
        /// 當 <c>rcb_Oem_Re_Print</c> 被勾選時，將 <c>txt_reprint_Oem_Sn</c> 設為不可編輯，否則恢復可編輯。
        /// 此方法僅負責切換 UI 控制項的啟用狀態，不涉及任何業務邏輯或資料存取。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 <c>rcb_Oem_Re_Print</c> 控制項。</param>
        /// <param name="e">事件參數。</param>
        private void Rcb_Oem_Re_Print_Click(object sender, EventArgs e)
        {
            // 當 OEM/ODM 重列印核取方塊被勾選時，將重列印序號欄位設為不可編輯；否則恢復可編輯
            if (rcb_Oem_Re_Print.Checked == true)
            {
                // 勾選時禁止輸入
                txt_reprint_Oem_Sn.Enabled = false;
            }
            else
            {
                // 取消勾選時允許輸入
                txt_reprint_Oem_Sn.Enabled = true;
            }
        }

        /// <summary>
        /// 選擇PHIL_EPC_WHL的檔案路徑
        /// </summary>
        /// <param name="modeltype">選擇01~08R</param>
        /// <param name="modelLab">選擇大張或小張Lab</param>
        /// <returns></returns>
        public string Phl_EPCWHL_Path(string modeltype, string modelLab)
        {
            string EPCWHL_Model = "EPC-WHL-43-C1-";
            string labFolder = download_Path + "\\" + EPCWHL_Model + modeltype + "\\";
            //Debug\Btw_Folder\EPC-WHL-43-C1-01R\E2090000809R.btw
            string path = "";
            switch (modeltype)
            {
                case "01R":
                    path = labFolder + "E2090000809R.btw";
                    break;
                case "02R":
                    path = labFolder + "E2090000809R.btw";
                    break;
                case "03R":
                    path = labFolder + "E2090000809R.btw";
                    break;
                case "04R":
                    path = labFolder + "E2090000809R.btw";
                    break;
                case "05R":
                    path = labFolder + "E2090000809R_170.btw";
                    break;
                case "06R":
                    path = labFolder + "E2090000809R_171.btw";
                    break;
                case "07R":
                    path = labFolder + "E2090000809R_172.btw";
                    break;
                case "08R":
                    path = labFolder + "E2090000809R_173.btw";
                    break;
                default:
                    break;
            }
            if (modelLab == "Small")
            {
                path = PhilEPC;
            }
            return path;
        }
        /// <summary>
        /// 處理 PHIL_EPC_WHL 標籤列印按鈕的 Click 事件。
        /// 根據使用者選擇的大張或小張標籤模式，取得工單相關資訊，並呼叫 PhilEPC_Print.PrintLabel 進行列印。
        /// 列印完成後更新訊息列表，若發生例外則顯示錯誤訊息。
        /// </summary>
        /// <param name="sender">事件來源物件（通常為按鈕）。</param>
        /// <param name="e">事件參數（EventArgs）。</param>
        private void btn_P_EPC_Click(object sender, EventArgs e)
        {
            try
            {
                // 取得工單相關資訊（大張標籤模式）
                string partNumber = "";   // 機種料號
                string ecn = "";         // ECN 編號
                string biosVer = "";     // BIOS 版本
                string bioscs = "";      // BIOS CS
                string shiftStr = "";    // 工單前10碼

                // 若選擇大張標籤，則依工單查詢 SFIS 資料
                if (rdb_Big.IsChecked)
                {
                    shiftStr = radTxt_Wono.Text.Trim().Substring(0, 10); // 取工單前10碼
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("{");
                    sb.AppendLine("\"Key\":\"" + "@Avalue.ZMO.SOP" + "\",");
                    sb.AppendLine("\"moid\":\"" + shiftStr.Trim() + "\",");
                    sb.AppendLine("}");
                    var test = SFISToJson.reDt3(sb); // 查詢 SFIS 資料
                    SFIS descJsonStu = JsonConvert.DeserializeObject<SFIS>(test.ToString());// 反序列化取得工單資訊

                    // 解析工單資訊
                    string[] PN = descJsonStu.productid.Split('-'); // 料號拆分
                    partNumber = PN[PN.Length - 1];                 // 取得料號
                    ecn = descJsonStu.Ecn.Trim().ToUpper();         // 取得 ECN 編號
                    biosVer = descJsonStu.biosVer.Trim();           // 取得 BIOS 版本
                    bioscs = descJsonStu.BIOSCS.Trim().ToUpper();   // 取得 BIOS CS
                }

                // 依據使用者選擇的標籤模式（大張或小張），設定 choiceRdb 變數
                string choiceRdb = "";
                if (rdb_Big.IsChecked)
                {
                    choiceRdb = "Big"; // 若選擇大張，設定為 "Big"
                }
                if (rdb_Small.IsChecked)
                {
                    choiceRdb = "Small"; // 若選擇小張，設定為 "Small"
                }

                // 取得對應的標籤檔案路徑
                string printlabe = Phl_EPCWHL_Path(partNumber, choiceRdb);
                #region Phil_EPC_WHL

                // 判斷 PhilEPC_Print 列印是否成功，成功則清空訊息並顯示「列印完成.....」
                if (PhilEPC_Print.PrintLabel(partNumber, choiceRdb, printlabe, ecn, biosVer, bioscs, radTxt_SN.Text, radTxt_Mac.Text))
                {
                    list_Phil_Msg.Items.Clear(); // 清空訊息列表
                    list_Phil_Msg.Items.Add("列印完成....."); // 顯示列印完成訊息
                }

                #endregion
            }
            catch (Exception ex)
            {
                // 發生例外時，將錯誤訊息加入訊息列表
                list_Phil_Msg.Items.Add("1694行\r\n" + ex.Message);
            }
        }

        /// <summary>
        /// 處理工單輸入欄位的 KeyPress 事件。
        /// 當使用者在 radTxt_Wono 輸入工單並按下 Enter 鍵時，
        /// 若尚未選擇標籤模式（大張或小張），則顯示提示訊息。
        /// 否則將焦點移至序號輸入欄位。
        /// 禁止使用者按下 Backspace 鍵以避免誤刪。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 radTxt_Wono 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void radTxt_Wono_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 禁止 Backspace 鍵
            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = true;
            }
            // 按下 Enter 鍵時執行
            if (e.KeyChar == 13)
            {
                // 若未選擇標籤模式則提示
                if (!rdb_Small.IsChecked && !rdb_Big.IsChecked)
                {
                    MessageBox.Show("請先選擇套版Label");
                }
                else
                {
                    // 移至序號輸入欄位
                    radTxt_SN.Focus();
                }
            }
        }

        /// <summary>
        /// 處理序號輸入欄位的 KeyPress 事件。
        /// 當使用者在 radTxt_SN 輸入序號並按下 Enter 鍵時，執行以下流程：
        /// 1. 在 btn_P_EPC_Click 觸發前，彈出 MessageBox 讓使用者確認是否要列印
        /// 2. 只有當使用者按下「確定」(DialogResult.OK)時才執行 btn_P_EPC_Click 的內容
        /// 3. 若按下「取消」則不執行列印流程
        /// 1. 檢查序號長度是否為 11 位，若不符則顯示錯誤訊息並清空欄位。
        /// 2. 禁止使用者按下 Backspace 鍵以避免誤刪。
        /// 3. 若序號長度正確且按下 Enter，判斷是否已選擇標籤模式（大張或小張），若未選擇則提示。
        /// 4. 若選擇大張標籤且工單欄位為空，則提示並將焦點移至工單欄位。
        /// 5. 若選擇小張標籤則直接觸發列印流程，否則將焦點移至 MAC 輸入欄位。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 radTxt_SN 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void radTxt_SN_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 檢查序號長度是否正確
            if (e.KeyChar == 13 && radTxt_SN.Text.Length != 11)
            {
                MessageBox.Show("序號長度錯誤, 請確認後重新輸入");
                radTxt_SN.Text = "";
                return;
            }

            // 禁止 Backspace 鍵
            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = true;
            }

            // 按下 Enter 且序號長度正確時執行
            if (e.KeyChar == 13 && radTxt_SN.Text.Length == 11)
            {
                // 未選擇標籤模式則提示
                if (!rdb_Small.IsChecked && !rdb_Big.IsChecked)
                {
                    MessageBox.Show("請先選擇套版Label");
                }
                else
                {
                    // 若工單欄位為空且選擇大張標籤，則提示並移至工單欄位
                    if (string.IsNullOrWhiteSpace(radTxt_Wono.Text) && rdb_Big.IsChecked)
                    {
                        radTxt_Wono.Focus();
                        MessageBox.Show("工單未輸入!");
                    }
                    else
                    {
                        // 若選擇小張標籤則直接列印，否則移至 MAC 欄位
                        if (rdb_Small.IsChecked)
                        {
                            // 新增防呆：彈出確認視窗 20250912 By Jesse
                            var result = MessageBox.Show("是否確認要列印？", "列印確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                            if (result == DialogResult.OK)
                            {
                                btn_P_EPC_Click(sender, e);
                            }
                            // 按取消則不執行
                        }
                        else
                        {
                            // 移至 MAC 輸入欄位
                            radTxt_Mac.Focus();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 處理「清除」按鈕的 Click 事件。
        /// 當使用者點擊此按鈕時，會清空 PHIL_EPC_WHL 標籤列印頁面的工單、序號、MAC 輸入欄位。
        /// 此方法主要用於重設 PHIL_EPC_WHL 標籤列印頁面的 UI 狀態，方便使用者重新操作。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為「清除」按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void btn_C_EPC_Click(object sender, EventArgs e)
        {
            // 清空 PHIL_EPC_WHL 標籤列印頁面的工單、序號、MAC 輸入欄位
            radTxt_Wono.Text = "";    // 工單欄位設為空
            radTxt_SN.Text = "";      // 序號欄位設為空
            radTxt_Mac.Text = "";     // MAC欄位設為空
        }

        /// <summary>
        /// 處理 rdb_Big 與 rdb_Small 標籤模式切換事件。
        /// 當使用者切換至「大張」模式時，啟用工單與序號輸入欄位並將焦點移至工單欄位。
        /// 當切換至「小張」模式時，停用工單與序號欄位並將焦點移至 MAC 輸入欄位。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 rdb_Big 或 rdb_Small 控制項。</param>
        /// <param name="e">事件參數。</param>
        private void rdb_Big_CheckStateChanged(object sender, EventArgs e)
        {
            // 當使用者切換至「大張」模式時，啟用工單與序號輸入欄位並將焦點移至工單欄位
            if (rdb_Big.IsChecked)
            {
                radTxt_SN.Enabled = true;
                radTxt_Wono.Enabled = true;
                radTxt_Wono.Focus();
            }
            // 當切換至「小張」模式時，停用工單與序號欄位並將焦點移至 MAC 輸入欄位
            if (rdb_Small.IsChecked)
            {
                radTxt_SN.Enabled = false;
                radTxt_Wono.Enabled = false;
                radTxt_Mac.Focus();
            }
        }

        /// <summary>
        /// 處理序號輸入欄位的 KeyDown 事件。
        /// 當使用者在 radTxt_SN 輸入序號時按下 Delete 鍵，
        /// 會禁止刪除動作以避免誤刪序號內容。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 radTxt_SN 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void radTxt_SN_KeyDown(object sender, KeyEventArgs e)
        {
            // 當使用者在序號輸入欄位按下 Delete 鍵時，禁止刪除動作以避免誤刪序號內容
            if (e.KeyCode == Keys.Delete)
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// 處理工單輸入欄位的 KeyDown 事件。
        /// 當使用者在 radTxt_Wono 輸入工單時按下 Delete 鍵，
        /// 會禁止刪除動作以避免誤刪工單內容。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 radTxt_Wono 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void radTxt_Wono_KeyDown(object sender, KeyEventArgs e)
        {
            // 當使用者在工單輸入欄位按下 Delete 鍵時，禁止刪除動作以避免誤刪工單內容
            if (e.KeyCode == Keys.Delete)
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// 處理 radTxt_Mac 控制項的 KeyDown 事件。
        /// 當使用者在 MAC 輸入欄位按下 Delete 鍵時，
        /// 會禁止刪除動作以避免誤刪 MAC 內容。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 radTxt_Mac 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void radTxt_Mac_KeyDown(object sender, KeyEventArgs e)
        {
            // 當使用者在 MAC 輸入欄位按下 Delete 鍵時，禁止刪除動作以避免誤刪 MAC 內容
            if (e.KeyCode == Keys.Delete)
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// 處理安勤標籤列印按鈕的 Click 事件。
        /// 此方法會根據使用者輸入的 MAC 及序號資訊，設定 BarTender 標籤檔案的子字串，
        /// 並執行列印流程。若有輸入 MAC，則列印兩份標籤，分別帶入 MAC 與序號。
        /// 列印前會於 radListControl1 顯示「列印中.....」提示。
        /// </summary>
        /// <param name="sender">事件來源物件（通常為按鈕）。</param>
        /// <param name="e">事件參數（EventArgs）。</param>
        private void btn_Advantech_Click(object sender, EventArgs e)
        {
            try
            {
                #region Phil_EPC_WHL
                // 列印流程：初始化列印份數、顯示列印中訊息、啟動 BarTender Engine 並開啟標籤檔案
                int printCount = 1;
                radListControl1.Items.Add("列印中.....");
                Engine engine = null;
                LabelFormatDocument btFormat = null;
                engine = new Engine();
                engine.Start();
                btFormat = engine.Documents.Open(download_Path + "\\" + DLfilename);

                // 若有輸入 MAC，則設定標籤子字串並列印兩份
                if (!string.IsNullOrWhiteSpace(txt_Oem_Mac.Text.Trim()))
                {
                    printCount = 2; // 列印份數設為2
                    btFormat.SubStrings["mac_n"].Value = txt_Oem_Mac.Text.Trim(); // 設定MAC欄位
                    btFormat.SubStrings["sn_n"].Value = Sn_Oem_GV.Rows[0].Cells["SN"].Value.ToString().Trim(); // 設定序號欄位
                    btFormat.SubStrings["SN1"].Value = Sn_Oem_GV.Rows[0].Cells["SN"].Value.ToString().Trim(); // 標籤檔中SN1欄位
                    btFormat.SubStrings["SN2"].Value = txt_Oem_Mac.Text.Trim();  // 標籤檔中SN2欄位
                }
                #endregion
            }
            catch (Exception ex)
            {
                List_Oem_Msg.Items.Add("1922行\r\n" + ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 處理安勤標籤條碼輸入欄位的 KeyPress 事件。
        /// 當使用者在 txt_AdvantechBarcode 輸入條碼並按下 Enter 鍵時，
        /// 會根據條碼查詢 advantechDt 資料表，取得對應的 PART_NO，並呼叫 AdvantechPrint.PrintLabel 進行列印。
        /// 列印完成後會更新 radListControl1 狀態，若查無資料則顯示提示訊息。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_AdvantechBarcode 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void txt_AdvantechBarcode_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 禁止使用者按下 Backspace 鍵
            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = true;
            }
            // 列印旗標
            bool printflag = false;

            // 按下 Enter 鍵時執行列印流程
            if (e.KeyChar == 13)
            {
                try
                {
                    // 依條碼查詢資料表，取得所有符合條件的 PART_NO
                    var s = from a in advantechDt.AsEnumerable()
                        .Where(o => o.Field<string>("BARCODE_NO") == txt_AdvantechBarcode.Text)
                            select a.Field<string>("PART_NO");
                    if (s != null)
                    {
                        radListControl1.Items.Add("列印中.....");
                        string printlabe = AdvantechLabel;

                        // 逐一列印所有符合條件的 PART_NO
                        // 依條碼查詢資料表，取得所有符合條件的 PART_NO，並逐一列印
                        foreach (var item in s)
                        {
                            // 呼叫 AdvantechPrint.PrintLabel 方法進行安勤標籤列印
                            // item.ToUpper()：將料號轉為大寫
                            // printlabe：標籤檔案路徑
                            // List_Msg：訊息顯示控制項
                            printflag = AdvantechPrint.PrintLabel(item.ToUpper(), printlabe, List_Msg);
                        }
                        // 根據列印結果更新 UI 狀態
                        if (printflag)
                        {
                            radListControl1.Items.Clear();                // 清空訊息列表
                            radListControl1.Items.Add("列印完成.....");   // 顯示列印完成訊息
                            txt_AdvantechBarcode.Text = "";               // 清空條碼輸入欄位
                        }
                        else
                        {
                            radListControl1.Items.Clear();                // 清空訊息列表
                            radListControl1.Items.Add("列印失敗.....");   // 顯示列印失敗訊息
                        }
                    }
                    else
                    {
                        MessageBox.Show("此條碼在文件資料庫中查無紀錄。");
                    }
                }
                catch (Exception er)
                {
                    radListControl1.Items.Add("1989行\r\n" + er.Message);
                }
            }
        }
        /// <summary>
        /// 讀取 Excel 檔案並將指定工作表 ("ITEM_TMP") 內容轉換為 DataTable。
        /// 支援 .xlsx (Excel 2007 以上) 及 .xls (Excel 97) 格式。
        /// 會自動將第一列作為欄位名稱，並依據儲存格型別自動判斷資料型態。
        /// 日期型別會轉為 yyyy-MM-dd 字串，數值型別直接存入，其他型別則以字串存入。
        /// 若讀取過程發生例外，會以訊息視窗顯示錯誤訊息。
        /// </summary>
        /// <param name="xlsFilename">Excel 檔案完整路徑。</param>
        /// <returns>回傳 DataTable，內容為 "ITEM_TMP" 工作表的所有資料。</returns>
        public DataTable LoadExcelAsDataTable(String xlsFilename)
        {
            // 取得檔案資訊
            FileInfo fi = new FileInfo(xlsFilename);
            // 開啟檔案串流
            using (FileStream fstream = new FileStream(fi.FullName, FileMode.Open))
            {
                // 根據副檔名選擇適當的 NPOI 工作簿類別
                IWorkbook wb;
                if (fi.Extension == ".xlsx")
                    wb = new XSSFWorkbook(fstream); // excel2007
                else
                    wb = new HSSFWorkbook(fstream); // excel97

                // 只取第一個sheet。
                ISheet sheet = wb.GetSheet("ITEM_TMP");

                // target
                DataTable table = new DataTable();

                // 由第一列取標題做為欄位名稱
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum; // 取得欄位數量
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    // 依照 Excel 標題建立 DataTable 欄位（以字串型別為主）
                    table.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue));
                }
                try
                {
                    // 略過第零列(標題列)，一直處理至最後一列
                    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;

                        DataRow dataRow = table.NewRow();

                        // 依先前取得的欄位數逐一設定欄位內容
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell != null)
                            {
                                // 根據儲存格型別分別處理
                                switch (cell.CellType)
                                {
                                    case CellType.Numeric:
                                        if (DateUtil.IsCellDateFormatted(cell))
                                        {   // 日期格式
                                            dataRow[j] = cell.DateCellValue.ToString("yyyy-MM-dd");
                                        }
                                        else
                                        {   // 數值格式
                                            if (cell.CellStyle.DataFormat == 0)
                                            {   // 一般數值
                                                dataRow[j] = cell.NumericCellValue;
                                            }
                                            else
                                            {   // 其他數值格式，如百分比、貨幣等
                                                dataRow[j] = cell.ToString();
                                            }
                                        }
                                        break;
                                    default: // 字串型別
                                        // 直接轉成字串
                                        dataRow[j] = cell.ToString();
                                        break;
                                }
                            }
                        }
                        table.Rows.Add(dataRow);
                    }
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
                return table;
            }
        }

        /// <summary>
        /// 處理 radTxt_Mac 控制項的 KeyPress 事件。
        /// 當使用者在 MAC 輸入欄位按下 Enter 鍵時，根據所選標籤模式（大張/小張）及 MAC 長度判斷是否可進行列印。
        /// 若條件不符則顯示錯誤訊息，否則進行防呆確認後觸發列印流程。
        /// 禁止使用者按下 Backspace 鍵以避免誤刪 MAC 內容。
        /// 若未選擇標籤模式則提示使用者。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 radTxt_Mac 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void radTxt_Mac_KeyPress(object sender, KeyPressEventArgs e)
        {
            //00045F + 6碼  安勤  ;拆分大小張
            // 判斷 MAC 輸入欄位的 KeyPress 事件，依據標籤模式與 MAC 長度進行防呆檢查
            bool flag = false;

            // 禁止使用者按下 Backspace 鍵
            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = true;
            }
            // 按下 Enter 鍵時執行
            if (e.KeyChar == 13)
            {
                // 若選擇小張標籤且 MAC 長度為 6，則通過檢查
                if (rdb_Small.IsChecked && radTxt_Mac.Text.Length == 6)
                {
                    //radTxt_Mac.Text = "";
                    flag = true;
                }
                // 若選擇大張標籤且 MAC 長度為 12，則通過檢查
                if (rdb_Big.IsChecked && radTxt_Mac.Text.Length == 12)
                {
                    //radTxt_Mac.Text = "";
                    flag = true;
                }
                // 若 flag 為 false，代表未選擇套版或 MAC 長度錯誤
                if (flag)
                {
                    // 若序號欄位為空且選擇大張標籤，則提示序號未輸入
                    if (string.IsNullOrWhiteSpace(radTxt_SN.Text) && rdb_Big.IsChecked)
                    {
                        radTxt_SN.Focus();
                        MessageBox.Show("序號未輸入!");
                    }
                    // 若工單欄位為空且選擇大張標籤，則提示工單未輸入
                    else if (string.IsNullOrWhiteSpace(radTxt_Wono.Text) && rdb_Big.IsChecked)
                    {
                        radTxt_Wono.Focus();
                        MessageBox.Show("工單未輸入!");
                    }
                    // 若 MAC 欄位為空，則提示 MAC 未輸入
                    else if (string.IsNullOrWhiteSpace(radTxt_Mac.Text))
                    {
                        radTxt_Mac.Focus();
                        MessageBox.Show("MAC未輸入!");
                    }
                    else
                    {
                        // 新增防呆：彈出確認視窗，僅在使用者按下「確定」時才執行列印  20250912 By Jesse
                        var result = MessageBox.Show("是否確認要列印？", "列印確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                        if (result == DialogResult.OK)
                        {
                            btn_P_EPC_Click(sender, e);
                        }
                        // 按取消則不執行
                    }
                }
                else
                {
                    // MAC 長度錯誤時提示使用者重新輸入
                    MessageBox.Show("MAC長度錯誤, 請確認後重新輸入");
                }
            }
            // 若未選擇標籤模式則提示使用者
            else if (!rdb_Small.IsChecked && !rdb_Big.IsChecked)
            {
                MessageBox.Show("請先選擇套版Label");
                return;
            }
        }

        /// <summary>
        /// 處理 OEM/ODM 數量輸入欄位的 KeyPress 事件。
        /// 僅允許使用者輸入數字（0-9）與 Backspace，其他按鍵將被禁止。
        /// 當輸入非數字或非 Backspace 時，會自動啟用序號輸入欄位 (txt_Oem_Sn)。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Oem_Qty 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void Txt_Oem_Qty_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允許輸入數字（0-9）和 Backspace，其餘按鍵禁止，並啟用序號輸入欄位
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
                txt_Oem_Sn.Enabled = true;
            }
        }

        /// <summary>
        /// 處理 OEM/ODM MAC 輸入欄位的 KeyPress 事件。
        /// 當使用者在 <c>txt_Oem_Mac</c> 輸入 MAC 並按下 Enter 鍵時，
        /// 會自動將焦點移至序號輸入欄位 <c>txt_Oem_Sn</c>，方便連續輸入序號。
        /// 此方法僅負責 UI 控制項的焦點切換，不涉及任何業務邏輯或資料存取。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 <c>txt_Oem_Mac</c> 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void txt_Oem_Mac_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 當使用者在 txt_Oem_Mac 輸入欄位按下 Enter 鍵時，自動將焦點移至序號輸入欄位 txt_Oem_Sn
            if (e.KeyChar == 13)
            {
                txt_Oem_Sn.Focus();
            }
        }

        /// <summary>
        /// 處理 OEM/ODM 重列印序號輸入欄位的 KeyPress 事件。
        /// 當使用者在 <c>txt_reprint_Oem_Sn</c> 輸入序號並按下 Enter 鍵時，
        /// 會清空 OEM/ODM 相關欄位，查詢資料庫取得綁定序號的所有 SN、數量、重量等資訊，
        /// 並自動填入序號列表與欄位，若序號數量已達設定值則自動觸發列印流程。
        /// 若查無資料則顯示錯誤訊息，並將焦點移回輸入欄位。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 <c>txt_reprint_Oem_Sn</c> 控制項。</param>
        /// <param name="e">事件參數，包含按鍵資訊。</param>
        private void Txt_reprint_Oem_Sn_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (Convert.ToInt32(e.KeyChar) == 13)
                {
                    // 清空 OEM/ODM 相關欄位與序號列表
                    #region 清空欄位
                    txt_Oem_Sn.Enabled = false;           // 禁用序號輸入欄位
                    txt_Oem_Mac.Text = string.Empty;      // 清空 MAC 欄位
                    txt_Oem_Bios.Text = string.Empty;     // 清空 BIOS 欄位
                    txt_Oem_Qty.Text = string.Empty;      // 清空數量欄位
                    txt_Oem_Weight.Text = string.Empty;   // 清空重量欄位
                    List_Oem_Msg.Items.Clear();           // 清空訊息顯示區
                    Sn_Oem_GV.Rows.Clear();               // 清空序號列表
                    #endregion

                    // 檢查是否有輸入第一筆 SN
                    if (!string.IsNullOrWhiteSpace(txt_reprint_Oem_Sn.Text.Trim()))
                    {
                        // 查詢資料庫取得所有綁定序號的 SN、數量、重量等資訊
                        string sqlCmd = "SELECT * FROM [Print_SN_Bind_Carton_Table] where Bind_SN = '" + txt_reprint_Oem_Sn.Text.Trim() + "' ";
                        DataSet ds = db.reDs(sqlCmd);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            // 將查詢到的數量與重量填入欄位
                            txt_Oem_Qty.Text = ds.Tables[0].Rows[0]["Quantity"].ToString();
                            txt_Oem_Weight.Text = ds.Tables[0].Rows[0]["Weight"].ToString();

                            // 逐筆填入序號列表
                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                if (Sn_Oem_GV.Rows.Count == 0)
                                {
                                    Item = "1";
                                }
                                else
                                {
                                    Item = (Sn_Oem_GV.Rows.Count + 1).ToString();
                                }
                                Sn_Oem_GV.Rows.Add(new object[] { Item, ds.Tables[0].Rows[i]["SN"].ToString(), ds.Tables[0].Rows[i]["Record_Time"].ToString() });
                            }
                            // 若序號數量已達設定值則自動觸發列印流程
                            if (Item == txt_Oem_Qty.Text.Trim())
                            {
                                // 新增防呆：彈出確認視窗 20250912 By Jesse
                                var result = MessageBox.Show("是否確認要列印？", "列印確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                                if (result == DialogResult.OK)
                                {
                                    Btn_Oem_Print_Click(sender, e);
                                }
                                // 按取消則不執行
                            }
                        }
                        else
                        {
                            // 查無資料時顯示錯誤訊息
                            List_Oem_Msg.Items.Add("ERROR!!查無第一筆SN資料");
                        }
                    }
                    else
                    {
                        // 未輸入第一筆 SN 時顯示錯誤訊息並回復焦點
                        List_Oem_Msg.Items.Add("ERROR!!請輸入第一筆SN ");
                        txt_reprint_Oem_Sn.Focus();
                    }
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString());
            }
        }


        /// <summary>
        /// 處理「上傳套版」按鈕的 Click 事件。
        /// 當使用者點擊此按鈕時，會依序檢查套版檔案路徑、工單欄位、標籤料號欄位是否有輸入，
        /// 若有任一欄位為空則顯示提示訊息並回復焦點。
        /// 若所有欄位皆有輸入，則啟動 FTP_Up_Btw_thread 進行套版檔案的背景上傳作業。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為「上傳套版」按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void Btn_Up_File_Click(object sender, EventArgs e)
        {
            // 檢查上傳套版相關欄位，若有空值則提示並回復焦點，否則啟動 FTP 上傳背景執行緒
            List2_Msg.Items.Clear();
            string msg = string.Empty;
            // 檢查套版檔案路徑是否有輸入
            if (string.IsNullOrWhiteSpace(txt_Btw_Path.Text))
            {
                msg = msg + "尚未選擇套版檔案" + "\r\n";
                MessageBox.Show(msg);
                txt_Btw_Path.Focus();
                return;
            }
            // 檢查工單欄位是否有輸入
            if (string.IsNullOrWhiteSpace(txt_Up_WO.Text))
            {
                msg = msg + "工單空白" + "\r\n";
                MessageBox.Show(msg);
                txt_Up_WO.Focus();
                return;
            }
            // 檢查標籤料號欄位是否有輸入
            if (string.IsNullOrWhiteSpace(txt_Up_Pn.Text))
            {
                msg = msg + "標籤料號空白" + "\r\n";
                MessageBox.Show(msg);
                txt_Up_Pn.Focus();
                return;
            }

            // 若所有欄位皆有輸入，則啟動 FTP 上傳背景執行緒
            if (string.IsNullOrWhiteSpace(msg))
            {
                this.FTP_Up_Btw_thread.WorkerSupportsCancellation = true; //允許中斷
                this.FTP_Up_Btw_thread.RunWorkerAsync(); //呼叫背景程式
            }
        }





        /// <summary>
        /// 查詢指定工單號 (MOID) 是否存在於 Print_AvSFIS_Table 資料表。
        /// 若查詢到資料則回傳 true，否則回傳 false。
        /// 查詢結果會存入全域 dataSet 變數。
        /// </summary>
        /// <returns>
        /// 若查詢結果有資料則回傳 true，否則回傳 false。
        /// </returns>
        private bool AvSFIS_Method()
        {
            string sql = "select * from Print_AvSFIS_Table where MOID='" + txt_Wo.Text.Trim() + "' ";
            dataSet = db.reDs(sql);
            return (dataSet.Tables[0].Rows.Count == 0) ? false : true;
        }







        /// <summary>
        /// FTP_Dl_Btw_thread_DoWork 方法負責從 FTP 伺服器下載指定的套版檔案。
        /// 1. 先初始化下載結果旗標為 false。
        /// 2. 取得 FTP 連線資訊（伺服器、帳號、密碼）。
        /// 3. 建立 FTP 下載請求，設定下載檔案路徑與認證。
        /// 4. 取得 FTP 回應並開啟串流，將檔案內容分段寫入本機指定路徑。
        /// 5. 下載完成後關閉串流與回應物件，並將結果旗標設為 true。
        /// 6. 若發生例外則顯示錯誤訊息並將結果旗標設為 false。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 BackgroundWorker。</param>
        /// <param name="e">事件參數，包含 DoWork 相關資訊。</param>
        private void FTP_Dl_Btw_thread_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                TempUploadResult = false; // 初始化下載結果旗標為 false
                Getftp("Print"); // 取得 FTP 連線資訊

                // 建立 FTP 下載請求
                FtpWebRequest ftpRequest = (FtpWebRequest)FtpWebRequest.Create("ftp://" + ftpServer + "/" + DLfilename);
                NetworkCredential ftpCredential = new NetworkCredential(ftpuser, ftppassword);
                ftpRequest.Credentials = ftpCredential;
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                // 取得 FTP 回應並開啟串流
                FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                Stream ftpStream = ftpResponse.GetResponseStream();
                using (FileStream fileStream = new FileStream(download_Path + "\\" + DLfilename, FileMode.Create))
                {
                    int bufferSize = 2048; // 設定緩衝區大小
                    int readCount;
                    byte[] buffer = new byte[bufferSize];

                    readCount = ftpStream.Read(buffer, 0, bufferSize); // 讀取 FTP 串流資料
                    int allbye = (int)fileStream.Length;
                    Form.CheckForIllegalCrossThreadCalls = false;

                    // 分段寫入檔案內容到本機
                    while (readCount > 0)
                    {
                        fileStream.Write(buffer, 0, readCount); // 寫入本機檔案
                        readCount = ftpStream.Read(buffer, 0, bufferSize); // 繼續讀取下一段
                    }
                }
                ftpStream.Close(); // 關閉 FTP 串流
                ftpResponse.Close(); // 關閉 FTP 回應
                TempUploadResult = true; // 設定下載結果為成功
            }
            catch (Exception ex)
            {
                TempUploadResult = false; // 設定下載結果為失敗
                MessageBox.Show(ex.ToString()); // 顯示錯誤訊息
            }
        }

        /// <summary>
        /// FTP_Dl_Btw_thread_RunWorkerCompleted 事件處理函式。
        /// 當 FTP 套版下載背景執行緒完成時，根據下載結果更新 UI 狀態與提示訊息。
        /// 若下載成功，依據目前頁籤顯示「列印套版下載完成」訊息；
        /// 若下載失敗，則顯示「列印檔案下載失敗」訊息。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 BackgroundWorker。</param>
        /// <param name="e">包含事件資料的 RunWorkerCompletedEventArgs。</param>
        private void FTP_Dl_Btw_thread_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FTP_Dl_Btw_thread.Dispose();
            if (TempUploadResult == true)
            {
                if (Print_Page.SelectedPage.Name == "Standard_Page")
                {
                    List_Msg.Items.Add("列印套版下載完成");
                }
                else
                {
                    List_Oem_Msg.Items.Add("列印套版下載完成");
                    //DL_PictureBox.Image = null;
                    //using (Engine btEngine = new Engine(true))
                    //{
                    //    LabelFormatDocument labelFormat = btEngine.Documents.Open(download_Path + "\\" + DLfilename);

                    //    if (labelFormat != null)
                    //    {
                    //        Seagull.BarTender.Print.Messages m;
                    //        labelFormat.ExportPrintPreviewToFile(Application.StartupPath, @"\exp.bmp", ImageType.JPEG, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(300, 300), System.Drawing.Color.White, OverwriteOptions.Overwrite, true, true, out m);
                    //        labelFormat.ExportImageToFile(_bmp_path, ImageType.JPEG, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(300, 300), OverwriteOptions.Overwrite);

                    //        Image image = Image.FromFile(_bmp_path);
                    //        Bitmap NmpImage = new Bitmap(image);
                    //        DL_PictureBox.Image = NmpImage;
                    //        image.Dispose();
                    //    }
                    //    else
                    //    {
                    //        MessageBox.Show("生成圖片錯誤", "操作提示");
                    //    }
                    //}
                }
            }
            else
            {
                List2_Msg.Items.Add("列印檔案下載失敗");
            }
        }

        /// <summary>
        /// FTP_Up_Btw_thread_DoWork 方法負責將套版檔案上傳至 FTP 伺服器。
        /// 1. 先初始化上傳總數與結果旗標。
        /// 2. 檢查是否被中斷，若中斷則取消作業。
        /// 3. 檢查本機檔案是否存在，若存在則依序執行兩次 FTP 上傳：
        ///    (1) 上傳至 Print 伺服器（主要用途）。
        ///    (2) 上傳至 PrintBarTender 伺服器（同步備份用途）。
        /// 4. 每次上傳皆以 2048 bytes 為單位分段寫入，並於完成後關閉串流。
        /// 5. 若發生例外則顯示錯誤訊息並將結果旗標設為 false。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 BackgroundWorker。</param>
        /// <param name="e">事件參數，包含 DoWork 相關資訊。</param>
        private void FTP_Up_Btw_thread_DoWork(object sender, DoWorkEventArgs e)
        {

            //20171117 Jim 上傳總數初始化
            Sum_Of_SQLfile_size = 0;

            // 以下程式碼片段已加入中文註解，說明每個步驟的用途與邏輯
            TempUploadResult = false; // 初始化上傳結果旗標為 false
            if (FTP_Up_Btw_thread.CancellationPending) // 如果背景執行緒被要求中斷
                e.Cancel = true; // 取消本次作業
            this.FTP_Up_Btw_thread.WorkerReportsProgress = true; // 設定可回報進度
            BackgroundWorker worker = (BackgroundWorker)sender; // 取得觸發事件的 BackgroundWorker 物件
            //string temp_path = System.IO.Path.GetDirectoryName(dialog.FileName); // 取得檔案路徑（註解範例）

            // 以下區塊為 FTP 上傳套版檔案的主要流程，包含主伺服器與同步備份伺服器
            if (File.Exists(txt_Btw_Path.Text) == true)
            {
                // 取得檔案資訊
                FileInfo finfo = new FileInfo(txt_Btw_Path.Text);
                // 檔案名稱重新命名，加入時間戳記以避免覆蓋
                UPfilename = DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + filename;
                try
                {
                    // 取得 Print 伺服器 FTP 連線資訊
                    Getftp("Print");

                    // 建立 FTP 上傳請求
                    // 這段程式碼用於建立 FTP 上傳檔案的請求，並設定相關參數
                    FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://" + ftpServer + "/" + UPfilename); // 建立 FTP 請求物件，指定上傳路徑與檔名
                    request.KeepAlive = true; // 設定連線保持存活，避免中途斷線
                    request.UseBinary = true; // 設定傳輸模式為二進位，確保檔案正確上傳
                    request.Credentials = new NetworkCredential(ftpuser, ftppassword); // 設定 FTP 帳號與密碼
                    request.Method = WebRequestMethods.Ftp.UploadFile; // 設定 FTP 方法為上傳檔案
                    request.ContentLength = finfo.Length; // 設定上傳檔案的大小（位元組數）

                    // 取得 FTP 回應物件
                    FtpWebResponse response = request.GetResponse() as FtpWebResponse;
                    int buffLength = 2048;
                    byte[] buffer = new byte[buffLength];
                    int contentLen;
                    FileStream fs = File.OpenRead(txt_Btw_Path.Text);
                    Stream ftpstream = request.GetRequestStream();
                    contentLen = fs.Read(buffer, 0, buffer.Length);
                    int allbye = (int)finfo.Length;
                    Form.CheckForIllegalCrossThreadCalls = false;

                    int startbye = 0;
                    // 分段寫入檔案至 FTP
                    while (contentLen != 0)
                    {
                        startbye = contentLen + startbye;
                        ftpstream.Write(buffer, 0, contentLen);
                        // 若有進度條可在此更新進度
                        contentLen = fs.Read(buffer, 0, buffLength);
                    }
                    fs.Close();
                    ftpstream.Close();
                    response.Close();

                    TempUploadResult = true;
                }
                catch (Exception ftp)
                {
                    TempUploadResult = false;
                    MessageBox.Show(ftp.Message);
                }
                #region 套版同步上傳
                try
                {
                    // 取得 PrintBarTender 備份伺服器 FTP 連線資訊
                    Getftp("PrintBarTender");

                    ftpPutFile = "Eversun";
                    FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://" + ftpServer + "/" + ftpPutFile + "/" + UPfilename);
                    request.KeepAlive = true;
                    request.UseBinary = true;
                    request.Credentials = new NetworkCredential(ftpuser, ftppassword);
                    request.Method = WebRequestMethods.Ftp.UploadFile;
                    request.ContentLength = finfo.Length; // 設定上傳檔案大小

                    FtpWebResponse response = request.GetResponse() as FtpWebResponse;
                    int buffLength = 2048;
                    byte[] buffer = new byte[buffLength];
                    int contentLen;
                    FileStream fs = File.OpenRead(txt_Btw_Path.Text);
                    Stream ftpstream = request.GetRequestStream();
                    contentLen = fs.Read(buffer, 0, buffer.Length);
                    int allbye = (int)finfo.Length;
                    Form.CheckForIllegalCrossThreadCalls = false;

                    int startbye = 0;
                    // 分段寫入檔案至 FTP 備份
                    while (contentLen != 0)
                    {
                        startbye = contentLen + startbye;
                        ftpstream.Write(buffer, 0, contentLen);
                        // 若有進度條可在此更新進度
                        contentLen = fs.Read(buffer, 0, buffLength);
                    }
                    fs.Close();
                    ftpstream.Close();
                    response.Close();

                    TempUploadResult = true;
                }
                catch (Exception ftp)
                {
                    TempUploadResult = false;
                    MessageBox.Show(ftp.Message);
                }
                #endregion
            }

        }

        /// <summary>
        /// FTP_Up_Btw_thread_RunWorkerCompleted 方法在 FTP 上傳套版檔案完成後執行。
        /// 根據上傳結果，更新 UI 顯示、清空欄位，並將套版相關設定寫入資料庫。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">包含事件資料的 RunWorkerCompletedEventArgs。</param>
        /// 上傳套版
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FTP_Up_Btw_thread_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FTP_Up_Btw_thread.Dispose();
            if (TempUploadResult == true)
            {
                #region Label變數設定
                string Qty_Set = string.Empty // 數量設定
                    , Weight_Set = string.Empty // 重量設定
                    , Sn_Set = string.Empty // 序號設定
                    , Mac_Set = string.Empty // MAC設定
                    , Bios_Set = string.Empty; // BIOS設定

                List2_Msg.Items.Add("列印檔案上傳完成");
                // 若有勾選數量，設定為 Yes
                if (rck_Qty.Checked == true)
                {
                    Qty_Set = "Yes";
                }
                else
                {
                    // 未勾選則設定為 No
                    Qty_Set = "No";
                }
                if (rck_Weight.Checked == true)
                {
                    Weight_Set = "Yes";// 若有勾選重量，設定為 Yes
                }
                else
                {
                    Weight_Set = "No";// 未勾選則設定為 No
                }
                if (rck_Sn.Checked == true)
                {
                    Sn_Set = "Yes";// 若有勾選序號，設定為 Yes
                }
                else
                {
                    Sn_Set = "No";// 未勾選則設定為 No
                }
                if (rck_mac.Checked == true)
                {
                    Mac_Set = "Yes";// 若有勾選 MAC，設定為 Yes
                }
                else
                {
                    Mac_Set = "No";// 未勾選則設定為 No
                }
                if (rck_Bios.Checked == true)
                {
                    Bios_Set = "Yes";// 若有勾選 BIOS，設定為 Yes
                }
                else
                {
                    Bios_Set = "No";// 未勾選則設定為 No
                }
                #endregion
                string InsSql = " INSERT INTO [Print_Carton_Table] (" +
                                "Time,Filename,Pn,Wo,Qty_Set,Weight_Set,Sn_Set,Mac_Set,Bios_Set) VALUES("
                            + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                            + "'" + UPfilename + "',"
                            + "'" + txt_Up_Pn.Text.Trim() + "',"
                            + "'" + txt_Up_WO.Text.Trim() + "',"
                            + "'" + Qty_Set + "',"
                            + "'" + Weight_Set + "',"
                            + "'" + Sn_Set + "',"
                            + "'" + Mac_Set + "',"
                            + "'" + Bios_Set + "')";
                if (db.Exsql(InsSql) == true)
                {
                    string Print_Type = "Oem_On_Line";
                    InsSql = " INSERT INTO [Print_Wo_Setting_Table] (" +
                        "Record_Time,Work_Order,Print_Type) VALUES("
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                    + "'" + txt_Up_WO.Text.Trim() + "',"
                    + "'" + Print_Type + "')";

                    if (db.Exsql(InsSql) == true)
                    {

                        List2_Msg.Items.Add(txt_Up_WO.Text + " 工單設定完成");
                        txt_Up_WO.Text = string.Empty;

                    }
                    else
                    {
                        List2_Msg.Items.Add("資料庫寫入錯誤");
                    }

                    txt_Btw_Path.Text = string.Empty;
                    txt_Up_Pn.Text = string.Empty;
                    txt_Up_WO.Text = string.Empty;
                    UP_PictureBox.Image = null;

                }
                else
                {
                    List2_Msg.Items.Add("資料庫寫入錯誤");
                }
            }
            else
            {
                List2_Msg.Items.Add("列印檔案上傳失敗");
            }

        }

        /// <summary>
        /// 列印 BarTender 標籤檔案並產生預覽圖。
        /// </summary>
        /// <param name="isPreView">是否僅產生預覽圖（true：只產生預覽圖，不列印；false：產生預覽圖並列印）。</param>
        /// <remarks>
        /// 此方法會使用 BarTender Engine 開啟指定路徑的標籤檔案，
        /// 若 <paramref name="isPreView"/> 為 true，則僅產生預覽圖並顯示於 UP_PictureBox。
        /// 若 <paramref name="isPreView"/> 為 false，則會依據 _PrinterName 設定列印機並執行列印。
        /// 若未設定印表機名稱則顯示提示訊息。
        /// </remarks>
        void PrintBar(bool isPreView = false)
        {

            using (Engine btEngine = new Engine(true))
            {
                LabelFormatDocument labelFormat = btEngine.Documents.Open(txt_Btw_Path.Text);

                try
                {

                    //labelFormat.SubStrings.SetSubString("WO", name_textBox.Text);
                    //labelFormat.SubStrings.SetSubString("PN", age_textBox.Text);
                    //labelFormat.SubStrings.SetSubString("PC", id_textBox.Text);
                    //labelFormat.SubStrings.SetSubString("QTY", num_textBox.Text);

                }

                catch (Exception ex)
                {
                    MessageBox.Show("修改内容出错 " + ex.Message, "操作提示");
                }

                if (labelFormat != null)
                {
                    // 產生標籤預覽圖並顯示於 UP_PictureBox
                    labelFormat.ExportImageToFile(txt_Btw_Path.Text, ImageType.BMP, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(407, 407
                        ), OverwriteOptions.Overwrite);

                    System.Drawing.Image image = System.Drawing.Image.FromFile(txt_Btw_Path.Text);
                    Bitmap NmpImage = new Bitmap(image);
                    UP_PictureBox.Image = NmpImage;
                    image.Dispose();
                }
                else
                {
                    MessageBox.Show("生成圖片錯誤", "操作提示");
                }

                if (isPreView) return;

                if (_PrinterName != "")
                {
                    // 設定印表機並執行列印
                    labelFormat.PrintSetup.PrinterName = _PrinterName;
                    labelFormat.Print("BarPrint" + DateTime.Now, 3 * 1000);
                }
                else
                {
                    MessageBox.Show("請先選擇印表機", "操作提示");
                }
            }
        }

        /// <summary>
        /// 取得 FTP 伺服器連線資訊。
        /// 根據指定的 FTP 伺服器名稱，查詢資料庫 [i_Program_FtpServer_Table]，
        /// 並將伺服器 IP、帳號、密碼、工廠名稱等資訊存入對應欄位。
        /// </summary>
        /// <param name="Ftp_Server_name">FTP 伺服器名稱（如 "Print"、"PrintBarTender"）。</param>
        /// <remarks>
        /// 此方法會直接更新全域欄位 ftpServer、ftpuser、ftppassword、ftpdlfactory。
        /// 若查無資料則不做任何設定。
        /// </remarks>
        private void Getftp(string Ftp_Server_name)
        {
            string sqlCmd = @"SELECT [Ftp_Server_Ip],[Ftp_Server_OA_Ip],[Ftp_Username],[Ftp_Password],[Ftp_Server_name],[Ftp_Factory] " +
                            $"FROM [i_Program_FtpServer_Table] where [Ftp_Server_name] ='" + Ftp_Server_name + "' ";
            DataSet ds = db.reDs(sqlCmd);
            if (ds.Tables[0].Rows.Count != 0)
            {
                ftpServer = ds.Tables[0].Rows[0]["Ftp_Server_OA_Ip"].ToString().Trim();
                ftpuser = ds.Tables[0].Rows[0]["Ftp_Username"].ToString().Trim();
                ftppassword = ds.Tables[0].Rows[0]["Ftp_Password"].ToString().Trim();
                ftpdlfactory = ds.Tables[0].Rows[0]["Ftp_Factory"].ToString().Trim();
            }
        }

        /// <summary>
        /// 取得指定程式在 TE_Program_Table 的版本號。
        /// </summary>
        /// <param name="tool">程式名稱（例如 "PrintProgram64"）。</param>
        /// <returns>
        /// 回傳版本字串；若查無或發生錯誤則回傳空字串。
        /// </returns>
        /// <remarks>
        /// 此方法僅讀取資料庫回傳的第一筆版本欄位值，不會變更類別層級的欄位值。
        /// 呼叫端可將回傳值指派給需要的欄位（例如 version_new）。
        /// </remarks>
        /// <exception cref="Exception">在發生資料庫或其他例外時會被捕捉並以訊息視窗提示。</exception>
        private string selectVerSQL_new(string tool)
        {
            string version = string.Empty;

            try
            {
                if (string.IsNullOrWhiteSpace(tool))
                {
                    return version;
                }

                // 避免簡單的 SQL 注入（若 db 支援參數化請改用參數化查詢）
                string safeTool = tool.Replace("'", "''");
                string sqlCmd = $"SELECT * FROM TE_Program_Table WHERE [Program_Name] = '{safeTool}'";

                DataSet ds = db.reDs(sqlCmd);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    DataTable table = ds.Tables[0];

                    // 若有名為 "Version" 的欄位，直接讀取；否則嘗試大小寫不敏感比對欄位名稱
                    if (table.Columns.Contains("Version"))
                    {
                        version = table.Rows[0]["Version"]?.ToString() ?? string.Empty;
                    }
                    else
                    {
                        foreach (DataColumn col in table.Columns)
                        {
                            if (string.Equals(col.ColumnName, "Version", StringComparison.OrdinalIgnoreCase))
                            {
                                version = table.Rows[0][col]?.ToString() ?? string.Empty;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 顯示簡短錯誤訊息以利除錯
                MessageBox.Show($"讀取版本資訊失敗：{ex.Message}", "錯誤");
            }

            return version;
        }

        /// <summary>
        /// 判斷指定名稱的 Mutex 是否已存在（即程式是否已執行）。
        /// </summary>
        /// <param name="prgname">要檢查的 Mutex 名稱（通常為程式名稱）。</param>
        /// <returns>
        /// 若 Mutex 已存在（程式正在執行），則回傳 <c>true</c>；否則回傳 <c>false</c>。
        /// </returns>
        /// <remarks>
        /// 此方法會建立一個新的 Mutex，並透過 <paramref name="prgname"/> 判斷是否已有相同名稱的 Mutex 存在於系統。
        /// 若回傳 <c>true</c>，表示程式已執行中，不應重複啟動。
        /// </remarks>
        /// <example>
        /// if (IsMyMutex("PrintProgram64")) { MessageBox.Show("程式正在執行中!!"); }
        /// </example>
        private bool IsMyMutex(string prgname)
        {
            bool IsExist;
            m = new Mutex(true, prgname, out IsExist);
            GC.Collect();
            if (IsExist)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// 執行自動更新：啟動位於啟動目錄的 AutoUpdate.exe，啟動後關閉目前視窗。
        /// </summary>
        /// <remarks>
        /// 1. 會先確認 AutoUpdate.exe 是否存在於 Application.StartupPath。
        /// 2. 使用 ProcessStartInfo 啟動外部更新程式並以非同步方式關閉目前表單，避免阻塞 UI 執行緒。
        /// 請確保更新程式具有必要權限可被啟動，且不會依賴本程序仍在執行的資源。
        /// </remarks>
        /// <example>
        /// autoupdate();
        /// </example>
        private void autoupdate()
        {
            string updaterPath = Path.Combine(Application.StartupPath, "AutoUpdate.exe");

            if (!File.Exists(updaterPath))
            {
                MessageBox.Show($"找不到更新程式：{updaterPath}", "更新失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = updaterPath,
                    WorkingDirectory = Application.StartupPath,
                    UseShellExecute = true,
                    // 可視需求加入 Verb = "runas" 以使用提升權限
                };

                using (var p = new Process { StartInfo = startInfo })
                {
                    p.Start();
                }

                // 在 UI 執行緒安全地關閉目前表單
                if (this.InvokeRequired)
                {
                    this.BeginInvoke((MethodInvoker)(() => this.Close()));
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                // 顯示簡短錯誤訊息，避免泄漏過多內部資訊
                MessageBox.Show($"啟動更新程式失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
