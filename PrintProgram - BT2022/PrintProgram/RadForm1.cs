using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.UI;
using Seagull.BarTender.Print;
using Newtonsoft.Json;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace PrintProgram
{

    public partial class RadForm1 : Telerik.WinControls.UI.RadForm
    {
        public class SFIS
        {
            public string itemNo { get; set; }
            public string productid { get; set; }
            public string powercord { get; set; }
            public string biosVer { get; set; }
            public string BIOSCS { get; set; }
            public string ProductID_MF { get; set; }
            public string Ecn { get; set; }
        }
        public class AvSFIS
        {
            public string itemNo { get; set; }
        }
        public class EversunWoNo
        {
            public string wipNO { get; set; }
            public string EngSr { get; set; }
            public string startNO { get; set; }
            public string endNO { get; set; }
        }

        #region Fields
        // Common strings.
        private const string appName = "PrintProgram";
        private const string dataSourced = "Data Sourced";

        private Engine engine = null; // The BarTender Print Engine
        private LabelFormatDocument format = null; // The currently open Format
        private bool isClosing = false; // Set to true if we are closing. This helps discontinue thumbnail loading.

        // Label browser
        private string[] browsingFormats; // The list of filenames in the current folder
        Hashtable listItems; // A hash table containing ListViewItems and indexed by format name.
                             // It keeps track of what formats have had their image loaded.
        Queue<int> generationQueue; // A queue containing indexes into browsingFormats
                                    // to facilitate the generation of thumbnails

        // Label browser indexes.
        int topIndex; // The top visible index in the lstLabelBrowser
        int selectedIndex; // The selected index in the lstLabelBrowser
        #endregion
        //宣告變數
        #region
        string SN = string.Empty, Item = string.Empty, WO = string.Empty, PN = string.Empty, PC = string.Empty, Qty = string.Empty, VER = string.Empty, No_Number = string.Empty, ENGSR = string.Empty;
        string version_old = "", version_new = "";
        #endregion
        private string _bmp_path = Application.StartupPath + @"\exp.jpg";
        private string _btw_path = "";
        string _PrinterName = "";
        public long Sum_Of_SQLfile_size = 0;
        public bool TempUploadResult = false;
        public string Accpath, Boardpath, Boardpath2, Boardpath3, Boardpath4, Systempath, Systempath2 = string.Empty,PhilEPC,AdvantechLabel;
        //FTP使用
        public string ftpdlfactory;
        public string ftpServer;
        public string ftpuser;
        public string ftppassword;
        public string ftpfilepath, ftpPutFile, ftpGetFile, filename, ini_filepath;
        string download_Path = System.Windows.Forms.Application.StartupPath+"\\"+ "Btw_Folder";
        public string DLfilename, UPfilename;
        public long SQLfile_size;
        public string Cfgname = "Setup.ini";
        static Mutex m;
        SetupIniIP ini = new SetupIniIP();
        DataSet dataSet = new DataSet();
        DataTable advantechDt;
        public class SetupIniIP
        { //api ini
            public string path;
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            private static extern long WritePrivateProfileString(string section,
            string key, string val, string filePath);
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            private static extern int GetPrivateProfileString(string section,
            string key, string def, StringBuilder retVal,
            int size, string filePath);
            public void IniWriteValue(string Section, string Key, string Value, string inipath)
            {
                WritePrivateProfileString(Section, Key, Value, Application.StartupPath + "\\" + inipath);
            }
            public string IniReadValue(string Section, string Key, string inipath)
            {
                StringBuilder temp = new StringBuilder(255);
                int i = GetPrivateProfileString(Section, Key, "", temp, 255, Application.StartupPath + "\\" + inipath);
                return temp.ToString();
            }
        }

        public RadForm1()
        {
            InitializeComponent();
        }

        private void RadForm1_Load(object sender, EventArgs e)
        {
            Getftp("Print");
            if (IsMyMutex("PrintProgram64"))
            {
                MessageBox.Show("程式正在執行中!!");
                Dispose();//關閉
            }
            #region 系統更新
            version_old = ini.IniReadValue("Version", "version", Cfgname);
            version_new = selectVerSQL_new("PrintProgram64");
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
            txt_Pn.Enabled = false;
            txt_Qty.Enabled = false;
            txt_Pc.Enabled = false;
            txt_Bios_Ver.Enabled = false;
            txt_Sn.Enabled = false;
            txt_No_Number.Enabled = false;
            Accpath = download_Path +"\\"+ ini.IniReadValue("Option", "Acc_Carton_Label_Name", Cfgname);
            Boardpath = download_Path + "\\" + ini.IniReadValue("Option", "Board_Carton_Label_Name", Cfgname);
            Boardpath2 = download_Path + "\\" + ini.IniReadValue("Option", "Board_Carton_Label2_Name ", Cfgname);
            Boardpath3 = download_Path + "\\" + ini.IniReadValue("Option", "Board_Carton_Label3_Name ", Cfgname);
            Boardpath4 = download_Path + "\\" + ini.IniReadValue("Option", "Board_Carton_Label4_Name ", Cfgname);
            Systempath = download_Path + "\\" + ini.IniReadValue("Option", "System_Carton_Label_Name", Cfgname);
            Systempath2 = download_Path + "\\" + ini.IniReadValue("Option", "System_Carton_Label2_Name", Cfgname);
            PhilEPC = download_Path + "\\" + ini.IniReadValue("Option", "Phil_EPC_WHL_Small_Name", Cfgname);
            AdvantechLabel= download_Path + "\\" + ini.IniReadValue("Option", "AdvantechLabelName ", Cfgname);
            string Advantech = ini.IniReadValue("Option", "AdvantechMAC", Cfgname);
            string AdvantechPath= ini.IniReadValue("Option", "AdvantechMACPath", Cfgname);
            //version_old = ini.IniReadValue("Version", "version", filename);
            //version_new = selectVerSQL_new("E-SOP");

            radLabelElement1.Text = version_new;
            if (Advantech=="True")
            {
                advantechDt = new DataTable();
                this.Wo_Set_Page.Parent = null;
                this.Standard_Page.Parent = null;
                this.Oem_Page.Parent = null;
                this.Oem_Page_UP.Parent = null;
                this.radPageViewPage1.Parent = null;
                advantechDt=LoadExcelAsDataTable(AdvantechPath);
            }
            else
            {
                this.radPageViewPage2.Parent = null;
            }

        }
        private void Timer1_Tick(object sender, EventArgs e)
        {
            System_date_ID.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }
        private void Btn_Print_Click(object sender, EventArgs e)
        {
            try
            {
                List_Msg.Items.Clear();
                if (SN_GV.Rows.Count <1)
                {
                    if (!rdb_Acc.IsChecked)
                    {
                        List_Msg.Items.Add("尚未輸入序號.....");
                        txt_Sn.Focus();
                        return;
                    }
                }
                if(SN_GV.Rows.Count != int.Parse(txt_Qty.Text))
                {
                    if (!rdb_Acc.IsChecked)
                    {
                        List_Msg.Items.Add("數量與輸入序號數量不符");
                        txt_Sn.Focus();
                        return;
                    }
                }

                btn_Print.Enabled = false;
                txt_Sn.Text = string.Empty;
                txt_Sn.Enabled = false;
                List<string> SNList = new List<string>();
                List<int> NOList = new List<int>();

                string printlabe = string.Empty;
                bool printresult = false;
                #region 欄位轉換成值
                WO = txt_Wo.Text.Trim();
                PN = txt_Pn.Text.Trim();
                if (!string.IsNullOrEmpty(txt_Pc.Text.Trim()))
                {
                    PC = txt_Pc.Text.Trim().Replace(",", " ");
                }
                else
                {
                    PC = "NA";
                }
                if (string.IsNullOrEmpty(txt_Qty.Text))
                {
                    List_Msg.Items.Add("尚未輸入數量");
                    txt_Qty.Focus();
                    return;
                }
                else
                {
                    Qty = txt_Qty.Text.Trim();
                }
                if (!string.IsNullOrEmpty(txt_Bios_Ver.Text.Trim()))
                {
                    VER = txt_Bios_Ver.Text.Trim();
                }
                else
                {
                    VER = "NA";
                    txt_Bios_Ver.Text = "NA";
                }
                ENGSR = txt_engSR.Text.Trim();
                if (!string.IsNullOrEmpty(txt_No_Number.Text.Trim()))
                {
                    No_Number = "(" + txt_No_Number.Text.Trim().ToUpper() + ")";
                }
                else
                {
                    No_Number = string.Empty;
                }
                #endregion

                List_Msg.Items.Add("列印中.....");
                if (rdb_Acc.IsChecked)
                {
                    printlabe = Accpath;
                    if (AccPrint.PrintLabel(printlabe, WO, PN, Qty, No_Number) == true)
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
                    //20210818改system列印超過6組印第二頁
                    #region MyRegion
                    if (int.Parse(txt_Qty.Text) > 6)
                    {
                        for (int i = 0; i < 6; i++)
                        {
                            SNList.Add(SN_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }
                        if (SystemPrint.PrintLabel(printlabe, WO, PN, PC, VER, Qty, No_Number, ENGSR, SNList) == true)
                        {
                            SNList.Clear();
                            for (int i = 6; i < SN_GV.Rows.Count; i++)
                            {
                                SNList.Add(SN_GV.Rows[i].Cells["SN"].Value.ToString().Trim());

                            }
                            if (SystemPrint2.PrintLabel(Systempath2, SNList) == true)
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
                            printresult = false;
                        }
                    }
                    #endregion
                    else
                    {
                        for (int i = 0; i < SN_GV.Rows.Count; i++)
                        {
                            SNList.Add(SN_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }
                        if (SystemPrint.PrintLabel(printlabe, WO, PN, PC, VER, Qty, No_Number, ENGSR, SNList) == true)
                        {
                            printresult = true;
                        }
                        else
                        {
                            printresult = false;
                        }
                    }
                }
                else
                {
                    printlabe = Boardpath;
                    int number = int.Parse(txt_Qty.Text);     // 輸入的數字
                    int groupSize = 36; // 每組的大小

                    if (number <= 26)
                    {
                        //小於 = 26情況 ,直接印
                        for (int i = 0; i < SN_GV.Rows.Count; i++)
                        {
                            SNList.Add(SN_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }
                        if (BoardPrint.PrintLabel(printlabe, WO, PN, VER, Qty, No_Number, ENGSR, SNList) == true)
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
                        //大於26情況, 先印1~26
                        for (int i = 0; i < 26; i++)
                        {
                            SNList.Add(SN_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }
                        if (BoardPrint.PrintLabel(printlabe, WO, PN, VER, Qty, No_Number, ENGSR, SNList) == true)
                        {
                            //groupSize=36 ,每36為一組
                            SNList.Clear();
                            int startNumber = 27; // 第一組的起始數字
                            int endNumber = 26 + groupSize; // 第一組的結束數字
                            int group = 2; // 組數

                            while (number > endNumber)
                            {
                                // 處理該組的數字
                                NOList.Clear();
                                SNList.Clear();

                                for (int i = startNumber; i <= endNumber; i++)
                                {
                                    NOList.Add(i);
                                    SNList.Add(SN_GV.Rows[i - 1].Cells["SN"].Value.ToString().Trim());
                                }

                                if (BoardPrint4.PrintLabel(Boardpath4, SNList, NOList) == true)
                                {
                                    printresult = true;
                                }
                                else
                                {
                                    printresult = false;
                                    break; // 停止處理後面的組數
                                }
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

                                for (int i = startNumber; i <= endNumber; i++)
                                {
                                    NOList.Add(i);
                                    if (SN_GV.Rows.Count >= i && SN_GV.Rows[i - 1].Cells["SN"].Value != null)
                                    {
                                        SNList.Add(SN_GV.Rows[i - 1].Cells["SN"].Value.ToString().Trim());
                                    }
                                    else
                                    {
                                        SNList.Add(string.Empty);
                                    }
                                }

                                if (BoardPrint4.PrintLabel(Boardpath4, SNList, NOList) == true)
                                {
                                    printresult = true;
                                }
                                else
                                {
                                    printresult = false;
                                }
                            }
                        }
                    }
                }

                if (printresult == true)
                {
                    btn_Print.Enabled = true;
                    if (!rdb_Acc.IsChecked)//更新Print_Carton_Number_Table目前Carton編號
                    {

                        SN_GV.Rows.Clear();
                        List_Msg.Items.Clear();
                        List_Msg.Items.Add("列印完成.....");
                        txt_Sn.Enabled = true;

                    }
                    else
                    {
                        //Acc/bcc列印不做Carton_Number更新
                        SN_GV.Rows.Clear();
                        List_Msg.Items.Clear();
                        List_Msg.Items.Add("列印完成.....");
                    }
                }
                else
                {
                    btn_Print.Enabled = true;
                    List_Msg.Items.Add("列印失敗.....");
                }
            }
            catch (Exception ex)
            {
                btn_Print.Enabled = true;
            }
        }
        private void Btn_Search_File_Click(object sender, EventArgs e)
        {

            UP_PictureBox.Image = null;
            List2_Msg.Items.Clear();
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//該值確定是否可以選擇多個檔案
            dialog.Title = "請選擇列印檔案";
            dialog.Filter = "列印檔案(*.btw)|*.btw";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txt_Btw_Path.Text = dialog.FileName;
                _btw_path = dialog.FileName;
                filename = dialog.SafeFileName;
                //PrintBar(true);
                UP_PictureBox.Image = null;
                using (Engine btEngine = new Engine(true))
                {
                    LabelFormatDocument labelFormat = btEngine.Documents.Open(_btw_path);

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

        }
        private void Btn_Oem_Search_Click(object sender, EventArgs e)
        {
            try
            {
                List_Oem_Msg.Items.Clear();
                //DL_PictureBox.Image = null;
                txt_Oem_Pn.Text = string.Empty;

                List_Oem_Msg.Items.Add("列印套版下載中......");
                string sqlCmd = "";

                if (!string.IsNullOrEmpty(txt_Oem_Wo.Text.Trim()))
                {
                    sqlCmd = "SELECT * FROM [Print_Carton_Table] where Wo = '" + txt_Oem_Wo.Text.Trim() + "' order by time desc ";
                    DataSet ds = db.reDs(sqlCmd);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        txt_Oem_Pn.Text = ds.Tables[0].Rows[0]["Pn"].ToString();
                        DLfilename = ds.Tables[0].Rows[0]["Filename"].ToString();
                        if (ds.Tables[0].Rows[0]["Qty_Set"].ToString() == "Yes")
                        {
                            txt_Oem_Qty.Enabled = true;
                        }
                        else
                        {
                            txt_Oem_Qty.Enabled = false;

                        }
                        if (ds.Tables[0].Rows[0]["Weight_Set"].ToString() == "Yes")
                        {
                            txt_Oem_Weight.Enabled = true;
                        }
                        else
                        {
                            txt_Oem_Weight.Enabled = false;
                        }
                        if (ds.Tables[0].Rows[0]["Sn_Set"].ToString() == "Yes")
                        {
                            txt_Oem_Sn.Enabled = true;
                        }
                        else
                        {
                            txt_Oem_Sn.Enabled = false;
                        }
                        if (ds.Tables[0].Rows[0]["Mac_Set"].ToString() == "Yes")
                        {
                            txt_Oem_Mac.Enabled = true;
                            txt_Oem_Mac.Focus();
                        }
                        else
                        {
                            txt_Oem_Mac.Enabled = false;
                        }
                        if (ds.Tables[0].Rows[0]["Bios_Set"].ToString() == "Yes")
                        {
                            txt_Oem_Bios.Enabled = true;
                        }
                        else
                        {
                            txt_Oem_Bios.Enabled = false;
                        }
                        this.FTP_Dl_Btw_thread.WorkerSupportsCancellation = true; //允許中斷
                        this.FTP_Dl_Btw_thread.RunWorkerAsync(); //呼叫背景程式
                    }
                    else
                    {
                        List_Oem_Msg.Items.Add("查無" + txt_Oem_Wo.Text.Trim() + "列印套版");
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());

            }
        }
        private void Btn_Oem_Print_Click(object sender, EventArgs e)
        {
            try
            {
                List<string> SNList = new List<string>();
                #region OEM-STD_Bios
                if (txt_Oem_Bios.Enabled == true)
                {
                    string printlabe = download_Path + "\\" + DLfilename;
                    if (int.Parse(txt_Oem_Qty.Text) > 26)
                    {
                        for (int i = 0; i < 26; i++)
                        {
                            SNList.Add(Sn_Oem_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }
                        if (OemPrint2.PrintLabel(printlabe, txt_Oem_Wo.Text, txt_Oem_Qty.Text.Trim(), txt_Oem_Bios.Text.Trim(), SNList) == true)
                        {
                            SNList.Clear();
                            for (int i = 26; i < Sn_Oem_GV.Rows.Count; i++)
                            {
                                SNList.Add(Sn_Oem_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                            }
                            BoardPrint2.PrintLabel(Boardpath2, SNList);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < Sn_Oem_GV.Rows.Count; i++)
                        {
                            SNList.Add(Sn_Oem_GV.Rows[i].Cells["SN"].Value.ToString().Trim());
                        }
                        OemPrint2.PrintLabel(printlabe, txt_Oem_Wo.Text.Trim(), txt_Oem_Qty.Text, txt_Oem_Bios.Text.Trim(), SNList);
                        //if (OemPrint2.PrintLabel(printlabe, WO, SNList) == true)
                    }
                }
                #endregion

                #region 標準列印以外
                else
                {
                    List_Oem_Msg.Items.Clear();
                    if (!string.IsNullOrEmpty(txt_Oem_Qty.Text) && Sn_Oem_GV.Rows.Count < 1)
                    {
                        List_Oem_Msg.Items.Add("尚未輸入序號.....");
                        txt_Oem_Sn.Focus();
                        return;
                    }
                    if (!string.IsNullOrEmpty(txt_Oem_Qty.Text) && Sn_Oem_GV.Rows.Count != int.Parse(txt_Oem_Qty.Text))
                    {
                        List_Oem_Msg.Items.Add("數量與輸入序號數量不符");
                        txt_Oem_Sn.Focus();
                        return;
                    }
                    int printCount = 1;
                    btn_Oem_Print.Enabled = false;
                    txt_Oem_Sn.Text = string.Empty;
                    txt_Oem_Sn.Enabled = false;

                    List_Oem_Msg.Items.Add("列印中.....");
                    Engine engine = null;
                    LabelFormatDocument btFormat = null;
                    engine = new Engine();
                    engine.Start();
                    btFormat = engine.Documents.Open(download_Path + "\\" + DLfilename);

                    if (!string.IsNullOrEmpty(txt_Oem_Qty.Text.Trim()))
                    {
                        btFormat.SubStrings["QTY"].Value = txt_Oem_Qty.Text.Trim();
                    }
                    if (!string.IsNullOrEmpty(txt_Oem_Weight.Text.Trim()))
                    {
                        btFormat.SubStrings["Weight"].Value = txt_Oem_Weight.Text.Trim();
                    }

                    if (!string.IsNullOrEmpty(txt_Oem_Mac.Text.Trim()))
                    {
                        printCount = 2;
                        btFormat.SubStrings["mac_n"].Value = txt_Oem_Mac.Text.Trim();
                        btFormat.SubStrings["sn_n"].Value = Sn_Oem_GV.Rows[0].Cells["SN"].Value.ToString().Trim();
                        btFormat.SubStrings["SN1"].Value = Sn_Oem_GV.Rows[0].Cells["SN"].Value.ToString().Trim(); //標籤檔中所設定的欄位名稱 。
                        btFormat.SubStrings["SN2"].Value = txt_Oem_Mac.Text.Trim();  //標籤檔中所設定的欄位名稱 。
                    }
                    //if (!string.IsNullOrEmpty(txt_Oem_Wo.Text.Trim()))
                    //{
                    //    btFormat.SubStrings["WO"].Value = txt_Oem_Wo.Text.Trim();
                    //    btFormat.SubStrings["ENGSR"].Value = "(" + txt_Oem_Pn.Text.Trim() + ")";
                    //}
                    /*
                     * btFormat.SubStrings["WO"].Value = WO;
                     * btFormat.SubStrings["ENGSR"].Value = "(" + ENGSR + ")";
                     */
                    if (Sn_Oem_GV.Rows.Count > 0 && txt_Oem_Qty.Enabled == true)
                    {
                        for (int i = 0; i < Sn_Oem_GV.Rows.Count; i++)
                        {
                            string SN_Name = "SN" + (i + 1).ToString();
                            btFormat.SubStrings[SN_Name].Value = Sn_Oem_GV.Rows[i].Cells["SN"].Value.ToString().Trim(); //標籤檔中所設定的欄位名稱 。
                        }
                    }

                    btFormat.PrintSetup.IdenticalCopiesOfLabel = printCount;//int.Parse("1"); //列印標籤數
                    btFormat.Print();
                    engine.Stop();
                }
                #endregion

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
                btn_Oem_Print.Enabled = true;
            }
        }

        private void Rcb_Re_Print_Click(object sender, EventArgs e)
        {
            if(rcb_Re_Print.Checked ==true)
            {
                txt_reprint_Sn.Enabled = false;
            }
            else
            {
                txt_reprint_Sn.Enabled = true;

            }
        }

        private void Btn_Wo_Set_Click(object sender, EventArgs e)
        {
            List_Wo_Set_Msg.Items.Clear();
            string msg = string.Empty, Print_Type = string.Empty;
            //To do ini upload ftp
            if (txt_Wo_Set.Text == "")
            {
                msg = msg + "尚未輸入工單" + "\r\n";
                MessageBox.Show(msg);
                txt_Wo_Set.Focus();
                return;
            }

            if(rdb_Standard.IsChecked)
            {
                Print_Type = "Standard";
            }
            else if(rdb_Oem_Of.IsChecked)
            {
                Print_Type = "Oem_Off_Line";
            }
            else if(rdb_Oem_On.IsChecked)
            {
                Print_Type = "Oem_On_Line";
            }
            else
            {
                List_Wo_Set_Msg.Items.Add("請設定列印模式");
                txt_Wo_Set.Focus();
                return;

            }

            string InsSql = " INSERT INTO [Print_Wo_Setting_Table] (Record_Time,Work_Order,Print_Type) VALUES("
                                                       + "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                                                       + "'" + txt_Wo_Set.Text.Trim() + "',"
                                                       + "'" + Print_Type + "')";
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

        private void Btn_Wo_Serch_Click(object sender, EventArgs e)
        {
            List_Wo_Search_Msg.Items.Clear();
            string msg = string.Empty, Print_Type = string.Empty, Work_Order = string.Empty;
            //To do ini upload ftp
            if (txt_Wo_Serach.Text == "")
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

                if(Print_Type == "Standard")
                {
                    List_Wo_Search_Msg.Items.Add("此工單為標準品列印");
                    txt_Wo.Text = txt_Wo_Serach.Text;
                    Print_Page.SelectedPage = Standard_Page;
                    txt_Wo_Serach.Text = string.Empty;
                }
                else if(Print_Type == "Oem_On_Line")
                {
                    List_Wo_Search_Msg.Items.Add("此工單為OEM/ODM ON Line列印");
                    txt_Oem_Wo.Text = txt_Wo_Serach.Text;
                    Print_Page.SelectedPage = Oem_Page;
                    txt_Wo_Serach.Text=string.Empty;
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

        private void Btn_Oem_Clear_Click(object sender, EventArgs e)
        {
            txt_Oem_Wo.Text = string.Empty;
            txt_Oem_Pn.Text = string.Empty;
            txt_Oem_Qty.Text = string.Empty;
            txt_Oem_Weight.Text = string.Empty;
            txt_reprint_Oem_Sn.Text = string.Empty;
            rcb_Oem_Re_Print.Checked = false;
            txt_Oem_Qty.Enabled = false;
            txt_Oem_Weight.Enabled = false;
            txt_reprint_Oem_Sn.Enabled = false;
            Sn_Oem_GV.Rows.Clear();

            List_Oem_Msg.Items.Clear();
        }

        private void Btn_Up_Clear_Click(object sender, EventArgs e)
        {
            txt_Up_WO.Text = string.Empty;
            txt_Up_Pn.Text = string.Empty;
            txt_Btw_Path.Text = string.Empty;
            UP_PictureBox.Image = null;
            List2_Msg.Items.Clear();
        }

        private void Btn_Out_Print_Click(object sender, EventArgs e)
        {
            try
            {
                List_Msg.Items.Clear();
                btn_Print.Enabled = false;
                btn_Out_Print.Enabled = false;
                txt_Sn.Text = string.Empty;
                txt_Sn.Enabled = false;


                string printlabe = string.Empty;
                bool printresult = false;
                #region 欄位轉換成值
                WO = txt_Wo.Text.Trim();
                PN = txt_Pn.Text.Trim();
                if (!string.IsNullOrEmpty(txt_Pc.Text.Trim()))
                {
                    PC = txt_Pc.Text.Trim().Replace(",", " ");
                }
                else
                {
                    PC = "NA";
                }
                if (string.IsNullOrEmpty(txt_Qty.Text))
                {
                    List_Msg.Items.Add("尚未輸入數量");
                    txt_Qty.Focus();
                    return;
                }
                else
                {
                    Qty = txt_Qty.Text.Trim();
                }

                VER = txt_Bios_Ver.Text.Trim();
                if (!string.IsNullOrEmpty(txt_No_Number.Text.Trim()))
                {
                    No_Number = "(" + txt_No_Number.Text.Trim().ToUpper() + ")";
                }
                else
                {
                    No_Number = string.Empty;
                }
                #endregion

                List_Msg.Items.Add("列印中.....");
                if (rdb_Acc.IsChecked)
                {
                    printlabe = Accpath;
                    if (AccPrint.PrintLabel(printlabe, WO, PN, Qty, No_Number) == true)
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

                    if (SystemPrint.Out_PrintLabel(printlabe, WO, PN, PC, VER, Qty, No_Number, ENGSR) == true)
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
                btn_Print.Enabled = true;
            }
        }
        private void Btn_Clear_Click(object sender, EventArgs e)
        {
            txt_Wo.Text = string.Empty;
            txt_Pn.Text = string.Empty;
            txt_Pc.Text = string.Empty;
            txt_Qty.Text = string.Empty;
            txt_Bios_Ver.Text = string.Empty;
            txt_No_Number.Text = string.Empty;
            txt_Sn.Text= string.Empty;
            txt_reprint_Sn.Text = string.Empty;
            txt_reprint_Sn.Enabled = false;
            rcb_Re_Print.Checked = false;
            SN_GV.Rows.Clear();
            List_Msg.Items.Clear();

        }
        public static bool IsInRange(object input, object min, object max)
        {
            if (input is long inputNumber && min is long minNumber && max is long maxNumber)
            {
                return inputNumber >= minNumber && inputNumber <= maxNumber;
            }
            else if (input is string inputString && min is string minString && max is string maxString)
            {
                return string.Compare(inputString, minString) >= 0 && string.Compare(inputString, maxString) <= 0;
            }
            else
            {
                throw new ArgumentException("输入的数据类型不支持。");
            }
        }
        private void Txt_Oem_Sn_KeyPress(object sender, KeyPressEventArgs e)
        {
            bool AddSN = true;
            string Bind_SN = string.Empty;
            if (string.IsNullOrEmpty(txt_Oem_Qty.Text) && txt_Oem_Qty.Enabled==true)
            {
                List_Oem_Msg.Items.Add("請先輸入數量 ");
                txt_Oem_Qty.Focus();
                return;
            }

            if (Convert.ToInt32(e.KeyChar) == 13)
            {
                #region IGT-OEM
                if (txt_Oem_Mac.Enabled)
                {
                    //string sqlCmd = "SELECT * FROM Avalue_MOID_ShipmentSN_Table where MOID = '" + txt_Oem_Wo.Text.Trim() + "' ";
                    string sqlCmd = "select top(1)Eversun_WoNo from Print_CustomWONO_Table where Avalue_WoNo = '" + txt_Oem_Wo.Text.Trim() + "' order by sno desc";
                    DataSet ds = db.reDs(sqlCmd);
                    string eversun_wono = ds.Tables[0].Rows[0][0].ToString();
                    string WipInfo = Auto_Route.WipbarcodeOther(eversun_wono.Trim());

                    EversunWoNo descJsonStu = JsonConvert.DeserializeObject<EversunWoNo>(WipInfo.ToString());//反序列化

                    string inputString = txt_Oem_Sn.Text.Trim();
                    bool isInRange = false;
                    if (long.TryParse(inputString, out long inputNumber))
                    {
                        // 判断 long 类型输入是否在范围内
                        isInRange = IsInRange(inputNumber, descJsonStu.startNO, descJsonStu.endNO);
                    }
                    else
                    {
                        // 判断 string 类型输入是否在范围内
                        isInRange = IsInRange(inputString, descJsonStu.startNO, descJsonStu.endNO);
                    }

                    // 输出结果
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
                if (!string.IsNullOrEmpty(txt_Oem_Sn.Text.Trim().ToUpper()))
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
                        string InsSql = " INSERT INTO [Print_SN_Bind_Carton_Table] (Record_Time,Work_Order,Quantity,SN,Bind_SN,Weight,MAC) VALUES("
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
                            //輸入序號數量與數量欄位一樣(列印)
                            if (Item == txt_Oem_Qty.Text.Trim() || txt_Oem_Mac.Enabled==true)
                            {
                                Btn_Oem_Print_Click(sender, e);
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

        private void Sn_Oem_GV_UserDeletingRow(object sender, GridViewRowCancelEventArgs e)
        {
            string DelSN = Sn_Oem_GV.CurrentRow.Cells["SN"].Value.ToString();
            string DelTime = Convert.ToDateTime(Sn_Oem_GV.CurrentRow.Cells["Time"].Value).ToString("yyyy-MM-dd HH:mm:ss");
            string DelSql = " Delete  [Print_SN_Bind_Carton_Table] where Sn = '" + DelSN + "' and Record_Time = '" + DelTime + "'";
            if (db.Exsql(DelSql) == true)
            {


            }
            else
            {
                List_Msg.Items.Add("資料庫刪除錯誤");
            }
        }

        private void Rcb_Oem_Re_Print_Click(object sender, EventArgs e)
        {
            if (rcb_Oem_Re_Print.Checked == true)
            {
                txt_reprint_Oem_Sn.Enabled = false;
            }
            else
            {
                txt_reprint_Oem_Sn.Enabled = true;

            }
        }

        /// <summary>
        /// 選擇PHIL_EPC_WHL的檔案路徑
        /// </summary>
        /// <param name="modeltype">選擇01~08R</param>
        /// <param name="modelLab">選擇大張或小張Lab</param>
        /// <returns></returns>
        public string Phl_EPCWHL_Path(string modeltype,string modelLab)
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
            if (modelLab=="Small")
            {
                path = PhilEPC;
            }
            return path;
        }
        private void btn_P_EPC_Click(object sender, EventArgs e)
        {
            try
            {
                //List<string> SNList = new List<string>();
                string partNumber = "";
                string ecn = "";
                string biosVer = "";
                string bioscs = "";
                string shiftStr = "";
                if (rdb_Big.IsChecked)
                {
                    shiftStr = radTxt_Wono.Text.Trim().Substring(0, 10);
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("{");
                    sb.AppendLine("\"Key\":\"" + "@Avalue.ZMO.SOP" + "\",");
                    sb.AppendLine("\"moid\":\"" + shiftStr.Trim() + "\",");
                    sb.AppendLine("}");
                    var test = SFISToJson.reDt3(sb);
                    SFIS descJsonStu = JsonConvert.DeserializeObject<SFIS>(test.ToString());//反序列化

                    //

                    string[] PN = descJsonStu.productid.Split('-');
                    partNumber = PN[PN.Length - 1];
                    ecn = descJsonStu.Ecn.Trim().ToUpper();
                    biosVer = descJsonStu.biosVer.Trim();
                    bioscs = descJsonStu.BIOSCS.Trim().ToUpper();
                }

                string choiceRdb = "";
                if (rdb_Big.IsChecked)
                {
                    choiceRdb = "Big";
                }
                if (rdb_Small.IsChecked)
                {
                    choiceRdb = "Small";

                }

                string printlabe = Phl_EPCWHL_Path(partNumber, choiceRdb);

                #region Phil_EPC_WHL

                if (PhilEPC_Print.PrintLabel(partNumber, choiceRdb, printlabe, ecn, biosVer, bioscs, radTxt_SN.Text, radTxt_Mac.Text))
                {
                    list_Phil_Msg.Items.Clear();
                    list_Phil_Msg.Items.Add("列印完成.....");
                }

                #endregion
            }
            catch (Exception ex)
            {
                list_Phil_Msg.Items.Add(ex.Message);
            }
        }
        private void radTxt_Wono_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (true)
            {

            }
            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = true;
            }
            if (e.KeyChar == 13)
            {
                if (!rdb_Small.IsChecked && !rdb_Big.IsChecked)
                {
                    MessageBox.Show("請先選擇套版Label");
                }
                else
                {
                    radTxt_SN.Focus();
                }
            }
        }
        private void radTxt_SN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && radTxt_SN.Text.Length != 11)
            {
                MessageBox.Show("序號長度錯誤, 請確認後重新輸入");
                radTxt_SN.Text = "";
                return;
            }

            if (e.KeyChar== (char)Keys.Back )
            {
                e.Handled = true;
            }

            if (e.KeyChar==13 && radTxt_SN.Text.Length==11)
            {
                if (!rdb_Small.IsChecked && !rdb_Big.IsChecked)
                {
                    MessageBox.Show("請先選擇套版Label");
                }
                else
                {
                    if (string.IsNullOrEmpty(radTxt_Wono.Text) && rdb_Big.IsChecked)//Wono不得為空
                    {
                        radTxt_Wono.Focus();
                        MessageBox.Show("工單未輸入!");
                    }
                    else
                    {
                        if (rdb_Small.IsChecked)
                        {
                            btn_P_EPC_Click(sender, e);
                        }
                        else
                        {
                            radTxt_Mac.Focus();
                        }
                    }

                }
            }
        }

        private void btn_C_EPC_Click(object sender, EventArgs e)
        {
            radTxt_Wono.Text = "";
            radTxt_SN.Text = "";
            radTxt_Mac.Text = "";
        }

        private void rdb_Small_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void rdb_Big_CheckStateChanged(object sender, EventArgs e)
        {
            if (rdb_Big.IsChecked)
            {
                radTxt_SN.Enabled = true;
                radTxt_Wono.Enabled = true;
                radTxt_Wono.Focus();
            }
            if (rdb_Small.IsChecked)
            {
                radTxt_SN.Enabled = false;
                radTxt_Wono.Enabled = false;
                radTxt_Mac.Focus();
            }
        }

        private void radTxt_SN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                e.Handled = true;
            }
        }

        private void radTxt_Wono_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                e.Handled = true;
            }
        }

        private void radTxt_Mac_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                e.Handled = true;
            }
        }

        private void btn_Advantech_Click(object sender, EventArgs e)
        {
            try
            {

                #region Phil_EPC_WHL



                int printCount = 1;
                radListControl1.Items.Add("列印中.....");
                Engine engine = null;
                LabelFormatDocument btFormat = null;
                engine = new Engine();
                engine.Start();
                btFormat = engine.Documents.Open(download_Path + "\\" + DLfilename);

                if (!string.IsNullOrEmpty(txt_Oem_Mac.Text.Trim()))
                {
                    printCount = 2;
                    btFormat.SubStrings["mac_n"].Value = txt_Oem_Mac.Text.Trim();
                    btFormat.SubStrings["sn_n"].Value = Sn_Oem_GV.Rows[0].Cells["SN"].Value.ToString().Trim();
                    btFormat.SubStrings["SN1"].Value = Sn_Oem_GV.Rows[0].Cells["SN"].Value.ToString().Trim(); //標籤檔中所設定的欄位名稱 。
                    btFormat.SubStrings["SN2"].Value = txt_Oem_Mac.Text.Trim();  //標籤檔中所設定的欄位名稱 。
                }




                #endregion
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void txt_AdvantechBarcode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = true;
            }
            bool printflag = false;
            if (e.KeyChar == 13)
            {
                try
                {
                   // var s = advantechDt.AsEnumerable()
                   //.Where(o => o.Field<string>("BARCODE_NO") == txt_AdvantechBarcode.Text)
                   //.Select(o => o.Field<string>("PART_NO"))
                   //.FirstOrDefault();  // 取得第一個符合條件的結果

                    var s = from a in advantechDt.AsEnumerable()
                        .Where(o => o.Field<string>("BARCODE_NO") == txt_AdvantechBarcode.Text)
                            select a.Field<string>("PART_NO");
                    if (s != null)
                    {
                        radListControl1.Items.Add("列印中.....");
                        string printlabe = AdvantechLabel;

                        foreach (var item in s)
                        {
                            printflag = AdvantechPrint.PrintLabel(item.ToUpper(), printlabe);
                        }
                        if (printflag)
                        {
                            radListControl1.Items.Clear();
                            radListControl1.Items.Add("列印完成.....");
                            txt_AdvantechBarcode.Text = "";
                        }
                        else
                        {
                            radListControl1.Items.Clear();
                            radListControl1.Items.Add("列印失敗.....");
                        }
                    }
                    else
                    {
                        MessageBox.Show("This record cannot be found in the document database.");
                    }
                }
                catch (Exception er)
                {
                    radListControl1.Items.Add(er.Message);
                }

            }
        }
        public DataTable LoadExcelAsDataTable(String xlsFilename)
        {
            FileInfo fi = new FileInfo(xlsFilename);
            using (FileStream fstream = new FileStream(fi.FullName, FileMode.Open))
            {
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
                int cellCount = headerRow.LastCellNum; // 取欄位數
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    //table.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue, typeof(double)));
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

                        //依先前取得的欄位數逐一設定欄位內容
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell != null)
                            {
                                //如要針對不同型別做個別處理，可善用.CellType判斷型別
                                //再用.StringCellValue, .DateCellValue, .NumericCellValue...取值

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
                                    default: // String
                                             // 此處只簡單轉成字串
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

                //table.PrimaryKey = new DataColumn[] { table.Columns["日期"] };
                // success
                return table;
            }
        }

        private void Print_Page_SelectedPageChanged(object sender, EventArgs e)
        {

        }

        private void radTxt_Mac_KeyPress(object sender, KeyPressEventArgs e)
        {
            //00045F + 6碼  安勤  ;拆分大小張
            bool flag = false;

            if (e.KeyChar == (char)Keys.Back)
            {
                e.Handled = true;
            }
            if (e.KeyChar == 13)
            {
                if (rdb_Small.IsChecked && radTxt_Mac.Text.Length == 6)
                {
                    //radTxt_Mac.Text = "";
                    flag = true;
                }
                if (rdb_Big.IsChecked && radTxt_Mac.Text.Length == 12)
                {
                    //radTxt_Mac.Text = "";
                    flag = true;
                }
                //如果flag =false 代表未選擇套版或MAC長度錯誤
                if (flag)
                {
                    if (string.IsNullOrEmpty(radTxt_SN.Text) && rdb_Big.IsChecked)//SN不得為空
                    {
                        radTxt_SN.Focus();
                        MessageBox.Show("序號未輸入!");
                    }
                    else if (string.IsNullOrEmpty(radTxt_Wono.Text) && rdb_Big.IsChecked)//Wono不得為空
                    {
                        radTxt_Wono.Focus();
                        MessageBox.Show("工單未輸入!");
                    }
                    else if (string.IsNullOrEmpty(radTxt_Mac.Text))//MAC不得為空
                    {
                        radTxt_Mac.Focus();
                        MessageBox.Show("MAC未輸入!");
                    }
                    else
                    {
                        btn_P_EPC_Click(sender, e);
                    }
                }
                else
                {
                    MessageBox.Show("MAC長度錯誤, 請確認後重新輸入");
                }
            }
            else if (!rdb_Small.IsChecked && !rdb_Big.IsChecked)
            {
                MessageBox.Show("請先選擇套版Label");
                return;
            }

        }

        private void Txt_Oem_Qty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;

                txt_Oem_Sn.Enabled = true;


            }
        }

        private void txt_Oem_Mac_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar==13)
            {
                txt_Oem_Sn.Focus();
            }
        }

        private void Txt_reprint_Oem_Sn_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (Convert.ToInt32(e.KeyChar) == 13)
                {

                    #region 清空欄位

                    txt_Oem_Sn.Enabled = false;
                    txt_Oem_Mac.Text = string.Empty;
                    txt_Oem_Bios.Text = string.Empty;
                    txt_Oem_Qty.Text = string.Empty;
                    txt_Oem_Weight.Text = string.Empty;
                    List_Oem_Msg.Items.Clear();
                    Sn_Oem_GV.Rows.Clear();
                    #endregion

                    if (!string.IsNullOrEmpty(txt_reprint_Oem_Sn.Text.Trim()))
                    {
                        string sqlCmd = "SELECT * FROM [Print_SN_Bind_Carton_Table] where Bind_SN = '" + txt_reprint_Oem_Sn.Text.Trim() + "' ";
                        DataSet ds = db.reDs(sqlCmd);
                        if (ds.Tables[0].Rows.Count > 0)
                        {


                            txt_Oem_Qty.Text = ds.Tables[0].Rows[0]["Quantity"].ToString();
                            txt_Oem_Weight.Text = ds.Tables[0].Rows[0]["Weight"].ToString();


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
                            if (Item == txt_Oem_Qty.Text.Trim())
                            {
                                Btn_Oem_Print_Click(sender, e);
                            }

                        }
                        else
                        {
                            List_Oem_Msg.Items.Add("ERROR!!查無第一筆SN資料");
                        }
                    }
                    else
                    {
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



        private void Btn_Up_File_Click(object sender, EventArgs e)
        {
            List2_Msg.Items.Clear();
            string msg = string.Empty;
            //To do ini upload ftp
            if (txt_Btw_Path.Text == "")
            {
                msg = msg + "尚未選擇套版檔案" + "\r\n";
                MessageBox.Show(msg);
                txt_Btw_Path.Focus();
                return;
            }
            if (txt_Up_WO.Text == "")
            {
                msg = msg + "工單空白" + "\r\n";
                MessageBox.Show(msg);
                txt_Up_WO.Focus();
                return;
            }
            if (txt_Up_Pn.Text == "")
            {
                msg = msg + "標籤料號空白" + "\r\n";
                MessageBox.Show(msg);
                txt_Up_Pn.Focus();
                return;
            }

            if (msg == string.Empty)
            {

                this.FTP_Up_Btw_thread.WorkerSupportsCancellation = true; //允許中斷
                this.FTP_Up_Btw_thread.RunWorkerAsync(); //呼叫背景程式
            }
        }
        private void SN_GV_UserDeletingRow(object sender, GridViewRowCancelEventArgs e)
        {

            string DelSN = SN_GV.CurrentRow.Cells["SN"].Value.ToString();
            string DelTime = Convert.ToDateTime(SN_GV.CurrentRow.Cells["Time"].Value).ToString("yyyy-MM-dd HH:mm:ss");
            string DelSql = " Delete  [Print_SN_Bind_Carton_Table] where Sn = '" + DelSN + "' and Record_Time = '" + DelTime + "'";
            if (db.Exsql(DelSql) == true)
            {


            }
            else
            {
                List_Msg.Items.Add("資料庫刪除錯誤");
            }
        }
        private void Txt_reprint_Sn_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (Convert.ToInt32(e.KeyChar) == 13)
                {
                    if (rdb_Board.IsChecked || rdb_System.IsChecked || rdb_Acc.IsChecked)
                    {

                    }
                    else
                    {
                        List_Msg.Items.Add("ERROR!!請先選擇列印製程");
                        radGroupBox7.Focus();
                        return;
                    }
                    #region 清空欄位

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

                    if (!string.IsNullOrEmpty(txt_reprint_Sn.Text.Trim()))
                    {
                        string sqlCmd = "SELECT * FROM [Print_SN_Bind_Carton_Table] where Bind_SN = '" + txt_reprint_Sn.Text.Trim() + "' ";
                        DataSet ds = db.reDs(sqlCmd);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            txt_Wo.Text = ds.Tables[0].Rows[0]["Work_Order"].ToString();
                            txt_Pn.Text = ds.Tables[0].Rows[0]["Part_No"].ToString();
                            txt_Pc.Text = ds.Tables[0].Rows[0]["Power_Code"].ToString();
                            txt_Qty.Text = ds.Tables[0].Rows[0]["Quantity"].ToString();
                            txt_Bios_Ver.Text = ds.Tables[0].Rows[0]["Bios_Ver"].ToString();
                            txt_No_Number.Text = ds.Tables[0].Rows[0]["No_Number"].ToString();

                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                if (SN_GV.Rows.Count == 0)
                                {
                                    Item = "1";

                                }
                                else
                                {
                                    Item = (SN_GV.Rows.Count + 1).ToString();

                                }
                                SN_GV.Rows.Add(new object[] { Item, ds.Tables[0].Rows[i]["SN"].ToString(), ds.Tables[0].Rows[i]["Record_Time"].ToString() });
                            }
                            if (Item == txt_Qty.Text.Trim())
                            {
                                Btn_Print_Click(sender, e);
                            }

                        }
                        else
                        {
                            List_Msg.Items.Add("ERROR!!查無第一筆SN資料");
                        }
                    }
                    else
                    {
                        List_Msg.Items.Add("ERROR!!請輸入第一筆SN ");
                        txt_reprint_Sn.Focus();
                    }

                }

            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString());
            }
        }
        private void Txt_Wo_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (Convert.ToInt32(e.KeyChar) == 13)
                {
                    if (rdb_Board.IsChecked || rdb_System.IsChecked || rdb_Acc.IsChecked)
                    {
                        if(rdb_Board.IsChecked)
                        {
                            if (!File.Exists(Boardpath))
                            {
                                List_Msg.Items.Add("列印套版下載中......");

                                DLfilename = ini.IniReadValue("Option", "Board_Carton_Label_Name", Cfgname);
                                this.FTP_Dl_Btw_thread.WorkerSupportsCancellation = true; //允許中斷
                                this.FTP_Dl_Btw_thread.RunWorkerAsync(); //呼叫背景程式
                                txt_Sn.ReadOnly = false;
                            }
                        }
                        else if (rdb_System.IsChecked)
                        {
                            if (!File.Exists(Systempath))
                            {
                                DLfilename = ini.IniReadValue("Option", "System_Carton_Label_Name", Cfgname);
                                this.FTP_Dl_Btw_thread.WorkerSupportsCancellation = true; //允許中斷
                                this.FTP_Dl_Btw_thread.RunWorkerAsync(); //呼叫背景程式
                            }
                        }
                        else
                        {
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
                        List_Msg.Items.Add("ERROR!!請先選擇列印製程");
                        radGroupBox7.Focus();
                        return;
                    }
                    #region 清空欄位
                    txt_Sn.ReadOnly = true;
                    txt_Pn.Text = string.Empty;
                    txt_Pc.Text = string.Empty;
                    txt_Qty.Text = string.Empty;
                    txt_Bios_Ver.Text = string.Empty;
                    txt_No_Number.Text = string.Empty;
                    List_Msg.Items.Clear();
                    #endregion

                    if (!string.IsNullOrEmpty(txt_Wo.Text.Trim()))
                    {
                        if (Convert.ToInt32(txt_Wo.Text.Trim().Length) == 10)
                        {
                            #region 開啟欄位
                            txt_Pn.Enabled = true;
                            txt_Qty.Enabled = true;
                            txt_Pc.Enabled = true;
                            txt_Bios_Ver.Enabled = true;
                            txt_Sn.Enabled = true;
                            btn_Print.Enabled = true;
                            btn_Out_Print.Enabled = true;
                            txt_No_Number.Enabled = true;

                            //if (txt_Wo.Text.Substring(0, 3) == "960")
                            //{
                            //    txt_No_Number.Enabled = true;
                            //}
                            //else
                            //{
                            //    txt_No_Number.Enabled = false;
                            //}
                            #endregion

                            string WipInfo = string.Empty;
                            string[] relatedwip;
                            //WipInfo = Auto_Route.WipSystem(txt_Wo.Text.Trim());

                            //改為透過安勤工單查昶亨工單;在透過昶亨工單查詢必要資訊
                            //[0] : wono ; [1] : engsr
                            relatedwip = Auto_Route.WipInfoByRelatedWoNo(txt_Wo.Text.Trim()).Split('&');
                            if (rdb_Board.IsChecked)
                            {
                                WipInfo = Auto_Route.WipBoard_Eve(relatedwip[0]);
                            }
                            else if (rdb_System.IsChecked)
                            {
                                WipInfo = Auto_Route.WipSystem_Eve(relatedwip[0]);
                            }

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
                                        string [] rows= row["ProductID_MF"].ToString().Trim().Split('-');
                                        ENGSR = rows[0];
                                        if (!string.IsNullOrEmpty(ENGSR))
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
                                if (!string.IsNullOrEmpty(descJsonStu.biosVer))
                                {
                                    txt_Bios_Ver.Text = descJsonStu.biosVer;
                                    if (string.IsNullOrEmpty(descJsonStu.biosVer))
                                    {
                                        txt_Bios_Ver.Text = "NA";
                                    }
                                }
                                if (!string.IsNullOrEmpty(descJsonStu.powercord))
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
                                if (!string.IsNullOrEmpty(relatedwip[1]))
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
                                if (!string.IsNullOrEmpty(JsonStu.itemNo))
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

            }

        }

        private bool AvSFIS_Method()
        {
            string sql = "select * from Print_AvSFIS_Table where MOID='" + txt_Wo.Text.Trim() + "' ";
            dataSet = db.reDs(sql);
            return (dataSet.Tables[0].Rows.Count == 0) ? false : true;
        }

        private void Txt_Pn_Leave(object sender, EventArgs e)
        {

            string sqlBTWCmd = "";
            //搜尋列印檔案
            if (!string.IsNullOrEmpty(txt_Pn.Text.Trim()))
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
        private void Txt_Sn_KeyPress(object sender, KeyPressEventArgs e)
        {
            bool AddSN = true;
            string Bind_SN = string.Empty;
            if (string.IsNullOrEmpty(txt_Qty.Text))
            {
                List_Msg.Items.Add( "請先輸入數量 ");
                txt_Qty.Focus();
                return;
            }

            if (Convert.ToInt32(e.KeyChar) == 13)
            {

                if (!string.IsNullOrEmpty(txt_Sn.Text.Trim().ToUpper()))
                {
                    for (int i = 0; i < SN_GV.Rows.Count; i++)
                    {
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

                    if (AddSN == true)
                    {
                        if (SN_GV.Rows.Count == 0)
                        {
                            Item = "1";
                            List_Msg.Items.Clear();
                            Bind_SN = txt_Sn.Text.Trim().ToUpper();
                        }
                        else
                        {
                            Item = (SN_GV.Rows.Count + 1).ToString();
                            Bind_SN = SN_GV.Rows[0].Cells["SN"].Value.ToString().ToUpper();
                            if (int.Parse(Item) > int.Parse(txt_Qty.Text.Trim()))
                            {
                                List_Msg.Items.Add(txt_Sn.Text.Trim().ToUpper() + "EEROR!! 數量滿箱");
                                return;
                            }
                        }
                        SN = txt_Sn.Text.Trim().ToUpper();
                        SN_GV.Rows.Add(new object[] { Item, SN, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") });
                        #region 資料庫寫入SN資訊
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
                        if (db.Exsql(InsSql) == true)
                        {
                            txt_Sn.Text = string.Empty;
                            //輸入序號數量與數量欄位一樣(列印)
                            if (Item == txt_Qty.Text.Trim())
                            {
                                Btn_Print_Click(sender, e);
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
        private void Txt_Qty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
                btn_Out_Print.Enabled = true;
                txt_Sn.Enabled = true;
                btn_Print.Enabled = true;
            }
        }

        private void FTP_Dl_Btw_thread_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                TempUploadResult = false;
                Getftp("Print");


                //FtpWebRequest
                FtpWebRequest ftpRequest = (FtpWebRequest)FtpWebRequest.Create("ftp://" + ftpServer + "/"  + DLfilename);
                NetworkCredential ftpCredential = new NetworkCredential(ftpuser, ftppassword);
                ftpRequest.Credentials = ftpCredential;
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                //FtpWebResponse
                FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                //Get Stream From FtpWebResponse
                Stream ftpStream = ftpResponse.GetResponseStream();
                using (FileStream fileStream = new FileStream(download_Path+"\\"+ DLfilename, FileMode.Create))
                {
                    int bufferSize = 2048;
                    int readCount;
                    byte[] buffer = new byte[bufferSize];

                    readCount = ftpStream.Read(buffer, 0, bufferSize);
                    int allbye = (int)fileStream.Length;
                    Form.CheckForIllegalCrossThreadCalls = false;

                    while (readCount > 0)
                    {

                        fileStream.Write(buffer, 0, readCount);

                        readCount = ftpStream.Read(buffer, 0, bufferSize);
                    }
                }
                ftpStream.Close();
                ftpResponse.Close();
                TempUploadResult = true;



            }
            catch (Exception ex)
            {
                TempUploadResult = false;
                MessageBox.Show(ex.ToString());

            }
        }
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
        private void FTP_Up_Btw_thread_DoWork(object sender, DoWorkEventArgs e)
        {


            //20171117 Jim 上傳總數初始化
            Sum_Of_SQLfile_size = 0;

            TempUploadResult = false;
            if (FTP_Up_Btw_thread.CancellationPending) //如果被中斷...
                e.Cancel = true;
            this.FTP_Up_Btw_thread.WorkerReportsProgress = true;
            BackgroundWorker worker = (BackgroundWorker)sender;
            //string temp_path = System.IO.Path.GetDirectoryName(dialog.FileName);

            if (File.Exists(txt_Btw_Path.Text) == true)
            {

                //File.Copy(ini_filepath, SQL_96level);

                //FileInfo fInfoBefore = new FileInfo(SQL_96level);
                //SizeBefore = fInfoBefore.Length;
                //ftpPutFile = SQL_96level;

                // ZIP.GhostToZip(ghostFileName, ghostFileName.Replace(".GHO", ".ZIP"));
                FileInfo finfo = new FileInfo(txt_Btw_Path.Text);
                UPfilename = DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + filename;
                try
                {
                    Getftp("Print");

                    //ftpPutFile = "SWM_INI";
                    //UPfilename = DateTime.Now.ToString("yyyyMMddHHmmss") +"_"+ filename;
                    FtpWebResponse response = null;
                    FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://" + ftpServer +  "/" + UPfilename);
                    request.KeepAlive = true;
                    request.UseBinary = true;
                    request.Credentials = new NetworkCredential(ftpuser, ftppassword);
                    request.Method = WebRequestMethods.Ftp.UploadFile;
                    request.ContentLength = finfo.Length;//指定上傳文件的大小
                    response = request.GetResponse() as FtpWebResponse;
                    int buffLength = 2048;
                    byte[] buffer = new byte[buffLength];
                    int contentLen;
                    FileStream fs = File.OpenRead(txt_Btw_Path.Text);
                    Stream ftpstream = request.GetRequestStream();
                    contentLen = fs.Read(buffer, 0, buffer.Length);
                    int allbye = (int)finfo.Length;
                    Form.CheckForIllegalCrossThreadCalls = false;

                    int startbye = 0;
                    while (contentLen != 0)
                    {
                        startbye = contentLen + startbye;
                        ftpstream.Write(buffer, 0, contentLen);
                        //更新進度
                        //if (toolStripProgressBar1 != null)
                        //{
                        //    toolStripProgressBar1.Value += contentLen;//更新進度條
                        //}
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
                    Getftp("PrintBarTender");

                    ftpPutFile = "Eversun";
                    //UPfilename = DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + filename;
                    FtpWebResponse response = null;
                    FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://" + ftpServer + "/" + ftpPutFile + "/" + UPfilename);
                    request.KeepAlive = true;
                    request.UseBinary = true;
                    request.Credentials = new NetworkCredential(ftpuser, ftppassword);
                    request.Method = WebRequestMethods.Ftp.UploadFile;
                    request.ContentLength = finfo.Length;//指定上傳文件的大小
                    response = request.GetResponse() as FtpWebResponse;
                    int buffLength = 2048;
                    byte[] buffer = new byte[buffLength];
                    int contentLen;
                    FileStream fs = File.OpenRead(txt_Btw_Path.Text);
                    Stream ftpstream = request.GetRequestStream();
                    contentLen = fs.Read(buffer, 0, buffer.Length);
                    int allbye = (int)finfo.Length;
                    Form.CheckForIllegalCrossThreadCalls = false;

                    int startbye = 0;
                    while (contentLen != 0)
                    {
                        startbye = contentLen + startbye;
                        ftpstream.Write(buffer, 0, contentLen);
                        //更新進度
                        //if (toolStripProgressBar1 != null)
                        //{
                        //    toolStripProgressBar1.Value += contentLen;//更新進度條
                        //}
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
        private void FTP_Up_Btw_thread_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FTP_Up_Btw_thread.Dispose();
            if (TempUploadResult == true)
            {
                #region Label變數設定
                string Qty_Set = string.Empty, Weight_Set = string.Empty, Sn_Set = string.Empty, Mac_Set = string.Empty, Bios_Set = string.Empty;
                List2_Msg.Items.Add("列印檔案上傳完成");
                if(rck_Qty.Checked == true)
                {
                    Qty_Set = "Yes";
                }
                else
                {
                    Qty_Set = "No";
                }
                if (rck_Weight.Checked == true)
                {
                    Weight_Set = "Yes";
                }
                else
                {
                    Weight_Set = "No";
                }
                if (rck_Sn.Checked == true)
                {
                    Sn_Set = "Yes";
                }
                else
                {
                    Sn_Set = "No";
                }
                if (rck_mac.Checked == true)
                {
                    Mac_Set = "Yes";
                }
                else
                {
                    Mac_Set = "No";
                }
                if (rck_Bios.Checked == true)
                {
                    Bios_Set = "Yes";
                }
                else
                {
                    Bios_Set = "No";
                }
                #endregion
                string InsSql = " INSERT INTO [Print_Carton_Table] (Time,Filename,Pn,Wo,Qty_Set,Weight_Set,Sn_Set,Mac_Set,Bios_Set) VALUES("
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
                    InsSql = " INSERT INTO [Print_Wo_Setting_Table] (Record_Time,Work_Order,Print_Type) VALUES("
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
                    //Generate a thumbnail for it.
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
                    labelFormat.PrintSetup.PrinterName = _PrinterName;
                    labelFormat.Print("BarPrint" + DateTime.Now, 3 * 1000);
                }
                else
                {
                    MessageBox.Show("請先選擇印表機", "操作提示");
                }
            }
        }
        private void Getftp(string Ftp_Server_name)//ftp資訊
        {


            string sqlCmd = "SELECT [Ftp_Server_Ip],[Ftp_Server_OA_Ip],[Ftp_Username],[Ftp_Password],[Ftp_Server_name],[Ftp_Factory] FROM [i_Program_FtpServer_Table] where [Ftp_Server_name] ='" + Ftp_Server_name + "' ";
            DataSet ds = db.reDs(sqlCmd);
            if (ds.Tables[0].Rows.Count != 0)
            {
                    ftpServer = ds.Tables[0].Rows[0]["Ftp_Server_OA_Ip"].ToString().Trim();
                    ftpuser = ds.Tables[0].Rows[0]["Ftp_Username"].ToString().Trim();
                    ftppassword = ds.Tables[0].Rows[0]["Ftp_Password"].ToString().Trim();
                    ftpdlfactory = ds.Tables[0].Rows[0]["Ftp_Factory"].ToString().Trim();
            }
        }
        private string selectVerSQL_new(string tool)//Version Check new
        {
            string sqlCmd = "";
            bool result = false;
            try
            {
                sqlCmd = "select *  FROM TE_Program_Table where [Program_Name] ='" + tool + "'";
                DataSet ds = db.reDs(sqlCmd);
                if (ds.Tables[0].Rows.Count != 0)
                {
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        if (ds.Tables[0].Columns[i].ToString() == "Version")
                        {
                            version_new = ds.Tables[0].Rows[0][i].ToString();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("更新異常");
            }
            return version_new;
        }
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
        public void autoupdate()//自動更新
        {
            //寫入目前版本與程式名後執行更新

            Process p = new Process();
            p.StartInfo.FileName = System.Windows.Forms.Application.StartupPath + "\\AutoUpdate.exe";
            p.StartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath; //檔案所在的目錄
            p.Start();
            this.Close();
        }
    }
}
