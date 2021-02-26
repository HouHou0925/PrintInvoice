using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Xml;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ZXing;
using ZXing.QrCode;
using ZXing.Common;
using ZXing.QrCode.Internal;


namespace PrintInvoice
{
    class Program
    {
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool FreeConsole();
        static string[] cmd = {"-i","-p"};
        static string[] cmdval = {"",""};
        
        static string appPath;
        
        static DataSetXml dx1 = new DataSetXml();

        static void Main(string[] args)
        {
            
            //隱藏式窗
            FreeConsole();

            string vName;
            int poi = -1;
            //分析參數
            for (int i = 0; i < args.Length; i++)
            {
                vName = args[i].ToLower();
                int j;
                //檢查是否為指令
                for (j = 0; j < cmd.Length; j++)
                {
                    if (cmd[j] == vName)
                    {
                        //當符合指令時，跳出迴圈，此時j=指令索引
                        break;
                    }
                }
                //如果j>=指令索引陣列長度(沒找到)，代表不是指令，屬於前一個指令資料的附加字串
                //但是也必須在有前一個指令的狀況下 (poi >= 0)
                if (j == cmd.Length && poi >= 0)
                {
                    //附加到前一個指令的資料中
                    if (cmdval[poi] == "")
                    {
                        cmdval[poi] = args[i];
                    }
                    else
                    {
                        cmdval[poi] = cmdval[poi] + " " + args[i];
                    }
                    //處理下一個
                    continue;
                }
                //poi指定為指令索引位置
                poi = j;

                
            }
            //取得系統執行路徑
            appPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            
            //檢查參數
            if (!File.Exists(cmdval[0]))
            {
                ErrorLog("資料檔案" + cmdval[0] + "不存在 -- 終止執行1");
                return;
            }
            //載入資料
            if (LoadData(cmdval[0]) == false)
            {
                ErrorLog("載入檔案" + cmdval[0] + "無資料 -- 終止執行");
                return;
            }
            
            //產生列印物件
            PrintDocument printDoc = new PrintDocument();
            if (cmdval[1] != "")
            {
                printDoc.PrinterSettings.PrinterName = cmdval[1];
            }

            if (!printDoc.PrinterSettings.IsValid)
            {
                ErrorLog("印表機未備妥!!");
            }
            else
            {
                // 綁定事件
                printDoc.BeginPrint += new PrintEventHandler(BeginPrint);
                printDoc.EndPrint += new PrintEventHandler(EndPrint);
                printDoc.QueryPageSettings += (o, e) =>
                {
                    ErrorLog("QueryPageSettings[]");
                };

                printDoc.PrintPage += new PrintPageEventHandler(PrintPage);
                
                DataRow row = dx1.Tables["Table1"].Rows[0];
                //預先執行外部程式
                if (!String.IsNullOrEmpty(row["startrun"].ToString()))
                {
                    if (File.Exists(row["startrun"].ToString()))
                    {
                        Process p1 = Process.Start(row["startrun"].ToString());
                        p1.WaitForExit();
                    }
                    else
                    {
                        ErrorLog("指定預先執行的外部程序不存在：" + row["startrun"].ToString());
                    }
                }
                // 列印開始
                printDoc.Print();

                //預先執行外部程式
                if (!String.IsNullOrEmpty(row["endrun"].ToString()))
                {
                    if (File.Exists(row["endrun"].ToString()))
                    {
                        Process p1 = Process.Start(row["endrun"].ToString());
                        
                    }
                    else
                    {
                        ErrorLog("指定結束時執行的外部程序不存在：" + row["startrun"].ToString());
                    }
                }
            }

        }
        static void BeginPrint(object sender, PrintEventArgs e)
        {
        }
        static void EndPrint(object sender, PrintEventArgs e)
        {
        }
        /// <summary>
        /// DocumentPrinter的PrintPage事件處理程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void PrintPage(object sender, PrintPageEventArgs e)
        {
            DataRow row = dx1.Tables["Table1"].Rows[0];
            //取得印表機解析度，計算定位使用(因為等會用長度互乘，但是精確度不夠誤差會很大)
            float dpiX = (float)e.PageSettings.PrinterSettings.DefaultPageSettings.PrinterResolution.X;
            float dpiY = (float)e.PageSettings.PrinterSettings.DefaultPageSettings.PrinterResolution.Y;
            if (dpiX == 0 && dpiY == 0)
            {
                dpiX = 200;
                dpiY = 200;
                ErrorLog("無法取得印表機X,Y解析度，強迫使用200DPI");
            }
            if (dpiX == 0)
            {
                dpiX = dpiY;
                ErrorLog("無法取得印表機X解析度，強迫使用Y解析度(" + dpiY.ToString() + "DPI)");
            }
            if (dpiY == 0)
            {
                dpiY = dpiX;
                ErrorLog("無法取得印表機Y解析度，強迫使用X解析度(" + dpiX.ToString() + "DPI)");
            }

            float ItoC = 2.54f;
            //基本參考設定
            StringFormat centerFormat = new StringFormat();
            centerFormat.Alignment = StringAlignment.Center;
            centerFormat.LineAlignment = StringAlignment.Near;
            StringFormat normalFormat = new StringFormat();
            //設定列印解析度參考，Pixel處會直接參考目前印表機預設的解析度
            //以ZEBRA機器來說通常都是預設203DPI
            //參考來源為e.PageSettings.PrinterSettings.DefaultPageSettings.PrinterResolution
            e.Graphics.PageUnit = GraphicsUnit.Pixel;
            
            //畫外框(5.7cm X 9.0cm)
            if (row["refbox"].ToString() == "1")
            {
                Rectangle paperRect = new Rectangle(
                    0,
                    0,
                    (int)((5.7f / ItoC) * dpiX),
                    (int)((9.0f / ItoC) * dpiY));
                float[] dashVal = { 5, 10 };
                Pen myPen = new Pen(Brushes.Black);
                myPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                myPen.DashPattern = dashVal;
                e.Graphics.DrawRectangle(myPen, paperRect);
            }
            //畫LOGO
            Image logopic = Image.FromFile(row["logo"].ToString());
            Rectangle logoRect = new Rectangle(
                (int)((0.3f / ItoC) * dpiX),
                (int)((0.5f / ItoC) * dpiY),
                (int)((5.0f / ItoC) * dpiX),
                (int)((1.3f / ItoC) * dpiY));
            e.Graphics.DrawImage(
                logopic, 
                logoRect, 
                0, 0, logopic.Width, logopic.Height, 
                GraphicsUnit.Pixel,
                BWThreshold(logopic,0.9f));
            //標題文字
            Rectangle tRect = new Rectangle(
               (int)((0.1f / ItoC) * dpiX),
               (int)((1.79f / ItoC) * dpiY),
               (int)((5.5f / ItoC) * dpiX),
               (int)((0.57f / ItoC) * dpiY));
            Font titleFont = new Font(@"新細明體", 16);
            
            e.Graphics.DrawString(
                "電子發票證明聯" + row["fixtext"].ToString(),
                titleFont,
                Brushes.Black,
                tRect, centerFormat);
            //e.Graphics.DrawRectangle(Pens.Black, tRect);
            //繪製年月份
            Rectangle bRect = new Rectangle(
               (int)((0.3f / ItoC) * dpiX),
               (int)((2.22f / ItoC) * dpiY),
               (int)((5.0f / ItoC) * dpiX),
               (int)((0.68f / ItoC) * dpiY));
            Font barFont = new Font(@"微軟正黑體",16,FontStyle.Bold);
            e.Graphics.DrawString(
                row["year"].ToString() + "年" + row["months"].ToString() + "月",
                barFont,
                Brushes.Black,
                bRect,
                centerFormat);
            //e.Graphics.DrawRectangle(Pens.Black, bRect);
            //繪製發票號碼
            Rectangle iRect = new Rectangle(
                (int)((0.3f / ItoC) * dpiX),
                (int)((2.746f / ItoC) * dpiY),
                (int)((5.0f / ItoC) * dpiX),
                (int)((0.68f / ItoC) * dpiY));
            //Font iFont = new Font("Arial Black", 16);
            e.Graphics.DrawString(
                row["invoice"].ToString(),
                barFont,
                Brushes.Black,
                iRect,
                centerFormat);
            //繪製列印日期
            Font mFont = new Font(@"微軟正黑體", 9);
            DateTime now = DateTime.Now;
            e.Graphics.DrawString(now.ToString("yyyy-MM-dd HH:mm:ss"), mFont, Brushes.Black,
                (0.4f / ItoC) * dpiX,
                (3.38f / ItoC) * dpiY);
            //繪製附加文字
            e.Graphics.DrawString(row["atttext"].ToString(), mFont, Brushes.Black,
                (4.0f / ItoC) * dpiX,
                (3.38f / ItoC) * dpiY);
            //隨機碼、總計
            e.Graphics.DrawString("隨機碼:" + row["randcode"].ToString(), mFont, Brushes.Black,
                (0.4f / ItoC) * dpiX,
                (3.76f / ItoC) * dpiY);
            e.Graphics.DrawString("總計:" + row["total"].ToString(), mFont, Brushes.Black,
                (2.98f / ItoC) * dpiX,
                (3.76f / ItoC) * dpiY);
            //賣方、買方
            e.Graphics.DrawString("賣方:" + row["sellerid"].ToString(), mFont, Brushes.Black,
                (0.4f / ItoC) * dpiX,
                (4.1f / ItoC) * dpiY);
            e.Graphics.DrawString("買方:" + row["byerid"].ToString(), mFont, Brushes.Black,
                (2.98f / ItoC) * dpiX,
                (4.1f / ItoC) * dpiY);
            //條碼
            Image c39Img = GetCode39(row["barcode"].ToString(),50);
            //計算條碼寬度產生的比例(會依照印表機DPI值變化)
            //可列印寬度(兩側留白0.3以上) = 5.7 - (0.3 * 2) = 5.1cm
            float widthlimit = (5.7f / ItoC) * dpiX;
            float rate = 0f;
            float newWidth = 0f;
            do
            {
                rate ++;
                newWidth = c39Img.Width * rate;

            }while((c39Img.Width * (rate + 1.0f)) <= widthlimit);
            //計算X定位點(條碼置中)
            float newX = (((5.7f / ItoC) * dpiX) / 2.0f) - (newWidth / 2.0f);

            Rectangle c39Rect = new Rectangle(
                (int)newX,
                (int)((4.5f / ItoC) * dpiY),
                (int)newWidth,
                (int)((0.65f / ItoC) * dpiY));
            e.Graphics.DrawImage(c39Img, c39Rect, 0, 0, c39Img.Width, c39Img.Height, System.Drawing.GraphicsUnit.Pixel);
            //QR碼
            //補足120byte長度
            //先取得實際長度後，計算應補字元數，然後在padright時,以程式內的length取得的長度加上應補字元數才會正確
            int dataLen = Encoding.Default.GetByteCount(row["qrcode1"].ToString());
            int subd = 0;
            if (dataLen < 120)
            {
                subd = 120 - dataLen;
                row["qrcode1"] = row["qrcode1"].ToString().PadRight(row["qrcode1"].ToString().Length + subd);
            }
            dataLen = Encoding.Default.GetByteCount(row["qrcode2"].ToString());
            if (dataLen < 120)
            {
                subd = 120 - dataLen;
                row["qrcode2"] = row["qrcode2"].ToString().PadRight(row["qrcode2"].ToString().Length + subd);
            }
            //利用matrix來計算產生QR碼的實際Size(去白邊)
            var hints = new Dictionary<EncodeHintType, object> { { EncodeHintType.CHARACTER_SET, "UTF-8" } };
            var matrix = new MultiFormatWriter().encode(row["qrcode1"].ToString(), BarcodeFormat.QR_CODE, 140, 140, hints);
            var matrix2 = new MultiFormatWriter().encode(row["qrcode2"].ToString(), BarcodeFormat.QR_CODE, 140, 140, hints);

            matrix = CutWhiteBorder(matrix);
            matrix2 = CutWhiteBorder(matrix2);
            //把QR碼實際Size給BarcodeWriter參考產生
            var qr1Writer = new BarcodeWriter
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new QrCodeEncodingOptions
                {
                    Height = matrix.Width,
                    Width = matrix.Height,
                    CharacterSet = "utf-8",
                    Margin=0,
                    ErrorCorrection = ErrorCorrectionLevel.L
                }
            };
            
            var qr2Writer = new BarcodeWriter
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new QrCodeEncodingOptions
                {
                    Height = matrix2.Width,
                    Width = matrix2.Height,
                    CharacterSet = "utf-8",
                    Margin = 0,
                    ErrorCorrection = ErrorCorrectionLevel.L
                }
            };
            //e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;
            //e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            //QR碼至此產生的大小為不含白邊的原圖大小
            Image qr1Img = qr1Writer.Write(row["qrcode1"].ToString());
            Image qr2Img = qr2Writer.Write(row["qrcode2"].ToString());
            
            //要把QR碼強制產生為規定的大小
            Rectangle q1Rect = new Rectangle(
                (int)((0.9f / ItoC) * dpiX),
                (int)((5.3f / ItoC) * dpiY),
                (int)((1.7f / ItoC) * dpiX),
                (int)((1.7f / ItoC) * dpiY));
            Rectangle q2Rect = new Rectangle(
                (int)((3.19f / ItoC) * dpiX),
                (int)((5.3f / ItoC) * dpiY),
                (int)((1.7f / ItoC) * dpiX),
                (int)((1.7f / ItoC) * dpiY));
            //e.Graphics.FillRectangle(Brushes.Black, q1Rect);
            //e.Graphics.FillRectangle(Brushes.Black, q2Rect);
            e.Graphics.DrawImage(
                qr1Img, 
                q1Rect, 
                0, 0, qr1Img.Width, qr1Img.Height, 
                System.Drawing.GraphicsUnit.Pixel,
                BWThreshold(qr1Img,0.9f));
            e.Graphics.DrawImage(
                qr2Img, 
                q2Rect, 
                0, 0, qr2Img.Width, qr2Img.Height, 
                System.Drawing.GraphicsUnit.Pixel,
                BWThreshold(qr1Img, 0.9f));
            //e.Graphics.DrawRectangle(Pens.Black, q1Rect);
            //e.Graphics.DrawRectangle(Pens.Black, q2Rect);
            //註記1
            e.Graphics.DrawString(row["note1"].ToString(), mFont, Brushes.Black,
                (0.4f / ItoC) * dpiX,
                (7.2f / ItoC) * dpiY);
            //註記2
            e.Graphics.DrawString(row["note2"].ToString(), mFont, Brushes.Black,
                (0.4f / ItoC) * dpiX,
                (7.5f / ItoC) * dpiY);
            //註記3
            e.Graphics.DrawString(row["note3"].ToString(), mFont, Brushes.Black,
                (0.4f / ItoC) * dpiX,
                (7.8f / ItoC) * dpiY);
            //畫線
            Point p1 = new Point(
                (int)((0.4f / ItoC) * dpiX),
                (int)((8.17f / ItoC) * dpiY));
            Point p2 = new Point(
                (int)((5.0f / ItoC) * dpiX),
                (int)((8.17f / ItoC) * dpiY));
            e.Graphics.DrawLine(new Pen(Brushes.Black, 2), p1, p2);
            //註記4
            e.Graphics.DrawString(row["note4"].ToString(), mFont, Brushes.Black,
                (0.4f / ItoC) * dpiX,
                (8.2f / ItoC) * dpiY);
            //註記5
            e.Graphics.DrawString(row["note5"].ToString(), mFont, Brushes.Black,
                (0.4f / ItoC) * dpiX,
                (8.5f / ItoC) * dpiY);

            //告知PrintDocument列印結束(此後EndPrint事件才能執行)
            e.HasMorePages = false;

            ErrorLog("PrintPage[]");
        }
        /// <summary>
        /// 測量QR碼產生時的實際大小，將來給BarcodeWriter產生參考Width、Height時，
        /// 可以避免白邊的產生。
        /// </summary>
        /// <param name="matrix">紀錄描述用的點陣空間</param>
        /// <returns>返回實際大小的點陣空間</returns>
        static BitMatrix CutWhiteBorder(BitMatrix matrix)
        {
            int[] rec = matrix.getEnclosingRectangle();
            int resWidth = rec[2] + 1;
            int resHeight = rec[3] + 1;
            BitMatrix resMatrix = new BitMatrix(resWidth + 1, resHeight + 1);
            resMatrix.clear();
            for (int i = 0; i < resWidth; i++)
            {
                for (int j = 0; j < resHeight; j++)
                {
                    resMatrix.flip(i + 1, j + 1);
                }
            }
            return resMatrix;
        }
        /// <summary>
        /// 寫入錯誤日誌
        /// </summary>
        /// <param name="message">訊息文字</param>
        static void ErrorLog(string message)
        {
            string logFile = appPath + @"\Error.log";
            string text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\t" + message;
            using (StreamWriter fs = new StreamWriter(logFile, true, Encoding.GetEncoding("big5")))
            {
                fs.WriteLine(text);
            }
        }
        /// <summary>
        /// 將XML檔案載入DataSetXML裡面，其XML欄位名稱會自動配對DataSetXML的欄位
        /// 全部以文字型態載入;DataSetXML是個Global變數
        /// </summary>
        /// <param name="afile">指定載入的XML檔案</param>
        /// <returns>成功true/失敗false</returns>
        static bool LoadData(string afile)
        {
            XmlDocument xmlDOc = new XmlDocument();
            //string xmlText = File.ReadAllText(afile, Encoding.GetEncoding("big5"));
            xmlDOc.Load(afile);
            //xmlDOc.LoadXml(xmlText);
            XmlNodeList nodeList = xmlDOc.SelectNodes(@"data/row/col");
            int row = 0;

            DataRow dr = dx1.Tables["Table1"].NewRow(); 

            //判斷節點存在
            if (nodeList == null)
            {
                ErrorLog("載入的XML檔案" + afile + ",不存在data/row/col結構! -- 終止執行");
                return false;
            }
            
            foreach (XmlNode oneNode in nodeList)
            {
                row++;
                if (oneNode.Attributes["name"] != null)
                {
                    string name = oneNode.Attributes["name"].Value;
                    if(dr[name] != null)
                    {
                        dr[name]=oneNode.InnerText;
                    
                    }else{
                        ErrorLog("欄位" + name + "不存在DataSetXml.Table1中");
                    }
                    
                }
                else
                {
                    ErrorLog("第" + row.ToString() + "個col發現沒有name屬性");
                }
            }

            dx1.Tables["Table1"].Rows.Add(dr);
            if (dx1.Tables["Table1"].Rows.Count <= 0)
            {
                return false;
            }
            return true;

        }
        /// <summary>
        /// 產生code39條碼影像，影像粗細比為 2:1 ，細線以1bit為主
        /// 因此寬度無法自訂，會依照編碼字串長度發生變化，
        /// 其產生寬度為 = ((寬線 * 3) + (細線 * 7) * (字串長度 + 2)) + (左邊界5 * 2)
        /// 若將來影像縮放，最好依照整數比例縮放，以免影響讀取效果
        /// </summary>
        /// <param name="strSource">編碼字串(不含前後*號)</param>
        /// <param name="barHeight">影像高度</param>
        /// <returns>Code39 影像</returns>
        static Bitmap GetCode39(string strSource,int barHeight)
        {
            int x = 5; //左邊界
            int y = 0; //上邊界
            int WidLength = 2; //粗BarCode長度
            int NarrowLength = 1; //細BarCode長度
            int BarCodeHeight = barHeight; //BarCode高度
            int intSourceLength = strSource.Length;
            string strEncode = "010010100"; //編碼字串 初值為 起始符號 *

            string AlphaBet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*"; //Code39的字母

            string[] Code39 = //Code39的各字母對應碼
            {
       /* 0 */ "000110100",
       /* 1 */ "100100001",
       /* 2 */ "001100001",
       /* 3 */ "101100000",
       /* 4 */ "000110001",
       /* 5 */ "100110000",
       /* 6 */ "001110000",
       /* 7 */ "000100101",
       /* 8 */ "100100100",
       /* 9 */ "001100100",
       /* A */ "100001001",
       /* B */ "001001001",
       /* C */ "101001000",
       /* D */ "000011001",
       /* E */ "100011000",
       /* F */ "001011000",
       /* G */ "000001101",
       /* H */ "100001100",
       /* I */ "001001100",
       /* J */ "000011100",
       /* K */ "100000011",
       /* L */ "001000011",
       /* M */ "101000010",
       /* N */ "000010011",
       /* O */ "100010010",
       /* P */ "001010010",
       /* Q */ "000000111",
       /* R */ "100000110",
       /* S */ "001000110",
       /* T */ "000010110",
       /* U */ "110000001",
       /* V */ "011000001",
       /* W */ "111000000",
       /* X */ "010010001",
       /* Y */ "110010000",
       /* Z */ "011010000",
       /* - */ "010000101",
       /* . */ "110000100",
       /*' '*/ "011000100",
       /* $ */ "010101000",
       /* / */ "010100010",
       /* + */ "010001010",
       /* % */ "000101010",
       /* * */ "010010100"
            };


            strSource = strSource.ToUpper();

            //實作圖片
            Bitmap objBitmap = new Bitmap(
              ((WidLength * 3 + NarrowLength * 7) * (intSourceLength + 2)) + (x * 2),
              BarCodeHeight + (y * 2));
            objBitmap.SetResolution(200f, 200f);

            Graphics objGraphics = Graphics.FromImage(objBitmap); //宣告GDI+繪圖介面

            //填上底色
            objGraphics.FillRectangle(Brushes.White, 0, 0, objBitmap.Width, objBitmap.Height);

            for (int i = 0; i < intSourceLength; i++)
            {

                if (AlphaBet.IndexOf(strSource[i]) == -1 || strSource[i] == '*') //檢查是否有非法字元
                {
                    objGraphics.DrawString("含有非法字元", SystemFonts.DefaultFont, Brushes.Red, x, y);
                    return objBitmap;
                }
                //查表編碼
                strEncode = string.Format("{0}0{1}", strEncode, Code39[AlphaBet.IndexOf(strSource[i])]);
            }

            strEncode = string.Format("{0}0010010100", strEncode); //補上結束符號 *

            int intEncodeLength = strEncode.Length; //編碼後長度
            int intBarWidth;

            for (int i = 0; i < intEncodeLength; i++) //依碼畫出Code39 BarCode
            {
                intBarWidth = strEncode[i] == '1' ? WidLength : NarrowLength;
                objGraphics.FillRectangle(i % 2 == 0 ? Brushes.Black : Brushes.White,
                  x, y, intBarWidth, BarCodeHeight);
                x += intBarWidth;
            }
            return objBitmap;
        }
        /// <summary>
        /// 取得具有可設定Threshold值的黑白影像ImageAttributes
        /// </summary>
        /// <param name="sourceImage">參考的影像檔</param>
        /// <param name="ThresholdLevel">Threshold值</param>
        /// <returns>黑白影像的ImageAttributes</returns>
        static ImageAttributes BWThreshold(Image sourceImage, float ThresholdLevel)
        {
            var gray_matrix = new float[][] { 
                new float[] { 0.299f, 0.299f, 0.299f, 0, 0 }, 
                new float[] { 0.587f, 0.587f, 0.587f, 0, 0 }, 
                new float[] { 0.114f, 0.114f, 0.114f, 0, 0 }, 
                new float[] { 0,      0,      0,      1, 0 }, 
                new float[] { 0,      0,      0,      0, 1 } 
            };
            var ia = new System.Drawing.Imaging.ImageAttributes();
            using (Graphics gr = Graphics.FromImage(sourceImage))
            {
                ia.SetColorMatrix(new System.Drawing.Imaging.ColorMatrix(gray_matrix));
                ia.SetThreshold(ThresholdLevel); // Change this threshold as needed
            }
            return ia;
        }
    }
}
