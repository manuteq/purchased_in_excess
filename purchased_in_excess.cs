using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace CreateExcelReport
{
    class Program
    {
        private const string sqlselect = @"
        SELECT             
            [ITEM_CODE] AS 品號,
            [ITEM_DESCRIPTION] AS 品名,
            [PR_D].[ITEM_SPECIFICATION] AS 規格,                       
            [WAREHOUSE].[WAREHOUSE_NAME] AS 倉別,
            UNIT.UNIT_NAME AS 單位,
            [BUSINESS_QTY] AS 入庫數量,                                 
            FORMAT(PR.ApproveDate, 'yyyy/MM/dd') AS 進貨核准日,
            [PR].[DOC_NO] AS 進貨單號,
            [SequenceNumber] AS 序,                        
            [PR_D].[REMARK] AS 備註,   
            PR.[REMARK]	AS 入庫單備註
        FROM [PURCHASE_RECEIPT_D] PR_D
            INNER JOIN [PURCHASE_RECEIPT] PR ON PR_D.[PURCHASE_RECEIPT_ID] = PR.[PURCHASE_RECEIPT_ID]
            LEFT JOIN ITEM ON PR_D.[ITEM_ID] = ITEM.[ITEM_BUSINESS_ID]
            LEFT JOIN UNIT ON PR_D.BUSINESS_UNIT_ID = UNIT.UNIT_ID
            LEFT JOIN WAREHOUSE ON PR_D.[WAREHOUSE_ID] = WAREHOUSE.[WAREHOUSE_ID]
        WHERE 
            PR.ApproveStatus = 'Y' AND
            PR.ApproveDate >= DATEADD(day, -180, GETDATE()) AND
            ITEM_ID NOT IN (
                SELECT [ITEM_ID] FROM (
                    SELECT [ITEM_ID]          
                    FROM [ISSUE_RECEIPT_D] IR_D
                    INNER JOIN [ISSUE_RECEIPT] IR ON IR_D.[ISSUE_RECEIPT_ID] = IR.[ISSUE_RECEIPT_ID]
                    INNER JOIN DOC ON IR.[DOC_ID] = DOC.DOC_ID
                    WHERE 
                        IR.[ApproveStatus] = 'Y' AND
                        IR.[ApproveDate] >= DATEADD(day, -180, GETDATE()) AND
                        DOC.[CATEGORY] = '56'    
                    UNION ALL
                    SELECT [ITEM_ID]          
                    FROM TRANSFER_DOC_D TD_D
                    INNER JOIN TRANSFER_DOC TD ON TD_D.[TRANSFER_DOC_ID] = TD.[TRANSFER_DOC_ID]
                    INNER JOIN DOC ON TD.[DOC_ID] = DOC.DOC_ID
                    WHERE 
                        TD.[ApproveStatus] = 'Y' AND
                        TD.[ApproveDate] >= DATEADD(day, -180, GETDATE()) AND
                        DOC.[CATEGORY] = '16'
                ) AS CombinedResults
            )
        ORDER BY [ITEM_DESCRIPTION], [PR].[DOC_NO], [SequenceNumber]                        
        ";

        public const string ConnectionStringManuteq = "Data Source=10.10.10.111;Initial Catalog=VWT_E10;Persist Security Info=True;User ID=sa;Password=m60246598;Connect Timeout=3600;";
        public const string ConnectionStringWK151 = "Data Source=192.168.0.112;Initial Catalog=VWT_E10;Persist Security Info=True;User ID=sa;Password=m60246598;Connect Timeout=3600;";
        public const string ConnectionStringRelease = "Data Source=192.168.0.135;Initial Catalog=VWT_E10;Persist Security Info=True;User ID=manuteq;Password=Mn60246598;Connect Timeout=3600;";
        public const string DefaultFileName = "近180天已入庫未領用_物料明細";
        public const string DefaultSavePath = "\\\\192.168.0.159\\VisionWide_File\\Z-資料交換暫存區\\digi_DailyReport\\Inventory\\";

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            CultureInfo currentCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            string hostname = Environment.MachineName;
            string connectionString;
            string saveFileName;

            switch (hostname)
            {
                case "WK151":
                    connectionString = ConnectionStringWK151;
                    saveFileName = "C:\\manuteq" + "\\" + DefaultFileName + ".xlsx";
                    break;
                case "DESKTOP-JENTSO":
                    connectionString = ConnectionStringManuteq;
                    saveFileName = "C:\\manuteq" + "\\" + DefaultFileName + ".xlsx";
                    break;
                default:
                    connectionString = ConnectionStringRelease;
                    saveFileName = DefaultSavePath + DefaultFileName + ".xlsx";
                    break;
            }

            try
            {
                FileInfo oldFile = new FileInfo(saveFileName);
                if (oldFile.Exists)
                {
                    int retries = 10; // 設定重試次數
                    while (retries > 0)
                    {
                        try
                        {
                            oldFile.Delete();
                            break; // 刪除成功，跳出迴圈
                        }
                        catch (IOException ex)
                        {
                            Console.WriteLine("Error removing old Excel report: " + ex.Message);
                            retries--; // 重試次數 -1
                            Thread.Sleep(5000); // 暫停 5 秒再次嘗試
                        }
                    }
                    if (retries == 0)
                    {
                        Console.WriteLine("Excel report file is locked, program will exit.");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error removing old Excel report: " + ex.Message);
                return;
            }

            try
            {
                using var package = new ExcelPackage(new FileInfo(saveFileName));

                var worksheet = package.Workbook.Worksheets.Add(DefaultFileName);
                worksheet.Cells[1, 1].Value = DefaultFileName + "_(" + DateTime.Now.ToString("yyyy-MM-dd") + ")";
                worksheet.Cells["A1:K1"].Merge = true;
                worksheet.Cells["A1:K1"].Style.Font.Bold = true;
                worksheet.Cells["A1:K1"].Style.Font.Name = "微軟正黑體";
                worksheet.Cells["A1:K1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                int row = 2;
                try
                {
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        using (SqlCommand cmd = new SqlCommand(sqlselect, conn))
                        {
                            using (SqlDataReader dr = cmd.ExecuteReader())
                            {
                                if (dr.HasRows)
                                {
                                    for (int col = 1; col <= dr.FieldCount; col++)
                                    {
                                        worksheet.Cells[row, col].Value = dr.GetName(col - 1);
                                        worksheet.Cells[row, col].Style.Font.Bold = true;
                                        worksheet.Cells[row, col].Style.Font.Name = "微軟正黑體";
                                    }
                                    row++;
                                }
                                while (dr.Read())
                                {
                                    for (int col = 1; col <= dr.FieldCount; col++)
                                    {
                                        if (col == 22 || col == 24)
                                        {
                                            if (!dr.IsDBNull(col - 1))
                                            {
                                                DateTime date = dr.GetDateTime(col - 1);
                                                string dateString = date.ToString("yyyy/MM/dd"); // 格式為 yyyy/MM/dd
                                                worksheet.Cells[row, col].Value = dateString;
                                            }

                                        }
                                        else
                                        {

                                            worksheet.Cells[row, col].Value = dr.GetValue(col - 1);

                                        }
                                        worksheet.Cells[row, col].Style.Font.Name = "微軟正黑體";
                                        worksheet.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                    }
                                    row++;
                                }
                            }
                        }
                    }
                }
                catch (SqlException ex)
                {
                    Console.WriteLine("Error connecting to database: " + ex.Message);
                    return;
                }
                catch (IOException ex)
                {
                    Console.WriteLine("Error reading or writing file: " + ex.Message);
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    return;
                }

                worksheet.Cells["A2:K" + (row - 1)].Style.Font.Size = 10;
                worksheet.Cells["A2:K" + (row - 1)].AutoFitColumns();
                worksheet.Cells["B2:K" + (row - 1)].Style.WrapText = true;
                worksheet.Cells["A2:K" + (row - 1)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A2:K2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells["A2:K2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#D9D9D9"));
                worksheet.Cells["A2:K2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFF00"));

                worksheet.Cells["A2:K" + (row - 1)].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells["A2:K" + (row - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells["A2:K" + (row - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells["A2:K" + (row - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                worksheet.Cells["A2:A" + (row - 1)].AutoFitColumns(20);
                worksheet.Cells["B2:C" + (row - 1)].AutoFitColumns(24);
                worksheet.Cells["D2:D" + (row - 1)].AutoFitColumns(20);
                worksheet.Cells["E2:E" + (row - 1)].AutoFitColumns(8);
                worksheet.Cells["F2:G" + (row - 1)].AutoFitColumns(12);
                worksheet.Cells["H2:H" + (row - 1)].AutoFitColumns(20);
                worksheet.Cells["I2:I" + (row - 1)].AutoFitColumns(6);
                worksheet.Cells["J2:K" + (row - 1)].AutoFitColumns(24);

                worksheet.Cells["F3:F" + (row - 1)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["I3:I" + (row - 1)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J2:K" + (row - 1)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                // 設定凍結視窗
                worksheet.View.FreezePanes(3, 1);


                // Save the workbook
                package.Save();

                Console.WriteLine("Excel report has been created on your desktop");


                if (hostname.Equals("WK176"))
                {
                    // do nothing
                }
                else
                {
                    if (System.Windows.Forms.MessageBox.Show("Would you like to open it?", "Created Excel report", System.Windows.Forms.MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(saveFileName);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error creating Excel report: " + ex.Message);
            }

            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = currentCulture;
            }
        }


    }

}



