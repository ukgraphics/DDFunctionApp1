using GrapeCity.Documents.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;

namespace DDFunctionApp1
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
            [Blob("template/SimpleInvoiceJP.xlsx", FileAccess.Read, Connection = "AzureWebJobsStorage")]Stream inputBlob,
            [Blob("output/ResultInvoiceJP.xlsx", FileAccess.Write, Connection = "AzureWebJobsStorage")] Stream outxlsxBlob,
            [Blob("output/ResultInvoiceJP.pdf", FileAccess.Write, Connection = "AzureWebJobsStorage")] Stream outpdfBlob,
            ILogger log)
        {
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject<Data>(requestBody);

            if (data is null)
                return new BadRequestObjectResult("Please pass a data in the request body");
            else
            {
                //トライアル用
                string key = System.Environment.GetEnvironmentVariable("DdExcelLicenseString", EnvironmentVariableTarget.Process);
                Workbook.SetLicenseKey(key);

                //新しいワークブックを生成
                var workbook = new GrapeCity.Documents.Excel.Workbook();

                //テンプレートを読み込む
                workbook.Open(inputBlob);

                //ワークシートの取得
                var worksheet = workbook.ActiveSheet;

                //発行元情報をセルに指定
                CultureInfo culture = new CultureInfo("ja-JP");
                worksheet.Range["I1"].Value = DateTime.Now.ToString("D", culture); //発行日
                var ticks = DateTime.Now.Ticks.ToString();
                worksheet.Range["H3"].Value = DateTime.Now.Day.ToString("00")
                           + "-" + ticks.Substring(ticks.Length - 4, 4);  //伝票番号
                worksheet.Range["I3"].Value = data.publisher.representative; //担当者
                worksheet.Range["G8"].Value = data.publisher.companyname; //発行元
                worksheet.Range["G9"].Value = "〒" + data.publisher.postalcode; //郵便番号
                worksheet.Range["G10"].Value = data.publisher.address1 + data.publisher.address2; //所在地
                worksheet.Range["H12"].Value = data.publisher.tel;//電話番号
                worksheet.Range["G13"].Value = data.publisher.bankname; //銀行名
                worksheet.Range["H13"].Value = data.publisher.bankblanch; //支店名
                worksheet.Range["H14"].Value = data.publisher.account; //口座番号 

                //顧客情報をセルに指定
                worksheet.Range["A3"].Value = "〒" + data.customer.postalcode; //郵便番号
                worksheet.Range["A4"].Value = data.customer.address1; //住所1
                worksheet.Range["A5"].Value = data.customer.address2; //住所2
                worksheet.Range["A6"].Value = data.customer.companyname; //会社名
                worksheet.Range["A8"].Value = data.customer.name; //氏名

                //明細の開始位置を指定
                var dt_init_row = 17;
                var lines_mun = 2;  //1明細で2行利用
                                    //明細データ分の繰り返し
                for (int i = 0; i < data.customer.detail.Length; i++)
                {
                    var this_item = i * lines_mun;
                    worksheet.Range[dt_init_row + this_item, 0].Value = (string)data.customer.detail[i].sku; //商品番号
                    worksheet.Range[(dt_init_row + 1) + this_item, 0].Value = data.customer.detail[i].name; //商品名
                    worksheet.Range[dt_init_row + this_item, 4].Value = data.customer.detail[i].price; //単価
                    worksheet.Range[dt_init_row + this_item, 5].Value = data.customer.detail[i].unit; //数量
                    worksheet.Range[dt_init_row + this_item, 7].Value = data.customer.detail[i].remark; //備考
                }

                //Excelファイルに保存
                workbook.Save(outxlsxBlob);

                //PDFファイルに保存
                workbook.Save(outpdfBlob, SaveFileFormat.Pdf);

                return new OkObjectResult("Finished.");
            }
        }
    }
}
