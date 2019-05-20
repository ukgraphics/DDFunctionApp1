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
                //�g���C�A���p
                string key = System.Environment.GetEnvironmentVariable("DdExcelLicenseString", EnvironmentVariableTarget.Process);
                Workbook.SetLicenseKey(key);

                //�V�������[�N�u�b�N�𐶐�
                var workbook = new GrapeCity.Documents.Excel.Workbook();

                //�e���v���[�g��ǂݍ���
                workbook.Open(inputBlob);

                //���[�N�V�[�g�̎擾
                var worksheet = workbook.ActiveSheet;

                //���s�������Z���Ɏw��
                CultureInfo culture = new CultureInfo("ja-JP");
                worksheet.Range["I1"].Value = DateTime.Now.ToString("D", culture); //���s��
                var ticks = DateTime.Now.Ticks.ToString();
                worksheet.Range["H3"].Value = DateTime.Now.Day.ToString("00")
                           + "-" + ticks.Substring(ticks.Length - 4, 4);  //�`�[�ԍ�
                worksheet.Range["I3"].Value = data.publisher.representative; //�S����
                worksheet.Range["G8"].Value = data.publisher.companyname; //���s��
                worksheet.Range["G9"].Value = "��" + data.publisher.postalcode; //�X�֔ԍ�
                worksheet.Range["G10"].Value = data.publisher.address1 + data.publisher.address2; //���ݒn
                worksheet.Range["H12"].Value = data.publisher.tel;//�d�b�ԍ�
                worksheet.Range["G13"].Value = data.publisher.bankname; //��s��
                worksheet.Range["H13"].Value = data.publisher.bankblanch; //�x�X��
                worksheet.Range["H14"].Value = data.publisher.account; //�����ԍ� 

                //�ڋq�����Z���Ɏw��
                worksheet.Range["A3"].Value = "��" + data.customer.postalcode; //�X�֔ԍ�
                worksheet.Range["A4"].Value = data.customer.address1; //�Z��1
                worksheet.Range["A5"].Value = data.customer.address2; //�Z��2
                worksheet.Range["A6"].Value = data.customer.companyname; //��Ж�
                worksheet.Range["A8"].Value = data.customer.name; //����

                //���ׂ̊J�n�ʒu���w��
                var dt_init_row = 17;
                var lines_mun = 2;  //1���ׂ�2�s���p
                                    //���׃f�[�^���̌J��Ԃ�
                for (int i = 0; i < data.customer.detail.Length; i++)
                {
                    var this_item = i * lines_mun;
                    worksheet.Range[dt_init_row + this_item, 0].Value = (string)data.customer.detail[i].sku;
                    worksheet.Range[(dt_init_row + 1) + this_item, 0].Value = data.customer.detail[i].name;
                    worksheet.Range[dt_init_row + this_item, 4].Value = data.customer.detail[i].price;
                    worksheet.Range[dt_init_row + this_item, 5].Value = data.customer.detail[i].unit;
                    worksheet.Range[dt_init_row + this_item, 7].Value = data.customer.detail[i].remark;
                }

                //Excel�t�@�C���ɕۑ�
                workbook.Save(outxlsxBlob);

                //PDF�t�@�C���ɕۑ�
                workbook.Save(outpdfBlob, SaveFileFormat.Pdf);

                return new OkObjectResult("Finished.");
            }
        }
    }
}
