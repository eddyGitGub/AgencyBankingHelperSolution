using NPOI.XSSF.UserModel;
using System.Globalization;
using System.Text.Json;
using NodaTime;
using NodaTime.Extensions;
using Org.BouncyCastle.Asn1.Ocsp;

namespace AgencyBankingHelperSolution
{
	public static class Endpoints
	{
		public static void MapNotificationEndPoint(this WebApplication app)
		{
			app.MapPost("/transaction/notification", UploadAndSendFile).DisableAntiforgery();
		}

		static async Task<IResult> UploadAndSendFile(IFormFile file, IHttpClientFactory httpClientFactory)
		{
			if (file == null || file.Length == 0)
			{
				return TypedResults.BadRequest(GetEmptyResult());
			}

			using var stream = file.OpenReadStream();
			var excelData = await ReadExcelData(stream);
			//var json= JsonSerializer.Serialize(excelData.FirstOrDefault());

			if (excelData == null || excelData.Count == 0)
			{
				return TypedResults.BadRequest(GetEmptyResult());
			}

			var (failedData, successfulCount) = await SendDataToApiAsync(excelData, httpClientFactory);

			var result = new
			{
				failedData,
				successfulCount
			};

			return TypedResults.Ok(result);
		}

		

		static async Task<(List<TransactionDto> failedData, int successfulCount)> SendDataToApiAsync(List<TransactionDto> excelData, IHttpClientFactory httpClientFactory)
		{
			var failedData = new List<TransactionDto>();
			var successfulCount = 0;
			var httpClient = httpClientFactory.CreateClient();
			var apiUrl = "https://tms-trans-notification.arcamoney.com/transaction/notification"; // Replace with your destination API URL

			foreach (var dataRecord in excelData)
			{
				var response = await httpClient.PostAsJsonAsync(apiUrl, dataRecord);

				if (!response.IsSuccessStatusCode)
				{
					failedData.Add(dataRecord);
				}
				else
				{
					successfulCount++;
				}
			}

			return (failedData, successfulCount);
		}
		static object GetEmptyResult()
		{
			return new
			{
				failedData = Enumerable.Empty<TransactionDto>(),
				successfulCount = 0,
			};
		}
		private static async Task<List<TransactionDto>> ReadExcelData(Stream stream)
		{
			var excelDataList = new List<TransactionDto>();

			using (var workbook = new XSSFWorkbook(stream))
			{
				var sheet = workbook.GetSheetAt(0);

				for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
				{
					var row = sheet.GetRow(rowIndex);

					if (row != null)
					{

                        var dateTime = DateTime.UtcNow;
                        var dateString = row.GetCell(6) != null ? row.GetCell(6).ToString() : dateTime.ToString("o"); // "o" format specifier for round-trip date/time pattern
						//var dateOh = DateTime.Parse(dateString, null, System.Globalization.DateTimeStyles.RoundtripKind);
                        DateTime dateOh;
                        if (DateTime.TryParseExact(dateString, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateOh))
                        {
                            //Console.WriteLine("Date: " + dateValue.ToString());
                        }
                        else
                        {
                           // Console.WriteLine("Invalid date format. Using default date: " + defaultDate.ToString());
                            dateOh = dateTime;
                        }

                        // Ensure the DateTime object is in UTC before saving to PostgreSQL
                        if (dateOh.Kind != DateTimeKind.Utc)
                        {
                            dateOh = dateOh.ToUniversalTime();
                        }
						var customFormattedDate = ConvertToPostgresDateTime(dateTime);// dateOh.ToString("yyyy-MM-ddTHH:mm:sszzz");


                        var timeString = row.GetCell(10) != null ? row.GetCell(10)?.ToString() : dateTime.ToString("HHmmss");
						if (string.IsNullOrEmpty(timeString))
						{

						}

                        var excelData = new TransactionDto
						{
							AgentAccountId = Convert.ToInt64(row.GetCell(0)?.ToString()),
							MsgType = row.GetCell(1)?.ToString() ?? string.Empty,
							CardNo = row.GetCell(2)?.ToString() ?? string.Empty,
							ProcCode = row.GetCell(3)?.ToString() ?? string.Empty,
							Balance = row.GetCell(4)?.ToString() ?? string.Empty,
							Amt = GetAmtInKobo(row.GetCell(5)?.ToString() ?? string.Empty),
							TransmissionDateTime = ConvertToPostgresDateTime(dateOh),
							TransactionType = row.GetCell(7)?.ToString() ?? string.Empty,
							CustomerNum = row.GetCell(8)?.ToString() ?? string.Empty,
							Stan = row.GetCell(9)?.ToString() ?? string.Empty,
							LocalTime = timeString,
							//LocalTime = Convert.ToDateTime(row.GetCell(10)?.ToString()).ToString("HHmmss"),
							LocalDate = row.GetCell(11)?.ToString(),
							// = row.GetCell(11)?.ToString(),
							MerchType = row.GetCell(12)?.ToString() ?? string.Empty, 
							PosEntryMode = row.GetCell(13)?.ToString() ?? string.Empty,
							CardSequenceNo = row.GetCell(14)?.ToString() ?? string.Empty,
							PosConditionCode = row.GetCell(15)?.ToString() ?? string.Empty,
							PosPinCaptureCode = row.GetCell(16)?.ToString() ?? string.Empty,
							Surcharge = row.GetCell(17)?.ToString() ?? string.Empty,
							AcqInstId = row.GetCell(18)?.ToString() ?? string.Empty,
							AcquirerId = row.GetCell(19)?.ToString() ?? string.Empty,
							FwdInstId = row.GetCell(20)?.ToString() ?? string.Empty,
							RetRefNo = row.GetCell(21)?.ToString() ?? string.Empty,
							Track2 = row.GetCell(22)?.ToString() ?? string.Empty,
							ServiceRestrictionCode = row.GetCell(23)?.ToString() ?? string.Empty,
							TerminalId = row.GetCell(24)?.ToString() ?? string.Empty,
							MerchantName = row.GetCell(25)?.ToString() ?? string.Empty,
							MerchantLoc = row.GetCell(26)?.ToString() ?? string.Empty,
							MerchantAddress = row.GetCell(27)?.ToString() ?? string.Empty,
							MerchantExtId = row.GetCell(28)?.ToString() ?? string.Empty,
							StatusCode = row.GetCell(29)?.ToString() ?? string.Empty,
							ExpDate = row.GetCell(30)?.ToString() ?? string.Empty,
							CurrencyCode = row.GetCell(31)?.ToString() ?? string.Empty,
							PinData = row.GetCell(32)?.ToString() ?? string.Empty,
							IccData = row.GetCell(33)?.ToString() ?? string.Empty,
							MsgReasonCode = row.GetCell(34)?.ToString() ?? string.Empty,
							PosDataCode = row.GetCell(35)?.ToString() ?? string.Empty,
							ResponseCode = row.GetCell(36)?.ToString() ?? string.Empty,
							AuthNum = row.GetCell(37)?.ToString() ?? string.Empty,
							Reversed = row.GetCell(38)?.ToString() ?? string.Empty,
							Completed = row.GetCell(39)?.ToString() ?? string.Empty,
							CreatedOn = row.GetCell(40)?.ToString() ?? string.Empty,
							MaskedPan = row.GetCell(41)?.ToString() ?? string.Empty,
							CardHolderName = row.GetCell(42)?.ToString() ?? string.Empty,
							CardTypeName = row.GetCell(43)?.ToString() ?? string.Empty,
							AccountType = row.GetCell(44)?.ToString() ?? string.Empty,
							AuthenticationMethod = row.GetCell(45)?.ToString() ?? string.Empty,
							Notified = row.GetCell(46)?.ToString() ?? string.Empty,
							Latency = row.GetCell(47)?.ToString() ?? string.Empty,
							TotalSales = row.GetCell(48)?.ToString() ?? string.Empty,
							SalesAmt = row.GetCell(49)?.ToString() ?? string.Empty,
							SalesId = row.GetCell(50)?.ToString() ?? string.Empty,
							UserName = row.GetCell(51)?.ToString() ?? string.Empty,
							ResponseMessage = row.GetCell(52)?.ToString() ?? string.Empty,
							PosResponseCode = row.GetCell(53)?.ToString() ?? string.Empty,
							PosResponseMessage = row.GetCell(54)?.ToString() ?? string.Empty,
							ProcessorResponseCode = row.GetCell(55)?.ToString() ?? string.Empty,
							OverallStatusCode = row.GetCell(56)?.ToString() ?? string.Empty,
							ReceiptPrinted = Convert.ToBoolean(row.GetCell(57)?.ToString()),
							AppChannel = row.GetCell(58)?.ToString() ?? string.Empty,
							Id = Convert.ToInt64(row.GetCell(59)?.ToString())
                        };

						excelDataList.Add(excelData);
					}
				}
			}

			return await Task.FromResult(excelDataList);
		}

		private static decimal GetAmtInKobo(string amt)
		{
			if(string.IsNullOrEmpty(amt)) return 0;
			return Convert.ToDecimal(amt) * 100;
		}

        public static DateTime ConvertToPostgresDateTime(DateTime dateTime)
        {
            // Assuming dateTime is in UTC and you want to keep it that way
            Instant instant = Instant.FromDateTimeUtc(dateTime);
            var localDateTime = instant.InUtc().ToDateTimeUtc();

            return localDateTime;
        }
    }

	//public record TransactionDto(long Id, long AgentAccountId, string MsgType, string CardNo, string ProcCode, string Balance, decimal Amt, DateTime TransmissionDateTime,
	//							string TransactionType, string CustomerNum, string Stan, string LocalTime, string LocalDate, string MerchType, string PosEntryMode,
	//							string CardSequenceNo, string PosConditionCode, string PosPinCaptureCode, string Surcharge, string AcqInstId, string AcquirerId,
	//							string FwdInstId, string RetRefNo, string Track2, string ServiceRestrictionCode, string TerminalId, string MerchantId, string MerchantName,
	//							string MerchantLoc, string MerchantAddress, string MerchantExtId, string StatusCode, string ExpDate, string CurrencyCode, string PinData, string IccData,
	//						   string MsgReasonCode, string PosDataCode, string ResponseCode, string AuthNum, string Reversed, string Completed, string CreatedOn, string MaskedPan,
	//						   string CardHolderName, string CardTypeName, string AccountType, string AuthenticationMethod, string Notified, string Latency, string TotalSales, string SalesAmt,
	//						   string SalesId, string UserName, string ResponseMessage, string PosResponseCode, string PosResponseMessage, string ProcessorResponseCode, string OverallStatusCode,
	//						   bool ReceiptPrinted, string AppChannel);

	public class TransactionDto
	{
		public long Id { get; set; }
        public long AgentAccountId {  get; set; }

        public string MsgType { get; set; }
		public string CardNo { get; set; }
		public string ProcCode { get; set; }
		public string Balance { get; set; }
		public decimal Amt { get; set; }
		public DateTime TransmissionDateTime { get; set; }
		public string TransactionType { get; set; }
		public string CustomerNum { get; set; }
		public string Stan { get; set; }
		public string LocalTime { get; set; }
		public string LocalDate { get; set; }
		public string MerchType { get; set; }
		public string PosEntryMode { get; set; }
		public string CardSequenceNo { get; set; }
		public string PosConditionCode { get; set; }
		public string PosPinCaptureCode { get; set; }
		public string Surcharge { get; set; }
		public string AcqInstId { get; set; }
		public string AcquirerId { get; set; }
		public string FwdInstId { get; set; }
		public string RetRefNo { get; set; }
		public string Track2 { get; set; }
		public string ServiceRestrictionCode { get; set; }
		public string TerminalId { get; set; }
		public string MerchantName { get; set; }
		public string MerchantLoc { get; set; }
		public string MerchantAddress { get; set; }
		public string MerchantExtId { get; set; }
		public string StatusCode { get; set; }
		public object ExpDate { get; set; }
		public string CurrencyCode { get; set; }
		public string PinData { get; set; }
		public string IccData { get; set; }
		public string MsgReasonCode { get; set; }
		public string PosDataCode { get; set; }
		public string ResponseCode { get; set; }
		public string AuthNum { get; set; }
		public string Reversed { get; set; }
		public string Completed { get; set; }
		public string CreatedOn { get; set; }
		public string MaskedPan { get; set; }
		public string CardHolderName { get; set; }
		public string CardTypeName { get; set; }
		public string AccountType { get; set; }
		public string AuthenticationMethod { get; set; }
		public string Notified { get; set; }
		public string Latency { get; set; }
		public string TotalSales { get; set; }
		public string SalesAmt { get; set; }
		public string SalesId { get; set; }
		public string UserName { get; set; }
		public string ResponseMessage { get; set; }
		public string PosResponseCode { get; set; }
		public string PosResponseMessage { get; set; }
		public string ProcessorResponseCode { get; set; }
		public string OverallStatusCode { get; set; }
		public bool ReceiptPrinted { get; set; }
		public string AppChannel { get; set; }
	}
}
