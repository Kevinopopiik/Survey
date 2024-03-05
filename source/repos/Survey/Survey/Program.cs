using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Data.SqlClient;
using Google.Apis.Util.Store;

namespace SurveyResponseApp
{
    class Program
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        static string SpreadsheetId = "19pvbCLDVcjungsUA5CbjzFzbeRNiG7lVv8lqUD1sXVU"; // 替換為您的Google試算表ID
        static string Range = "SurveyResponses!B2:I";// 替換為您的範圍
        static string connectionString = "Server=.\\SQLEXPRESS01;Database=SurveyDB;Trusted_Connection=True;"; // 替換為您的SQL Server連接字符串

        static void Main(string[] args)
        {
            try
            {
                UserCredential credential;

                using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                SpreadsheetsResource.ValuesResource.GetRequest request =
                        service.Spreadsheets.Values.Get(SpreadsheetId, Range);

                ValueRange response = request.Execute();
                IList<IList<Object>> values = response.Values;

                if (values != null && values.Count > 0)
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        foreach (var row in values)
                        {
                            if (row.Count < 7)
                            {
                                // 如果數據不足，可以選擇跳過此行或者添加日誌
                                Console.WriteLine("Data row is missing information. Skipping this row.");
                                continue; // 跳過這一行
                            }

                            // 注意這裡的 SQL 語句已經移除了對 Email 的引用
                            string sql = @"
                INSERT INTO SurveyResponses 
                (Name, Gender, AgeRange, OverallSatisfaction, ProfessionalismSatisfaction, WaitingTimeSatisfaction, ClarityOfResultsSatisfaction, AdditionalFeedback) 
                VALUES 
                (@Name, @Gender, @AgeRange, @OverallSatisfaction, @ProfessionalismSatisfaction, @WaitingTimeSatisfaction, @ClarityOfResultsSatisfaction, @AdditionalFeedback)";

                            using (SqlCommand command = new SqlCommand(sql, connection))
                            {
                                // 添加參數到 SqlCommand，已經移除了對 Email 的引用
                                command.Parameters.AddWithValue("@Name", row[0]?.ToString() ?? "");
                                command.Parameters.AddWithValue("@Gender", row[1]?.ToString() ?? "");
                                command.Parameters.AddWithValue("@AgeRange", row[2]?.ToString() ?? "");
                                command.Parameters.AddWithValue("@OverallSatisfaction", row[3]?.ToString() ?? "");
                                command.Parameters.AddWithValue("@ProfessionalismSatisfaction", row[4]?.ToString() ?? "");
                                command.Parameters.AddWithValue("@WaitingTimeSatisfaction", row[5]?.ToString() ?? "");
                                command.Parameters.AddWithValue("@ClarityOfResultsSatisfaction", row[6]?.ToString() ?? "");
                                var additionalFeedback = row.Count > 7 ? row[7]?.ToString() : null;
                                command.Parameters.AddWithValue("@AdditionalFeedback", string.IsNullOrEmpty(additionalFeedback) ? (object)DBNull.Value : additionalFeedback);
                                command.ExecuteNonQuery();
                            }
                        }
                    }
                    Console.WriteLine("回答已成功儲存到資料庫中。");
                }
                else
                {
                    Console.WriteLine("No data found.");
                }
            }
            catch (Google.GoogleApiException gex)
            {
                Console.WriteLine("Google API Exception: " + gex.Message);
            }
            catch (SqlException sqlex)
            {
                Console.WriteLine("SQL Exception: " + sqlex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("General Exception: " + ex.Message);
            }
        }
    }
}
