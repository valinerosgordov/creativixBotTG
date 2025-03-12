using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using Hangfire;
using Hangfire.MemoryStorage;

class Program
{
    private static readonly string BotToken = "8051109468:AAHamj381V59fley_3yhSukuuzFW4W0nvPY";
    private static readonly long AdminChatId = -1002453582408; // ID чата для отчетов
    private static TelegramBotClient botClient = new TelegramBotClient(BotToken);
    private static readonly Dictionary<long, List<string>> userResponses = new();
    private static readonly HashSet<long> submittedUsers = new();
    private static readonly List<long> allUsers = new();
    private static readonly object fileLock = new();
    
    private static readonly Dictionary<long, int> userLastMessageIds = new();
    private static readonly Dictionary<string, string> SalaryTypes = new()
    {
        { "Грим Москва 1", "moskvarium" },
        { "Грим Москва 0", "moskvarium" },
        { "Грим Москва 3", "moskvarium" },
        { "Авиапарк", "base" },
        { "Фантазия", "base" },
        { "МК Москва", "base" },
        { "Аир Парк", "base" },
        { "Луномосик", "base" },
        { "Мультпарк", "multpark" }
    };

    static async Task Main()
    {
        var botInfo = await botClient.GetMe();
        Console.WriteLine($"✅ Бот запущен! Имя: {botInfo.FirstName}, Логин: @{botInfo.Username}");

        await botClient.DeleteWebhook();
        Console.WriteLine("🔄 Webhook удален. Polling включен.");

        GlobalConfiguration.Configuration.UseMemoryStorage();
        using (var server = new BackgroundJobServer())
        {
            RecurringJob.AddOrUpdate("morning_greeting", () => SendMorningGreeting(), Cron.Daily(8, 0));
            RecurringJob.AddOrUpdate("reminder_12pm", () => SendReminderMessages(), Cron.Daily(12, 0));
            RecurringJob.AddOrUpdate("reminder_6pm", () => SendReminderMessages(), Cron.Daily(18, 0));
            RecurringJob.AddOrUpdate("final_reminder", () => SendFinalReminder(), Cron.Daily(22, 0));

            var cts = new CancellationTokenSource();
            botClient.StartReceiving(
                HandleUpdateAsync,
                HandleErrorAsync,
                new ReceiverOptions { AllowedUpdates = Array.Empty<UpdateType>() },
                cts.Token
            );

            await Task.Delay(-1);
        }
    }

    private static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
    {
        try
        {
            if (update.Type != UpdateType.Message || update.Message?.Text == null)
                return;

            var message = update.Message;
            long chatId = message.Chat.Id;
            Console.WriteLine($"📩 Получено сообщение: \"{message.Text}\" от {chatId}");

            if (!allUsers.Contains(chatId)) allUsers.Add(chatId);

            if (!userResponses.ContainsKey(chatId))
            {
                userResponses[chatId] = new List<string>();
                await botClient.SendMessage(chatId, "🌸 Как тебя зовут?");
            }
            else
            {
                userResponses[chatId].Add(message.Text);
                await ProcessUserResponses(chatId);
            }
            if (message.Text == "/start")
            {
                userResponses[chatId] = new List<string>(); // Очистка состояния
                await botClient.SendMessage(chatId, "🌸 Как тебя зовут?");
                return;
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка в HandleUpdateAsync: {ex.Message}");
        }

        
    }

    private static async Task SendLocationSelection(long chatId)
    {
        var keyboard = new InlineKeyboardMarkup(new[]
        {
            new[]
            {
                InlineKeyboardButton.WithCallbackData("Грим Москва 1"),
                InlineKeyboardButton.WithCallbackData("Грим Москва 0"),
                InlineKeyboardButton.WithCallbackData("Грим Москва 3")
            },
            new[]
            {
                InlineKeyboardButton.WithCallbackData("Мультпарк"),
                InlineKeyboardButton.WithCallbackData("Авиапарк"),
                InlineKeyboardButton.WithCallbackData("Фантазия")
            },
            new[]
            {
                InlineKeyboardButton.WithCallbackData("МК Москва"),
                InlineKeyboardButton.WithCallbackData("Аир Парк"),
                InlineKeyboardButton.WithCallbackData("Луномосик")
            }
        });

        await botClient.SendTextMessageAsync(chatId, "📍 Выбери свое местоположение:", replyMarkup: keyboard);
    }

    private static async Task DeleteAndSendNewMessage(long chatId, string message)
{
    // Delete the user's previous message if it exists
    if (userLastMessageIds.ContainsKey(chatId))
    {
        try
        {
            await botClient.DeleteMessageAsync(chatId, userLastMessageIds[chatId]);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠ Ошибка при удалении сообщения: {ex.Message}");
        }
    }

    // Send the new message and store its ID
    var sentMessage = await botClient.SendTextMessageAsync(chatId, message);
    userLastMessageIds[chatId] = sentMessage.MessageId;
}

        
        


    private static async Task ProcessUserResponses(long chatId)
{
    var responses = userResponses[chatId];

    switch (responses.Count)
    {
        case 1:
            await SendLocationSelection(chatId); // Location selection with buttons only
            return;
        case 2:
            await DeleteAndSendNewMessage(chatId, "📊 Введи сумму выручки:");
            return;
        case 3:
            await DeleteAndSendNewMessage(chatId, "💵 Введи сумму наличных денег:");
            return;
        case 4:
            await DeleteAndSendNewMessage(chatId, "🔄 Введи сумму переводов:");
            return;
        case 5:
            await DeleteAndSendNewMessage(chatId, "💸 Введи сумму дополнительных трат:");
            return;
        case 6:
            if (!decimal.TryParse(responses[2], out decimal revenue) ||
                !decimal.TryParse(responses[3], out decimal cash) ||
                !decimal.TryParse(responses[4], out decimal transfers) ||
                !decimal.TryParse(responses[5], out decimal expenses))
            {
                userResponses[chatId].RemoveAt(responses.Count - 1); // Remove incorrect input
                await DeleteAndSendNewMessage(chatId, "❌ Ошибка: Введи корректные числа для выручки, наличных, переводов и доп. трат!");
                return;
            }

            decimal income = revenue - (cash + transfers + expenses);
            decimal salary = CalculateSalary(responses[1], income);

            await SaveToExcel(chatId, responses, revenue, income, salary);

            string reportMessage = $"📊 *Отчет за день:*\n" +
                                $"👤 *Имя:* {responses[0]}\n" +
                                $"📍 *Локация:* {responses[1]}\n" +
                                $"📊 *Выручка:* {revenue} руб.\n" +
                                $"💵 *Наличные:* {cash} руб.\n" +
                                $"🔄 *Переводы:* {transfers} руб.\n" +
                                $"💸 *Доп. Траты:* {expenses} руб.";

            await botClient.SendTextMessageAsync(AdminChatId, reportMessage, parseMode: ParseMode.Markdown);
            await botClient.SendTextMessageAsync(chatId, "✨ Спасибо за отчет! 💖\n📊 Он отправлен в общий чат.", parseMode: ParseMode.Markdown);

            userResponses.Remove(chatId); // Clear user state
            return;
    }
}

private static async Task HandleCallbackQueryAsync(ITelegramBotClient botClient, CallbackQuery callbackQuery)
{
    long chatId = callbackQuery.Message.Chat.Id;
    string location = callbackQuery.Data;

    if (!SalaryTypes.ContainsKey(location))
    {
        await botClient.AnswerCallbackQueryAsync(callbackQuery.Id, "❌ Некорректный выбор!");
        return;
    }

    // Save user's location selection
    userResponses[chatId].Add(location);
    
    // Acknowledge button press
    await botClient.AnswerCallbackQueryAsync(callbackQuery.Id, "✅ Локация выбрана!");
    
    // Delete inline keyboard message
    await botClient.DeleteMessageAsync(chatId, callbackQuery.Message.MessageId);

    // Move to the next question (Revenue input)
    await DeleteAndSendNewMessage(chatId, "📊 Введи сумму выручки:");
}



    private static decimal CalculateSalary(string location, decimal income)
    {
        string type = SalaryTypes.ContainsKey(location) ? SalaryTypes[location] : "base";

        decimal salary = type switch
        {
            "base" => 2000 + (income * 0.1m),
            "multpark" => (income <= 10000) ? 2500 + (income * 0.1m) : 2500 + (income * 0.3m),
            "moskvarium" => (income <= 40000) ? 2000 + (income * 0.1m) : 2000 + (income * 0.15m),
            _ => 2000
        };

        return salary;
    }


    public static void SendMorningGreeting()
    {
        Task.Run(async () =>
        {
            Console.WriteLine("🌞 Доброе утро! Отправка утреннего сообщения...");
            await botClient.SendMessage(AdminChatId, "🌞 Доброе утро! Пусть день будет лёгким и продуктивным! 💖");
        }).Wait();
    }

public static async Task SendReminderMessages()
    {
        Console.WriteLine("🔔 Отправка напоминания...");

        foreach (var userId in allUsers.Except(submittedUsers))
        {
            await botClient.SendMessage(userId, "⏳ Напоминание! 💖 Пожалуйста, заполни форму сегодня!");
        }
    }



    public static void SendFinalReminder()
    {
        Task.Run(async () =>
        {
            Console.WriteLine("📢 Отправка финального напоминания...");
            foreach (var userId in allUsers.Except(submittedUsers))
            {
                await botClient.SendMessage(userId, "🌙 Вечерний чек-ин! 💕 Пожалуйста, заполни форму, если не успела!");
            }
        }).Wait();
    }

        private static async Task SaveToExcel(long chatId, List<string> responses, decimal revenue, decimal income, decimal salary)
{
    await Task.Run(() =>
    {
        try
        {
            // Get user's name from responses
            string userName = responses[0];

            // Sanitize file name (remove invalid characters)
            string safeUserName = string.Concat(userName.Split(Path.GetInvalidFileNameChars()));

            // Create filename with user's name
            string userFile = $"{safeUserName}.xlsx"; 

            lock (fileLock)
            {
                FileInfo file = new FileInfo(userFile);
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet;

                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        worksheet = package.Workbook.Worksheets.Add("Daily Report");

                        // Set column headers
                        worksheet.Cells[1, 1].Value = "📅 Дата";
                        worksheet.Cells[1, 2].Value = "👤 Имя";
                        worksheet.Cells[1, 3].Value = "📍 Локация";
                        worksheet.Cells[1, 4].Value = "📊 Выручка (руб)";
                        worksheet.Cells[1, 5].Value = "💵 Наличные (руб)";
                        worksheet.Cells[1, 6].Value = "🔄 Переводы (руб)";
                        worksheet.Cells[1, 7].Value = "💸 Доп. Траты (руб)";
                        worksheet.Cells[1, 8].Value = "📈 Чистый доход (руб)";
                        worksheet.Cells[1, 9].Value = "💰 Зарплата (руб)";

                        // Apply formatting
                        using (var range = worksheet.Cells["A1:I1"])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }
                    }
                    else
                    {
                        worksheet = package.Workbook.Worksheets[0];
                    }

                    // Find the next empty row
                    int row = (worksheet.Dimension?.Rows ?? 1) + 1;

                    // Fill in data
                    worksheet.Cells[row, 1].Value = DateTime.Now.ToString("yyyy-MM-dd"); // Date
                    worksheet.Cells[row, 2].Value = userName; // Name
                    worksheet.Cells[row, 3].Value = responses[1]; // Location
                    worksheet.Cells[row, 4].Value = revenue; // Revenue
                    worksheet.Cells[row, 5].Value = decimal.Parse(responses[3]); // Cash
                    worksheet.Cells[row, 6].Value = decimal.Parse(responses[4]); // Transfers
                    worksheet.Cells[row, 7].Value = decimal.Parse(responses[5]); // Expenses
                    worksheet.Cells[row, 8].Value = income; // Net Income
                    worksheet.Cells[row, 9].Value = salary; // Salary

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    package.Save();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Ошибка при сохранении Excel: {ex.Message}");
        }
    });
}




    private static Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
    {
        Console.WriteLine($"❌ Ошибка бота: {exception.Message}");
        return Task.CompletedTask;
    }

}
