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
    private static readonly string BotToken = "YOUR_BOT_TOKEN_HERE";
    private static readonly long AdminChatId = -1002453582408;
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

    static async Task Main(string[] args)
    {
        await RunBot();
    }

    static async Task RunBot()
    {
        var bot = new TelegramBotClient("YOUR_BOT_TOKEN_HERE");
        Console.WriteLine("Bot is running...");
        await Task.Delay(-1); // Keeps the bot running
    }
    
    private static readonly Dictionary<string, long> LocationChatIds = new()
    {
        { "Грим Москва 1", -1002397750170 },//done
        { "Грим Москва 0", -4633844539 },//done
        { "Грим Москва 3", -4617470799 }, //done
        { "Авиапарк", -1002307194245 },//done
        { "Фантазия", -4783982885 },//done
        { "МК Москва", -4654198477 },//done
        { "Аир Парк", -4711552893 }, //done
        { "Луномосик", -1002495223375 },//done
        { "Мультпарк", -1002413575599 } //done
    };

    static async Task ProcessUserResponses(long chatId)
    {
        var responses = userResponses[chatId];
        string[] prompts =
        {
            "📍 Выбери свое местоположение:",
            "📊 Введи сумму выручки:",
            "💵 Введи сумму наличных денег:",
            "🏦 Введи сумму безналичного расчета:",
            "🔄 Введи сумму переводов:",
            "💰 Введи сумму размена:" 
        };
        
        if (responses.Count == 1)
        {
            await SendLocationSelection(chatId);
            return;
        }
        
        if (responses.Count > 1 && responses.Count <= prompts.Length)
        {
            await botClient.SendMessage(chatId, prompts[responses.Count - 1]);
            return;
        }
        
        if (responses.Count == 6)
        {
            if (!decimal.TryParse(responses[2], out decimal revenue) ||
                !decimal.TryParse(responses[3], out decimal cash) ||
                !decimal.TryParse(responses[4], out decimal cashless) ||
                !decimal.TryParse(responses[5], out decimal transfers) ||
                !decimal.TryParse(responses[6], out decimal exchange))
            {
                userResponses[chatId].RemoveAt(responses.Count - 1);
                await botClient.SendMessage(chatId, "❌ Ошибка: Введи корректные числа!");
                return;
            }
            
            decimal income = revenue - (cash + cashless + transfers + exchange);
            await SaveToExcelAndSendReport(chatId, responses, revenue, cash, cashless, transfers, exchange, income);
            userResponses.Remove(chatId);
        }
    }

    static async Task SendLocationSelection(long chatId)
    {
        var keyboard = new InlineKeyboardMarkup(new[]
        {
            new[] { InlineKeyboardButton.WithCallbackData("Грим Москва 1"), InlineKeyboardButton.WithCallbackData("Грим Москва 0"), InlineKeyboardButton.WithCallbackData("Грим Москва 3") },
            new[] { InlineKeyboardButton.WithCallbackData("Мультпарк"), InlineKeyboardButton.WithCallbackData("Авиапарк"), InlineKeyboardButton.WithCallbackData("Фантазия") },
            new[] { InlineKeyboardButton.WithCallbackData("МК Москва"), InlineKeyboardButton.WithCallbackData("Аир Парк"), InlineKeyboardButton.WithCallbackData("Луномосик") }
        });
        await botClient.SendMessage(chatId, "📍 Выбери свое местоположение:", replyMarkup: keyboard);
    }

    static async Task SaveToExcelAndSendReport(long chatId, List<string> responses, decimal revenue, decimal cash, decimal cashless, decimal transfers, decimal exchange, decimal income)
    {
        string userName = responses[0];
        string location = responses[1];
        string safeUserName = string.Concat(userName.Split(Path.GetInvalidFileNameChars()));
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
                }
                else
                {
                    worksheet = package.Workbook.Worksheets[0];
                }
                
                int row = (worksheet.Dimension?.Rows ?? 1) + 1;
                worksheet.Cells[row, 1].Value = DateTime.Now.ToString("yyyy-MM-dd");
                worksheet.Cells[row, 2].Value = userName;
                worksheet.Cells[row, 3].Value = location;
                worksheet.Cells[row, 4].Value = revenue;
                worksheet.Cells[row, 5].Value = cash;
                worksheet.Cells[row, 6].Value = cashless;
                worksheet.Cells[row, 7].Value = transfers;
                worksheet.Cells[row, 8].Value = exchange;
                worksheet.Cells[row, 9].Value = income;
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                package.Save();
            }
        }
        
        if (LocationChatIds.TryGetValue(location, out long locationChatId))
        {
            string reportMessage = $"📊 *Отчет за день:*\n📅 {DateTime.Now:yyyy-MM-dd}\n👤 *Имя:* {userName}\n📍 *Локация:* {location}\n📊 *Выручка:* {revenue} руб.\n💵 *Наличные:* {cash} руб.\n🏦 *Безнал:* {cashless} руб.\n🔄 *Переводы:* {transfers} руб.\n💰 *Размен:* {exchange} руб.\n📈 *Чистый доход:* {income} руб.";
            await botClient.SendMessage(locationChatId, reportMessage, parseMode: ParseMode.Markdown);
        }
    }
}
