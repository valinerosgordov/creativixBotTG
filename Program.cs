using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

partial class Program
{
    private static readonly string BotToken = "8051109468:AAHamj381V59fley_3yhSukuuzFW4W0nvPY";
    private static readonly ITelegramBotClient botClient = new TelegramBotClient(BotToken);

    private static readonly Dictionary<long, List<string>> userResponses = new();
    private static readonly Dictionary<string, long> LocationChatIds = new()
    {
        { "Грим Москва 1", -1002397750170 },
        { "Грим Москва 0", -4633844539 },
        { "Грим Москва 3", -4617470799 },
        { "Авиапарк", -1002307194245 },
        { "Фантазия", -4783982885 },
        { "МК Москва", -4654198477 },
        { "Аир Парк", -4711552893 },
        { "Луномосик", -1002495223375 },
        { "Мультпарк", -1002413575599 }
    };

    private static readonly object fileLock = new();

    static async Task Main(string[] args)
    {
        Console.WriteLine("✅ Бот запущен...");
        using var cts = new CancellationTokenSource();

        botClient.StartReceiving(
            HandleUpdateAsync,
            HandleErrorAsync,
            new ReceiverOptions { AllowedUpdates = { } },
            cancellationToken: cts.Token
        );

        await Task.Delay(-1);
    }

    static async Task HandleUpdateAsync(ITelegramBotClient bot, Update update, CancellationToken token)
    {
        if (update.CallbackQuery is { } callback && callback.Message?.Chat?.Id is long callbackId)
        {
            string? location = callback.Data;
            if (string.IsNullOrEmpty(location)) return;

            if (!userResponses.ContainsKey(callbackId))
                userResponses[callbackId] = new List<string>();

            userResponses[callbackId].Add(location);
            await bot.AnswerCallbackQuery(callback.Id);
            await ProcessUserResponses(callbackId);
            return;
        }

        if (update.Message is not { } message || message.Text is not { } messageText || message.Chat?.Id is not long chatId)
            return;

        var text = messageText.Trim();
        Console.WriteLine($"📩 {chatId}: {text}");

        if (text.ToLower() is "/reset" or "сброс" or "🌸 отправить ещё один отчёт")
        {
            userResponses[chatId] = new List<string>();
            await bot.SendMessage(chatId, "🔄 Всё сбросили 🌸 Введи своё имя, милашка:");
            return;
        }

        if (!userResponses.ContainsKey(chatId))
{
    userResponses[chatId] = new List<string>();
    await bot.SendMessage(chatId, "🌼 Привет, солнышко! Давай начнём с твоего имени 💕");
    return;
}


        userResponses[chatId].Add(text);
        await ProcessUserResponses(chatId);
    }

    static async Task ProcessUserResponses(long chatId)
    {
        if (!userResponses.TryGetValue(chatId, out var responses)) return;

        string[] prompts =
        {
            "📍 Выбери свое местоположение:",
            "📊 Введи сумму выручки:",
            "💵 Введи сумму наличных денег:",
            "🏦 Введи сумму безналичного расчета:",
            "🔄 Введи сумму по СБП:",
            "🔁 Введи сумму переводов:",
            "💰 Введи сумму размена:",
            "📝 Добавь комментарий (если нужно), или напиши '-' чтобы пропустить:"
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

        if (responses.Count == 9)
        {
            if (!decimal.TryParse(responses[2], out var revenue) ||
                !decimal.TryParse(responses[3], out var cash) ||
                !decimal.TryParse(responses[4], out var cashless) ||
                !decimal.TryParse(responses[5], out var sbp) ||
                !decimal.TryParse(responses[6], out var transfers) ||
                !decimal.TryParse(responses[7], out var exchange))
            {
                userResponses[chatId].RemoveAt(responses.Count - 1);
                await botClient.SendMessage(chatId, "❌ Ошибка: Введи корректные числа!");
                return;
            }

            decimal income = revenue - (cash + cashless + sbp + transfers + exchange);
            await botClient.SendMessage(chatId, "🌷 Спасибо, котик! Сейчас всё красиво оформим... 💫");
            await SaveToExcelAndSendReport(chatId, responses, revenue, cash, cashless, sbp, transfers, exchange, income);
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

    static async Task SaveToExcelAndSendReport(long chatId, List<string> responses, decimal revenue, decimal cash, decimal cashless, decimal sbp, decimal transfers, decimal exchange, decimal income)
    {
        string userName = responses[0];
        string location = responses[1];
        string fileName = $"{string.Concat(userName.Split(Path.GetInvalidFileNameChars()))}.xlsx";

        lock (fileLock)
        {
            using var package = new ExcelPackage(new FileInfo(fileName));
            var sheet = package.Workbook.Worksheets.Count == 0
                ? package.Workbook.Worksheets.Add("Daily Report")
                : package.Workbook.Worksheets[0];

            int row = (sheet.Dimension?.Rows ?? 0) + 1;
            sheet.Cells[row, 1].Value = DateTime.Now.ToString("yyyy-MM-dd");
            sheet.Cells[row, 2].Value = userName;
            sheet.Cells[row, 3].Value = location;
            sheet.Cells[row, 4].Value = revenue;
            sheet.Cells[row, 5].Value = cash;
            sheet.Cells[row, 6].Value = cashless;
            sheet.Cells[row, 7].Value = sbp;
            sheet.Cells[row, 8].Value = transfers;
            sheet.Cells[row, 9].Value = exchange;
            sheet.Cells[row, 10].Value = income;
            sheet.Cells[row, 11].Value = string.IsNullOrWhiteSpace(responses[8]) ? "-" : responses[8];

            if (sheet.Dimension?.Rows == 1)
            {
                sheet.Cells[1, 1].Value = "Дата";
                sheet.Cells[1, 2].Value = "Имя";
                sheet.Cells[1, 3].Value = "Локация";
                sheet.Cells[1, 4].Value = "Выручка";
                sheet.Cells[1, 5].Value = "Наличные";
                sheet.Cells[1, 6].Value = "Безнал";
                sheet.Cells[1, 7].Value = "СБП";
                sheet.Cells[1, 8].Value = "Переводы";
                sheet.Cells[1, 9].Value = "Размен";
                sheet.Cells[1, 10].Value = "Чистый доход";
                sheet.Cells[1, 11].Value = "Комментарий";
            }

            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            package.Save();
        }

        if (!string.IsNullOrEmpty(location) && LocationChatIds.TryGetValue(location, out var locChatId))
        {
            string msg = $"🌸 *Твой милый отчёт готов, умничка!* 🌸📅 *Дата:* {DateTime.Now:yyyy-MM-dd}👩 *Имя:* {userName}📍 *Локация:* {location}💖 *Выручка:* {revenue} руб.💵 *Наличные:* {cash} руб.🏦 *Безнал:* {cashless} руб.💳 *СБП:* {sbp} руб.🔁 *Переводы:* {transfers} руб.💰 *Размен:* {exchange} руб.📝 *Комментарий:* {(string.IsNullOrWhiteSpace(responses[8]) ? "-" : responses[8])}";
            await botClient.SendMessage(locChatId, msg, parseMode: ParseMode.Markdown);
        }

        var replyMarkup = new ReplyKeyboardMarkup(new[]
        {
            new KeyboardButton("🌸 Отправить ещё один отчёт")
        })
        {
            ResizeKeyboard = true,
            OneTimeKeyboard = true
        };

        await botClient.SendMessage(chatId, "✅ Отчет отправлен! Хочешь заполнить ещё один? 😊", replyMarkup: replyMarkup);
    }

    static Task HandleErrorAsync(ITelegramBotClient bot, Exception exception, CancellationToken token)
    {
        Console.WriteLine($"❌ Ошибка: {exception.Message}");
        return Task.CompletedTask;
    }
}
