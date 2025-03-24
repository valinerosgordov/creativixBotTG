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

    private static readonly long adminId = 6947043193;
    private static readonly Dictionary<long, List<string>> userResponses = new();
    private static readonly Dictionary<string, long> LocationChatIds = new()
{
    { "Грим Мск (0/1/3)", -1002397750170 }, // объединённый ID
    { "МК Москва", -1002257906715 },
    { "Авиапарк", -1002307194245 },
    { "Фантазия", -1002475162608 },
    { "Аир Парк", -1002422847564 },
    { "Луномосик", -1002495223375 },
    { "Мультпарк", -1002413575599 }
};


    private static readonly object fileLock = new();

    private static async Task ReportError(Exception ex, long? chatId = null)
{
    string errorText = $"[{DateTime.Now}] ❌ Ошибка: {ex.Message}\n{ex.StackTrace}\n";
    System.IO.File.AppendAllText("error.log", errorText);

    try
    {
        string message = $"🚨 *Ошибка в боте:*\n`{ex.Message}`";
        if (chatId != null)
            message += $"\n👤 Chat ID: `{chatId}`";

        await botClient.SendMessage(adminId, message, parseMode: ParseMode.Markdown);
    }
    catch
    {
        // если даже отправка ошибки упала — молча
    }
}




    private static async Task TrySendMessage(long chatId, string message, IReplyMarkup? markup = null)
{
    try
    {
        await botClient.SendMessage(chatId, message, replyMarkup: markup);
    }
    catch (Exception ex)
    {
        string errorText = $"[{DateTime.Now}] ❌ Ошибка отправки сообщения: {ex.Message}\n{ex.StackTrace}";
        System.IO.File.AppendAllText("error.log", errorText + "\n");

        try
        {
            await botClient.SendMessage(adminId, $"❗️ *Ошибка отправки сообщения:*\n`{ex.Message}`", parseMode: ParseMode.Markdown);
        }
        catch
        {
            // если отправка администратору не удалась — тоже молчим
        }
    }
}


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
        try
        {
            if (update.CallbackQuery is { } callback && callback.Message?.Chat?.Id is long callbackId)
            {
                if (callback.Data == "redo")
                {
                    userResponses[callbackId] = new List<string>();
                    await bot.SendMessage(callbackId, "🔄 Давай начнём заново! Напиши своё имя, пожалуйста 😊");
                    return;
                }

                if (callback.Data == "submit")
                {
                    var responses = userResponses[callbackId];
                    if (responses.Count < 10) return;

                    if (!decimal.TryParse(responses[2], out var revenue) ||
                        !decimal.TryParse(responses[3], out var cash) ||
                        !decimal.TryParse(responses[4], out var cashless) ||
                        !decimal.TryParse(responses[5], out var sbp) ||
                        !decimal.TryParse(responses[6], out var transfers) ||
                        !decimal.TryParse(responses[7], out var extra) ||
                        !decimal.TryParse(responses[8], out var exchange))
                    {
                        await bot.SendMessage(callbackId, "❌ Ошибка при подтверждении: данные невалидны.");
                        return;
                    }

                    decimal income = revenue - (cash + cashless + sbp + transfers + extra + exchange);
                    await SaveToExcelAndSendReport(callbackId, responses, revenue, cash, cashless, sbp, transfers, extra, exchange, income);
                    userResponses.Remove(callbackId);
                    return;
                }

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

            if (text.ToLower() is "/reset" or "сброс")
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
        catch (Exception ex)
        {
            long? fallbackChatId = update.Message?.Chat?.Id ?? update.CallbackQuery?.Message?.Chat?.Id;
            await ReportError(ex, fallbackChatId);
        }
    }
    static async Task ProcessUserResponses(long chatId)
    {
        try
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
                "🧾 Введи сумму дополнительных трат:",
                "💰 Введи сумму размена:",
                "📝 Добавь комментарий (если нужно), или напиши любой символ, чтобы пропустить:"
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

            if (responses.Count == 10)
            {
                var confirmKeyboard = new InlineKeyboardMarkup(new[]
                {
                    new[]
                    {
                        InlineKeyboardButton.WithCallbackData("✅ Отправить отчёт", "submit"),
                        InlineKeyboardButton.WithCallbackData("🔄 Переделать", "redo")
                    }
                });

                await botClient.SendMessage(chatId, "✨ Всё готово! Хочешь отправить отчёт или переделать? 😊", replyMarkup: confirmKeyboard);
                return;
            }
        }
        catch (Exception ex)
        {
            await ReportError(ex, chatId);
        }
    }

    static async Task SendLocationSelection(long chatId)
    {
        try
        {
            var keyboard = new InlineKeyboardMarkup(new[]
            {
                new[] { InlineKeyboardButton.WithCallbackData("Грим Мск (0/1/3)") },
                new[] { InlineKeyboardButton.WithCallbackData("МК Москва"), InlineKeyboardButton.WithCallbackData("Авиапарк") },
                new[] { InlineKeyboardButton.WithCallbackData("Фантазия"), InlineKeyboardButton.WithCallbackData("Аир Парк") },
                new[] { InlineKeyboardButton.WithCallbackData("Луномосик"), InlineKeyboardButton.WithCallbackData("Мультпарк") }
            });

            await botClient.SendMessage(chatId, "📍 Выбери свое местоположение:", replyMarkup: keyboard);
        }
        catch (Exception ex)
        {
            await ReportError(ex, chatId);
        }
    }

    static async Task SaveToExcelAndSendReport(long chatId, List<string> responses, decimal revenue, decimal cash, decimal cashless, decimal sbp, decimal transfers, decimal extra, decimal exchange, decimal income)
    {
        try
        {
            string userName = responses[0];
            string location = responses[1];
            string comment = responses[9];
            string fileName = $"{DateTime.Now:yyyy-MM-dd}_{userName}.xlsx";

            lock (fileLock)
            {
                using var package = new ExcelPackage(new FileInfo(fileName));
                var sheet = package.Workbook.Worksheets.Count == 0
                    ? package.Workbook.Worksheets.Add("Daily Report")
                    : package.Workbook.Worksheets[0];

                int row = (sheet.Dimension?.Rows ?? 0);
                if (row == 0)
                {
                    sheet.Cells[1, 1].Value = "Дата";
                    sheet.Cells[1, 2].Value = "Имя";
                    sheet.Cells[1, 3].Value = "Локация";
                    sheet.Cells[1, 4].Value = "Выручка";
                    row = 2;
                }
                else row++;

                sheet.Cells[row, 1].Value = DateTime.Now.ToString("yyyy-MM-dd");
                sheet.Cells[row, 2].Value = userName;
                sheet.Cells[row, 3].Value = location;
                sheet.Cells[row, 4].Value = revenue;

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                package.Save();
            }

            if (!string.IsNullOrEmpty(location) && LocationChatIds.TryGetValue(location, out var locChatId))
            {
                string msg = $"🌸 *Твой милый отчёт готов, умничка!* 🌸\n" +
                             $"📅 *Дата:* {DateTime.Now:yyyy-MM-dd}\n" +
                             $"👩 *Имя:* {userName}\n" +
                             $"📍 *Локация:* {location}\n" +
                             $"💖 *Выручка:* {revenue} руб.\n" +
                             $"💵 *Наличные:* {cash} руб.\n" +
                             $"🏦 *Безнал:* {cashless} руб.\n" +
                             $"💳 *СБП:* {sbp} руб.\n" +
                             $"🔁 *Переводы:* {transfers} руб.\n" +
                             $"🧾 *Доп. траты:* {extra} руб.\n" +
                             $"💰 *Размен:* {exchange} руб.\n" +
                             $"📜 *Комментарий:* {(string.IsNullOrWhiteSpace(comment) ? "-" : comment)}";

                await botClient.SendMessage(locChatId, msg, parseMode: ParseMode.Markdown);
            }

            await TrySendMessage(chatId, "✅ Отчёт готов!");

        }
        catch (Exception ex)
        {
            await ReportError(ex, chatId);
        }
    }

    static Task HandleErrorAsync(ITelegramBotClient bot, Exception exception, CancellationToken token)
    {
        Console.WriteLine($"❌ Ошибка: {exception.Message}");
        return Task.CompletedTask;
    }
}
