using System;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows;
using System.Linq;
using System.IO;
// using System.Management; // removed: WMI not used after refactor

namespace VC__
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            BtnClean.Click += async (s, e) => await CleanOldVCAsync();
            BtnInstall.Click += async (s, e) => await InstallVCAsync();
            BtnUpdate.Click += async (s, e) => await UpdateVCAsync();
        }

        private static bool IsNoUpgrade(string output)
        {
            if (string.IsNullOrWhiteSpace(output)) return false;
            // common phrases indicating no upgrade available
            var phrases = new[]
            {
                "No available upgrade found",
                "No newer package versions are available from the configured sources",
                "No packages found matching input criteria",
                "No installed package found"
            };

            return phrases.Any(p => output.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void Log(string text)
        {
            // Ensure logging happens on UI thread
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => Log(text));
                return;
            }

            // Append new line
            LogBox.AppendText(text + Environment.NewLine);

            // Cap total log size to keep memory low (храним только последние ~150 KB)
            const int maxChars = 150_000;
            if (LogBox.Text.Length > maxChars)
            {
                // Удаляем старую часть, оставляя последние maxChars/2 символов
                int keep = maxChars / 2;
                var trimmed = LogBox.Text.Substring(LogBox.Text.Length - keep, keep);
                LogBox.Text = trimmed;
                LogBox.CaretIndex = LogBox.Text.Length;
            }

            LogBox.ScrollToEnd();
        }

        private async Task CleanOldVCAsync()
        {
            Log("=== Очистка VC++ 2005-2022 ===");

            var packages = VcList();

            foreach (var pkg in packages)
            {
                Log($"Проверка: {pkg.Name}");

                // Check if installed using `winget list --id <id>` (suppress logging)
                string listResult = await RunWingetAsync($"list --id {pkg.Id}", logOutput: false);

                if (string.IsNullOrWhiteSpace(listResult) ||
                    listResult.Contains("No installed package found matching input criteria") ||
                    listResult.Contains("No installed package found"))
                {
                    Log($"  ✓ Не найдено: {pkg.Name}");
                    continue;
                }

                Log($"Удаление: {pkg.Name}");
                try
                {
                    // Try uninstall in user scope first to avoid elevation (no admin prompt)
                    string result = await RunWingetAsync($"uninstall --id {pkg.Id} --silent --disable-interactivity --accept-source-agreements --all-versions --scope user", logOutput: false);
                    if (result.Contains("Successfully uninstalled", StringComparison.OrdinalIgnoreCase))
                    {
                        Log("  ✓ Удалено (user scope)");
                        continue;
                    }

                    // If multiple versions or other issues, try uninstall without user scope but with force/all-versions (may still need elevation)
                    result = await RunWingetAsync($"uninstall --id {pkg.Id} --silent --disable-interactivity --accept-source-agreements --all-versions --force", logOutput: false);
                    if (result.Contains("Successfully uninstalled", StringComparison.OrdinalIgnoreCase))
                    {
                        Log("  ✓ Удалено (force)");
                        continue;
                    }

                    // If we reach here, uninstall failed (possibly exit code 1603). Log brief reason and do not attempt interactive elevation.
                    var brief = SummarizeWingetResult(result);
                    Log("  ⚠ Не удалось удалить: " + brief);
                }
                catch (Exception ex)
                {
                    Log("  ⚠ Ошибка: " + ex.Message);
                }
            }

            Log("Очистка завершена.");
        }

        private async Task InstallVCAsync()
        {
            Log("=== Установка VC++ ===");

            var packages = VcList();

            foreach (var pkg in packages)
            {
                Log($"Установка: {pkg.Name}");

                // Check if already installed using `winget list --id <id>` (suppress logging)
                string listResult = await RunWingetAsync($"list --id {pkg.Id}", logOutput: false);
                if (!string.IsNullOrWhiteSpace(listResult) && !listResult.Contains("No installed package found matching input criteria") && !listResult.Contains("No installed package found"))
                {
                    Log("  ✓ Уже установлен");
                    continue;
                }

                try
                {
                    // Try installing in user scope to avoid admin prompt
                    string result = await RunWingetAsync($"install --id {pkg.Id} --silent --disable-interactivity --accept-package-agreements --accept-source-agreements --scope user", logOutput: false);
                    if (result.Contains("Successfully installed", StringComparison.OrdinalIgnoreCase))
                    {
                        Log("  ✓ Готово (user scope)");
                        continue;
                    }

                    // Detect messages that indicate installer will prompt for elevation
                    if (result.IndexOf("request to run as administrator", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        result.IndexOf("will request to run as administrator", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Log("  ⚠ Требуется администратор, пропущено (чтобы избежать запроса UAC)");
                        continue;
                    }

                    // Fallback: try machine install but do not force interactive elevation — attempt and check
                    result = await RunWingetAsync($"install --id {pkg.Id} --silent --disable-interactivity --accept-package-agreements --accept-source-agreements", logOutput: false);
                    if (result.Contains("Successfully installed", StringComparison.OrdinalIgnoreCase))
                    {
                        Log("  ✓ Готово");
                        continue;
                    }

                    Log("  ⚠ Установка не подтверждена: " + SummarizeWingetResult(result));
                }
                catch (Exception ex)
                {
                    Log("  ⚠ Ошибка: " + ex.Message);
                }
            }

            Log("Установка завершена.");
        }

        private async Task UpdateVCAsync()
        {
            Log("=== Обновление VC++ ===");

            var packages = VcList();

            foreach (var pkg in packages)
            {
                Log($"Проверка обновлений: {pkg.Name}");

                string result = await RunWingetAsync($"upgrade --id {pkg.Id} --accept-package-agreements --accept-source-agreements --scope user", logOutput: false);

                if (IsNoUpgrade(result))
                {
                    Log("  ✓ Обновлений нет");
                    continue;
                }

                // If user-scope upgrade didn't perform and indicates admin required, skip to avoid prompt
                if (result.IndexOf("request to run as administrator", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    result.IndexOf("will request to run as administrator", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Log("  ⚠ Требуется администратор для обновления, пропущено");
                    continue;
                }

                Log("  → Обновляем...");
                try
                {
                    string res2 = await RunWingetAsync($"upgrade --id {pkg.Id} --silent --disable-interactivity --accept-package-agreements --accept-source-agreements", logOutput: false);
                    if (res2.Contains("Successfully installed", StringComparison.OrdinalIgnoreCase) || res2.Contains("Successfully upgraded", StringComparison.OrdinalIgnoreCase) || res2.Contains("Successfully updated", StringComparison.OrdinalIgnoreCase) || res2.Contains("Successfully updated", StringComparison.OrdinalIgnoreCase))
                        Log("  ✓ Обновлено");
                    else
                        Log("  ⚠ Обновление не подтверждено: " + SummarizeWingetResult(res2));
                }
                catch (Exception ex)
                {
                    Log("  ⚠ Ошибка: " + ex.Message);
                }
            }

            Log("Обновление завершено.");
        }

        private async Task<string> RunWingetAsync(string args, bool logOutput = true)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "winget",
                Arguments = args,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            using var p = Process.Start(psi);
            if (p == null)
            {
                Log("⚠ Не удалось запустить winget.");
                return string.Empty;
            }

            var lastLines = new Queue<string>();
            const int maxKeepLines = 200;

            async Task ReadStreamAsync(StreamReader reader)
            {
                string? line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    if (IsSpinnerLine(line))
                        continue;

                    var trimmed = line.Trim();
                    if (logOutput && !string.IsNullOrEmpty(trimmed))
                        Log(trimmed);

                    lock (lastLines)
                    {
                        lastLines.Enqueue(trimmed);
                        if (lastLines.Count > maxKeepLines)
                            lastLines.Dequeue();
                    }
                }
            }

            var outTask = ReadStreamAsync(p.StandardOutput);
            var errTask = ReadStreamAsync(p.StandardError);

            await Task.WhenAll(outTask, errTask, p.WaitForExitAsync());

            lock (lastLines)
            {
                return string.Join("\n", lastLines);
            }
        }

        private static string SummarizeWingetResult(string output)
        {
            if (string.IsNullOrWhiteSpace(output))
                return string.Empty;

            // Remove spinner/animation characters and long repeated whitespace lines
            var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrEmpty(l))
                .Where(l => !IsSpinnerLine(l))
                .ToArray();

            // Return only first few meaningful lines to keep logs short
            return string.Join("\n", lines.Take(6));
        }

        private static bool IsSpinnerLine(string line)
        {
            // Lines that consist mostly of progress characters can be ignored
            if (line.Length < 2) return string.Empty.Equals(line);
            var spinnerChars = new[] { '-', '\\', '/', '|', '█', '▒', '░' };
            var significant = line.Count(c => !char.IsWhiteSpace(c) && !spinnerChars.Contains(c));
            return significant == 0;
        }

        private static (string Id, string Name)[] VcList()
        {
            return
            [
                ("Microsoft.VCRedist.2005.x86", "VC++ 2005 x86"),
                ("Microsoft.VCRedist.2005.x64", "VC++ 2005 x64"),
                ("Microsoft.VCRedist.2008.x86", "VC++ 2008 x86"),
                ("Microsoft.VCRedist.2008.x64", "VC++ 2008 x64"),
                ("Microsoft.VCRedist.2010.x86", "VC++ 2010 x86"),
                ("Microsoft.VCRedist.2010.x64", "VC++ 2010 x64"),
                ("Microsoft.VCRedist.2012.x86", "VC++ 2012 x86"),
                ("Microsoft.VCRedist.2012.x64", "VC++ 2012 x64"),
                ("Microsoft.VCRedist.2013.x86", "VC++ 2013 x86"),
                ("Microsoft.VCRedist.2013.x64", "VC++ 2013 x64"),
                ("Microsoft.VCRedist.2015+.x86", "VC++ 2015–2022 x86"),
                ("Microsoft.VCRedist.2015+.x64", "VC++ 2015–2022 x64")
            ];
        }
    }
}