using ManagedCommon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using Wox.Plugin;
using Wox.Plugin.Logger;

namespace plugin;

/// <summary>
/// Main class of this plugin that implement all used interfaces.
/// </summary>
public class Main : IPlugin, IContextMenu, IDisposable
{
    /// <summary>
    /// ID of the plugin.
    /// </summary>
    public static string PluginID => "A8F26C057BF74EF6B722F7FE13EC138D";

    /// <summary>
    /// Name of the plugin.
    /// </summary>
    public string Name => "plugin";

    /// <summary>
    /// Description of the plugin.
    /// </summary>
    public string Description => "plugin Description";

    private PluginInitContext Context { get; set; }

    private string IconPath { get; set; }

    private bool Disposed { get; set; }

    /// <summary>
    /// Return a filtered list, based on the given query.
    /// </summary>
    /// <param name="query">The query to filter the list.</param>
    /// <returns>A filtered list, can be empty when nothing was found.</returns>
    public List<Result> Query(Query query)
    {
        var search = query.Search;
        Log.Info($"Query invoked: '{search}'", GetType());

        var cfg = MailConfig.Load();
        var parts = (search ?? string.Empty).Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 2 && parts[0].Equals("mail", StringComparison.OrdinalIgnoreCase))
        {
            Log.Info($"Mail command detected: '{search}'", GetType());
            var cmd = parts[1].ToLowerInvariant();
            if (cmd == "start")
            {
                return new List<Result>
                {
                    new Result
                    {
                        QueryTextDisplay = search,
                        IcoPath = IconPath,
                        Title = "Send start mail",
                        SubTitle = "Send start notification via Outlook",
                        Action = _ =>
                        {
                            try
                            {
                                var subject = $"{cfg.SubjectPrefix} {DateTime.Now.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture)} start";
                                var body = cfg.StartBody;
                                Log.Info($"Sending start mail: subject='{subject}', recipients={string.Join(';', cfg.Recipients ?? Array.Empty<string>())}", GetType());
                                MailSender.SendMail(cfg.Recipients, subject, body);
                            }
                            catch (Exception ex)
                            {
                                Log.Error($"Error sending start mail: {ex}", GetType());
                            }

                            return true;
                        }
                    }
                };
            }

            if (cmd == "stop")
            {
                // optional hours parameter: mail stop 8h  or mail stop 8
                var hours = "";
                if (parts.Length >= 3)
                {
                    hours =  parts[2].Trim().ToLowerInvariant();
                }

                return new List<Result>
                {
                    new Result
                    {
                        QueryTextDisplay = search,
                        IcoPath = IconPath,
                        Title = "Send stop mail",
                        SubTitle = "Send stop notification via Outlook",
                        Action = _ =>
                        {
                            try
                            {
                                var subject = $"{cfg.SubjectPrefix} {DateTime.Now.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture)} stop";
                                var body = cfg.StopBody;
                                if (!string.IsNullOrEmpty(hours.ToString()))
                                {
                                    body = cfg.StopBody + " - " + hours;
                                }
                                Log.Info($"Sending stop mail: subject='{subject}', recipients={string.Join(';', cfg.Recipients ?? Array.Empty<string>())}, hours={hours}", GetType());
                                MailSender.SendMail(cfg.Recipients, subject, body);
                            }
                            catch (Exception ex)
                            {
                                Log.Error($"Error sending stop mail: {ex}", GetType());
                            }

                            return true;
                        }
                    }
                };
            }
        }

        return new List<Result>
        {
            new Result
            {
                QueryTextDisplay = search,
                IcoPath = IconPath,
                Title = "Title: " + search,
                SubTitle = "SubTitle",
                ToolTipData = new ToolTipData("Title", "Text"),
                Action = _ =>
                {
                    Clipboard.SetDataObject(search);
                    return true;
                },
                ContextData = search,
            }
        };
    }

    private class MailConfig
    {
        public string[] Recipients { get; set; } = Array.Empty<string>();
        public string SubjectPrefix { get; set; } = "Notification";
        public string StartBody { get; set; } = "Start notification";
        public string StopBody { get; set; } = "Stop notification";

        public static MailConfig Load()
        {
            try
            {
                var assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? AppContext.BaseDirectory;
                var cfgPath = Path.Combine(assemblyPath, "mailconfig.json");
                Log.Info($"Loading mail config from: {cfgPath}", typeof(MailConfig));
                if (!File.Exists(cfgPath))
                {
                    Log.Info("mailconfig.json not found, using defaults", typeof(MailConfig));
                    return new MailConfig();
                }

                var json = File.ReadAllText(cfgPath);
                var cfg = JsonSerializer.Deserialize<MailConfig>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                Log.Info($"mailconfig.json loaded, recipients={cfg?.Recipients?.Length ?? 0}", typeof(MailConfig));
                return cfg ?? new MailConfig();
            }
            catch
            {
                Log.Error("Failed to load mailconfig.json, using defaults", typeof(MailConfig));
                return new MailConfig();
            }
        }
    }

    private static class MailSender
    {
        public static void SendMail(string[] recipients, string subject, string body)
        {
            try
            {
                var assemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? AppContext.BaseDirectory;
                var scriptPath = Path.Combine(assemblyPath, "sendmail.ps1");
                
                if (!File.Exists(scriptPath))
                {
                    Log.Error($"sendmail.ps1 not found at {scriptPath}", typeof(MailSender));
                    return;
                }

                var to = string.Join(";", recipients ?? Array.Empty<string>());
                var args = $"-NoProfile -ExecutionPolicy Bypass -File \"{scriptPath}\" -To \"{to}\" -Subject \"{subject}\" -Body \"{body}\"";
                
                Log.Info($"MailSender: sending mail to {to} subject='{subject}'", typeof(MailSender));
                
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = args,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(psi))
                {
                    process?.WaitForExit();
                    if (process?.ExitCode == 0)
                    {
                        Log.Info("MailSender: mail sent successfully", typeof(MailSender));
                    }
                    else
                    {
                        var error = process?.StandardError.ReadToEnd();
                        Log.Error($"MailSender: PowerShell exit code {process?.ExitCode}, error: {error}", typeof(MailSender));
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error($"MailSender error: {ex}", typeof(MailSender));
            }
        }
    }

    /// <summary>
    /// Initialize the plugin with the given <see cref="PluginInitContext"/>.
    /// </summary>
    /// <param name="context">The <see cref="PluginInitContext"/> for this plugin.</param>
    public void Init(PluginInitContext context)
    {
        Context = context ?? throw new ArgumentNullException(nameof(context));
        Context.API.ThemeChanged += OnThemeChanged;
        UpdateIconPath(Context.API.GetCurrentTheme());
    }

    /// <summary>
    /// Return a list context menu entries for a given <see cref="Result"/> (shown at the right side of the result).
    /// </summary>
    /// <param name="selectedResult">The <see cref="Result"/> for the list with context menu entries.</param>
    /// <returns>A list context menu entries.</returns>
    public List<ContextMenuResult> LoadContextMenus(Result selectedResult)
    {
        if (selectedResult.ContextData is string search)
        {
            return
            [
                new ContextMenuResult
                {
                    PluginName = Name,
                    Title = "Copy to clipboard (Ctrl+C)",
                    Glyph = "\xE8C8", // Copy
                    FontFamily = "Segoe Fluent Icons,Segoe MDL2 Assets",
                    AcceleratorKey = Key.C,
                    AcceleratorModifiers = ModifierKeys.Control,
                    Action = _ =>
                    {
                        Clipboard.SetDataObject(search);
                        return true;
                    },
                }
            ];
        }

        return [];
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Wrapper method for <see cref="Dispose()"/> that dispose additional objects and events form the plugin itself.
    /// </summary>
    /// <param name="disposing">Indicate that the plugin is disposed.</param>
    protected virtual void Dispose(bool disposing)
    {
        if (Disposed || !disposing)
        {
            return;
        }

        if (Context?.API != null)
        {
            Context.API.ThemeChanged -= OnThemeChanged;
        }

        Disposed = true;
    }

    private void UpdateIconPath(Theme theme) => IconPath = theme == Theme.Light || theme == Theme.HighContrastWhite ? "Images/plugin.light.png" : "Images/plugin.dark.png";

    private void OnThemeChanged(Theme currentTheme, Theme newTheme) => UpdateIconPath(newTheme);
}
