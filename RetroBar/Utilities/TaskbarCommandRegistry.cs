using ManagedShell.Common.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;

namespace RetroBar.Utilities
{
    internal sealed class TaskbarCommandRegistry : IDisposable
    {
        private static readonly string CommandsFilePath = Path.Combine(AppContext.BaseDirectory, "Settings\\commands.txt");

        private readonly Timer _reloadTimer;
        private readonly FileSystemWatcher _watcher;
        private readonly object _reloadLock = new();

        private TaskbarCommandDefinition[] _commands = [];
        private bool _disposed;

        public IReadOnlyList<TaskbarCommandDefinition> Commands => _commands;

        public event EventHandler CommandsChanged;

        public TaskbarCommandRegistry()
        {
            _reloadTimer = new Timer(_ => ReloadCommands(), null, Timeout.Infinite, Timeout.Infinite);

            string commandsDirectory = Path.GetDirectoryName(CommandsFilePath) ?? AppContext.BaseDirectory;
            Directory.CreateDirectory(commandsDirectory);

            LoadCommands(logMissingFile: true);

            _watcher = new FileSystemWatcher(commandsDirectory, Path.GetFileName(CommandsFilePath))
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.Size
            };

            _watcher.Changed += Watcher_OnChanged;
            _watcher.Created += Watcher_OnChanged;
            _watcher.Deleted += Watcher_OnChanged;
            _watcher.Renamed += Watcher_OnRenamed;
            _watcher.EnableRaisingEvents = true;
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            _watcher.EnableRaisingEvents = false;
            _watcher.Changed -= Watcher_OnChanged;
            _watcher.Created -= Watcher_OnChanged;
            _watcher.Deleted -= Watcher_OnChanged;
            _watcher.Renamed -= Watcher_OnRenamed;
            _watcher.Dispose();
            _reloadTimer.Dispose();
        }

        private void Watcher_OnChanged(object sender, FileSystemEventArgs e)
        {
            ScheduleReload();
        }

        private void Watcher_OnRenamed(object sender, RenamedEventArgs e)
        {
            ScheduleReload();
        }

        private void ScheduleReload()
        {
            if (_disposed)
            {
                return;
            }

            _reloadTimer.Change(250, Timeout.Infinite);
        }

        private void ReloadCommands()
        {
            if (_disposed)
            {
                return;
            }

            if (LoadCommands(logMissingFile: false))
            {
                CommandsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        private bool LoadCommands(bool logMissingFile)
        {
            lock (_reloadLock)
            {
                string[] lines;
                if (!TryReadAllLines(out lines, logMissingFile))
                {
                    return false;
                }

                List<TaskbarCommandDefinition> loadedCommands = [];

                for (int i = 0; i < lines.Length; i++)
                {
                    string rawLine = lines[i];
                    string line = rawLine.Trim();

                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#") || line.StartsWith(";"))
                    {
                        continue;
                    }

                    if (!TryParseCommand(line, out ParsedTaskbarCommand command, out string error))
                    {
                        ShellLogger.Warning($"TaskbarCommandRegistry: Ignoring invalid line {i + 1} in '{CommandsFilePath}': {error}");
                        continue;
                    }

                    loadedCommands.Add(Create(command.Name, command.ProgramPath, command.Arguments));
                }

                _commands = loadedCommands
                    .OrderBy(command => command.Name, StringComparer.OrdinalIgnoreCase)
                    .ToArray();

                ShellLogger.Info($"TaskbarCommandRegistry: Loaded {_commands.Length} command(s) from '{CommandsFilePath}'");
                return true;
            }
        }

        private static bool TryReadAllLines(out string[] lines, bool logMissingFile)
        {
            for (int attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    if (!File.Exists(CommandsFilePath))
                    {
                        if (logMissingFile)
                        {
                            ShellLogger.Warning($"TaskbarCommandRegistry: Commands file was not found at '{CommandsFilePath}'");
                        }

                        lines = [];
                        return true;
                    }

                    lines = File.ReadAllLines(CommandsFilePath);
                    return true;
                }
                catch (IOException ex) when (attempt < 4)
                {
                    Thread.Sleep(100);
                    ShellLogger.Debug($"TaskbarCommandRegistry: Waiting to reload commands file '{CommandsFilePath}': {ex.Message}");
                }
                catch (UnauthorizedAccessException ex) when (attempt < 4)
                {
                    Thread.Sleep(100);
                    ShellLogger.Debug($"TaskbarCommandRegistry: Waiting to reload commands file '{CommandsFilePath}': {ex.Message}");
                }
                catch (Exception ex)
                {
                    ShellLogger.Error($"TaskbarCommandRegistry: Failed to read commands file '{CommandsFilePath}': {ex.Message}", ex);
                    lines = [];
                    return false;
                }
            }

            try
            {
                lines = File.ReadAllLines(CommandsFilePath);
                return true;
            }
            catch (Exception ex)
            {
                ShellLogger.Error($"TaskbarCommandRegistry: Failed to read commands file '{CommandsFilePath}' after retrying: {ex.Message}", ex);
                lines = [];
                return false;
            }
        }

        private static bool TryParseCommand(string line, out ParsedTaskbarCommand command, out string error)
        {
            command = default;

            int separatorIndex = line.IndexOf('|');
            if (separatorIndex <= 0)
            {
                error = "expected the format Name|\"Full program path\" optional_arguments";
                return false;
            }

            string name = line[..separatorIndex].Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                error = "command name cannot be empty";
                return false;
            }

            string commandText = line[(separatorIndex + 1)..].Trim();
            if (!commandText.StartsWith('"'))
            {
                error = "program path must be wrapped in double quotes";
                return false;
            }

            int closingQuoteIndex = commandText.IndexOf('"', 1);
            if (closingQuoteIndex <= 1)
            {
                error = "program path is missing its closing double quote";
                return false;
            }

            string programPath = commandText.Substring(1, closingQuoteIndex - 1);
            if (string.IsNullOrWhiteSpace(programPath))
            {
                error = "program path cannot be empty";
                return false;
            }

            if (!Path.IsPathRooted(programPath))
            {
                error = "program path must be an absolute path";
                return false;
            }

            string workingDirectory = Path.GetDirectoryName(programPath);
            if (string.IsNullOrWhiteSpace(workingDirectory))
            {
                error = "program path must include a parent directory";
                return false;
            }

            command = new ParsedTaskbarCommand
            {
                Name = name,
                ProgramPath = programPath,
                Arguments = commandText[(closingQuoteIndex + 1)..].TrimStart()
            };

            error = null;
            return true;
        }

        private TaskbarCommandDefinition Create(string name, string programPath, string arguments)
        {
            return new TaskbarCommandDefinition
            {
                Name = name,
                Execute = () => ExecuteCommand(name, programPath, arguments)
            };
        }

        private static void ExecuteCommand(string name, string programPath, string arguments)
        {
            try
            {
                ProcessStartInfo startInfo = new()
                {
                    FileName = programPath,
                    Arguments = arguments,
                    WorkingDirectory = Path.GetDirectoryName(programPath),
                    UseShellExecute = true
                };

                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                ShellLogger.Error($"TaskbarCommandRegistry: Command '{name}' failed: {ex.Message}", ex);
                MessageBox.Show($"Command '{name}' failed.\n{ex.Message}", "RetroBar", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private readonly struct ParsedTaskbarCommand
        {
            public string Name { get; init; }

            public string ProgramPath { get; init; }

            public string Arguments { get; init; }
        }
    }

    internal sealed class TaskbarCommandDefinition
    {
        public string Name { get; init; }

        public Action Execute { get; init; }
    }
}
