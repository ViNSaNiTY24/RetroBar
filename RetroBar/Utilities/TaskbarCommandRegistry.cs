using ManagedShell.Common.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;

namespace RetroBar.Utilities
{
    internal sealed class TaskbarCommandRegistry
    {
        private readonly Taskbar _taskbar;
        private readonly List<TaskbarCommandDefinition> _commands;

        public IReadOnlyList<TaskbarCommandDefinition> Commands => _commands;

        public TaskbarCommandRegistry(Taskbar taskbar)
        {
            _taskbar = taskbar;

            // Add new taskbar commands here. Each entry pairs the visible command
            // name with the method that should run when that command is chosen.
            _commands =
            [
                Create("Aria2c Download", Aria2cDownload),
                Create("HardDriveFileLister", HardDriveFileLister),
                Create("Page File", ShowPageFileSize),
                Create("Sound Settings", SoundSettings),
                Create("SportsRssQbt", SportsRssQbt),
                Create("Timer", Timer),
                Create("Uninstall Tool", UninstallTool),
                Create("Volume Mixer", VolumeMixer),
                Create("WinSpyDark", WinSpyDark)
            ];
        }

        private TaskbarCommandDefinition Create(string name, Action execute)
        {
            return new TaskbarCommandDefinition
            {
                Name = name,
                Execute = () => ExecuteCommand(name, execute)
            };
        }

        private static void ExecuteCommand(string name, Action execute)
        {
            try
            {
                execute();
            }
            catch (Exception ex)
            {
                ShellLogger.Error($"TaskbarCommandRegistry: Command '{name}' failed: {ex.Message}", ex);
                MessageBox.Show($"Command '{name}' failed.\n{ex.Message}", "RetroBar", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static void Aria2cDownload()
        {
            var P = new Process();
            P.StartInfo.FileName = @"E:\Misc. Stuff\Python\aria2c.pyw";
            P.StartInfo.WorkingDirectory = @"E:\Misc. Stuff\Python";
            P.StartInfo.UseShellExecute = true;
            P.Start();
        }

        private static void ShowPageFileSize()
        {
            var P = new Process();
            P.StartInfo.FileName = @"E:\Misc. Stuff\AutoHotkey\GetPageFile.ahk";
            P.StartInfo.WorkingDirectory = @"E:\Misc. Stuff\AutoHotkey";
            P.StartInfo.UseShellExecute = true;
            P.Start();
        }

        private static void HardDriveFileLister()
        {
            var P = new Process();
            P.StartInfo.FileName = @"E:\Misc. Stuff\AutoHotkey\HardDriveFileLister.ahk";
            P.StartInfo.WorkingDirectory = @"E:\Misc. Stuff\AutoHotkey";
            P.StartInfo.UseShellExecute = true;
            P.Start();
        }

        private static void SoundSettings()
        {
            var P = new Process();
            P.StartInfo.FileName = @"mmsys.cpl";
            P.StartInfo.UseShellExecute = true;
            P.Start();
        }

        private static void SportsRssQbt()
        {
            var P = new Process();
            P.StartInfo.FileName = @"E:\Misc. Stuff\AutoHotkey\SportsRssQbt.ahk";
            P.StartInfo.WorkingDirectory = @"E:\Misc. Stuff\AutoHotkey";
            P.StartInfo.UseShellExecute = true;
            P.Start();
        }

        private static void Timer()
        {
            var P = new Process();
            P.StartInfo.FileName = @"E:\Misc. Stuff\AutoHotkey\Timer.ahk";
            P.StartInfo.WorkingDirectory = @"E:\Misc. Stuff\AutoHotkey";
            P.StartInfo.UseShellExecute = true;
            P.Start();
        }

        private static void UninstallTool()
        {
            var P = new Process();
            P.StartInfo.FileName = @"C:\Programs\U\UninstallTool\UninstallToolPortable.exe";
            P.StartInfo.WorkingDirectory = @"C:\Programs\U\UninstallTool";
            P.StartInfo.UseShellExecute = true;
            P.Start();
        }

        private static void VolumeMixer()
        {
            var P = new Process();
            P.StartInfo.FileName = @"E:\Misc. Stuff\AutoHotkey\VolumeMixer.ahk";
            P.StartInfo.WorkingDirectory = @"E:\Misc. Stuff\AutoHotkey";
            P.StartInfo.UseShellExecute = true;
            P.Start();
        }

        private static void WinSpyDark()
        {
            var P = new Process();
            P.StartInfo.FileName = @"E:\Misc. Stuff\AutoHotkey\WinSpyDark.ahk";
            P.StartInfo.WorkingDirectory = @"E:\Misc. Stuff\AutoHotkey";
            P.StartInfo.UseShellExecute = true;
            P.Start();
        }
    }

    internal sealed class TaskbarCommandDefinition
    {
        public string Name { get; init; }

        public Action Execute { get; init; }
    }
}
