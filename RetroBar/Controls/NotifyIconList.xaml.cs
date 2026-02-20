using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using ManagedShell.Common.Helpers;
using ManagedShell.WindowsTray;
using RetroBar.Extensions;
using RetroBar.Utilities;

namespace RetroBar.Controls
{
    /// <summary>
    /// Interaction logic for NotifyIconList.xaml
    /// </summary>
    public partial class NotifyIconList : UserControl
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        private bool _isLoaded;
        private CollectionViewSource allNotifyIconsSource;
        private CollectionViewSource pinnedNotifyIconsSource;
        private ObservableCollection<ManagedShell.WindowsTray.NotifyIcon> promotedIcons = [];

        // Sorted collections
        private ObservableCollection<ManagedShell.WindowsTray.NotifyIcon> sortedAllIcons = [];
        private ObservableCollection<ManagedShell.WindowsTray.NotifyIcon> sortedPinnedIcons = [];

        public static DependencyProperty NotificationAreaProperty = DependencyProperty.Register(nameof(NotificationArea), typeof(NotificationArea), typeof(NotifyIconList), new PropertyMetadata(NotificationAreaChangedCallback));

        public NotificationArea NotificationArea
        {
            get { return (NotificationArea)GetValue(NotificationAreaProperty); }
            set { SetValue(NotificationAreaProperty, value); }
        }

        public NotifyIconList()
        {
            InitializeComponent();
        }

        #region Ghost Icon Cleanup (Event-Driven)

        private void MonitorIconProcess(ManagedShell.WindowsTray.NotifyIcon icon)
        {
            if (icon == null || icon.HWnd == IntPtr.Zero) return;

            GetWindowThreadProcessId(icon.HWnd, out uint processId);

            if (processId == 0) return;

            try
            {
                Process process = Process.GetProcessById((int)processId);

                // If the process is already dead, clean it up immediately
                if (process.HasExited)
                {
                    CleanUpGhostIcon(icon);
                    process.Dispose();
                    return;
                }

                process.EnableRaisingEvents = true;
                process.Exited += (sender, args) =>
                {
                    // The process just crashed or closed! Clean up the icon.
                    CleanUpGhostIcon(icon);

                    // Clean up the process object from memory
                    if (sender is Process p) p.Dispose();
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[RetroBar] Could not monitor process {processId}: {ex.Message}");
                // If we get an Access Denied exception (e.g. for an elevated app), the process is running 
                // but we can't hook it. Only clean up if the window itself is actually dead.
                if (!IsWindow(icon.HWnd))
                {
                    CleanUpGhostIcon(icon);
                }
            }
        }

        private void CleanUpGhostIcon(ManagedShell.WindowsTray.NotifyIcon icon)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (icon != null && icon.HWnd != IntPtr.Zero && !IsWindow(icon.HWnd))
                {
                    // Simulating a mouse event forces ManagedShell to verify the window handle.
                    // When it detects the window is dead, it will remove the stale icon.
                    icon.IconMouseMove(MouseHelper.GetCursorPositionParam());
                }
            }), DispatcherPriority.Background);
        }

        #endregion

        #region Custom Sorting

        /// <summary>
        /// Gets the sort priority for a NotifyIcon based on its Title.
        /// Lower numbers appear first.
        /// </summary>
        private int GetIconSortPriority(ManagedShell.WindowsTray.NotifyIcon icon)
        {
            if (icon == null || string.IsNullOrEmpty(icon.Title))
                return 1000;

            string title = icon.Title;

            if (title.Contains("GB") && icon.Path.Contains("AutoHotkey64.exe") || title.Contains("Twitch.ahk") && icon.Path.Contains("AutoHotkey64.exe"))
            {
                return 1;
            }

            if (title.Contains("VSTHost"))
            {
                return 2;
            }

            if (title.Contains("Safely Remove Hardware"))
            {
                return 9998;
            }

            if (title.Contains("Speakers:"))
            {
                return 9999;
            }

            return 100;
        }

        /// <summary>
        /// Extracts NotifyIcons from any enumerable source
        /// </summary>
        private List<ManagedShell.WindowsTray.NotifyIcon> ExtractIcons(IEnumerable source)
        {
            var result = new List<ManagedShell.WindowsTray.NotifyIcon>();
            if (source == null) return result;

            foreach (object item in source)
            {
                if (item is ManagedShell.WindowsTray.NotifyIcon icon)
                {
                    result.Add(icon);
                }
            }
            return result;
        }

        /// <summary>
        /// Sorts a list of icons by priority
        /// </summary>
        private List<ManagedShell.WindowsTray.NotifyIcon> SortIconList(List<ManagedShell.WindowsTray.NotifyIcon> icons)
        {
            return icons
                .OrderBy(GetIconSortPriority)
                .ThenBy(i => i?.Title ?? "")
                .ToList();
        }

        /// <summary>
        /// Rebuilds the sorted all icons collection
        /// </summary>
        private void RebuildSortedAllIcons()
        {
            if (NotificationArea == null) return;

            try
            {
                var unpinned = ExtractIcons(NotificationArea.UnpinnedIcons);
                var pinned = ExtractIcons(NotificationArea.PinnedIcons);

                // Apply filter to unpinned AND exclude items
                var filteredUnpinned = unpinned
                    .Where(icon => UnpinnedNotifyIcons_Filter(icon) && !ShouldExcludeIcon(icon))
                    .ToList();

                // Also exclude from pinned
                var filteredPinned = pinned
                    .Where(icon => !ShouldExcludeIcon(icon))
                    .ToList();

                var all = filteredUnpinned.Concat(filteredPinned).ToList();
                var sorted = SortIconList(all);

                sortedAllIcons.Clear();
                foreach (var icon in sorted)
                {
                    sortedAllIcons.Add(icon);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[RetroBar] RebuildSortedAllIcons error: {ex.Message}");
            }
        }

        /// <summary>
        /// Rebuilds the sorted pinned icons collection
        /// </summary>
        private void RebuildSortedPinnedIcons()
        {
            if (NotificationArea == null) return;

            try
            {
                var pinned = ExtractIcons(NotificationArea.PinnedIcons);
                var promoted = promotedIcons.ToList();

                // Exclude items from both collections
                var filteredPinned = pinned
                    .Where(icon => !ShouldExcludeIcon(icon))
                    .ToList();

                var filteredPromoted = promoted
                    .Where(icon => !ShouldExcludeIcon(icon))
                    .ToList();

                var all = filteredPromoted.Concat(filteredPinned).ToList();
                var sorted = SortIconList(all);

                sortedPinnedIcons.Clear();
                foreach (var icon in sorted)
                {
                    sortedPinnedIcons.Add(icon);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[RetroBar] RebuildSortedPinnedIcons error: {ex.Message}");
            }
        }

        private void RebuildAllSortedCollections()
        {
            RebuildSortedAllIcons();
            RebuildSortedPinnedIcons();
        }

        #endregion

        private void Settings_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(Settings.CollapseNotifyIcons))
            {
                if (Settings.Instance.CollapseNotifyIcons)
                {
                    RebuildSortedPinnedIcons();
                    NotifyIcons.ItemsSource = sortedPinnedIcons;
                    SetToggleVisibility();
                }
                else
                {
                    NotifyIconToggleButton.IsChecked = false;
                    NotifyIconToggleButton.Visibility = Visibility.Collapsed;

                    RebuildSortedAllIcons();
                    NotifyIcons.ItemsSource = sortedAllIcons;
                }
            }
            else if (e.PropertyName == nameof(Settings.InvertIconsMode) || e.PropertyName == nameof(Settings.InvertNotifyIcons))
            {
                NotifyIcons.ItemsSource = null;

                if (Settings.Instance.CollapseNotifyIcons && NotifyIconToggleButton.IsChecked != true)
                {
                    RebuildSortedPinnedIcons();
                    NotifyIcons.ItemsSource = sortedPinnedIcons;
                }
                else
                {
                    RebuildSortedAllIcons();
                    NotifyIcons.ItemsSource = sortedAllIcons;
                }
            }
        }

        private void SetNotificationAreaCollections()
        {
            if (!_isLoaded && NotificationArea != null)
            {
                // Keep original setup for filter functionality
                CompositeCollection allNotifyIcons =
                [
                    new CollectionContainer { Collection = NotificationArea.UnpinnedIcons },
                    new CollectionContainer { Collection = NotificationArea.PinnedIcons },
                ];
                allNotifyIconsSource = new CollectionViewSource { Source = allNotifyIcons };
                NotificationArea.UnpinnedIcons.Filter = UnpinnedNotifyIcons_Filter;

                CompositeCollection pinnedNotifyIcons =
                [
                    new CollectionContainer { Collection = promotedIcons },
                    new CollectionContainer { Collection = NotificationArea.PinnedIcons },
                ];
                pinnedNotifyIconsSource = new CollectionViewSource { Source = pinnedNotifyIcons };

                // Subscribe to changes
                NotificationArea.UnpinnedIcons.CollectionChanged += UnpinnedIcons_CollectionChanged;
                NotificationArea.PinnedIcons.CollectionChanged += PinnedIcons_CollectionChanged;
                NotificationArea.NotificationBalloonShown += NotificationArea_NotificationBalloonShown;
                promotedIcons.CollectionChanged += PromotedIcons_CollectionChanged;

                Settings.Instance.PropertyChanged += Settings_PropertyChanged;

                // Monitor initially loaded icons for ghost cleanup
                foreach (var icon in ExtractIcons(NotificationArea.UnpinnedIcons)) MonitorIconProcess(icon);
                foreach (var icon in ExtractIcons(NotificationArea.PinnedIcons)) MonitorIconProcess(icon);

                // Build sorted collections
                RebuildAllSortedCollections();

                if (Settings.Instance.CollapseNotifyIcons)
                {
                    NotifyIcons.ItemsSource = sortedPinnedIcons;
                    SetToggleVisibility();

                    if (NotifyIconToggleButton.IsChecked == true)
                    {
                        NotifyIconToggleButton.IsChecked = false;
                    }
                }
                else
                {
                    NotifyIcons.ItemsSource = sortedAllIcons;
                }

                _isLoaded = true;
            }
        }

        private void PinnedIcons_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            // Monitor newly added icons for ghost cleanup
            if (e.Action == NotifyCollectionChangedAction.Add && e.NewItems != null)
            {
                foreach (ManagedShell.WindowsTray.NotifyIcon icon in e.NewItems)
                {
                    MonitorIconProcess(icon);
                }
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                RebuildAllSortedCollections();
            }), DispatcherPriority.Background);
        }

        private void PromotedIcons_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                RebuildSortedPinnedIcons();
            }), DispatcherPriority.Background);
        }

        private static void NotificationAreaChangedCallback(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            if (sender is NotifyIconList notifyIconList && e.OldValue == null && e.NewValue != null)
            {
                notifyIconList.SetNotificationAreaCollections();
            }
        }

        private bool UnpinnedNotifyIcons_Filter(object obj)
        {
            if (obj is ManagedShell.WindowsTray.NotifyIcon notifyIcon)
            {
                // Exclude items with certain text
                if (ShouldExcludeIcon(notifyIcon))
                {
                    return false;
                }

                return !notifyIcon.IsPinned && !notifyIcon.IsHidden && notifyIcon.GetBehavior() != NotifyIconBehavior.Remove;
            }

            return true;
        }

        /// <summary>
        /// Returns true if the icon should be excluded/hidden from the list
        /// </summary>
        private bool ShouldExcludeIcon(ManagedShell.WindowsTray.NotifyIcon icon)
        {
            if (icon == null || string.IsNullOrEmpty(icon.Title))
                return false;

            string title = icon.Title;

            // Add your exclusion rules here
            if (title.Contains("apps are using your microphone"))
            {
                return true;
            }

            // Add more exclusion rules as needed:
            // if (title.Contains("Some other text"))
            // {
            //     return true;
            // }

            return false;
        }

        private void NotificationArea_NotificationBalloonShown(object sender, NotificationBalloonEventArgs e)
        {
            if (NotificationArea == null)
            {
                return;
            }

            ManagedShell.WindowsTray.NotifyIcon notifyIcon = e.Balloon.NotifyIcon;

            var pinnedIcons = ExtractIcons(NotificationArea.PinnedIcons);
            if (pinnedIcons.Contains(notifyIcon))
            {
                return;
            }

            if (notifyIcon.GetBehavior() != NotifyIconBehavior.HideWhenInactive)
            {
                return;
            }

            if (promotedIcons.Contains(notifyIcon))
            {
                return;
            }

            promotedIcons.Add(notifyIcon);

            DispatcherTimer unpromoteTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(e.Balloon.Timeout + 500)
            };
            unpromoteTimer.Tick += (object s, EventArgs args) =>
            {
                if (promotedIcons.Contains(notifyIcon))
                {
                    promotedIcons.Remove(notifyIcon);
                }
                unpromoteTimer.Stop();
            };
            unpromoteTimer.Start();
        }

        private void NotifyIconList_Loaded(object sender, RoutedEventArgs e)
        {
            SetNotificationAreaCollections();
        }

        private void NotifyIconList_OnUnloaded(object sender, RoutedEventArgs e)
        {
            if (!_isLoaded)
            {
                return;
            }

            Settings.Instance.PropertyChanged -= Settings_PropertyChanged;

            if (NotificationArea != null)
            {
                NotificationArea.UnpinnedIcons.CollectionChanged -= UnpinnedIcons_CollectionChanged;
                NotificationArea.PinnedIcons.CollectionChanged -= PinnedIcons_CollectionChanged;
                NotificationArea.NotificationBalloonShown -= NotificationArea_NotificationBalloonShown;
            }

            promotedIcons.CollectionChanged -= PromotedIcons_CollectionChanged;

            _isLoaded = false;
        }

        private void UnpinnedIcons_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            SetToggleVisibility();

            // Monitor newly added icons for ghost cleanup
            if (e.Action == NotifyCollectionChangedAction.Add && e.NewItems != null)
            {
                foreach (ManagedShell.WindowsTray.NotifyIcon icon in e.NewItems)
                {
                    MonitorIconProcess(icon);
                }
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                RebuildAllSortedCollections();
            }), DispatcherPriority.Background);
        }

        private void NotifyIconToggleButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (NotifyIconToggleButton.IsChecked == true)
            {
                RebuildSortedAllIcons();
                NotifyIcons.ItemsSource = sortedAllIcons;
            }
            else
            {
                RebuildSortedPinnedIcons();
                NotifyIcons.ItemsSource = sortedPinnedIcons;
            }
        }

        private void SetToggleVisibility()
        {
            if (!Settings.Instance.CollapseNotifyIcons) return;

            if (NotificationArea.UnpinnedIcons.IsEmpty)
            {
                NotifyIconToggleButton.Visibility = Visibility.Collapsed;

                if (NotifyIconToggleButton.IsChecked == true)
                {
                    NotifyIconToggleButton.IsChecked = false;
                }
            }
            else
            {
                NotifyIconToggleButton.Visibility = Visibility.Visible;
            }
        }
    }
}