using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ExcelDiff.App.Helpers;

public static class ScrollSyncHelper
{
    private static bool _isSyncing = false;

    public static readonly DependencyProperty ScrollGroupProperty =
        DependencyProperty.RegisterAttached(
            "ScrollGroup",
            typeof(string),
            typeof(ScrollSyncHelper),
            new PropertyMetadata(null, OnScrollGroupChanged));

    public static string GetScrollGroup(DependencyObject obj)
    {
        return (string)obj.GetValue(ScrollGroupProperty);
    }

    public static void SetScrollGroup(DependencyObject obj, string value)
    {
        obj.SetValue(ScrollGroupProperty, value);
    }

    private static void OnScrollGroupChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not DataGrid dataGrid)
            return;

        if (e.OldValue != null)
        {
            dataGrid.Loaded -= DataGrid_Loaded;
        }

        if (e.NewValue != null)
        {
            dataGrid.Loaded += DataGrid_Loaded;
        }
    }

    private static void DataGrid_Loaded(object sender, RoutedEventArgs e)
    {
        var dataGrid = (DataGrid)sender;
        var scrollViewer = GetScrollViewer(dataGrid);

        if (scrollViewer != null)
        {
            scrollViewer.ScrollChanged += (s, args) => ScrollViewer_ScrollChanged(dataGrid, args);
        }
    }

    private static void ScrollViewer_ScrollChanged(DataGrid sourceGrid, ScrollChangedEventArgs e)
    {
        if (_isSyncing)
            return;

        var scrollGroup = GetScrollGroup(sourceGrid);
        if (string.IsNullOrEmpty(scrollGroup))
            return;

        _isSyncing = true;

        try
        {
            var window = Window.GetWindow(sourceGrid);
            if (window == null)
                return;

            SyncScrollForGroup(window, scrollGroup, sourceGrid, e.HorizontalOffset, e.VerticalOffset);
        }
        finally
        {
            _isSyncing = false;
        }
    }

    private static void SyncScrollForGroup(DependencyObject parent, string scrollGroup, DataGrid sourceGrid, double horizontalOffset, double verticalOffset)
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);

            if (child is DataGrid targetGrid && targetGrid != sourceGrid)
            {
                var targetScrollGroup = GetScrollGroup(targetGrid);
                if (targetScrollGroup == scrollGroup)
                {
                    var scrollViewer = GetScrollViewer(targetGrid);
                    scrollViewer?.ScrollToHorizontalOffset(horizontalOffset);
                    scrollViewer?.ScrollToVerticalOffset(verticalOffset);
                }
            }

            SyncScrollForGroup(child, scrollGroup, sourceGrid, horizontalOffset, verticalOffset);
        }
    }

    private static ScrollViewer? GetScrollViewer(DependencyObject obj)
    {
        if (obj is ScrollViewer scrollViewer)
            return scrollViewer;

        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
        {
            var child = VisualTreeHelper.GetChild(obj, i);
            var result = GetScrollViewer(child);
            if (result != null)
                return result;
        }

        return null;
    }
}
