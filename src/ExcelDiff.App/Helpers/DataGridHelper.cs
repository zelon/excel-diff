using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using ExcelDiff.App.Converters;
using ExcelDiff.App.ViewModels;

namespace ExcelDiff.App.Helpers;

public static class DataGridHelper
{
    public static readonly DependencyProperty AutoGenerateCellColumnsProperty =
        DependencyProperty.RegisterAttached(
            "AutoGenerateCellColumns",
            typeof(bool),
            typeof(DataGridHelper),
            new PropertyMetadata(false, OnAutoGenerateCellColumnsChanged));

    public static bool GetAutoGenerateCellColumns(DependencyObject obj)
    {
        return (bool)obj.GetValue(AutoGenerateCellColumnsProperty);
    }

    public static void SetAutoGenerateCellColumns(DependencyObject obj, bool value)
    {
        obj.SetValue(AutoGenerateCellColumnsProperty, value);
    }

    private static void OnAutoGenerateCellColumnsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is DataGrid dataGrid && (bool)e.NewValue)
        {
            dataGrid.Loaded += (s, args) => GenerateColumns(dataGrid);

            ((INotifyCollectionChanged)dataGrid.Items).CollectionChanged += (s, args) =>
            {
                if (args.Action == NotifyCollectionChangedAction.Reset ||
                    args.Action == NotifyCollectionChangedAction.Add)
                {
                    GenerateColumns(dataGrid);
                }
            };
        }
    }

    private static void GenerateColumns(DataGrid dataGrid)
    {
        if (dataGrid.Items.Count == 0)
            return;

        var firstRow = dataGrid.Items[0] as RowViewModel;
        if (firstRow == null || firstRow.Cells.Count == 0)
            return;

        // 컬럼이 이미 생성되었으면 스킵
        if (dataGrid.Columns.Count > 0)
            return;

        var converter = new DiffTypeToColorConverter();

        for (int i = 0; i < firstRow.Cells.Count; i++)
        {
            var columnIndex = i;
            var column = new DataGridTemplateColumn
            {
                Header = GetColumnName(i + 1),
                Width = new DataGridLength(100),
                CellTemplate = CreateCellTemplate(columnIndex, converter)
            };

            dataGrid.Columns.Add(column);
        }
    }

    private static DataTemplate CreateCellTemplate(int columnIndex, DiffTypeToColorConverter converter)
    {
        var factory = new FrameworkElementFactory(typeof(Border));
        factory.SetValue(Border.PaddingProperty, new Thickness(5, 2, 5, 2));
        factory.SetValue(Border.BorderThicknessProperty, new Thickness(0));

        var backgroundBinding = new Binding($"Cells[{columnIndex}].DiffType")
        {
            Converter = converter
        };
        factory.SetBinding(Border.BackgroundProperty, backgroundBinding);

        var textBlock = new FrameworkElementFactory(typeof(TextBlock));
        var textBinding = new Binding($"Cells[{columnIndex}].Value");
        textBlock.SetBinding(TextBlock.TextProperty, textBinding);

        var toolTipBinding = new Binding($"Cells[{columnIndex}].ToolTip");
        textBlock.SetBinding(TextBlock.ToolTipProperty, toolTipBinding);

        factory.AppendChild(textBlock);

        var template = new DataTemplate
        {
            VisualTree = factory
        };

        return template;
    }

    private static string GetColumnName(int columnNumber)
    {
        string columnName = "";
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }
        return columnName;
    }
}
