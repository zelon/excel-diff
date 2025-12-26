using System.Collections.ObjectModel;
using System.IO;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelDiff.Core.Models;
using ExcelDiff.Core.Services;
using Microsoft.Win32;

namespace ExcelDiff.App.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly IExcelReader _excelReader;
    private readonly IDiffEngine _diffEngine;

    [ObservableProperty]
    private string _oldFilePath = string.Empty;

    [ObservableProperty]
    private string _newFilePath = string.Empty;

    [ObservableProperty]
    private SheetComparisonViewModel? _selectedSheetComparison;

    [ObservableProperty]
    private bool _isLoading;

    [ObservableProperty]
    private string _statusMessage = "파일을 선택하고 비교 버튼을 클릭하세요.";

    public ObservableCollection<SheetComparisonViewModel> SheetComparisons { get; } = new();

    public MainViewModel()
    {
        _excelReader = new ExcelReaderService();
        _diffEngine = new DiffEngineService();
    }

    [RelayCommand]
    private void LoadOldFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = "Old 파일 선택"
        };

        if (dialog.ShowDialog() == true)
        {
            OldFilePath = dialog.FileName;
        }
    }

    [RelayCommand]
    private void LoadNewFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = "New 파일 선택"
        };

        if (dialog.ShowDialog() == true)
        {
            NewFilePath = dialog.FileName;
        }
    }

    [RelayCommand(CanExecute = nameof(CanCompare))]
    private async Task CompareAsync()
    {
        try
        {
            IsLoading = true;
            StatusMessage = "파일을 읽는 중...";
            SheetComparisons.Clear();

            var oldFile = await Task.Run(() => _excelReader.ReadExcelFile(OldFilePath));
            var newFile = await Task.Run(() => _excelReader.ReadExcelFile(NewFilePath));

            StatusMessage = "비교 중...";
            var diffResults = await Task.Run(() => _diffEngine.CompareExcelFiles(oldFile, newFile));

            StatusMessage = "결과 표시 중...";
            foreach (var result in diffResults)
            {
                var viewModel = new SheetComparisonViewModel(result, oldFile, newFile);
                SheetComparisons.Add(viewModel);
            }

            if (SheetComparisons.Count > 0)
            {
                SelectedSheetComparison = SheetComparisons[0];
            }

            var totalChanged = diffResults.Sum(r => r.Statistics.ChangedCells);
            StatusMessage = $"비교 완료! {SheetComparisons.Count}개 시트, {totalChanged}개 셀 변경됨.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"오류: {ex.Message}";
            System.Windows.MessageBox.Show($"오류가 발생했습니다: {ex.Message}", "오류", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
        finally
        {
            IsLoading = false;
        }
    }

    private bool CanCompare()
    {
        return !string.IsNullOrEmpty(OldFilePath) &&
               !string.IsNullOrEmpty(NewFilePath) &&
               File.Exists(OldFilePath) &&
               File.Exists(NewFilePath) &&
               !IsLoading;
    }

    partial void OnOldFilePathChanged(string value)
    {
        CompareCommand.NotifyCanExecuteChanged();
    }

    partial void OnNewFilePathChanged(string value)
    {
        CompareCommand.NotifyCanExecuteChanged();
    }

    partial void OnIsLoadingChanged(bool value)
    {
        CompareCommand.NotifyCanExecuteChanged();
    }

    /// <summary>
    /// 파일 경로를 설정하고 자동으로 비교를 실행합니다.
    /// 명령줄 인자로 파일을 받았을 때 사용됩니다.
    /// </summary>
    public async void LoadFilesAndCompare(string oldFilePath, string newFilePath)
    {
        OldFilePath = oldFilePath;
        NewFilePath = newFilePath;

        // UI가 업데이트되도록 잠시 대기
        await Task.Delay(100);

        // 비교 실행
        if (CompareCommand.CanExecute(null))
        {
            await CompareAsync();
        }
    }
}
