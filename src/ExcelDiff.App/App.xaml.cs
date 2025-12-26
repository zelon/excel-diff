using System.Configuration;
using System.Data;
using System.IO;
using System.Windows;
using ExcelDiff.App.ViewModels;

namespace ExcelDiff.App;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var mainWindow = new MainWindow();
        var viewModel = mainWindow.DataContext as MainViewModel;

        // 명령줄 인자가 2개 있으면 자동으로 비교 실행
        if (e.Args.Length >= 2 && viewModel != null)
        {
            var oldFilePath = e.Args[0];
            var newFilePath = e.Args[1];

            // 파일 존재 여부 확인
            if (File.Exists(oldFilePath) && File.Exists(newFilePath))
            {
                viewModel.LoadFilesAndCompare(oldFilePath, newFilePath);
            }
            else
            {
                MessageBox.Show(
                    $"파일을 찾을 수 없습니다.\nOld: {oldFilePath}\nNew: {newFilePath}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        mainWindow.Show();
    }
}

