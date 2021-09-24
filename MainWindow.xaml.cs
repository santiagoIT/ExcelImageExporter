using System;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Shapes;
using Cursors = System.Windows.Forms.Cursors;
using Window = System.Windows.Window;

namespace ImagesFromExcel
{
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {
    public MainWindow()
    {
      InitializeComponent();
    }

    private void Button_Click(object sender, RoutedEventArgs e)
    {
      TextBlockConsole.Text = string.Empty;
      if (!InputValid())
      {
        return;
      }

      var path = TextBoxOutputDirectory.Text;
      var excelFile = TextBoxExcelFile.Text;

      var currentCursor = Cursor;
      try
      {

        Cursor = System.Windows.Input.Cursors.Wait;

        LogMessage("Processing excel file. Please wait...");
        Workbook workbook = new Workbook();
        workbook.LoadFromFile(excelFile);

        Worksheet sheet = workbook.Worksheets[0];

        var exportedPictures = 0;
        foreach (var sheetPicture in sheet.Pictures)
        {
          var top = sheetPicture.Top;
          var test = sheetPicture as XlsShape;
          var topRow = test.TopRow;
          var topRowOffset = test.TopRowOffset;

          var fileName = sheet.Range[$"A{topRow}"].Text;
          sheetPicture.Picture.Save(System.IO.Path.Combine(path, $"{fileName}.png"), ImageFormat.Png);
          exportedPictures++;
        }

        LogMessage($"Exported {exportedPictures} pictures!");
      }
      catch (Exception exception)
      {
        LogMessage("An error has ocurred.");
        LogMessage(exception.Message);
      }
      finally
      {
        Cursor = currentCursor;
      }
    }

    private void ButtonSelectExcelFile_OnClick(object sender, RoutedEventArgs e)
    {
      var openFileDialog = new OpenFileDialog()
      {
        Filter = "Excel files (*.xlsx)|*.xlsx"
      };
      if (openFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
      {
        return;
      }
      
      TextBoxExcelFile.Text = openFileDialog.FileName;

      ButtonSaveImages.IsEnabled = InputValid();
    }

    private void ButtonBrowseOutputDirectory_OnClick(object sender, RoutedEventArgs e)
    {
      using (var dialog = new FolderBrowserDialog {
          Description = "Select Output Folder",
          SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                         + Path.DirectorySeparatorChar,
          ShowNewFolderButton = true
        }) 
      {
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
          TextBoxOutputDirectory.Text = dialog.SelectedPath;
        }
      };

      ButtonSaveImages.IsEnabled = InputValid();
    }

    private bool InputValid()
    {
      ButtonSaveImages.IsEnabled = false;
      if (string.IsNullOrWhiteSpace(TextBoxOutputDirectory.Text) || string.IsNullOrWhiteSpace(TextBoxExcelFile.Text))
      {
        return false;
      }

      if (!File.Exists(TextBoxExcelFile.Text) || !Directory.Exists(TextBoxOutputDirectory.Text))
      {
        return false;
      }

      ButtonSaveImages.IsEnabled = true;
      return true;
    }

    private void LogMessage(string msg)
    {
      if (string.IsNullOrWhiteSpace(TextBlockConsole.Text))
      {
        TextBlockConsole.Text = msg;
      }
      else
      {
        TextBlockConsole.Text += Environment.NewLine + msg;
      }
    }
  }
}
