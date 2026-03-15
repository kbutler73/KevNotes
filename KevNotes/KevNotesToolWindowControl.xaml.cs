using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System.Linq;
using Newtonsoft.Json;
using Task = System.Threading.Tasks.Task;

namespace KevNotes
{
    /// <summary>
    /// Interaction logic for KevNotesToolWindowControl.
    /// </summary>
    public partial class KevNotesToolWindowControl : UserControl
    {
        private Events _dteEvents;
        private SolutionEvents _slnEvents;
        private DTE2 _dte;
        private OutputWindowPane _outputPane;
        private bool _isLoading;

        private const string NotesFolderName = "kevNotes";
        private string _fileName = string.Empty;

        /// <summary>
        /// Initializes a new instance of the <see cref="KevNotesToolWindowControl"/> class.
        /// </summary>
        public KevNotesToolWindowControl()
        {
            Dispatcher.VerifyAccess();
            this.InitializeComponent();

            foreach (var fontFamily in Fonts.SystemFontFamilies.OrderBy(x => x.ToString()))
            {
                cboFontFamily.Items.Add(fontFamily);
            }

            try
            {
                _dte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));

                _dteEvents = _dte.Events;
                _slnEvents = _dteEvents.SolutionEvents;
                _slnEvents.Opened += OnSolutionOpened;
                _slnEvents.BeforeClosing += OnSolutionClosing;

                OutputWindowPanes panes = _dte.ToolWindows.OutputWindow.OutputWindowPanes;
                try
                {
                    _outputPane = panes.Item("KevNotes");
                }
                catch (ArgumentException)
                {
                    _outputPane = panes.Add("KevNotes");
                }

                OnSolutionOpened();
            }
            catch (Exception ex)
            {
                WriteOutputAsync(ex.Message);
            }
        }

        private async Task WriteOutputAsync(string message)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            _outputPane?.OutputString($"{message}{Environment.NewLine}");
        }

        private async void OnSolutionClosing()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            await WriteOutputAsync($"Closing {_dte.Solution.FullName}");
            await SaveAsync();
            tbNotes.Clear();
        }

        private async void OnSolutionOpened()
        {
            await LoadAsync();
        }

        private async Task LoadAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            if (_isLoading)
            {
                return;
            }

            try
            {
                var solutionName = _dte.Solution.FullName;
                if (string.IsNullOrWhiteSpace(solutionName)) return;

                _isLoading = true;
                _fileName = GetNotesFilePath(solutionName);

                await WriteOutputAsync($"Loading {solutionName}");
                await WriteOutputAsync($"Looking for notes in {_fileName}");
                if (!File.Exists(_fileName))
                {
                    tbNotes.Clear();
                    cboFontFamily.SelectedItem = null;
                    return;
                }

                var data = JsonConvert.DeserializeObject<NotesData>(File.ReadAllText(_fileName));
                if (data == null)
                {
                    return;
                }

                tbNotes.Text = data.Note ?? string.Empty;
                if (data.FontFamily != null)
                {
                    tbNotes.FontFamily = data.FontFamily;
                }
                if (data.FontSize > 0)
                {
                    tbNotes.FontSize = data.FontSize;
                }
                if (data.CaretIndex >= 0)
                {
                    tbNotes.CaretIndex = data.CaretIndex;
                    var lineIndex = tbNotes.GetLineIndexFromCharacterIndex(data.CaretIndex);
                    tbNotes.ScrollToLine(lineIndex);
                }

                if (data.FontFamily != null)
                {
                    cboFontFamily.SelectedItem = data.FontFamily;
                }
            }
            catch (Exception ex)
            {
                await WriteOutputAsync($"Error Loading KevNotes. {ex.Message}");
            }
            finally
            {
                _isLoading = false;
            }
        }

        private async Task SaveAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            if (_isLoading || string.IsNullOrWhiteSpace(_fileName))
            {
                return;
            }

            //if (string.IsNullOrWhiteSpace(tbNotes.Text))
            //{
            //    await WriteOutputAsync("Skipped saving because there is no text.");
            //    return;
            //}

            try
            {
                var data = new NotesData
                {
                    CaretIndex = tbNotes.CaretIndex,
                    FontFamily = tbNotes.FontFamily,
                    FontSize = tbNotes.FontSize,
                    Note = tbNotes.Text
                };

                var path = Path.GetDirectoryName(_fileName);
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                File.WriteAllText(_fileName, JsonConvert.SerializeObject(data));
#if DEBUG
                await WriteOutputAsync($"KevNotes Saved to {_fileName}");
#endif
            }
            catch (Exception ex)
            {
                await WriteOutputAsync($"Error saving KevNotes. {ex.Message}");
            }
        }

        private async void tbNotes_LostFocus(object sender, RoutedEventArgs e)
        {
            await SaveAsync();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            tbNotes.FontSize--;
            await SaveAsync();
        }

        private async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            tbNotes.FontSize++;
            await SaveAsync();
        }

        private async void cboFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isLoading)
            {
                return;
            }

            if (cboFontFamily.SelectedItem is FontFamily font)
            {
                tbNotes.FontFamily = font;
            }
            else
            {
                tbNotes.FontFamily = new FontFamily("Arial");
            }
            await SaveAsync();
        }

        private static string GetNotesFilePath(string solutionFullName)
        {
            var solutionDir = Path.GetDirectoryName(solutionFullName);
            var solutionFile = Path.GetFileNameWithoutExtension(solutionFullName);
            var jsonName = $"{solutionFile}.json";
            var solutionPath = solutionDir == null ? jsonName : Path.Combine(solutionDir, jsonName);
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, NotesFolderName, solutionPath.Replace(":", ""));
        }
    }
}