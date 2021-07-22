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

        //private readonly string _path;
        private string _fileName = "";

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

            _outputPane.OutputString($"{message}{Environment.NewLine}");
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
            try
            {
                var solutionName = _dte.Solution.FullName;
                if (string.IsNullOrWhiteSpace(solutionName)) return;
                
                var solutionDir = Path.Combine(Path.GetDirectoryName(solutionName), Path.GetFileNameWithoutExtension(solutionName) + ".json");

                await WriteOutputAsync($"Loading {solutionName}");
                if (solutionDir != null)
                {
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    _fileName = Path.Combine(appData, "kevNotes", solutionDir.Replace(":", ""));
                    if (!File.Exists(_fileName)) return;
                    var data = JsonConvert.DeserializeObject<NotesData>(File.ReadAllText(_fileName)); 

                    tbNotes.Text = data.Note;
                    tbNotes.FontFamily = data.FontFamily;
                    tbNotes.FontSize = data.FontSize;
                    tbNotes.CaretIndex = data.CaretIndex;
                    var lineIndex = tbNotes.GetLineIndexFromCharacterIndex(data.CaretIndex);
                    tbNotes.ScrollToLine(lineIndex);

                    cboFontFamily.SelectedItem = data.FontFamily;
                }
            }
            catch (Exception ex)
            {
                await WriteOutputAsync($"Error Loading KevNotes. {ex.Message}");
            }
        }

        private async Task SaveAsync()
        {
            if (tbNotes.Text.Length == 0)
            {
                await WriteOutputAsync("Skipped saving because there is no text.");
                return;
            }

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
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
            if (cboFontFamily.SelectedItem is FontFamily font)
            { 
                tbNotes.FontFamily = font; 
            }
            else
            {
                tbNotes.FontFamily = new FontFamily("Ariel");
            }
            await SaveAsync();
        }
    }
}