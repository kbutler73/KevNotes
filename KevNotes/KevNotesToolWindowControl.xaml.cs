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
using System.Text.RegularExpressions;
using Task = System.Threading.Tasks.Task;

namespace KevNotes
{
    /// <summary>
    /// Interaction logic for KevNotesToolWindowControl.
    /// </summary>
    public partial class KevNotesToolWindowControl : UserControl
    {
        private static readonly Lazy<FontFamily[]> FontFamilies =
            new Lazy<FontFamily[]>(() => Fonts.SystemFontFamilies.OrderBy(x => x.ToString()).ToArray());

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
            ShowFindBar(showReplace: true);

            foreach (var fontFamily in FontFamilies.Value)
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
            var snapshot = CreateSnapshot();
            await SaveSnapshotAsync(snapshot);
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
                if (!File.Exists(_fileName))
                {
                    return;
                }

                await WriteOutputAsync($"Loading Notes from {_fileName}");

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

            try
            {
                var data = CreateSnapshot();
                await SaveSnapshotAsync(data);
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

        private NotesData CreateSnapshot()
        {
            return new NotesData
            {
                CaretIndex = tbNotes.CaretIndex,
                FontFamily = tbNotes.FontFamily,
                FontSize = tbNotes.FontSize,
                Note = tbNotes.Text ?? string.Empty
            };
        }

        private async Task SaveSnapshotAsync(NotesData data)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            if (data == null || string.IsNullOrWhiteSpace(_fileName))
            {
                return;
            }

            var path = Path.GetDirectoryName(_fileName);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            File.WriteAllText(_fileName, JsonConvert.SerializeObject(data));
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

        private void tbNotes_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.F3)
            {
                if (e.KeyboardDevice.Modifiers == System.Windows.Input.ModifierKeys.Shift)
                {
                    FindPrevious();
                }
                else
                {
                    FindNext();
                }
                e.Handled = true;
            }
        }

        private void tbFind_TextChanged(object sender, TextChangedEventArgs e)
        {
        }

        private void tbFind_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                FindNext();
                e.Handled = true;
            }
        }

        private void tbReplace_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                ReplaceNext();
                e.Handled = true;
            }
        }

        private void FindNext_Click(object sender, RoutedEventArgs e)
        {
            FindNext();
        }

        private void Replace_Click(object sender, RoutedEventArgs e)
        {
            ReplaceNext();
        }

        private void ReplaceAll_Click(object sender, RoutedEventArgs e)
        {
            ReplaceAll();
        }

        private void ShowFindBar(bool showReplace)
        {
            ReplaceRow.Visibility = showReplace ? Visibility.Visible : Visibility.Collapsed;
            tbReplace.IsEnabled = showReplace;
        }

        private void FindNext()
        {
            if (!TryGetFindText(out var term))
            {
                return;
            }

            var comparison = GetFindComparison();
            var text = tbNotes.Text ?? string.Empty;
            var startIndex = Math.Max(tbNotes.SelectionStart + tbNotes.SelectionLength, 0);
            var index = text.IndexOf(term, startIndex, comparison);
            if (index < 0 && startIndex > 0)
            {
                index = text.IndexOf(term, 0, comparison);
            }

            if (index >= 0)
            {
                SelectMatch(index, term.Length);
            }
        }

        private void FindPrevious()
        {
            if (!TryGetFindText(out var term))
            {
                return;
            }

            var comparison = GetFindComparison();
            var text = tbNotes.Text ?? string.Empty;
            var startIndex = Math.Max(tbNotes.SelectionStart - 1, 0);
            var index = text.LastIndexOf(term, startIndex, comparison);
            if (index < 0 && text.Length > 0)
            {
                index = text.LastIndexOf(term, text.Length - 1, comparison);
            }

            if (index >= 0)
            {
                SelectMatch(index, term.Length);
            }
        }

        private void ReplaceNext()
        {
            if (!TryGetFindText(out var term))
            {
                return;
            }

            var comparison = GetFindComparison();
            var selection = tbNotes.SelectedText ?? string.Empty;
            if (selection.Length > 0 && string.Equals(selection, term, comparison))
            {
                var replaceWith = tbReplace.Text ?? string.Empty;
                var selectionStart = tbNotes.SelectionStart;
                tbNotes.SelectedText = replaceWith;
                tbNotes.SelectionStart = selectionStart;
                tbNotes.SelectionLength = replaceWith.Length;
                FindNext();
            }
            else
            {
                FindNext();
            }
        }

        private void ReplaceAll()
        {
            if (!TryGetFindText(out var term))
            {
                return;
            }

            var replaceWith = tbReplace.Text ?? string.Empty;
            var comparison = GetFindComparison();
            var options = comparison == StringComparison.CurrentCulture
                ? RegexOptions.None
                : RegexOptions.IgnoreCase;

            var text = tbNotes.Text ?? string.Empty;
            var regex = new Regex(Regex.Escape(term), options);
            tbNotes.Text = regex.Replace(text, replaceWith);
        }

        private bool TryGetFindText(out string term)
        {
            term = tbFind.Text;
            return !string.IsNullOrEmpty(term);
        }

        private StringComparison GetFindComparison()
        {
            return cbMatchCase.IsChecked == true
                ? StringComparison.CurrentCulture
                : StringComparison.CurrentCultureIgnoreCase;
        }

        private void SelectMatch(int index, int length)
        {
            tbNotes.Focus();
            tbNotes.SelectionStart = index;
            tbNotes.SelectionLength = length;
            tbNotes.ScrollToLine(tbNotes.GetLineIndexFromCharacterIndex(index));
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