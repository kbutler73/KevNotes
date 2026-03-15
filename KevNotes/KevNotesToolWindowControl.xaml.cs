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
using System.Windows.Documents;
using Task = System.Threading.Tasks.Task;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Input;

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
        private bool _isFormatting;
        private const string NotesFolderName = "kevNotes";
        private string _fileName = string.Empty;
        private static readonly Regex UrlRegex =
            new Regex(@"^(https?://|www\.)\S+$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

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

                SetNoteText(data.Note ?? string.Empty);
                if (data.FontFamily != null)
                {
                    rtbNotes.FontFamily = data.FontFamily;
                }
                if (data.FontSize > 0)
                {
                    rtbNotes.FontSize = data.FontSize;
                }
                if (data.CaretIndex >= 0)
                {
                    SetCaretIndex(data.CaretIndex);
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

        private async void rtbNotes_LostFocus(object sender, RoutedEventArgs e)
        {
            await SaveAsync();
        }

        private void rtbNotes_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isLoading || _isFormatting)
            {
                return;
            }
        }

        private void rtbNotes_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == 0)
            {
                return;
            }

            var textPointer = rtbNotes.GetPositionFromPoint(e.GetPosition(rtbNotes), true);
            if (textPointer == null)
            {
                return;
            }

            var word = GetWordAtPointer(textPointer);
            if (string.IsNullOrWhiteSpace(word))
            {
                return;
            }

            var url = NormalizeUrl(word);
            if (!UrlRegex.IsMatch(url))
            {
                return;
            }

            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true });
                e.Handled = true;
            }
            catch (Exception ex)
            {
                WriteOutputAsync($"Error opening link. {ex.Message}");
            }
        }

        private NotesData CreateSnapshot()
        {
            return new NotesData
            {
                CaretIndex = GetCaretIndex(),
                FontFamily = rtbNotes.FontFamily,
                FontSize = rtbNotes.FontSize,
                Note = GetNoteText()
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
            rtbNotes.FontSize--;
            await SaveAsync();
        }

        private async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            rtbNotes.FontSize++;
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
                rtbNotes.FontFamily = font;
            }
            else
            {
                rtbNotes.FontFamily = new FontFamily("Arial");
            }
            await SaveAsync();
        }

        private void rtbNotes_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
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
            var text = GetNoteText();
            var startIndex = Math.Max(GetSelectionStart() + GetSelectionLength(), 0);
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
            var text = GetNoteText();
            var startIndex = Math.Max(GetSelectionStart() - 1, 0);
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
            var selection = GetSelectedText();
            if (selection.Length > 0 && string.Equals(selection, term, comparison))
            {
                var replaceWith = tbReplace.Text ?? string.Empty;
                var selectionStart = GetSelectionStart();
                ReplaceSelectionText(replaceWith);
                SetSelection(selectionStart, replaceWith.Length);
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
            var text = GetNoteText();
            var comparison = GetFindComparison();
            var updated = ReplaceAllText(text, term, replaceWith, comparison);
            SetNoteText(updated);
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
            rtbNotes.Focus();
            SetSelection(index, length);
        }

        private string GetNoteText()
        {
            var range = new System.Windows.Documents.TextRange(rtbNotes.Document.ContentStart, rtbNotes.Document.ContentEnd);
            return NormalizeText(range.Text);
        }

        private void SetNoteText(string text)
        {
            _isFormatting = true;
            try
            {
                rtbNotes.Document.Blocks.Clear();
                rtbNotes.Document.Blocks.Add(new Paragraph(new Run(text)));
            }
            finally
            {
                _isFormatting = false;
            }
        }

        private int GetCaretIndex()
        {
            return GetTextOffset(rtbNotes.Document.ContentStart, rtbNotes.CaretPosition);
        }

        private void SetCaretIndex(int index)
        {
            var position = GetTextPointerAtOffset(index);
            if (position != null)
            {
                rtbNotes.CaretPosition = position;
            }
        }

        private int GetSelectionStart()
        {
            return GetTextOffset(rtbNotes.Document.ContentStart, rtbNotes.Selection.Start);
        }

        private int GetSelectionLength()
        {
            return Math.Max(0, GetTextOffset(rtbNotes.Selection.Start, rtbNotes.Selection.End));
        }

        private string GetSelectedText()
        {
            var range = new System.Windows.Documents.TextRange(rtbNotes.Selection.Start, rtbNotes.Selection.End);
            return NormalizeText(range.Text);
        }

        private void ReplaceSelectionText(string replacement)
        {
            rtbNotes.Selection.Text = replacement;
        }

        private void SetSelection(int start, int length)
        {
            var startPos = GetTextPointerAtOffset(start);
            var endPos = GetTextPointerAtOffset(start + length);
            if (startPos != null && endPos != null)
            {
                rtbNotes.Selection.Select(startPos, endPos);
            }
        }

        private static string NormalizeUrl(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return text;
            }

            if (text.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                text.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                return text;
            }

            if (text.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
            {
                return "https://" + text;
            }

            return text;
        }

        private static string GetWordAtPointer(TextPointer pointer)
        {
            var wordStart = pointer;
            var wordEnd = pointer;

            while (wordStart != null && !IsWordBoundary(wordStart, LogicalDirection.Backward))
            {
                wordStart = wordStart.GetPositionAtOffset(-1, LogicalDirection.Backward);
            }

            while (wordEnd != null && !IsWordBoundary(wordEnd, LogicalDirection.Forward))
            {
                wordEnd = wordEnd.GetPositionAtOffset(1, LogicalDirection.Forward);
            }

            if (wordStart == null || wordEnd == null)
            {
                return string.Empty;
            }

            var range = new System.Windows.Documents.TextRange(wordStart, wordEnd);
            return range.Text.Trim();
        }

        private static bool IsWordBoundary(TextPointer pointer, LogicalDirection direction)
        {
            var context = pointer.GetPointerContext(direction);
            if (context != TextPointerContext.Text)
            {
                return true;
            }

            var text = pointer.GetTextInRun(direction);
            if (string.IsNullOrEmpty(text))
            {
                return true;
            }

            var ch = direction == LogicalDirection.Forward ? text[0] : text[text.Length - 1];
            return char.IsWhiteSpace(ch) || ch == '(' || ch == ')' || ch == '[' || ch == ']' || ch == '{' || ch == '}' || ch == '<' || ch == '>' || ch == '"' || ch == '\'' || ch == ',';
        }

        private TextPointer GetTextPointerAtOffset(int offset)
        {
            var navigator = rtbNotes.Document.ContentStart;
            var count = 0;
            while (navigator != null)
            {
                if (navigator.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                {
                    var textRun = navigator.GetTextInRun(LogicalDirection.Forward);
                    if (count + textRun.Length >= offset)
                    {
                        return navigator.GetPositionAtOffset(offset - count);
                    }
                    count += textRun.Length;
                    navigator = navigator.GetPositionAtOffset(textRun.Length);
                }
                else
                {
                    navigator = navigator.GetNextContextPosition(LogicalDirection.Forward);
                }
            }

            return rtbNotes.Document.ContentEnd;
        }

        private int GetTextOffset(TextPointer start, TextPointer end)
        {
            if (start == null || end == null)
            {
                return 0;
            }

            return new System.Windows.Documents.TextRange(start, end).Text.Length;
        }

        private static string NormalizeText(string text)
        {
            if (text.EndsWith("\r\n", StringComparison.Ordinal))
            {
                return text.Substring(0, text.Length - 2);
            }

            return text;
        }

        private static string ReplaceAllText(string text, string term, string replaceWith, StringComparison comparison)
        {
            if (string.IsNullOrEmpty(term))
            {
                return text;
            }

            var comparisonType = comparison == StringComparison.CurrentCulture
                ? StringComparison.CurrentCulture
                : StringComparison.CurrentCultureIgnoreCase;

            var result = new StringBuilder(text.Length);
            var index = 0;
            while (true)
            {
                var matchIndex = text.IndexOf(term, index, comparisonType);
                if (matchIndex < 0)
                {
                    result.Append(text.Substring(index));
                    break;
                }

                result.Append(text.Substring(index, matchIndex - index));
                result.Append(replaceWith);
                index = matchIndex + term.Length;
            }

            return result.ToString();
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
