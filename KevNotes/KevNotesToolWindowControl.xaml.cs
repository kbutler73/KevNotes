using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using EnvDTE;
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
        private DTE _dte;

        //private readonly string _path;
        private const string _fileName = ".KevNotes.json";

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
                _dte = (DTE)ServiceProvider.GlobalProvider.GetService(typeof(DTE));

                _dteEvents = _dte.Events;
                _slnEvents = _dteEvents.SolutionEvents;
                _slnEvents.Opened += OnSolutionOpened;
                _slnEvents.BeforeClosing += OnSolutionClosing;

                OnSolutionOpened();
            }
            catch (Exception)
            {
            }
        }

        private async void OnSolutionClosing()
        {
            await SaveAsync();
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
                var solutionDir = Path.GetDirectoryName(_dte.Solution.FullName);
                if (solutionDir != null)
                {
                    var data = JsonConvert.DeserializeObject<NotesData>(File.ReadAllText(Path.Combine(solutionDir, _fileName)));

                    tbNotes.Text = data.Note;
                    tbNotes.FontFamily = data.FontFamily;
                    tbNotes.FontSize = data.FontSize;
                    tbNotes.CaretIndex = data.CaretIndex;
                    var lineIndex = tbNotes.GetLineIndexFromCharacterIndex(data.CaretIndex);
                    tbNotes.ScrollToLine(lineIndex);

                    cboFontFamily.SelectedItem = data.FontFamily;
                }
            }
            catch (Exception)
            {
            }
        }

        private async Task SaveAsync()
        {
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

                var solutionDir = Path.GetDirectoryName(_dte.Solution.FullName);
                File.WriteAllText(Path.Combine(solutionDir, _fileName), JsonConvert.SerializeObject(data));
            }
            catch (Exception)
            {
            }
        }

        private async void tbNotes_LostFocus(object sender, RoutedEventArgs e)
        {
            await SaveAsync();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            tbNotes.FontSize--;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            tbNotes.FontSize++;
        }

        private async void cboFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboFontFamily.SelectedItem is FontFamily font)
                //var font = new FontFamily(cboFontFamily.SelectedItem.ToString());
                tbNotes.FontFamily = font;
            else tbNotes.FontFamily = new FontFamily("Ariel");
        }
    }
}