using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace SelectInsideQuotes
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SelectInsideQuotesCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("51d58fdf-fb06-41ec-bfa7-8da09b084b09");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SelectInsideQuotesCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SelectInsideQuotesCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static SelectInsideQuotesCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command1's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SelectInsideQuotesCommand(package, commandService);
        }

        public static List<int> GetCharPositions(char c, string input)
        {
            var cleaned = input.Replace("\r\n", "\n").Replace("\r", "\n");

            List<int> positions = new List<int>();

            bool start = true; 
            for (int i = 0; i < cleaned.Length; i++)
            {
                if (cleaned[i] == c)
                {
                    if (start)
                    {
                        positions.Add(i+1);
                        start = false;
                    }
                    else
                    {
                        positions.Add((i+2));
                        start = true;
                    }
                }
            }

            return positions;
        }

        public static List<List<T>> ChunkList<T>(List<T> source, int chunkSize)
        {
            List<List<T>> chunks = new List<List<T>>();
            for (int i = 0; i < source.Count; i += chunkSize)
            {
                List<T> chunk = source.GetRange(i, Math.Min(chunkSize, source.Count - i));
                chunks.Add(chunk);
            }
            return chunks;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);

            var dte = (DTE2)ServiceProvider.GetServiceAsync(typeof(DTE)).Result;
            var document = dte.ActiveDocument;
            if (document != null)
            {
                var textDocument = (TextDocument)document.Object("TextDocument");

                var editPoint = textDocument.CreateEditPoint(textDocument.StartPoint);
                string allText = editPoint.GetText(textDocument.EndPoint);

                var doubleQuotes = ChunkList(GetCharPositions('"', allText), 2);

                var selection = textDocument.Selection;
                var startPoint = selection.ActivePoint.CreateEditPoint();
                var endPoint = selection.ActivePoint.CreateEditPoint();

                var curPos = selection.ActivePoint.AbsoluteCharOffset;
                var inDoubleQuotes = doubleQuotes.Any(x => x.Count == 2 && x[0] < curPos && x[1] > curPos);
                
                if (inDoubleQuotes)
                {
                    // Move to the start of the string
                    while (!startPoint.AtStartOfDocument)
                    {
                        var leftChar = startPoint.GetText(-1);
                        if (leftChar == "\"" || leftChar == "'")
                        {
                            break;
                        }
                        else
                        {
                            startPoint.CharLeft(1);
                        }
                    }

                    // Move to the end of the string
                    while (!endPoint.AtEndOfDocument)
                    {
                        var rightChar = endPoint.GetText(1);
                        if (rightChar == "\"" || rightChar == "'")
                        {
                            break;
                        }
                        else
                        {
                            endPoint.CharRight(1);
                        }
                    }

                    // Select text between the start and end points
                    selection.MoveToPoint(startPoint);
                    selection.MoveToPoint(endPoint, true);
                }
            }
        }
    }
}
