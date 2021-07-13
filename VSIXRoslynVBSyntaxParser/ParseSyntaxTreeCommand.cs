using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.LanguageServices;
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

namespace VSIXRoslynVBSyntaxParser
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ParseSyntaxTreeCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("37109e9e-a444-4e99-b0c9-580c1f4fc76b");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ParseSyntaxTreeCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ParseSyntaxTreeCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static ParseSyntaxTreeCommand Instance
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
            // Switch to the main thread - the call to AddCommand in ParseSyntaxTreeCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ParseSyntaxTreeCommand(package, commandService);
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

            var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
            var workspace = (Workspace)componentModel.GetService<VisualStudioWorkspace>();

            // Get documents in target project
            var documents = from project in workspace.CurrentSolution.Projects
                            from document in project.Documents
                            select document;

            foreach (var document in documents)
            {
                string title = document.Name;    
                var variables = new List<string>();
                var methodParameters = new List<string>();

                SyntaxTree syntaxTree = document.GetSyntaxTreeAsync().Result;
                SyntaxNode syntaxRoot = syntaxTree.GetRoot();
                var nodes = syntaxRoot.DescendantNodes();

                // Get classes
                foreach (var classSyntax in nodes.OfType<ClassBlockSyntax>())
                {
                    // Get variables
                    foreach (var fieldSyntax in classSyntax.Members.OfType<FieldDeclarationSyntax>())
                    {
                        var name = fieldSyntax.Declarators.First().Names;
                        var kind = fieldSyntax.Kind();
                        var type = fieldSyntax.Declarators.First().AsClause;
                        variables.Add($"Name: {name}, Kind: {kind}, Type: {type}");
                    }

                    // Get methods
                    foreach (var methodSyntax in classSyntax.Members.OfType<MethodBlockSyntax>())
                    {
                        MethodStatementSyntax methodStatement = methodSyntax.ChildNodes().First(x => x is MethodStatementSyntax) as MethodStatementSyntax;
                        foreach (var parameterSyntax in methodStatement.ParameterList.Parameters)
                        {
                            var name = parameterSyntax.Identifier;
                            var kind = parameterSyntax.Kind();
                            var type = parameterSyntax.AsClause;
                            methodParameters.Add($"Name: {name}, Kind: {kind}, Type: {type}");
                        }
                    }
                }
                if (variables.Count() == 0 || methodParameters.Count() == 0) continue;

                // Show a message box to prove we were here
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    $"variables: \n{string.Join(",\n", variables)} \n\n" +
                    $"methodParams: \n{string.Join(",\n", methodParameters)}",
                    title,
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }

        }
    }
}
