using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace VSIXProject1
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command1
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("ec6c7294-7f23-4646-b8c9-84c8f3101ae6");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Command1(AsyncPackage package, OleMenuCommandService commandService)
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
        public static Command1 Instance
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
            Instance = new Command1(package, commandService);
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

            GenerateProject(
                "C:\\Users\\kav128\\Documents\\Учеба\\5 семестр\\Инструментальные средства\\Lab1\\Lab\\bin\\Debug\\netcoreapp3.0\\InstrLab.xml",
                "C:\\Users\\kav128\\Documents\\Учеба\\5 семестр\\Инструментальные средства\\Lab1\\Lab\\bin\\Debug\\netcoreapp3.0\\InstrLab.dll");

            Thread thread = new Thread(CompileDocumentation);
            thread.Start();
        }

        private static string GenerateProject(string xmlSource, string binSource)
        {
            XNamespace xmlns = "http://schemas.microsoft.com/developer/msbuild/2003";
            XDocument doc = new XDocument(
                new XDeclaration("1.0", "utf-16", ""),
                new XElement(
                    xmlns + "Project",
                    new XAttribute("ToolsVersion", "14.0"),
                    new XAttribute("DefaultTargets", "Build"),
                    new XElement(
                        xmlns + "Import",
                        new XAttribute("Project", "$(MSBuildExtensionsPath)\\$(MSBuildToolsVersion)\\Microsoft.Common.props"),
                        new XAttribute("Condition", "Exists('$(MSBuildExtensionsPath)\\$(MSBuildToolsVersion)\\Microsoft.Common.props')")
                    ),
                    new XElement(
                        xmlns + "PropertyGroup",
                        new XElement(xmlns + "SchemaVersion", "2.0"),
                        new XElement(xmlns + "ProjectGuid", Guid.NewGuid()),
                        new XElement(xmlns + "SHFBSchemaVersion", "2017.9.26.0"),
                        new XElement(xmlns + "Name", "Documentation"),
                        new XElement(xmlns + "FrameworkVersion", ".NET Framework 4.5"),
                        new XElement(xmlns + "OutputPath", ".\\Help\\"),
                        new XElement(xmlns + "HtmlHelpName", "Documentation"),
                        new XElement(xmlns + "Language", "ru-RU"),
                        new XElement(
                            xmlns + "DocumentationSources",
                            new XElement(
                                xmlns + "DocumentationSource",
                                new XAttribute("sourceFile", xmlSource)
                            ),
                            new XElement(
                                xmlns + "DocumentationSource",
                                new XAttribute("sourceFile", binSource)
                            )
                        ),
                        new XElement(xmlns + "HelpFileFormat", "HtmlHelp1"),
                        new XElement(xmlns + "SyntaxFilters", "Standard"),
                        new XElement(xmlns + "PresentationStyle", "VS2013"),
                        new XElement(xmlns + "CleanIntermediates", true),
                        new XElement(xmlns + "KeepLogFile", true),
                        new XElement(xmlns + "DisableCodeBlockComponent", false),
                        new XElement(xmlns + "IndentHtml", true),
                        new XElement(xmlns + "BuildAssemblerVerbosity", "OnlyWarningsAndErrors"),
                        new XElement(xmlns + "SaveComponentCacheCapacity", 100)
                    ),
                    new XElement(
                        xmlns + "Import",
                        new XAttribute("Project", "$(MSBuildToolsPath)\\Microsoft.Common.targets"),
                        new XAttribute("Condition", "'$(MSBuildRestoreSessionId)' != ''")
                    ),
                    new XElement(
                        xmlns + "Import",
                        new XAttribute("Project", "$(SHFBROOT)\\SandcastleHelpFileBuilder.targets"),
                        new XAttribute("Condition", "'$(MSBuildRestoreSessionId)' == ''")
                    )

                )
            );
            doc.Root.ReplaceAttributes(null);
            doc.Save("test1.shfbproj");
            return "test1.shfbproj";
        }

        private static void CompileDocumentation()
        {
            Process process = new Process();
            ProcessStartInfo info = new ProcessStartInfo()
            {
                FileName = "C:\\Program Files (x86)\\MSBuild\\14.0\\Bin\\MSBuild.exe",
                Arguments = "test1.shfbproj",
                CreateNoWindow = false
            };
            process.StartInfo = info;
            process.Start();
            process.WaitForExit();

            Process.Start(".\\Help\\Documentation.chm");
        }
    }
}
