using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Tools;
using FlaUI.UIA3;
using Microsoft.Win32;
using NUnit.Framework;
using OfficeRibbonXEditor.Common;
using OfficeRibbonXEditor.UITests.Extensions;

namespace OfficeRibbonXEditor.UITests.Main;

[TestFixture]
[SingleThreaded]
[Apartment(ApartmentState.STA)]
[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "UI tests target Windows")]
public sealed class CustomUi14PowerShellE2ETests
{
    private const string CustomUi14Namespace = "http://schemas.microsoft.com/office/2009/07/customui";
    private const int XlOpenXmlWorkbookMacroEnabled = 52;
    private const int MsoAutomationSecurityLow = 1;
    private const int VbextCtStdModule = 1;

    private string _testFolder = string.Empty;

    [SetUp]
    public void SetUp()
    {
        _testFolder = Path.Combine(Path.GetTempPath(), "OfficeRibbonXEditorE2E", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_testFolder);
    }

    [TearDown]
    public void TearDown()
    {
        try
        {
            if (Directory.Exists(_testFolder))
            {
                Directory.Delete(_testFolder, true);
            }
        }
        catch (IOException)
        {
            TestContext.WriteLine($"Could not delete test folder '{_testFolder}'.");
        }
        catch (UnauthorizedAccessException)
        {
            TestContext.WriteLine($"Could not delete test folder '{_testFolder}'.");
        }
    }

    [Test]
    public void RibbonXBaselineCanBeUpdatedByPowerShellToolAndVerifiedInExcel()
    {
        Assume.That(IsExcelInstalled(), Is.True, "Excel is required for this end-to-end UI test.");
        Assume.That(IsVbaProjectAccessTrusted(), Is.True, "Excel Trust Center setting 'Trust access to the VBA project object model' is required.");

        var repositoryRoot = FindRepositoryRoot();
        var workbookPath = Path.Combine(_testFolder, "RibbonXBaseline.xlsm");
        var updatedWorkbookPath = Path.Combine(_testFolder, "RibbonXUpdated.xlsm");
        var definitionPath = Path.Combine(_testFolder, "update-def.txt");

        CreateMacroEnabledWorkbook(workbookPath);
        AddInitialRibbonXCustomUi(workbookPath);
        AddOffice2007CustomUiRegressionPart(workbookPath);

        AssertExcelRibbonButtons(
            workbookPath,
            tabLabel: "RX E2E",
            groupLabel: "Baseline Actions",
            expectedCallbacks: new Dictionary<string, string>
            {
                ["Export A"] = "ExportA",
                ["Sync A"] = "SyncA",
            });

        var office2007CustomUiBefore = ReadCustomPart(workbookPath, XmlPart.RibbonX12);
        WriteUpdateDefinition(definitionPath);

        var firstUpdate = RunPowerShellTool(repositoryRoot, workbookPath, definitionPath, updatedWorkbookPath);
        Assert.That(firstUpdate.ExitCode, Is.EqualTo(0), firstUpdate.AllOutput);
        Assert.That(File.Exists(updatedWorkbookPath), Is.True, "The PowerShell tool did not create the requested output workbook.");
        Assert.That(ReadCustomPart(workbookPath, XmlPart.RibbonX14), Does.Contain("Export A"), "The source workbook should stay unchanged when OutputPath is used.");
        Assert.That(ReadCustomPart(updatedWorkbookPath, XmlPart.RibbonX12), Is.EqualTo(office2007CustomUiBefore), "The Office 2007 customUI.xml part must not be modified.");

        AssertCustomUi14Buttons(
            updatedWorkbookPath,
            expectedTabLabel: "RX E2E Updated",
            expectedGroupLabel: "Updated Actions",
            expectedButtons: new[]
            {
                new ExpectedButton("btnExport", "Export B", "large", "FileSave", "ExportB"),
                new ExpectedButton("btnSync", "Sync B", "normal", "Repeat", "SyncB"),
                new ExpectedButton("btnClean", "Clean", "normal", "ClearFormatting", "CleanB"),
            });

        var secondUpdate = RunPowerShellTool(repositoryRoot, updatedWorkbookPath, definitionPath, updatedWorkbookPath, forceInPlace: true);
        Assert.That(secondUpdate.ExitCode, Is.EqualTo(0), secondUpdate.AllOutput);
        Assert.That(secondUpdate.AllOutput, Does.Contain("CreatedButtons: 0"));
        Assert.That(secondUpdate.AllOutput, Does.Contain("UpdatedButtons: 3"));

        AssertExcelRibbonButtons(
            updatedWorkbookPath,
            tabLabel: "RX E2E Updated",
            groupLabel: "Updated Actions",
            expectedCallbacks: new Dictionary<string, string>
            {
                ["Export B"] = "ExportB",
                ["Sync B"] = "SyncB",
                ["Clean"] = "CleanB",
            });

        PreserveWorkbookArtifact(repositoryRoot, updatedWorkbookPath);
    }

    [Test]
    public void PowerShellToolReturnsErrorsForInvalidDefinitionsAndLockedWorkbooks()
    {
        Assume.That(IsExcelInstalled(), Is.True, "Excel is required to create the macro-enabled workbook fixture.");
        Assume.That(IsVbaProjectAccessTrusted(), Is.True, "Excel Trust Center setting 'Trust access to the VBA project object model' is required.");

        var repositoryRoot = FindRepositoryRoot();
        var workbookPath = Path.Combine(_testFolder, "RibbonXBaseline.xlsm");
        CreateMacroEnabledWorkbook(workbookPath);
        AddInitialRibbonXCustomUi(workbookPath);

        var invalidSizeDefinition = Path.Combine(_testFolder, "invalid-size-def.txt");
        File.WriteAllText(
            invalidSizeDefinition,
            """
            #tab_id=tabRxE2E
            #tab_label=RX E2E
            #group_id=grpBaseline
            #group_label=Baseline Actions

            id;label;size;icon;onAction
            btnBad;Bad;big;FileSave;ExportB
            """);

        var invalidSizeResult = RunPowerShellTool(repositoryRoot, workbookPath, invalidSizeDefinition, Path.Combine(_testFolder, "InvalidSize.xlsm"));
        Assert.That(invalidSizeResult.ExitCode, Is.Not.EqualTo(0));
        Assert.That(invalidSizeResult.AllOutput, Does.Contain("size muss 'normal' oder 'large' sein"));

        var duplicateIdDefinition = Path.Combine(_testFolder, "duplicate-id-def.txt");
        File.WriteAllText(
            duplicateIdDefinition,
            """
            #tab_id=tabRxE2E
            #tab_label=RX E2E
            #group_id=grpBaseline
            #group_label=Baseline Actions

            id;label;size;icon;onAction
            btnSame;One;large;FileSave;ExportB
            btnSame;Two;normal;Repeat;SyncB
            """);

        var duplicateIdResult = RunPowerShellTool(repositoryRoot, workbookPath, duplicateIdDefinition, Path.Combine(_testFolder, "DuplicateId.xlsm"));
        Assert.That(duplicateIdResult.ExitCode, Is.Not.EqualTo(0));
        Assert.That(duplicateIdResult.AllOutput, Does.Contain("Doppelte id 'btnSame'"));

        var validDefinition = Path.Combine(_testFolder, "valid-def.txt");
        WriteUpdateDefinition(validDefinition);
        using var lockStream = new FileStream(workbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var lockedWorkbookResult = RunPowerShellTool(repositoryRoot, workbookPath, validDefinition, workbookPath, forceInPlace: true);
        Assert.That(lockedWorkbookResult.ExitCode, Is.Not.EqualTo(0));
        Assert.That(lockedWorkbookResult.AllOutput, Does.Contain("Excel"));
    }

    private static void CreateMacroEnabledWorkbook(string workbookPath)
    {
        using var session = ExcelSession.CreateVisible(visible: false);
        dynamic excel = session.Excel;
        excel.DisplayAlerts = false;

        dynamic workbook = excel.Workbooks.Add();
        session.Workbook = workbook;
        workbook.Worksheets[1].Range("A1").Value2 = string.Empty;

        dynamic module = workbook.VBProject.VBComponents.Add(VbextCtStdModule);
        module.CodeModule.AddFromString(
            """
            Option Explicit

            Public Sub ExportA(control)
                ThisWorkbook.Worksheets(1).Range("A1").Value = "ExportA"
            End Sub

            Public Sub SyncA(control)
                ThisWorkbook.Worksheets(1).Range("A1").Value = "SyncA"
            End Sub

            Public Sub ExportB(control)
                ThisWorkbook.Worksheets(1).Range("A1").Value = "ExportB"
            End Sub

            Public Sub SyncB(control)
                ThisWorkbook.Worksheets(1).Range("A1").Value = "SyncB"
            End Sub

            Public Sub CleanB(control)
                ThisWorkbook.Worksheets(1).Range("A1").Value = "CleanB"
            End Sub
            """);

        workbook.SaveAs(workbookPath, XlOpenXmlWorkbookMacroEnabled);
    }

    private static void AddInitialRibbonXCustomUi(string workbookPath)
    {
        using var document = new OfficeDocument(workbookPath);
        var part = document.RetrieveCustomPart(XmlPart.RibbonX14) ?? document.CreateCustomPart(XmlPart.RibbonX14);
        part.Save(
            """
            <?xml version="1.0" encoding="utf-8"?>
            <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
              <ribbon>
                <tabs>
                  <tab id="tabRxE2E" label="RX E2E" insertAfterMso="TabHome">
                    <group id="grpBaseline" label="Baseline Actions">
                      <button id="btnExport" label="Export A" size="large" imageMso="FileSaveAs" onAction="ExportA" />
                      <button id="btnSync" label="Sync A" size="normal" imageMso="RefreshAll" onAction="SyncA" />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>
            """);
        document.Save();
    }

    private static void AddOffice2007CustomUiRegressionPart(string workbookPath)
    {
        using var document = new OfficeDocument(workbookPath);
        var part = document.RetrieveCustomPart(XmlPart.RibbonX12) ?? document.CreateCustomPart(XmlPart.RibbonX12);
        part.Save(
            """
            <?xml version="1.0" encoding="utf-8"?>
            <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
              <ribbon>
                <tabs>
                  <tab id="tabLegacy" label="Legacy UI" />
                </tabs>
              </ribbon>
            </customUI>
            """);
        document.Save();
    }

    private static void WriteUpdateDefinition(string definitionPath)
    {
        File.WriteAllText(
            definitionPath,
            """
            #tab_id=tabRxE2E
            #tab_label=RX E2E Updated
            #group_id=grpBaseline
            #group_label=Updated Actions
            #insert_after_mso=TabHome

            id;label;size;icon;onAction
            btnExport;Export B;large;FileSave;ExportB
            btnSync;Sync B;normal;Repeat;SyncB
            btnClean;Clean;normal;ClearFormatting;CleanB
            """);
    }

    private static void AssertCustomUi14Buttons(string workbookPath, string expectedTabLabel, string expectedGroupLabel, IReadOnlyCollection<ExpectedButton> expectedButtons)
    {
        var customUi = ReadCustomPart(workbookPath, XmlPart.RibbonX14);
        var document = XDocument.Parse(customUi);
        var ns = XNamespace.Get(CustomUi14Namespace);
        var tab = document.Root?.Element(ns + "ribbon")?.Element(ns + "tabs")?.Elements(ns + "tab").SingleOrDefault(x => (string?)x.Attribute("id") == "tabRxE2E");
        Assert.That(tab, Is.Not.Null, "Expected CustomUI14 tab was not found.");
        Assert.That((string?)tab!.Attribute("label"), Is.EqualTo(expectedTabLabel));

        var group = tab.Elements(ns + "group").SingleOrDefault(x => (string?)x.Attribute("id") == "grpBaseline");
        Assert.That(group, Is.Not.Null, "Expected CustomUI14 group was not found.");
        Assert.That((string?)group!.Attribute("label"), Is.EqualTo(expectedGroupLabel));

        var actualButtons = group.Elements(ns + "button").ToList();
        Assert.That(actualButtons.Select(x => (string?)x.Attribute("id")).Distinct().Count(), Is.EqualTo(actualButtons.Count), "Button IDs must be unique.");
        Assert.That(actualButtons, Has.Count.EqualTo(expectedButtons.Count));

        foreach (var expectedButton in expectedButtons)
        {
            var button = actualButtons.SingleOrDefault(x => (string?)x.Attribute("id") == expectedButton.Id);
            Assert.That(button, Is.Not.Null, $"Button '{expectedButton.Id}' was not found.");
            Assert.That((string?)button!.Attribute("label"), Is.EqualTo(expectedButton.Label));
            Assert.That((string?)button.Attribute("size"), Is.EqualTo(expectedButton.Size));
            Assert.That((string?)button.Attribute("imageMso"), Is.EqualTo(expectedButton.ImageMso));
            Assert.That((string?)button.Attribute("onAction"), Is.EqualTo(expectedButton.OnAction));
        }
    }

    private static void AssertExcelRibbonButtons(string workbookPath, string tabLabel, string groupLabel, IReadOnlyDictionary<string, string> expectedCallbacks)
    {
        using var session = ExcelSession.Open(workbookPath);
        using var automation = new UIA3Automation();
        using var application = Application.Attach(session.ProcessId);

        var window = application.GetMainWindow(automation, TimeSpan.FromSeconds(20));
        Assert.That(window, Is.Not.Null, "Excel main window was not found.");
        window!.Focus();
        TryMaximize(window);

        var tab = FindExcelRibbonElement(window, tabLabel, ControlType.TabItem, TimeSpan.FromSeconds(20));
        Assert.That(tab, Is.Not.Null, $"Excel ribbon tab '{tabLabel}' was not found.");
        tab!.Click();

        var group = FindExcelRibbonElement(window, groupLabel, ControlType.Group, TimeSpan.FromSeconds(10)) ??
            FindExcelRibbonElementByName(window, groupLabel, TimeSpan.FromSeconds(5));
        Assert.That(group, Is.Not.Null, $"Excel ribbon group '{groupLabel}' was not found.");

        foreach (var (buttonLabel, expectedMarker) in expectedCallbacks)
        {
            var button = FindExcelRibbonElement(window, buttonLabel, ControlType.Button, TimeSpan.FromSeconds(10));
            Assert.That(button, Is.Not.Null, $"Excel ribbon button '{buttonLabel}' was not found.");

            session.SetMarker(string.Empty);
            button!.Click();
            var marker = Retry.While(
                () => session.Marker,
                value => value != expectedMarker,
                TimeSpan.FromSeconds(10),
                TimeSpan.FromMilliseconds(250)).Result;

            Assert.That(marker, Is.EqualTo(expectedMarker), $"Button '{buttonLabel}' did not run the expected VBA callback.");
        }
    }

    private static AutomationElement? FindExcelRibbonElement(Window window, string name, ControlType controlType, TimeSpan timeout)
    {
        return Retry.WhileNull(
            () => window.FindFirstDescendant(cf => cf.ByControlType(controlType).And(cf.ByName(name))),
            timeout,
            TimeSpan.FromMilliseconds(250)).Result;
    }

    private static AutomationElement? FindExcelRibbonElementByName(Window window, string name, TimeSpan timeout)
    {
        return Retry.WhileNull(
            () => window.FindFirstDescendant(cf => cf.ByName(name)),
            timeout,
            TimeSpan.FromMilliseconds(250)).Result;
    }

    private static void TryMaximize(Window window)
    {
        try
        {
            window.Patterns.Window.Pattern.SetWindowVisualState(WindowVisualState.Maximized);
        }
        catch (InvalidOperationException)
        {
            // Some Excel startup states do not expose the window pattern immediately.
        }
    }

    private static string ReadCustomPart(string workbookPath, XmlPart partType)
    {
        using var document = new OfficeDocument(workbookPath);
        var part = document.RetrieveCustomPart(partType);
        Assert.That(part, Is.Not.Null, $"Custom part '{partType}' was not found.");
        return part!.ReadContent();
    }

    private static ProcessResult RunPowerShellTool(string repositoryRoot, string workbookPath, string definitionPath, string outputPath, bool forceInPlace = false)
    {
        var scriptPath = Path.Combine(repositoryRoot, "scripts", "Update-CustomUI14.ps1");
        Assert.That(File.Exists(scriptPath), Is.True, $"PowerShell script was not found at '{scriptPath}'.");

        var arguments = $"-NoProfile -ExecutionPolicy Bypass -File {Quote(scriptPath)} -WorkbookPath {Quote(workbookPath)} -DefinitionPath {Quote(definitionPath)}";
        if (forceInPlace)
        {
            arguments += " -ForceInPlace";
        }
        else
        {
            arguments += $" -OutputPath {Quote(outputPath)}";
        }

        using var process = new Process();
        process.StartInfo = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = arguments,
            WorkingDirectory = repositoryRoot,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };

        process.Start();
        var standardOutput = process.StandardOutput.ReadToEnd();
        var standardError = process.StandardError.ReadToEnd();
        Assert.That(process.WaitForExit(60000), Is.True, "PowerShell tool did not exit within 60 seconds.");

        return new ProcessResult(process.ExitCode, standardOutput, standardError);
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\"", "\\\"", StringComparison.Ordinal) + "\"";
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(TestContext.CurrentContext.TestDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeRibbonXEditor.slnx")) &&
                File.Exists(Path.Combine(directory.FullName, "scripts", "Update-CustomUI14.ps1")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        Assert.Fail("Could not locate the repository root.");
        return string.Empty;
    }

    private static void PreserveWorkbookArtifact(string repositoryRoot, string sourceWorkbookPath)
    {
        var outputDirectory = Path.Combine(repositoryRoot, "tests", "UITests", "Output");
        Directory.CreateDirectory(outputDirectory);

        var artifactPath = Path.Combine(outputDirectory, "RibbonXUpdated.xlsm");
        File.Copy(sourceWorkbookPath, artifactPath, overwrite: true);
        TestContext.WriteLine($"Preserved workbook artifact: {artifactPath}");
    }

    private static bool IsExcelInstalled()
    {
        return Type.GetTypeFromProgID("Excel.Application") != null;
    }

    private static bool IsVbaProjectAccessTrusted()
    {
        using var securityKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Excel\Security");
        return Convert.ToInt32(securityKey?.GetValue("AccessVBOM") ?? 0) == 1;
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    private sealed record ExpectedButton(string Id, string Label, string Size, string ImageMso, string OnAction);

    private sealed record ProcessResult(int ExitCode, string StandardOutput, string StandardError)
    {
        public string AllOutput => StandardOutput + Environment.NewLine + StandardError;
    }

    private sealed class ExcelSession : IDisposable
    {
        private ExcelSession(object excel, int processId)
        {
            Excel = excel;
            ProcessId = processId;
        }

        public object Excel { get; }

        public object? Workbook { get; set; }

        public int ProcessId { get; }

        public string Marker
        {
            get
            {
                dynamic workbook = Workbook ?? throw new InvalidOperationException("Workbook is not open.");
                return Convert.ToString(workbook.Worksheets[1].Range("A1").Value2) ?? string.Empty;
            }
        }

        public static ExcelSession CreateVisible(bool visible)
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            Assume.That(excelType, Is.Not.Null, "Excel.Application COM type was not found.");

            var excel = Activator.CreateInstance(excelType!) ?? throw new InvalidOperationException("Could not start Excel.");
            dynamic dynamicExcel = excel;
            dynamicExcel.Visible = visible;
            dynamicExcel.DisplayAlerts = false;
            dynamicExcel.AutomationSecurity = MsoAutomationSecurityLow;

            var hwnd = new IntPtr(Convert.ToInt32(dynamicExcel.Hwnd));
            GetWindowThreadProcessId(hwnd, out var processId);
            return new ExcelSession(excel, Convert.ToInt32(processId));
        }

        public static ExcelSession Open(string workbookPath)
        {
            var session = CreateVisible(visible: true);
            dynamic excel = session.Excel;
            excel.DisplayAlerts = false;
            excel.AutomationSecurity = MsoAutomationSecurityLow;
            session.Workbook = excel.Workbooks.Open(workbookPath);
            excel.ActiveWindow.Activate();
            return session;
        }

        public void SetMarker(string value)
        {
            dynamic workbook = Workbook ?? throw new InvalidOperationException("Workbook is not open.");
            workbook.Worksheets[1].Range("A1").Value2 = value;
        }

        public void Dispose()
        {
            try
            {
                if (Workbook != null)
                {
                    dynamic workbook = Workbook;
                    workbook.Close(SaveChanges: false);
                }
            }
            catch (COMException)
            {
                // Excel may already be gone after a failing UI assertion.
            }
            finally
            {
                try
                {
                    dynamic excel = Excel;
                    excel.Quit();
                }
                catch (COMException)
                {
                    // Nothing useful to do in test cleanup.
                }

                ReleaseComObject(Workbook);
                ReleaseComObject(Excel);
            }
        }

        private static void ReleaseComObject(object? value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }
    }
}
