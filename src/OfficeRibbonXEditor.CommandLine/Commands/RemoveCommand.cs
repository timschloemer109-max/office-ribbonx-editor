using CommunityToolkit.Diagnostics;
using JetBrains.Annotations;
using McMaster.Extensions.CommandLineUtils;
using OfficeRibbonXEditor.Common;

namespace OfficeRibbonXEditor.CommandLine.Commands;

[Command(Description = "Removes Custom UI parts from an Office file")]
public class RemoveCommand(IConsole console) : BaseUpdateCommand(console)
{
    [Option("--type <TYPE>", Description = "The custom UI type to remove: 12, 14, or all")]
    [AllowedValues("12", "14", "all")]
    [UsedImplicitly]
    public string Type { get; set; } = "all";

    public override int OnExecute(CommandLineApplication app)
    {
        Guard.IsNotNull(OfficeFile);

        Log($"Opening Office file '{Path.GetFileName(OfficeFile)}'...");
        var doc = new OfficeDocument(OfficeFile);

        if (Type is "12" or "all")
        {
            Log("Removing Custom UI 2007 part...");
            doc.RemoveCustomPart(XmlPart.RibbonX12);
        }

        if (Type is "14" or "all")
        {
            Log("Removing Custom UI 2010+ part...");
            doc.RemoveCustomPart(XmlPart.RibbonX14);
        }

        doc.Save(OutputFile);
        return 0;
    }
}
