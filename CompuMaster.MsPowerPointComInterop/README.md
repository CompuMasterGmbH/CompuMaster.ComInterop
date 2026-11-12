# CompuMaster.MsPowerPointComInterop

Automate Microsoft PowerPoint through COM on Windows with PowerPoint installed. The wrapper gives C# code access to PowerPoint's object model and manages diagnostic logs from its automation process.

```csharp
using System;
using CompuMaster.MsPowerPointComInterop;

class Program
{
    [STAThread]
    static void Main()
    {
        using var powerpoint = new PowerPointApplication { PreserveDiagnosticLogFiles = true };
        using var presentations = powerpoint.CreateWrapperAndRegisterComChildForDisposal(
            powerpoint.InvokePropertyGet<object>("Presentations"));
        using var presentation = presentations.CreateWrapperAndRegisterComChildForDisposal(
            presentations.InvokeFunction<object>("Add", -1), // msoTrue opens a window
            null, deck => deck.InvokeMethod("Close"), null);

        Console.WriteLine(presentation.InvokePropertyGet<string>("Name"));
        foreach (var log in powerpoint.DiagnosticLogFiles)
            Console.WriteLine(log.FullName);
    }
}
```

Start the example with PowerPoint closed and read `DiagnosticLogFiles` while the presentation window is open. By default, the wrapper removes confirmed logs after a clean shutdown and trims trailing NUL bytes after an abnormal exit. Set `PreserveDiagnosticLogFiles` to `true` to keep the files intact.
