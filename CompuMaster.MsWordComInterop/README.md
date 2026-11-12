# CompuMaster.MsWordComInterop

Automate Microsoft Word through COM on Windows with Word installed. The wrapper gives C# code access to Word's object model and manages diagnostic logs from its automation process.

```csharp
using System;
using CompuMaster.MsWordComInterop;

class Program
{
    [STAThread]
    static void Main()
    {
        using var word = new WordApplication { PreserveDiagnosticLogFiles = true };
        using var documents = word.CreateWrapperAndRegisterComChildForDisposal(
            word.InvokePropertyGet<object>("Documents"));
        using var document = documents.CreateWrapperAndRegisterComChildForDisposal(
            documents.InvokeFunction<object>("Add", Type.Missing, Type.Missing, Type.Missing, true),
            null, doc => doc.InvokeMethod("Close", 0, Type.Missing, Type.Missing), null);

        Console.WriteLine(document.InvokePropertyGet<string>("Name"));
        foreach (var log in word.DiagnosticLogFiles)
            Console.WriteLine(log.FullName);
    }
}
```

Read `DiagnosticLogFiles` while the document window is open. By default, the wrapper removes confirmed logs after a clean shutdown and trims trailing NUL bytes after an abnormal exit. Set `PreserveDiagnosticLogFiles` to `true` to keep the files intact.
