

# FileCopyDrag

A static helper class for "Drag and Drop" and "Paste From Clipboard" functionality in C-Sharp .NET.


## Methods

**PasteFromClipboard**
A static method that fetches valid file data from your clipboard and saves it to your temporary directory. It returns a `string[]` that contains the filenames of the copied files or `null` if there are none.
```csharp
private void GetFilesFromClipboard()
{
    string[]? fileNames = FileCopyDrag.PasteFromClipboard();

    if (fileNames != null && fileNames.Count > 0) 
    {
        Debug.WriteLine($"Found {fileNames.Count} files from your clipboard.");
    }
}
```

**DragDrop**
A static method that requires `IDataObject` for drag and drop operations. Similar to the method above, it returns a `string[]` that contains the filenames of the dropped files or `null` if there are none.
```csharp
private void GetFilesFromClipboard()
{
    string[]? fileNames = FileCopyDrag.();

    if (fileNames != null && fileNames.Count > 0) 
    {
        Debug.WriteLine($"Found {fileNames.Count} files from your clipboard.");
    }
}
```

**ClearCache**
Clears all the files from your temporary directory that were created by this class. By default, whenever `PasteFromClipboard` and `DragDrop` methods are called, the filenames are stored so that you can dynamically delete the temporary files.


## Sample Usage for Outlook Message Files (Emails)
Sample code to get Outlook messages through drag-n-drop and copy-paste operation  directly from Outlook App.
```csharp
public Form1()
{
    InitializeComponent();
    ButtonPasteFromClipboard.Click += PasteFromClipboard;
    ListBoxMails.DragEnter += DragEnter;
    ListBoxMails.DragDrop += DragDrop;
}

private void PasteFromClipboard(object? sender, EventArgs e)
{
    string[]? filenames = FileCopyDrag.PasteFromClipboard();

    if (filenames != null)
    {
        int count = 0;
        foreach (string filename in filenames)
        {
            FileInfo fileInfo = new(filename);

            // Filter files to .msg (Outlook Message File) only.
            if (fileInfo.Exists && fileInfo.Extension.Equals(".msg", StringComparison.CurrentCultureIgnoreCase))
            {
                ListBoxMails.Items.Add(filename);
                count++;
            }
        }

        if (count == 0) MessageBox.Show("No valid data found on your clipboard. Please try again.", "Paste From Clipboard", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
    else
    {
        MessageBox.Show("No valid data found on your clipboard. Please try again.", "Paste From Clipboard", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

private void ListDragEnter(object? sender, DragEventArgs e)
{
    if (e.Data?.GetDataPresent(DataFormats.FileDrop) ?? false)
    {
        e.Effect = DragDropEffects.Copy;
    }
    else if (e.Data?.GetDataPresent("FileGroupDescriptor") ?? false)
    {
        e.Effect = DragDropEffects.Copy;
    }
    else
    {
        e.Effect = DragDropEffects.None;
    }
}

private void ListDragDrop(object? sender, DragEventArgs e)
{
    if (e.Data == null) return;

    string[]? filenames = FileCopyDrag.DragDrop(e.Data);

    if (filenames != null)
    {
        int count = 0;
        foreach (string filename in filenames)
        {
            FileInfo fileInfo = new(filename);

            // Filter files to .msg (Outlook Message File) only.
            if (fileInfo.Exists && fileInfo.Extension.Equals(".msg", StringComparison.CurrentCultureIgnoreCase))
            {
                ListBoxMails.Items.Add(filename);
                count++;
            }
        }

        if (count == 0) MessageBox.Show("No valid data found on the dropped data. Please try again.", "Drag and Drop", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
    else
    {
        MessageBox.Show("No valid data found on the dropped data. Please try again.", "Drag and Drop", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
```
