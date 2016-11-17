[![Build status](https://ci.appveyor.com/api/projects/status/n159uhltbd90i3rh?svg=true)](https://ci.appveyor.com/project/mathieumack/mvvx-plugins-open-xml-sdk)

# MvvX.Plugins.Open-XML-SDK

Using the Open-XML-SDK-Plugin for MvvmCross is quite simple. The plugin injects the IWordManager interface into the IoC container.
Each resolve to IWordManager from the Mvx.Resolve<IWordManager>() will create a new instance of the service.

### API

The API of IWordManager is very easy to understand and to use.

```c#
public interface IWordManager : IDisposable
{
	IDatabase Database { get; }
	bool CreateConnection(string workingFolderPath, string databaseName);
}
```
#### WordManager Open existing template

In order to open a template, call the OpenDocFromTemplate method
```c#
	
    var resourceName = "<Set full template file path here>"; // ex : C:\temp\template.dotx
    var finalFilePath = "<Set saved new document file path here>"; // ex : C:\temp\createdDoc.docx
	
    using (IWordManager word = Mvx.Resolve<IWordManager>())
    {
        word.OpenDocFromTemplate(resourceName, finalFilePath, true);

        word.SaveDoc();
        word.CloseDoc();
    }
	
```

#### WordManager Insert text on bookmark

Using the name of the database and the folder on the client device where to store database files:
```c#
	
    var resourceName = "<Set full template file path here>"; // ex : C:\temp\template.dotx
    var finalFilePath = "<Set saved new document file path here>"; // ex : C:\temp\createdDoc.docx
	
    using (IWordManager word = Mvx.Resolve<IWordManager>())
    {
        word.OpenDocFromTemplate(resourceName, finalFilePath, true);

        word.SaveDoc();
        word.CloseDoc();
    }
	
```

To be complete...

