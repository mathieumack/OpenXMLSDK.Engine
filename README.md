# MvvX.Plugins.Open-XML-SDK

Using the Open-XML-SDK-Plugin for MvvmCross is quite simple. The plugin injects the IWordManager interface into the IoC container.
Each resolve to IWordManager from the Mvx.Resolve<IWordManager>() will create a new instance of the service.


## Quality and packaging

[![Build status](https://dev.azure.com/mackmathieu/Github/_apis/build/status/MvvX.Plugins.OpenXML)](https://dev.azure.com/mackmathieu/Github/_build/latest?definitionId=4)
[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=github-MvvX.Plugins.OpenXML&metric=alert_status)](https://sonarcloud.io/dashboard?id=github-nosqlrepository)

![Nuget](https://img.shields.io/nuget/dt/NoSqlRepositories.Core.svg?label=MvvX.Plugins.Open-XML-SDK&logo=nuget)


### API

The API of WordManager is very easy to understand and to use.

#### WordManager Open existing template

In order to open a template, call the OpenDocFromTemplate method
```c#
	
    var resourceName = "<Set full template file path here>"; // ex : C:\temp\template.dotx
    var finalFilePath = "<Set saved new document file path here>"; // ex : C:\temp\createdDoc.docx
	
    using (var word = Mvx.Resolve<WordManager>())
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
	
    using (var word = Mvx.Resolve<WordManager>())
    {
        word.OpenDocFromTemplate(resourceName, finalFilePath, true);

        word.SaveDoc();
        word.CloseDoc();
    }
	
```

To be complete...

