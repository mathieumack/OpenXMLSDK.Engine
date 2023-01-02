# OpenXMLSDK.Engine

This library let you to create quickly some docx documents, based on the ooxml sdk of Microsoft.

By using the WordManager object, you will be able to geneate quickly your documents.


## Quality and packaging

[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=mathieumack_OpenXMLSDK.Engine&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=mathieumack_OpenXMLSDK.Engine)
[![.NET](https://github.com/mathieumack/OpenXMLSDK.Engine/actions/workflows/ci.yml/badge.svg)](https://github.com/mathieumack/OpenXMLSDK.Engine/actions/workflows/ci.yml)
[![NuGet package](https://buildstats.info/nuget/OpenXMLSDK.Engine?includePreReleases=true)](https://nuget.org/packages/OpenXMLSDK.Engine)


### API

The API of WordManager is very easy to understand and to use.

#### WordManager Open existing template

In order to open a template, call the OpenDocFromTemplate method
```c#
	
    var resourceName = "<Set full template file path here>"; // ex : C:\temp\template.dotx
    var finalFilePath = "<Set saved new document file path here>"; // ex : C:\temp\createdDoc.docx
	
    using (var word = new WordManager())
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
	
    using (var word = new WordManager())
    {
        word.OpenDocFromTemplate(resourceName, finalFilePath, true);

        word.SaveDoc();
        word.CloseDoc();
    }
	
```

# Contribute

## How to contribute

If you want to contribute to this project, you can do it in several ways:

- [Submit bugs and feature requests]
- [Review source code changes]
- [Review the documentation and make pull requests for anything from typos to new content]