OfficeFlow [![License][badges.license]][links.license] [![Implementation Status][badges.status.umbra]][links.andivionian-status-classifier] [![GitHub Actions Workflow Status][badges.build]][links.workflows] [![Code Coverage][badges.code-coverage]][links.code-coverage]
=
OfficeFlow is a completely free and open source library for reading, creating and editing [Word][links.word], [Excel][links.excel] and [PowerPoint][links.power-point] documents.

Motivation
-
There are three reasons why this library is being developed:
- Most open source solutions offer an inconvenient and ugly API
- Also, these open source solutions are limited in functionality
- Paid libraries have closed source code

My dream is to give the .NET world a free alternative that will fully cover the needs of all developers to interact with the [Office][links.office] software suite. This means that I'm not going to make money by selling access to this library.

Implementation Status
-
The library is not ready yet. However, opening and saving documents with `.docx` extension is already implemented. You can also create and read sections and paragraphs from a document:
```csharp
var document = new WordDocument(WordDocumentType.Docx);

var section = document.AppendSection();
var paragraph = section.AppendParagraph();
paragraph.AppendText("Hello, OfficeFlow!");

document.SaveTo("OfficeFlow.docx");

document.Close();
```

Implementation Dashboard
-
- [ ] Document object model (DOM)
  - [x] Abstractions ([Element][links.dashboard.element], [CompositeElement][links.dashboard.composite-element], [ElementCollection][links.dashboard.element-collection])
  - [ ] Measure units
    - [x] Absolute units 
    - [ ] Relative units
  - [ ] Elements
    - [ ] Paragraph
      - [ ] Text Formatting (partially implemented)
    - [ ] Section
      - [ ] Page settings
    - [ ] Table
- [ ] Open XML Format support
  - [ ] Word
  - [ ] Excel
  - [ ] PowerPoint
- [ ] Binary Format support
  - [ ] Word
  - [ ] Excel
  - [ ] PowerPoint
- [ ] Other
  - [ ] Conversion to PDF
  - [ ] Documentation

[badges.license]: https://img.shields.io/github/license/y0ung3r/OfficeFlow
[badges.status.umbra]: https://img.shields.io/badge/status-umbra-red.svg
[badges.build]: https://img.shields.io/github/actions/workflow/status/y0ung3r/OfficeFlow/main.yaml
[badges.code-coverage]: https://codecov.io/github/y0ung3r/OfficeFlow/graph/badge.svg?token=GEHBFGNYLT
[links.license]: https://github.com/y0ung3r/OfficeFlow/blob/main/LICENSE.md
[links.workflows]: https://github.com/y0ung3r/OfficeFlow/actions
[links.office]: https://en.wikipedia.org/wiki/Microsoft_Office
[links.word]: https://en.wikipedia.org/wiki/Microsoft_Word
[links.excel]: https://en.wikipedia.org/wiki/Microsoft_Excel
[links.power-point]: https://en.wikipedia.org/wiki/Microsoft_PowerPoint
[links.andivionian-status-classifier]: https://andivionian.fornever.me/v1/#status-umbra-
[links.code-coverage]: https://codecov.io/github/y0ung3r/OfficeFlow
[links.dashboard.element]: https://github.com/y0ung3r/OfficeFlow/blob/main/src/Common/OfficeFlow.DocumentObjectModel/Element.cs
[links.dashboard.composite-element]: https://github.com/y0ung3r/OfficeFlow/blob/main/src/Common/OfficeFlow.DocumentObjectModel/CompositeElement.cs
[links.dashboard.element-collection]: https://github.com/y0ung3r/OfficeFlow/blob/main/src/Common/OfficeFlow.DocumentObjectModel/ElementCollection.cs

