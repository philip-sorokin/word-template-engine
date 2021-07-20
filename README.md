# Word Template Engine

Word Template Engine is a PHP based template processor and converter that allows to create documents from DOCX templates ([download an example.docx](https://github.com/philip-sorokin/word-template-engine/blob/main/examples/template.docx?raw=true)) and convert them trough LibreOffice into the following formats: PDF, HTML, XHTML, mail-adapted HTML. This light-weight library will help you in creating invoices, contracts, waylists and other documents. Use variables inside your document and substitute them with values on the server side, replace and delete images, manage document sections (pages), make a document copy or a section copy inside the document, add rows to your tables. Output the documents to a user browser, attach the WORD or PDF files to emails, or create a HTML mail content from your template.

Documentation and examples: https://addondev.com/opensource/word-template-engine

## Changelog

### [1.0.1] - 2021-07-19

- Improved performance of processing large documents. More than 1000 variable replacements take a few seconds. Reduced memory and CPU usage. Speed increased by 6-8 times.

### [Unreleased]

- Fixed directory separators caused issues on Windows platform.
