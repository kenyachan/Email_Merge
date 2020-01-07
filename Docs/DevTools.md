# Developer Tools

The purpose of the *DevTools* module is to provide functions that will aid in the development of a VBA project. Such functions as exporting all Modules, Class Modules, and UserForms so that they can then be checked into source control.

## Setting up DevTools

The *DevTools.bas* module can be imported into any VBA project and used.

In order for the *DevTools* module to work, you will need to *Trust access to the VBA project Object model* via the *Trust Center Settings*, and include *Microsoft Visual Basic for Applications Extensibility 5.3* in the references.

In the Excel Trust Center Settings, select *Trust access to the VBA project object model*.

![Excel Trust Center Settings](https://i.imgur.com/chBbCVc.jpg)

Include *Microsoft Visual Basic for Applications Extensibility 5.3* library.

![VBA References](https://i.imgur.com/a4RJcw5.jpg)

## Using DevTools

### Exporting

There are two exporting functions available.

`Export` will export all Modules and UserForms to a *Modules* and *UserForms* folder located in the same directory as the active workbook. If the folders do not exist, it will create them.

`ExportSourceFiles(destinationPath As String)` will export all Modules and UserForms to a *Modules* and *UserForms* folder at `destinationPath`. If the folders do not exist, they will be created.
