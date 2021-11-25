# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [0.5.0-beta] - TBD

### Added
- Added asynchronous versions of the `GenerateSourceCode(...)` extenson methods for
  `OpenXmlElement`, `OpenXmlPart`, and `OpenXmlPackage` objects.

## [0.4.2-beta] - 2021-11-22

### Changed
- Refactored the variable name generation process to reuse existing variable names 
  when they become available.
- *\[Breaking Change\]:* Added `UseUniqueVariableNames` property to `ISerializeSettings`
  interface. This allows to switch between unique and reused variable names.
- *\[Breaking Change\]:* Changed `typeCounts` parameter to `types` in the 
  `IOpenXmlElementHandler.BuildCodeStatements(...)` method to account for the repurposing
  of existing variable name.
- Update DocumentFormat.OpenXml reference to 2.14.0.

### Fixed
- Use the correct `CodeExpression` classes for the `XmlNodeType` parameter of the
  `OpenXmlMiscNode` constructor.

## [0.4.1-beta] - 2021-08-08

### Changed
- Refactored the using directive generation logic to better match how the OpenXML SDK
  Productivity Tool used to create them.  Using directive aliases are now dynamically
  generated based on the OpenXml object that the code output is based on.
- *\[Breaking Change\]:* Updated the `namespace` parameter type in the `IOpenXmlElementHandler`
  and `IOpenXmlPartHandler` interface methods from `ISet<string>` to `IDictionary<string, string>` 
  to account for the new namespace/using directive generation logic.  The following 
  interface methods are impacted:
  - `IOpenXmlElementHandler.BuildCodeStatements(...)`
  - `IOpenXmlPartHandler.BuildEntryMethodCodeStatements(...)`
  - `IOpenXmlPartHandler.BuildHelperMethod(...)`
  
### Fixed

- Issue related to Hyperlink and external relationship references were not being added properly
  in all `OpenXmlPart` code creation scenarios.
- Use the right constructor parameters for `OpenXmlMiscNode` objects.

## [0.4.0-beta] - 2021-08-02

### Added

- New `ISerializeSettings` interface to allows greater flexibility in the source code generation.
- New `IOpenXmlHandler`, `IOpenXmlElementHandler`, and `IOpenXmlPartHandler` interfaces that will 
  allow developers to control how source code is created.

### Changed

- Change visibility of many of the static method helpers so developers can use them in their custom
  code generation.
- Update DocumentFormat.OpenXml reference to 2.13.0.
  
### Fixed

- Make sure that the return type of generated element methods include the namespace alias if 
  needed.
- Choose between the default method or contentType parameter method for the custom OpenXmlPart.AddNewPart
  methods (ex: pkg.AddExtendedFilePropertiesPart() or mainDocumentPart.AddImagePart("image/x-emf")) 

## [0.3.2-alpha] - 2020-07-30

### Changed

- Updated process to account for more OpenXmlPart classes that may require custom AddNewPart methods
  to initialize.
- Changed the `CreatePackage` method to take in a `String` parameter for the full file path of the target file
  instead of a `Stream` when generating code for `OpenXmlPackage` objects.  This was to avoid using a C# `ref` 
  parameter that made using the generated code in a C# project more difficult to use.

### Fixed

- TargetInvocationException/FormatException when trying to parse a value that is not valid for
  `OpenXmlSimpleType` derived types being evaluated. [See this](https://github.com/OfficeDev/Open-XML-SDK/issues/780)
  for more details.
- When encountering OpenXmlUnknownElement objects, make sure to initialize them with the appropriate `ctor` method.
- Correct the initialization parameters for the generated `AddExternalRelationship` method.
- Issue where AddPart methods for OpenXmlPart paths that have already been visited are generated on variables
  that do not exist.

## [0.3.1-alpha] - 2020-07-25

### Fixed

- TargetInvocationException/FormatException when trying to parse a value that is not valid for
  the `EnumValue` type being evaluated. [See this](https://github.com/OfficeDev/Open-XML-SDK/issues/780)
  for more details.

## [0.3.0-alpha] - 2020-07-20

### Changed

- Update DocumentFormat.OpenXml reference to 2.11.3.

### Fixed

- Ambiguous Match Exception occuring when trying to identify parts that need to use the
  `AddImagePart` initialization method.

## [0.2.1-alpha] - 2020-07-03

### Changed

- Change the parameters for all of the methods to `ref` parameters. This changes the generated
  VB code to create `byref` parameters instead of `byval` ones.

## [0.2.0-alpha] - 2020-06-27

### Added

- Added documentation output

### Changed

- Use the alias `AP` for DocumentFormat.OpenXml.ExtendedProperties namespace objects
- Use the `AddImagePart` method for initializing `ImagePart` objects.
- Included the content type parameter for the `AddNewPart` method for `EmbeddedPackagePart` objects.

## [0.1.0-alpha] - 2020-06-24

### Added

- Added initial project to convert OpenXml SDK based documents to source code files.
