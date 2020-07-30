# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [0.3.2-alpha] - TBD

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
