# VBA Functions

> A collection of reusable VBA functions — Path, File, Excel, Array, String and more.

## Structure

```
vba-functions/
├── Path/               # Path utilities
│   ├── GetPath.bas             # Get workbook path (OneDrive-aware)
│   ├── GetTemplatePath.bas     # Find template file by category/name/sub
│   ├── GetFilePath.bas        # Find data file by category/name/year/month
│   ├── CreatFilePath.bas      # Create file path + create dirs
│   ├── CreatePath.bas         # Recursive directory creation
│   └── FindExcelFilePathWithKeyword.bas  # Find Excel by keyword
├── File/
│   └── GetExcelFilePath.bas   # Open file dialog (multi-select support)
├── Excel/              # Excel operations (ListObject, Range, etc.)
├── Array/              # Array manipulation
├── String/             # String operations
└── Utils/              # General utilities
```

## Usage

Import `.bas` files into your Excel VBA project:

1. Open Excel → Alt + F11 (VBA Editor)
2. File → Import File → Select `.bas` file
3. Call functions from your modules

## Category

| Folder | Description |
|--------|------------|
| `Path/` | Path construction, OneDrive handling, directory creation |
| `File/` | File selection dialogs |
| `Excel/` | Excel operations, ListObject, Range |
| `Array/` | Array manipulation |
| `String/` | String operations |
| `Utils/` | General utilities |

## License

MIT License © NCLSFU
