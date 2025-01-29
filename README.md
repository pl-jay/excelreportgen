# Excel Report Generator

## Instructions to Build and Run the Application

### 1. Build the Application
Navigate to the project directory and run the following commands:

```bash
dotnet restore
dotnet build
```

This will restore all dependencies and compile the application.

---

### 2. Run the Application
Use the following syntax to execute the program:

```bash
dotnet run <TemplatePath1> <RowCount1> [<TemplatePath2> <RowCount2> ...]
```

#### Example:
```bash
dotnet run "Template1.xlsx" 1000 "Template2.xlsx" 2000 "Template3.xlsx" 1500
```

- **TemplatePath**: Path to the `.xlsx` template file. Must exist and be accessible.
- **RowCount**: Number of data rows to generate for the corresponding template.

### Output:
Generated Excel files will be saved in the same directory as the template, with filenames in the format:
```
<TemplateName>_<Timestamp>.xlsx
```

#### Example Output:
```
Template1_20250126_123456.xlsx
```

