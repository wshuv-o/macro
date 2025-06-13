# Excel Macro Runner Automation

This repository contains a **robust shell script** `run.sh` designed to automate running Excel macros on Windows systems (or WSL environments) using PowerShell. It supports running macros on single files or batches of Excel files, with logging, retries, and configurable options.

---

## Features

* Run Excel macros on `.xlsm` workbooks via PowerShell COM automation.
* Support single Excel file or batch runs via CSV file.
* Configurable retry logic with delay on failures.
* Verbose and timestamped logging to file and console.
* Supports showing or hiding Excel during macro execution.
* Load configuration from JSON files.
* Batch CSV expects a `filepath` column.
* Error handling with retries and detailed logs.
* Modular and extensible for preprocessing/postprocessing hooks.
* Cross-environment friendly: designed to work on Windows cmd/bash or WSL calling `powershell.exe`.

---

## Requirements

* Windows system with **Microsoft Excel** installed.
* **PowerShell** available in the system PATH (`powershell.exe`).
* Bash shell environment (Git Bash, WSL, Cygwin, or Linux/macOS with remote Windows).
* `jq` installed for JSON configuration file support (optional).
* `cygpath` available if using Cygwin or WSL for path translation (optional).

---

## Getting Started

### Clone the repository

```bash
git clone https://github.com/wshuv-o/macro.git
cd macro
```

### Make script executable

```bash
chmod +x run_excel_macro.sh
```

---

## Usage

### Run macro on a single Excel file

```bash
./run_excel_macro.sh -f "C:/Path/Workbook.xlsm" -m "Module1.MyMacro"
```

### Run macro on multiple files via CSV batch

Prepare a CSV file with a header and a `filepath` column, e.g.,

```csv
filepath
C:/Files/Workbook1.xlsm
C:/Files/Workbook2.xlsm
```

Then run:

```bash
./run_excel_macro.sh -b "batch_files.csv" -m "Module1.MyMacro"
```

### Show Excel UI during macro execution

Add `--visible` flag:

```bash
./run_excel_macro.sh -f "C:/Path/Workbook.xlsm" -m "Module1.MyMacro" --visible
```

### Enable verbose logging

Add `-v` flag:

```bash
./run_excel_macro.sh -f "C:/Path/Workbook.xlsm" -m "Module1.MyMacro" -v
```

### Use JSON configuration file

Create a JSON config file, e.g. `config.json`:

```json
{
  "file": "C:/Path/Workbook.xlsm",
  "macro": "Module1.MyMacro",
  "visible": true,
  "retries": 5,
  "delay": 10,
  "log": "custom_log.log"
}
```

Run script with:

```bash
./run_excel_macro.sh -c config.json
```

---

## Options Summary

| Option      | Description                                  | Example                   |
| ----------- | -------------------------------------------- | ------------------------- |
| `-f`        | Excel file path for macro execution          | `-f "C:/Files/file.xlsm"` |
| `-m`        | Macro name to run (required)                 | `-m "Module1.MyMacro"`    |
| `-b`        | CSV batch file with `filepath` column        | `-b "files.csv"`          |
| `-c`        | JSON config file                             | `-c config.json`          |
| `-v`        | Verbose logging                              | `-v`                      |
| `--visible` | Show Excel UI during macro execution         | `--visible`               |
| `-r`        | Number of retries on failure (default 3)     | `-r 5`                    |
| `-d`        | Delay seconds between retries (default 5)    | `-d 10`                   |
| `-l`        | Log file path (default `./macro_runner.log`) | `-l "/var/log/macro.log"` |
| `-h`        | Show help                                    | `-h`                      |

---

## Logging

* Logs are written to `macro_runner.log` by default (in script directory).
* Verbose mode (`-v`) prints debug info.
* Logs contain timestamps, status info, errors, and retry attempts.

---

## Troubleshooting

* Ensure Excel is installed and licensed on the Windows machine.
* Make sure PowerShell is accessible as `powershell.exe`.
* If running inside WSL or Cygwin, make sure `cygpath` is installed for path translation.
* If JSON config loading fails, install `jq` via your package manager.
* Check log file for detailed error messages.
* Macro must exist in the workbook and be accessible by the given name.

---

## Contributing

Feel free to open issues or pull requests to improve the script, add features, or fix bugs.

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Author

**Wahid Shuvo**

* GitHub: [wshuv-o](https://github.com/wshuv-o)

