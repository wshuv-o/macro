#!/bin/bash
#
# run_excel_macro.sh
#
# A comprehensive shell script to manage running Excel macros via PowerShell automation.
#
# Supports:
# - Single file macro run
# - Batch run from CSV file with retries and delays
# - Config JSON loading (jq required)
# - Logging to console and file
# - CLI argument parsing
# - Timestamped logs and progress
# - Error handling and retries
# - Placeholder hooks for preprocessing, postprocessing
#
# Usage example:
# ./run_excel_macro.sh -f "C:/Path/Workbook.xlsm" -m "Module1.MyMacro"
# ./run_excel_macro.sh -b "batch_files.csv" -m "Module1.MyMacro" --verbose
#

# -------------- Configurable Defaults --------------

LOG_FILE="./macro_runner.log"
VERBOSE=0
RETRIES=3
DELAY=5
VISIBLE=0
CONFIG_FILE=""
BATCH_FILE=""
EXCEL_FILE=""
MACRO_NAME=""

# Temporary directory for intermediate files
TMP_DIR="./tmp_macro_runner"
mkdir -p "$TMP_DIR"

# -------------- Utility functions ------------------

timestamp() {
  date "+%Y-%m-%d %H:%M:%S"
}

log_info() {
  echo "[$(timestamp)] INFO: $*"
  echo "[$(timestamp)] INFO: $*" >> "$LOG_FILE"
}

log_error() {
  echo "[$(timestamp)] ERROR: $*" >&2
  echo "[$(timestamp)] ERROR: $*" >> "$LOG_FILE"
}

log_debug() {
  if [ "$VERBOSE" -eq 1 ]; then
    echo "[$(timestamp)] DEBUG: $*"
    echo "[$(timestamp)] DEBUG: $*" >> "$LOG_FILE"
  fi
}

usage() {
  cat <<EOF
Usage: $0 [OPTIONS]

Options:
  -f FILE          Excel .xlsm file to run macro on
  -m MACRO         Macro name to run (e.g. Module1.MyMacro) [Required]
  -b CSVFILE       CSV file with list of Excel files to run macro on (column header: filepath)
  -c CONFIG        JSON config file to override options
  -v               Verbose logging
  --visible        Show Excel during macro execution (PowerShell)
  -r RETRIES       Number of retries on failure (default: $RETRIES)
  -d DELAY         Delay seconds between retries (default: $DELAY)
  -l LOGFILE       Log file path (default: $LOG_FILE)
  -h               Show this help message

Examples:
  $0 -f "C:/Files/MyWorkbook.xlsm" -m "Module1.MyMacro"
  $0 -b "files.csv" -m "Module1.MyMacro" --visible -v
  $0 -c "config.json"

EOF
}

# -------------- JSON config parsing using jq ------------

load_config() {
  local config_file="$1"
  if ! command -v jq >/dev/null 2>&1; then
    log_error "jq not found, required for JSON config parsing."
    exit 1
  fi

  if [ ! -f "$config_file" ]; then
    log_error "Config file '$config_file' does not exist."
    exit 1
  fi

  log_info "Loading configuration from $config_file"

  # Use jq to extract fields with defaults fallback
  EXCEL_FILE=$(jq -r '.file // empty' "$config_file")
  MACRO_NAME=$(jq -r '.macro // empty' "$config_file")
  BATCH_FILE=$(jq -r '.batch // empty' "$config_file")
  VISIBLE=$(jq -r '.visible // false' "$config_file")
  RETRIES=$(jq -r '.retries // '"$RETRIES" "$config_file")
  DELAY=$(jq -r '.delay // '"$DELAY" "$config_file")
  LOG_FILE=$(jq -r '.log // "'"$LOG_FILE"'"' "$config_file")

  # Convert jq boolean to 0/1 integer
  if [ "$VISIBLE" = "true" ]; then
    VISIBLE=1
  else
    VISIBLE=0
  fi

  log_debug "Config loaded: file=$EXCEL_FILE, macro=$MACRO_NAME, batch=$BATCH_FILE, visible=$VISIBLE, retries=$RETRIES, delay=$DELAY, log=$LOG_FILE"
}

# -------------- PowerShell command generator -------------

generate_powershell_script() {
  local excel_path="$1"
  local macro_name="$2"
  local visible_flag="$3"
  local ps_file="$4"

  cat > "$ps_file" <<EOF
# Auto-generated PowerShell script to run Excel macro
\$ErrorActionPreference = 'Stop'

try {
    Write-Host "Starting Excel application..."
    \$excel = New-Object -ComObject Excel.Application
    \$excel.Visible = $visible_flag
    \$excel.DisplayAlerts = \$false

    Write-Host "Opening workbook: $excel_path"
    \$workbook = \$excel.Workbooks.Open("$excel_path")

    Write-Host "Running macro: $macro_name"
    \$excel.Run("$macro_name")

    Write-Host "Saving workbook..."
    \$workbook.Save()

    Write-Host "Closing workbook and Excel..."
    \$workbook.Close(\$false)
    \$excel.Quit()
} catch {
    Write-Error "Error during Excel macro execution: \$_"
    exit 1
}
EOF
}

# -------------- Run macro on single file with retries ---------

run_macro_on_file() {
  local file_path="$1"
  local macro="$2"
  local retries_left="$3"
  local delay_sec="$4"

  local abs_path
  abs_path=$(cygpath -w "$file_path" 2>/dev/null || echo "$file_path")
  # If using WSL, convert path to Windows path; otherwise leave as is

  log_info "Running macro on file: $file_path"
  log_debug "Windows path for PowerShell: $abs_path"

  local ps_script="$TMP_DIR/run_macro_$(date +%s%N).ps1"
  generate_powershell_script "$abs_path" "$macro" "$VISIBLE" "$ps_script"

  local attempt=1
  while [ $attempt -le $((retries_left + 1)) ]; do
    log_info "Attempt $attempt of $((retries_left + 1))"
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$ps_script"
    local exit_code=$?

    if [ $exit_code -eq 0 ]; then
      log_info "Macro ran successfully on file: $file_path"
      rm -f "$ps_script"
      return 0
    else
      log_error "Macro failed on attempt $attempt (exit code $exit_code)."
      if [ $attempt -le $retries_left ]; then
        log_info "Retrying after $delay_sec seconds..."
        sleep $delay_sec
      fi
    fi
    attempt=$((attempt + 1))
  done

  log_error "All attempts failed for file: $file_path"
  rm -f "$ps_script"
  return 1
}

# -------------- Batch run from CSV ---------------------

run_batch_from_csv() {
  local csv_file="$1"
  local macro="$2"
  local total=0
  local success=0
  local fail=0

  if [ ! -f "$csv_file" ]; then
    log_error "Batch CSV file '$csv_file' not found."
    return 1
  fi

  # Detect header and find filepath column index
  local header
  header=$(head -1 "$csv_file")
  local filepath_col=0
  IFS=',' read -ra headers <<< "$header"
  for i in "${!headers[@]}"; do
    if [[ "${headers[i],,}" == "filepath" ]]; then
      filepath_col=$((i+1))
      break
    fi
  done

  if [ "$filepath_col" -eq 0 ]; then
    log_error "'filepath' column not found in CSV header."
    return 1
  fi

  # Skip header line, process each line
  tail -n +2 "$csv_file" | while IFS=',' read -r -a cols; do
    total=$((total + 1))
    local file="${cols[filepath_col-1]}"
    if [ -z "$file" ]; then
      log_error "Row $total has empty filepath, skipping."
      continue
    fi

    log_info "Batch processing file #$total: $file"
    if run_macro_on_file "$file" "$macro" "$RETRIES" "$DELAY"; then
      success=$((success + 1))
    else
      fail=$((fail + 1))
    fi
  done

  log_info "Batch run complete: Total=$total, Success=$success, Fail=$fail"
  [ "$fail" -eq 0 ] && return 0 || return 1
}

# -------------- Placeholder preprocessing -------------------

preprocess_placeholder() {
  log_info "Running preprocessing step... (placeholder)"
  # Add custom preprocessing here if needed
  sleep 1
}

# -------------- Main CLI argument parsing -------------------

parse_args() {
  while [[ $# -gt 0 ]]; do
    case $1 in
      -f)
        EXCEL_FILE="$2"
        shift 2
        ;;
      -m)
        MACRO_NAME="$2"
        shift 2
        ;;
      -b)
        BATCH_FILE="$2"
        shift 2
        ;;
      -c)
        CONFIG_FILE="$2"
        shift 2
        ;;
      -v)
        VERBOSE=1
        shift
        ;;
      --visible)
        VISIBLE=1
        shift
        ;;
      -r)
        RETRIES="$2"
        shift 2
        ;;
      -d)
        DELAY="$2"
        shift 2
        ;;
      -l)
        LOG_FILE="$2"
        shift 2
        ;;
      -h|--help)
        usage
        exit 0
        ;;
      *)
        log_error "Unknown option: $1"
        usage
        exit 1
        ;;
    esac
  done

  if [ -n "$CONFIG_FILE" ]; then
    load_config "$CONFIG_FILE"
  fi

  if [ -z "$MACRO_NAME" ]; then
    log_error "Macro name (-m) is required."
    usage
    exit 1
  fi

  if [ -z "$EXCEL_FILE" ] && [ -z "$BATCH_FILE" ]; then
    log_error "Either Excel file (-f) or batch CSV (-b) must be provided."
    usage
    exit 1
  fi

  log_debug "Final parameters:"
  log_debug "Excel file: $EXCEL_FILE"
  log_debug "Macro name: $MACRO_NAME"
  log_debug "Batch file: $BATCH_FILE"
  log_debug "Visible: $VISIBLE"
  log_debug "Retries: $RETRIES"
  log_debug "Delay: $DELAY"
  log_debug "Log file: $LOG_FILE"
}

# -------------- Main script logic ---------------------------

main() {
  # Clear or create log file
  : > "$LOG_FILE"

  parse_args "$@"

  preprocess_placeholder

  if [ -n "$BATCH_FILE" ]; then
    run_batch_from_csv "$BATCH_FILE" "$MACRO_NAME"
    local batch_result=$?
    if [ $batch_result -ne 0 ]; then
      log_error "Batch processing encountered errors."
      exit 2
    fi
  else
    if run_macro_on_file "$EXCEL_FILE" "$MACRO_NAME" "$RETRIES" "$DELAY"; then
      log_info "Macro execution succeeded."
      exit 0
    else
      log_error "Macro execution failed."
      exit 3
    fi
  fi
}

main "$@"
