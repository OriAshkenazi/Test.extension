# -*- coding: utf-8 -*-  # pragma: no cover
__title__ = "Mapping Validator"  # pragma: no cover
__doc__ = "Validate PBS mapping XLSX files."  # pragma: no cover

import os
from pathlib import Path

from pyrevit import script, forms

from lib.runner import temp_utils
from lib.runner.process_manager import ProcessManager
from lib.runner.data_exchange import ElementDataSerializer


def main():
    # Clear output window for clean run
    output = script.get_output()
    output.print_md("# PBS Mapping Validator")
    output.print_md("Starting mapping validation...")

    # Get XLSX file path from user
    output.print_md("**Step 1/4**: Selecting XLSX file to validate...")
    xlsx_filter = "Excel Files (*.xlsx)|*.xlsx"

    xlsx_path = forms.pick_file(
        file_ext="xlsx", files_filter=xlsx_filter, title="Select PBS Mapping XLSX File"
    )

    if not xlsx_path:
        output.print_md(" **Cancelled**: No XLSX file selected")
        return

    # Check if file exists
    if not Path(xlsx_path).exists():
        output.print_md(" **File Error**: Selected XLSX file does not exist")
        return

    output.print_md(" Selected file: `{}`".format(xlsx_path))

    # Initialize process manager
    proc_manager = ProcessManager(timeout=60)  # 1 minute timeout for validation

    # Validate Python environment first
    output.print_md("**Step 2/4**: Validating Python environment...")
    is_valid, error_msg = proc_manager.validate_python_environment()
    if not is_valid:
        output.print_md(" **Error**: {}".format(error_msg))
        forms.alert(
            "{}\n\nPlease install Python 3.12+ or set PBS_PYTHON_PATH environment variable.".format(
                error_msg
            ),
            title="Python Environment Error",
        )
        return
    output.print_md(" Python environment validated")

    with temp_utils.temporary_dir() as tmpdir:
        results_file = tmpdir / "validation_results.json"

        try:
            output.print_md("**Step 3/4**: Running validation...")

            # Find mapping script
            start = Path(__file__).resolve()
            default = start.parents[3] / "scripts" if len(start.parents) > 3 else None
            if default and default.is_dir():
                scripts_dir = default
            else:
                scripts_dir = temp_utils.find_scripts_dir(start)
                if scripts_dir != default:
                    output.print_md("Using scripts dir {}".format(scripts_dir))
            runner = scripts_dir / "mapping_runner.py"

            # Prepare arguments for validate command
            args = ["validate", str(xlsx_path), "--output", str(results_file)]
            if os.environ.get("PBS_DEBUG_VALIDATION") == "1":
                args.append("--verbose")
                output.print_md(
                    "Debug mode enabled - detailed validation information will be logged"
                )

            output.print_md(" Validating XLSX structure and PBS codes...")

            # Execute CPython script using process manager
            returncode, stdout, stderr = proc_manager.run_python_script(runner, args)

            # Log results
            log_file = temp_utils.get_log_path("mapping_validator.log")
            with log_file.open("a") as fh:
                fh.write(stdout or "")
                fh.write(stderr or "")
                fh.write("Process returned code: {}\n".format(returncode))

            if returncode in [0, 1]:  # 0 = valid, 1 = invalid but parseable
                # Read validation results
                is_valid, results = ElementDataSerializer.deserialize_validation_results(
                    results_file
                )

                output.print_md("**Step 4/4**: Processing validation results...")

                # Display results
                display_validation_results(output, results)

                # Show summary dialog
                show_validation_summary(results, xlsx_path)

            else:
                # Handle specific error cases based on return code and stderr
                error_details = (stderr or "Unknown error occurred").strip()
                if returncode == 2:
                    if "openpyxl library not available" in error_details:
                        output.print_md(" **Dependency Error**: openpyxl library not available")
                        output.print_md(
                            " **Solution**: Install openpyxl with: `pip install openpyxl`"
                        )
                        return
                    elif (
                        "File is corrupted" in error_details
                        or "corrupted" in error_details.lower()
                    ):
                        output.print_md(" **File Format Error**: File is corrupted")
                        return
                if "Validation failed" in error_details:
                    output.print_md(" **Validation Error**: Validation failed")
                else:
                    output.print_md(" **Error**: Validation process failed")
                    output.print_md("Error details: {}".format(error_details[:500]))

                forms.alert(
                    "Validation failed.\n\nError: {}\n\n"
                    "Check the log file for details.".format(error_details),
                    title="Validation Error",
                )

        except Exception as e:
            error_str = str(e)
            if "Deserialization failed" in error_str:
                output.print_md(" **Result Processing Error**: Deserialization failed")
            else:
                output.print_md(" **Critical Error**: {}".format(error_str))
            forms.alert(
                "Critical validation error: {}\n\n"
                "Please check your Python installation and try again.".format(error_str),
                title="Validation Error"
            )


def display_validation_results(output, results):
    """
    Display validation results in the PyRevit output window.

    Args:
        output: PyRevit output object
        results: Validation results dictionary
    """
    is_valid = results.get("valid", False)
    errors = results.get("errors", [])
    warnings = results.get("warnings", [])
    stats = results.get("statistics", {})

    # Display validation status
    if is_valid:
        output.print_md("##  Validation Passed")
        output.print_md("**File is valid** - All PBS codes are properly formatted")
    else:
        output.print_md("##  Validation Failed")
        output.print_md("**File contains errors** - Please review and fix issues below")

    # Display statistics
    if stats:
        output.print_md("###  Statistics")
        output.print_md("- **Total rows**: {}".format(stats.get("total_rows", 0)))
        output.print_md("- **Complete mappings**: {}".format(stats.get("complete_mappings", 0)))
        output.print_md("- **Pending mappings**: {}".format(stats.get("pending_mappings", 0)))
        output.print_md("- **Invalid mappings**: {}".format(stats.get("invalid_mappings", 0)))

    # Display errors
    if errors:
        output.print_md("###  Errors ({})".format(len(errors)))
        for i, error in enumerate(errors[:10]):  # Limit to first 10 errors
            error_type = error.get("type", "unknown")
            cell = error.get("cell", "unknown")
            message = error.get("message", "Unknown error")

            if error_type == "header_mismatch":
                output.print_md(
                    "- **{}**: Expected '{}', found '{}'".format(
                        cell, error.get("expected", ""), error.get("actual", "")
                    )
                )
            elif error_type == "invalid_pbs_format":
                value = error.get("value", "")
                output.print_md("- **{}**: Invalid PBS code '{}' - {}".format(cell, value, message))
            else:
                output.print_md("- **{}**: {}".format(cell, message))

        if len(errors) > 10:
            output.print_md("- *... and {} more errors*".format(len(errors) - 10))

    # Display warnings
    if warnings:
        output.print_md("###  Warnings ({})".format(len(warnings)))
        for warning in warnings[:5]:  # Limit to first 5 warnings
            warning_type = warning.get("type", "unknown")
            row = warning.get("row", "unknown")
            message = warning.get("message", "Unknown warning")

            if warning_type == "partial_mapping":
                output.print_md("- **Row {}**: {}".format(row, message))
            else:
                output.print_md("- **Row {}**: {}".format(row, message))

        if len(warnings) > 5:
            output.print_md("- *... and {} more warnings*".format(len(warnings) - 5))


def show_validation_summary(results, xlsx_path):
    """
    Show validation summary in a dialog box.

    Args:
        results: Validation results dictionary
        xlsx_path: Path to validated XLSX file
    """
    is_valid = results.get("valid", False)
    errors = results.get("errors", [])
    warnings = results.get("warnings", [])
    stats = results.get("statistics", {})

    if is_valid:
        # Success dialog
        complete = stats.get("complete_mappings", 0)
        pending = stats.get("pending_mappings", 0)

        message = "XLSX file validation passed!\n\n"
        message += "Complete mappings: {}\n".format(complete)
        message += "Pending mappings: {}\n".format(pending)
        message += "\nFile: {}".format(xlsx_path)

        if warnings:
            message += "\n\nNote: {} warnings found (see output window)".format(len(warnings))

        forms.alert(message, title="Validation Successful")
    else:
        # Error dialog
        error_count = len(errors)
        warning_count = len(warnings)

        message = "XLSX file validation failed!\n\n"
        message += "Errors: {}\n".format(error_count)
        if warning_count > 0:
            message += "Warnings: {}\n".format(warning_count)
        message += "\nPlease review the detailed results in the output window."
        message += "\n\nFile: {}".format(xlsx_path)

        forms.alert(message, title="Validation Failed")


if __name__ == "__main__":
    main()
