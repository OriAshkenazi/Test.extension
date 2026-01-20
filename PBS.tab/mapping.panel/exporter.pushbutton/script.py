# -*- coding: utf-8 -*-  # pragma: no cover
__title__ = "Mapping Exporter"  # pragma: no cover
__doc__ = "Export PBS mapping XLSX for the active document."  # pragma: no cover

import os
from pathlib import Path
from datetime import datetime

from pyrevit import script, forms

from lib.collector import element_sampler
from lib.collector.category_mapping import get_available_categories
from lib.ui.category_selector import select_categories, format_category_summary
from lib.runner import temp_utils
from lib.runner.process_manager import ProcessManager
from lib.runner.data_exchange import ElementDataSerializer

try:  # pragma: no cover - only defined in Revit
    DOC = __revit__.ActiveUIDocument.Document  # type: ignore[name]  # noqa: F821
except Exception:
    DOC = None


def main():
    if DOC is None:
        script.get_output().print_md("Not running inside Revit")
        return

    # Clear output window for clean run
    output = script.get_output()
    output.print_md("# PBS Mapping Exporter")
    output.print_md("Starting mapping export...")

    # Get output file path from user
    output.print_md("**Step 1/6**: Selecting output file...")
    xlsx_filter = "Excel Files (*.xlsx)|*.xlsx"
    suggested_name = "PBS_Mapping_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"

    xlsx_path = forms.save_file(
        file_ext="xlsx", default_name=suggested_name, files_filter=xlsx_filter
    )

    if not xlsx_path:
        output.print_md(" **Cancelled**: No output file selected")
        return

    output.print_md(" Output file: `{}`".format(xlsx_path))

    # Initialize process manager
    proc_manager = ProcessManager(timeout=120)  # 2 minute timeout

    # Validate Python environment first
    output.print_md("**Step 2/6**: Validating Python environment...")
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

    # Category selection step
    output.print_md("**Step 3/6**: Category selection...")
    available_categories = get_available_categories(DOC)

    if not available_categories:
        output.print_md(" **Error**: No categories found in the model")
        forms.alert(
            "No categories found in the model. The model may be empty or have no valid elements.",
            title="Mapping Exporter Error"
        )
        return

    output.print_md(" Found {} categories in model".format(len(available_categories)))

    # Show category selection dialog
    selected_categories = select_categories(
        DOC,
        title="Select Categories for PBS Mapping Export",
        multiselect=True,
        include_all_option=True
    )

    if not selected_categories:
        output.print_md(" **Cancelled**: No categories selected")
        return

    category_summary = format_category_summary(selected_categories)
    output.print_md(" Selected categories: {}".format(category_summary))

    output.print_md("**Step 4/6**: Collecting element data...")
    # Use new Schema v3.0 data extraction with category filtering
    element_data = element_sampler.sample_document(DOC, category_filter=selected_categories)
    total_elements = len(element_data)

    if total_elements == 0:
        output.print_md(" **Error**: No elements found in selected categories")
        forms.alert(
            "No elements found in the selected categories: {}\n\n"
            "Please select different categories or check your model.".format(category_summary),
            title="Mapping Exporter"
        )
        return

    output.print_md(
        " Collected data from {} elements using Schema v3.0 with category filtering".format(
            total_elements
        )
    )

    with temp_utils.temporary_dir() as tmpdir:
        data_file = tmpdir / "mapping_data.json"

        try:
            output.print_md("**Step 5/6**: Preparing mapping data...")

            # Serialize mapping data using new Schema v3.0 method with category filter metadata
            ElementDataSerializer.serialize_element_data(
                element_data, data_file, selected_categories
            )

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

            # Prepare arguments for create command
            args = ["create", str(data_file), "--output", str(xlsx_path)]
            if os.environ.get("PBS_DEBUG_MAPPING") == "1":
                args.append("--verbose")
                output.print_md("Debug mode enabled - detailed mapping information will be logged")

            output.print_md("**Step 6/6**: Generating XLSX mapping file...")
            output.print_md(
                " Processing {} elements from {} categories...".format(
                    total_elements, len(selected_categories)
                )
            )
            output.print_md(" This may take 30-60 seconds for large models...")

            # Execute CPython script using process manager
            returncode, stdout, stderr = proc_manager.run_python_script(runner, args)

            # Log results
            log_file = temp_utils.get_log_path("mapping_exporter.log")
            with log_file.open("a") as fh:
                fh.write(stdout or "")
                fh.write(stderr or "")
                fh.write("Process returned code: {}\n".format(returncode))

            if returncode == 0:
                # Validate XLSX file was created
                success, message = ElementDataSerializer.deserialize_mapping_results(xlsx_path)

                if success:
                    # Count unique combinations for user feedback
                    unique_combos = count_unique_combinations(element_data)

                    output.print_md("##  Export Complete")
                    output.print_md(
                        "**Created mapping for {} unique Category/Family/Type combinations** "
                        "from {} categories".format(unique_combos, len(selected_categories))
                    )
                    output.print_md("**Categories included**: {}".format(category_summary))
                    output.print_md("**XLSX file saved to**: `{}`".format(xlsx_path))
                    output.print_md(
                        "**Next steps**: Open the XLSX file in Excel to add PBS codes, "
                        "then use the Seeder to apply codes back to the model"
                    )

                    # Show success dialog
                    forms.alert(
                        "XLSX mapping file created successfully.\n\n"
                        "Categories: {}\n"
                        "Combinations: {}\n"
                        "File: {}".format(category_summary, unique_combos, xlsx_path),
                        title="Mapping Export Complete",
                    )
                else:
                    output.print_md(" **Error**: {}".format(message))
                    forms.alert(
                        "XLSX file creation failed: {}\n\n"
                        "Check the log file for details.".format(message),
                        title="Mapping Export Error",
                    )
            else:
                output.print_md(" **Error**: XLSX generation failed")
                error_details = (stderr or "Unknown error occurred")[:500]
                output.print_md("Error details: {}".format(error_details))

                forms.alert(
                    "XLSX generation failed.\n\nError: {}\n\n"
                    "Check the log file for details.".format(error_details),
                    title="Mapping Export Error",
                )

        except Exception as e:
            output.print_md(" **Critical Error**: {}".format(str(e)))
            forms.alert(
                "Critical processing error: {}\n\n"
                "Please check your Python installation and try again.".format(str(e)),
                title="Mapping Export Error",
            )


def count_unique_combinations(element_data):
    """
    Count unique Category/Family/Type combinations from Schema v3.0 data.

    Args:
        element_data: Dict of element data in Schema v3.0 format

    Returns:
        int: Number of unique combinations
    """
    combinations = set()

    for element_id, element_info in element_data.items():
        hierarchy = element_info.get("hierarchy", {})
        category = hierarchy.get("category", "Unknown")
        family = hierarchy.get("family", "Unknown")
        type_name = hierarchy.get("type", "Unknown")

        combo_key = (category, family, type_name)
        combinations.add(combo_key)

    return len(combinations)


if __name__ == "__main__":
    main()
