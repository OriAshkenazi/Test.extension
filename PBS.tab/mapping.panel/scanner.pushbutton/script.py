# -*- coding: utf-8 -*-  # pragma: no cover
__title__ = "Outlier Scanner"  # pragma: no cover
__doc__ = "Invoke the outlier scanner on the active document."  # pragma: no cover

import os
import shutil
import tempfile
import json
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
    output.print_md("# PBS Outlier Scanner")
    output.print_md("Starting outlier analysis...")

    # Initialize process manager  
    proc_manager = ProcessManager(timeout=60)  # Reduce timeout for faster debugging

    # Validate Python environment first
    output.print_md("**Step 1/3**: Validating Python environment...")
    is_valid, error_msg = proc_manager.validate_python_environment()
    if not is_valid:
        output.print_md("**Error**: {}".format(error_msg))
        forms.alert(
            "{}\n\nPlease install Python 3.12+ or set PBS_PYTHON_PATH environment variable.".format(
                error_msg
            ),
            title="Python Environment Error",
        )
        return
    output.print_md("Python environment validated")

    # Category selection step
    output.print_md("**Step 2/3**: Category selection...")
    available_categories = get_available_categories(DOC)

    if not available_categories:
        output.print_md(" **Error**: No categories found in the model")
        forms.alert(
            "No categories found in the model. The model may be empty or have no valid elements.",
            title="Outlier Scanner Error"
        )
        return

    output.print_md(" Found {} categories in model".format(len(available_categories)))

    # Show category selection dialog (pass discovered categories to avoid duplicate discovery)
    selected_categories = select_categories(
        DOC,
        title="Select Categories for Outlier Analysis",
        multiselect=True,
        include_all_option=True,
        available_categories=available_categories
    )

    if not selected_categories:
        output.print_md(" **Cancelled**: No categories selected")
        return

    category_summary = format_category_summary(selected_categories)
    output.print_md("Selected categories: {}".format(category_summary))

    output.print_md("**Step 3/3**: Collecting element data...")
    # Use new Schema v3.0 data extraction with category filtering
    element_data = element_sampler.sample_document(DOC, category_filter=selected_categories)
    total_elements = len(element_data)
    
    # DEBUG: Show comprehensive parameter extraction results
    if element_data:
        element_ids = list(element_data.keys())
        print("DEBUG: ========== COMPREHENSIVE PARAMETER EXTRACTION RESULTS ==========")
        
        # Show detailed results for first few elements
        for i, eid in enumerate(element_ids[:min(3, len(element_ids))]):
            elem = element_data[eid]
            print("DEBUG: ELEMENT {} (ID: {})".format(i+1, eid))
            
            built_in = elem.get("built_in_parameters", {})
            shared = elem.get("shared_parameters", {})
            project = elem.get("project_parameters", {})
            
            print("  Built-in parameters ({}): {}".format(len(built_in), built_in))
            print("  Shared parameters ({}): {}".format(len(shared), shared))  
            print("  Project parameters ({}): {}".format(len(project), project))
            print("  Hierarchy: {}".format(elem.get("hierarchy", {})))
            print("  " + "="*60)
        
        # Show summary across all elements
        total_built_in = sum(len(elem.get("built_in_parameters", {})) for elem in element_data.values())
        total_shared = sum(len(elem.get("shared_parameters", {})) for elem in element_data.values())
        total_project = sum(len(elem.get("project_parameters", {})) for elem in element_data.values())
        
        print("DEBUG: PARAMETER SUMMARY ACROSS ALL {} ELEMENTS:".format(len(element_data)))
        print("  Total built-in parameters: {}".format(total_built_in))
        print("  Total shared parameters: {}".format(total_shared))
        print("  Total project parameters: {}".format(total_project))
        avg_params = float(total_built_in + total_shared + total_project) / len(element_data) if element_data else 0.0
        print("  Average parameters per element: {:.1f}".format(avg_params))
        
        # Check if we need test data
        first_elem = element_data[element_ids[0]]
        if not any(first_elem.get(ptype, {}) for ptype in ["built_in_parameters", "shared_parameters", "project_parameters"]):
            if "built_in_parameters" not in first_elem:
                first_elem["built_in_parameters"] = {}
            first_elem["built_in_parameters"]["TEST_HEIGHT"] = 1000.0
            print("DEBUG: No real parameters found, injected single test outlier in element {}".format(element_ids[0]))
        else:
            print("DEBUG: Real parameters found - using actual data for outlier detection")
        
        print("DEBUG: ========== END PARAMETER EXTRACTION RESULTS ==========")
        print("")

    if total_elements == 0:
        output.print_md(" **Error**: No elements found in selected categories")
        forms.alert(
            "No elements found in the selected categories: {}\n\n"
            "Please select different categories or check your model.".format(category_summary),
            title="Outlier Scanner"
        )
        return

    output.print_md(
        "Collected data from {} elements using Schema v3.0 with category filtering".format(
            total_elements
        )
    )

    with temp_utils.temporary_dir() as tmpdir:
        # Ensure paths are strings for IronPython compatibility
        tmpdir_str = str(tmpdir)
        data_file = os.path.join(tmpdir_str, "scanner_data.json")
        out_csv = os.path.join(tmpdir_str, "out.csv")

        try:
            output.print_md("**Preparing analytics pipeline...**")

            # Serialize data using new Schema v3.0 data exchange module
            print("DEBUG: About to serialize {} elements".format(len(element_data)))
            try:
                ElementDataSerializer.serialize_element_data(
                    element_data, data_file, selected_categories
                )
                print("DEBUG: Serialization completed successfully")
                print("DEBUG: Data file created: {}".format(data_file))
                print("DEBUG: Data file exists: {}".format(os.path.exists(data_file)))
                if os.path.exists(data_file):
                    file_size = os.path.getsize(data_file)
                    print("DEBUG: Data file size: {} bytes".format(file_size))
                    
                    # Show first few lines of the data file
                    try:
                        with open(data_file, 'r') as f:
                            content = f.read()
                            print("DEBUG: First 1000 chars of data file:")
                            print(content[:1000])
                            
                            # Also look for our injected outliers
                            if "99999" in content:
                                print("DEBUG: Found HEIGHT outlier (99999) in data file")
                            if "88888" in content:
                                print("DEBUG: Found AREA outlier (88888) in data file")
                                
                    except Exception as read_error:
                        print("DEBUG: Could not read data file: {}".format(read_error))
                        
            except Exception as serialize_error:
                print("DEBUG: Serialization failed with error: {}".format(str(serialize_error)))
                print("DEBUG: Error type: {}".format(type(serialize_error).__name__))
                raise  # Re-raise to be caught by main exception handler

            # Find analytics script (using string paths for IronPython compatibility)
            current_file = str(__file__)
            current_dir = os.path.dirname(current_file)
            
            # Try to find scripts directory going up the directory tree
            scripts_dir = None
            search_dir = current_dir
            for _ in range(5):  # Search up to 5 levels up
                potential_scripts = os.path.join(search_dir, "scripts")
                if os.path.isdir(potential_scripts):
                    scripts_dir = potential_scripts
                    break
                parent = os.path.dirname(search_dir)
                if parent == search_dir:  # Reached filesystem root
                    break
                search_dir = parent
            
            if not scripts_dir:
                # Fallback: try using temp_utils with string conversion
                try:
                    start_path = Path(__file__).resolve()
                    scripts_path = temp_utils.find_scripts_dir(start_path)
                    scripts_dir = str(scripts_path)
                except Exception:
                    # Last resort: use relative path
                    scripts_dir = os.path.join(current_dir, "..", "..", "..", "scripts")
                    
            output.print_md("Using scripts dir {}".format(scripts_dir))
            runner = os.path.join(scripts_dir, "analytics_runner.py")

            # Prepare arguments (ensure all paths are strings)
            args = [data_file, "--csv-path", out_csv]
            # Enable debug mode for detailed analytics output
            args.append("--debug-grouping")
            output.print_md("Debug mode enabled - detailed grouping information will be logged")

            output.print_md("**Running 5-stage statistical analysis pipeline...**")
            output.print_md(
                "Analyzing {} elements from {} categories...".format(
                    total_elements, len(selected_categories)
                )
            )
            output.print_md("Each stage will be logged for debugging...")

            # Run staged debugging pipeline
            print("DEBUG: Starting staged debugging pipeline...")
            
            # Stage 1: JSON Loading and Validation
            print("DEBUG: === STAGE 1: JSON LOADING ===")
            stage1_script = os.path.join(scripts_dir, "debug_stage1_json_load.py")
            if not os.path.exists(stage1_script):
                raise Exception("Stage 1 script not found: {}".format(stage1_script))
            
            stage1_returncode, stage1_stdout, stage1_stderr = proc_manager.run_python_script(
                stage1_script, [data_file]
            )
            
            print("DEBUG: Stage 1 return code: {}".format(stage1_returncode))
            if stage1_stderr:
                print("DEBUG: Stage 1 messages:")
                print(stage1_stderr)
            
            if stage1_returncode != 0:
                print("DEBUG: Stage 1 failed - JSON loading/validation error")
                raise Exception("Stage 1 failed: {}".format(stage1_stderr[-500:]))
            
            # Parse stage 1 output file path
            if stage1_stdout and stage1_stdout.startswith("STAGE1_SUCCESS:"):
                stage1_output_file = stage1_stdout.split("STAGE1_SUCCESS:", 1)[1].strip()
                print("DEBUG: Stage 1 output file: {}".format(stage1_output_file))
            else:
                raise Exception("Stage 1 did not return valid output file path")
            
            # Stage 2: Schema Validation
            print("DEBUG: === STAGE 2: SCHEMA VALIDATION ===")
            stage2_script = os.path.join(scripts_dir, "debug_stage2_schema_validate.py")
            if not os.path.exists(stage2_script):
                raise Exception("Stage 2 script not found: {}".format(stage2_script))
            
            stage2_returncode, stage2_stdout, stage2_stderr = proc_manager.run_python_script(
                stage2_script, [stage1_output_file]
            )
            
            print("DEBUG: Stage 2 return code: {}".format(stage2_returncode))
            if stage2_stderr:
                print("DEBUG: Stage 2 messages:")
                print(stage2_stderr)
            
            if stage2_returncode != 0:
                print("DEBUG: Stage 2 failed - schema validation error")
                raise Exception("Stage 2 failed: {}".format(stage2_stderr[-500:]))
            
            # Parse stage 2 output file path
            if stage2_stdout and stage2_stdout.startswith("STAGE2_SUCCESS:"):
                stage2_output_file = stage2_stdout.split("STAGE2_SUCCESS:", 1)[1].strip()
                print("DEBUG: Stage 2 output file: {}".format(stage2_output_file))
            else:
                raise Exception("Stage 2 did not return valid output file path")
            
            # Stage 3: Element Grouping
            print("DEBUG: === STAGE 3: ELEMENT GROUPING ===")
            stage3_script = os.path.join(scripts_dir, "debug_stage3_element_grouping.py")
            if not os.path.exists(stage3_script):
                raise Exception("Stage 3 script not found: {}".format(stage3_script))
            
            stage3_returncode, stage3_stdout, stage3_stderr = proc_manager.run_python_script(
                stage3_script, [stage2_output_file]
            )
            
            print("DEBUG: Stage 3 return code: {}".format(stage3_returncode))
            if stage3_stderr:
                print("DEBUG: Stage 3 messages:")
                print(stage3_stderr)
            
            if stage3_returncode != 0:
                print("DEBUG: Stage 3 failed - element grouping error")
                raise Exception("Stage 3 failed: {}".format(stage3_stderr[-500:]))
            
            # Parse stage 3 output file path
            if stage3_stdout and stage3_stdout.startswith("STAGE3_SUCCESS:"):
                stage3_output_file = stage3_stdout.split("STAGE3_SUCCESS:", 1)[1].strip()
                print("DEBUG: Stage 3 output file: {}".format(stage3_output_file))
            else:
                raise Exception("Stage 3 did not return valid output file path")
            
            # Stage 4: Outlier Detection
            print("DEBUG: === STAGE 4: OUTLIER DETECTION ===")
            stage4_script = os.path.join(scripts_dir, "debug_stage4_outlier_detection.py")
            if not os.path.exists(stage4_script):
                raise Exception("Stage 4 script not found: {}".format(stage4_script))
            
            stage4_returncode, stage4_stdout, stage4_stderr = proc_manager.run_python_script(
                stage4_script, [stage3_output_file]
            )
            
            print("DEBUG: Stage 4 return code: {}".format(stage4_returncode))
            if stage4_stderr:
                print("DEBUG: Stage 4 messages:")
                print(stage4_stderr)
            
            if stage4_returncode != 0:
                print("DEBUG: Stage 4 failed - outlier detection error")
                raise Exception("Stage 4 failed: {}".format(stage4_stderr[-500:]))
            
            # Parse stage 4 output file path
            if stage4_stdout and stage4_stdout.startswith("STAGE4_SUCCESS:"):
                stage4_output_file = stage4_stdout.split("STAGE4_SUCCESS:", 1)[1].strip()
                print("DEBUG: Stage 4 output file: {}".format(stage4_output_file))
            else:
                raise Exception("Stage 4 did not return valid output file path")
            
            # Stage 5: CSV Output Generation
            print("DEBUG: === STAGE 5: CSV OUTPUT ===")
            stage5_script = os.path.join(scripts_dir, "debug_stage5_csv_output.py")
            if not os.path.exists(stage5_script):
                raise Exception("Stage 5 script not found: {}".format(stage5_script))
            
            stage5_returncode, stage5_stdout, stage5_stderr = proc_manager.run_python_script(
                stage5_script, [stage4_output_file, "--csv-path", out_csv]
            )
            
            print("DEBUG: Stage 5 return code: {}".format(stage5_returncode))
            if stage5_stderr:
                print("DEBUG: Stage 5 messages:")
                print(stage5_stderr)
            
            if stage5_returncode != 0:
                print("DEBUG: Stage 5 failed - CSV output error")
                raise Exception("Stage 5 failed: {}".format(stage5_stderr[-500:]))
            
            print("DEBUG: Analysis pipeline completed - all stages successful")
            
            # For CSV generation stage, check if CSV file was created
            if os.path.exists(out_csv):
                returncode = 0
                stdout = str(out_csv)
                stderr = ""
            else:
                returncode = 0
                stdout = "No outliers found"
                stderr = ""

            # Log results (using string paths for IronPython compatibility)
            log_file = str(temp_utils.get_log_path("scanner.log"))
            with open(log_file, "a") as fh:
                fh.write(stdout or "")
                fh.write(stderr or "")
                fh.write("Process returned code: {}\n".format(returncode))

            if returncode == 0:
                if os.path.exists(out_csv):
                    # Process results and count unique elements (not CSV rows)
                    count, results = ElementDataSerializer.deserialize_results(out_csv)
                    
                    # Count unique element IDs to avoid double-counting
                    unique_element_ids = set()
                    if results:
                        # Read CSV and extract unique element IDs
                        import csv
                        with open(out_csv, 'r') as csvfile:
                            reader = csv.DictReader(csvfile)
                            for row in reader:
                                unique_element_ids.add(row.get('ElementId', ''))
                    
                    unique_count = len(unique_element_ids)

                    # Move CSV to final location (using string paths)
                    dest_dir = os.path.join(tempfile.gettempdir(), "pbs-handler")
                    if not os.path.exists(dest_dir):
                        os.makedirs(dest_dir)
                    dest_path = os.path.join(dest_dir, 
                        "outliers_" + datetime.now().strftime("%Y%m%d%H%M%S") + ".csv"
                    )
                    shutil.move(out_csv, dest_path)

                    # Display results with proper formatting
                    if unique_count > 0:
                        output.print_md("## Analysis Complete")
                        output.print_md(
                            "**Found {} unique outlier elements** ({} total violations) "
                            "out of {} total elements from {} categories".format(
                                unique_count, count, total_elements, len(selected_categories)
                            )
                        )
                        output.print_md("**Analyzed categories**: {}".format(category_summary))
                        output.print_md("**Results saved to**: `{}`".format(dest_path))
                        output.print_md(
                            "**Next steps**: Open the CSV file to review flagged elements "
                            "and reasons"
                        )

                        # Show success dialog
                        forms.alert(
                            "Found {} unique outlier elements ({} total violations) "
                            "from {} categories.\n\n"
                            "Categories analyzed: {}\n\n"
                            "Results saved to:\n{}".format(
                                unique_count, count, len(selected_categories), category_summary, dest_path
                            ),
                            title="Outliers Detected",
                        )
                    else:
                        output.print_md("## Analysis Complete")
                        output.print_md(
                            "**No outliers detected** - all {} elements from {} categories "
                            "are within expected parameters".format(
                                total_elements, len(selected_categories)
                            )
                        )
                        output.print_md("**Analyzed categories**: {}".format(category_summary))
                        forms.alert(
                            "No outliers detected from {} categories.\n\n"
                            "Categories analyzed: {}\n\n"
                            "All elements are within expected statistical parameters.".format(
                                len(selected_categories), category_summary
                            ),
                            title="Outlier Scanner",
                        )
                else:
                    output.print_md("**Error**: No results file generated")
                    forms.alert(
                        "Analysis completed but no results file was generated.",
                        title="Outlier Scanner",
                    )
            else:
                output.print_md("**Error**: Analytics processing failed")
                error_details = (stderr or "Unknown error occurred")[:500]
                output.print_md("Error details: {}".format(error_details))

                forms.alert(
                    "Analytics processing failed.\n\nError: {}\n\n"
                    "Check the log file for details.".format(error_details),
                    title="Outlier Scanner Error",
                )

        except Exception as e:
            import traceback
            
            # Capture full error details
            error_details = {
                "error_message": str(e),
                "error_type": type(e).__name__,
                "traceback": traceback.format_exc()
            }
            
            # Log detailed error to file  
            # Save error log to PBS-Handler logs directory
            try:
                pbs_logs_dir = os.path.join(os.path.expanduser("~"), "pyRevit", "PBS-Handler", "logs")
                if not os.path.exists(pbs_logs_dir):
                    os.makedirs(pbs_logs_dir)
                error_log_path = os.path.join(pbs_logs_dir, "pbs_scanner_error_{}.txt".format(
                    datetime.now().strftime("%Y%m%d_%H%M%S")
                ))
            except Exception:
                # Fallback to temp
                error_log_path = tempfile.gettempdir() + "/pbs_scanner_error_{}.txt".format(
                    datetime.now().strftime("%Y%m%d_%H%M%S")
                )
            
            try:
                with open(error_log_path, "w") as f:
                    f.write("=" * 80 + "\n")
                    f.write("PBS HANDLER SCANNER ERROR LOG\n")
                    f.write("Time: {}\n".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    f.write("=" * 80 + "\n")
                    f.write("Error Type: {}\n".format(error_details["error_type"]))
                    f.write("Error Message: {}\n".format(error_details["error_message"]))
                    f.write("\nFull Traceback:\n")
                    f.write(error_details["traceback"])
                    f.write("\n" + "=" * 80 + "\n")
                    f.write("Total Elements Collected: {}\n".format(total_elements))
                    f.write("Selected Categories: {}\n".format(selected_categories))
                    f.write("=" * 80 + "\n")
                    
                print("\n" + "="*60)
                print("SCANNER ERROR LOG SAVED TO:")
                print("{}".format(error_log_path)) 
                print("="*60)
                print("\nPress ENTER to continue...")
                try:
                    raw_input()  # Python 2.7 compatible pause
                except NameError:
                    input()  # Python 3 fallback
                    
            except Exception:
                pass  # Don't fail if logging fails
            
            output.print_md(" **Critical Error**: {} (Error log: {})".format(str(e), error_log_path))
            forms.alert(
                "Critical processing error: {}\n\n"
                "Detailed error log saved to:\n{}\n\n"
                "Please check your Python installation and try again.".format(str(e), error_log_path),
                title="Scanner Error",
            )


if __name__ == "__main__":
    main()
