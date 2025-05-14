<p align="center">
  <img src="./icon2.png" alt="Field Profiler Logo" width="128"/>
</p>
QGIS Field Profiler Plugin

**Version:** 0.1.1
**QGIS Minimum Version:** 3.0
**Author:** ricks (rrcuario@gmail.com)

## Description

The Field Profiler is a QGIS plugin designed to provide an advanced and comprehensive analysis of attribute data for vector layers. It helps users understand data distributions, identify quality issues, and gain deep insights into their attribute fields.

This plugin goes beyond basic statistics, offering detailed metrics for numeric, text, and date/time fields, empowering users in data cleaning, validation, and exploration workflows.

![image](https://github.com/user-attachments/assets/fd0d0d3b-6de5-49a5-a584-8411295c80e2)

## Features

*   **Comprehensive Statistics:**
    *   **Numeric Fields:** Min, Max, Mean, Median, StdDev, Sum, Range, Quartiles (Q1, Q3), IQR, Null Count, Conversion Errors, Zeros, Positives, Negatives, Coefficient of Variation, Low Variance Flag.
        *   *Detailed Numeric:* Skewness, Kurtosis, Shapiro-Wilk Normality Test (p-value & flag), 1st/5th/95th/99th Percentiles, Integer/Decimal counts, Optimal Bins (Freedman-Diaconis), detailed Outlier counts (Min/Max Outlier, % Outliers). (Requires SciPy)
    *   **Text Fields:** Min/Max/Avg Length, Variety (distinct count), Empty String count, Leading/Trailing Space count.
        *   *Detailed Text:* Case analysis (% Upper/Lower/Title/Mixed), count of strings with internal multiple spaces, count of values occurring only once, count of strings with non-printable characters.
    *   **Date/Time Fields:** Min/Max Date, Common Years/Months/Days of Week, Dates Before/After Today, Unique Values.
        *   *Detailed Date/Time:* Common Hours, % Midnight/Noon times (for DateTime fields), % Weekend/Weekday dates.
*   **Data Type Mismatch Hints:** Provides suggestions if a field's content statistically resembles a different data type.
*   **Flexible Analysis Scope:**
    *   Analyze all features in a layer.
    *   Analyze only selected features.
*   **Interactive Results Table:**
    *   Clear presentation of statistics per field.
    *   Sortable columns.
    *   Visual cues for potential data quality issues.
*   **Feature Selection from Results:**
    *   Double-click on specific statistic cells (e.g., Null Count, Empty Strings, Conversion Errors, Outliers, Top Unique Value, Non-Printable Chars) to select the corresponding features on the map.
    *   Selection can intersect with existing selections if analysis was performed on selected features.
*   **Export & Copy:**
    *   Copy results table to clipboard.
    *   Export results table to CSV.
*   **Configurable Options:**
    *   Set the limit for "Top N" unique values displayed.
    *   Control the number of decimal places for numeric statistics.
    *   Enable/disable groups of detailed (potentially performance-intensive) statistics.
*   **User-Friendly Interface:** Dockable widget integrated into the QGIS environment.

## Installation

1.  **Download:**
    *   Download the latest release ZIP
2.  **QGIS Plugin Manager:**
    *   Open QGIS.
    *   Go to `Plugins` -> `Manage and Install Plugins...`.
    *   Select `Install from ZIP`.
    *   Browse to the downloaded ZIP file and click `Install Plugin`.
3.  **Enable the Plugin:**
    *   Find "Field Profiler" in the list of installed plugins and ensure its checkbox is ticked.

Alternatively, for development or manual installation:
1.  Clone or download this repository.
2.  Copy the `field_profiler` directory into your QGIS plugins directory:
    *   **Windows:** `C:\Users\<YourUser>\AppData\Roaming\QGIS\QGIS3\profiles\default\python\plugins\`
    *   **Linux:** `~/.local/share/QGIS/QGIS3[or QGIS]\profiles\default\python\plugins\`
    *   **macOS:** `~/Library/Application Support/QGIS/QGIS3/profiles/default/python/plugins\`
3.  Restart QGIS or use the Plugin Reloader plugin.
4.  Enable "Field Profiler" in the Plugin Manager.

## Dependencies

*   **QGIS:** Version 3.0 or higher.
*   **Python Libraries:**
    *   `numpy`: Usually bundled with QGIS.
    *   `scipy`: (Optional, but recommended for advanced numeric statistics like Skewness, Kurtosis, and Normality tests). If SciPy is not found, these specific statistics will be unavailable, and a warning will be shown.

## Usage

1.  Once installed and enabled, a toolbar icon for "Field Profiler" will appear, or you can find it under the `Vector` menu (or wherever you assigned it, e.g., `Plugins -> Field Profiler`). Click to open the dock widget.
2.  **Select Layer:** Choose a vector layer from the dropdown.
3.  **Select Field(s):** Select one or more fields from the list to analyze.
4.  **Settings:**
    *   Optionally, check "Analyze selected features only" to limit analysis to the currently selected features in the chosen layer.
    *   Adjust "Unique Values Limit" and "Numeric Decimal Places" as needed.
    *   Choose which "Detailed Analysis Options" to include. Disabling some may improve performance on very large datasets.
5.  **Analyze:** Click the "Analyze Selected Fields" button.
6.  **View Results:** The results table will populate with statistics.
    *   Double-click on certain cells (e.g., 'Null Count', 'Outliers (IQR)', 'Unique Values (Top)') to select the corresponding features on the map.
7.  **Export/Copy:** Use the buttons to copy the table to the clipboard or export it as a CSV file.

## Chagelogs

 Version 0.1.1 (2025-05-14):
 populate_results_table:
 - The loop iterates using original_stat_key in enumerate(stat_rows_ordered). This original_stat_key is the English, non-translated key.
 - stat_item = QTableWidgetItem(self.tr(original_stat_key)): The displayed text is translated.
 - stat_item.setData(Qt.UserRole, original_stat_key): The crucial line. The original, non-translated key is stored with the item.
 - All internal logic (tooltips, coloring, formatting conditions) now uses original_stat_key for comparisons.
 _on_cell_double_clicked:
 original_statistic_key = stat_name_item.data(Qt.UserRole): Retrieves the stored original key directly.
 - The old loop for key, trans_val in self.stat_tooltips.items(): ... has been removed.
 - All subsequent if/elif conditions for building expressions use this original_statistic_key.
 Minor improvements made during refactoring:
 - analyze_numeric_field_from_list:
 - Added more checks for data_np.size > 0 before tolist() for statistics.multimode.
 - dded try-except TypeError around statistics.multimode to handle potential NaNs and attempt processing without them.
 - Ensured % Outliers defaults to 0.0 if no outliers or no data.
 - More robust Optimal Bins logic, including fallback and handling count_val == 1.
 - Shapiro-Wilk error handling made slightly more generic for ValueError.
 - Percentile calculation checks if data_np is all NaNs before calling numpy.nanpercentile.
  analyze_text_field:
 - Ensure str_values handles None from raw_values by converting to empty string.
 - Clarified % Mixed Case calculation by explicitly counting strings that are not purely upper, lower, or title.
 - Kept hyphens in words for Top Words calculation.
 - Slightly more robust URL regex.
 - analyze_date_field_enhanced:
 - Added datetime import from datetime.
 - Clarified that q_obj_for_unique stores QDate or QDateTime for unique value counting.
 - Min/Max Date formatting now considers if the field was primarily DateTime or Date.
 - Translated day and month names for Common Days/Months.
 - Dates Before/After Today now correctly compares only the date part.
 - Unique Values (Top)_actual_first_value now correctly stores the QDate or QDateTime object itself.
 - Selection logic for QDateTime in _on_cell_double_clicked uses datetime('...') QGIS expression function with ISO string.
 _on_cell_double_clicked:
 - Handles actual_first_value is None for 'Unique Values (Top)' by creating an IS NULL expression.
 - Improved messaging for some selection scenarios.
 _select_features_by_expression / _select_features_by_ids:
 - Removed the experimental QgsAttributeTable update attempts as they are not reliable/standard. QGIS usually handles attribute table updates automatically when selections change.
 export_results_to_csv:
 - Added encoding='utf-8-sig' for better Excel compatibility.
 - Sanitized layer name for default CSV filename.
 General:
 Used Python 3 style super().__init__(parent).


## Future Enhancements

*   User-defined regex patterns for text analysis.
*   More granular character-level analysis for text.
*   Visualizations (e.g., small histograms/bar charts in the results).
*   Saving/loading analysis configurations.

## Contributing

Contributions, bug reports, and feature requests are welcome! Please open an issue or submit a pull request.

## License

This plugin is licensed under the GNU General Public License as published by the Free Software Foundation.
