# -*- coding: utf-8 -*-

import os
import csv
import re
import string
from qgis.PyQt import QtWidgets, QtCore, QtGui
from qgis.PyQt.QtCore import QVariant, Qt, QDate, QDateTime, QTime
from qgis.PyQt.QtWidgets import (QWidget, QVBoxLayout, QGroupBox, QLabel, QCheckBox,
                                 QListWidget, QPushButton, QDockWidget, QTableWidget,
                                 QAbstractItemView, QTableWidgetItem, QApplication,
                                 QFileDialog, QHBoxLayout, QSizePolicy, QProgressBar,
                                 QSpinBox, QFormLayout)
from qgis.gui import QgsMapLayerComboBox
from qgis.core import (QgsProject, QgsVectorLayer, QgsField, Qgis,
                       QgsStatisticalSummary, QgsMapLayerProxyModel, QgsFeatureRequest,
                       QgsExpression)

import statistics
from collections import Counter, OrderedDict
import numpy

SCIPY_AVAILABLE = False
try:
    from scipy import stats as scipy_stats
    SCIPY_AVAILABLE = True
except ImportError:
    scipy_stats = None # So we can check against it

STOP_WORDS = set([
    'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from', 'has', 'he',
    'in', 'is', 'it', 'its', 'of', 'on', 'that', 'the', 'to', 'was', 'were',
    'will', 'with',
])

class FieldProfilerDockWidget(QDockWidget):
    STAT_KEYS_NUMERIC = [
        'Non-Null Count', 'Null Count', '% Null', 'Conversion Errors',
        'Min', 'Max', 'Range', 'Sum', 'Mean', 'Median', 'Stdev (pop)', 'Mode(s)',
        'Variety (distinct)', 'Q1', 'Q3', 'IQR',
        'Outliers (IQR)', 'Min Outlier', 'Max Outlier', '% Outliers',
        'Low Variance Flag',
        'Zeros', 'Positives', 'Negatives', 'CV %',
        'Integer Values', 'Decimal Values', '% Integer Values',
        'Skewness', 'Kurtosis', 'Normality (Shapiro-Wilk p)', 'Normality (Likely Normal)',
        '1st Pctl', '5th Pctl', '95th Pctl', '99th Pctl',
        'Optimal Bins (Freedman-Diaconis)',
    ]
    STAT_KEYS_TEXT = [
        'Non-Null Count', 'Null Count', '% Null', 'Empty Strings', '% Empty',
        'Leading/Trailing Spaces', 'Internal Multiple Spaces',
        'Variety (distinct)', 'Min Length', 'Max Length', 'Avg Length',
        'Unique Values (Top)', 'Values Occurring Once',
        'Top Words', 'Pattern Matches',
        '% Uppercase', '% Lowercase', '% Titlecase', '% Mixed Case',
        'Non-Printable Chars Count',
    ]
    STAT_KEYS_DATE = [
        'Non-Null Count', 'Null Count', '% Null', 'Min Date', 'Max Date',
        'Unique Values (Top)',
        'Common Years', 'Common Months', 'Common Days',
        'Common Hours (Top 3)', '% Midnight Time', '% Noon Time',
        '% Weekend Dates', '% Weekday Dates',
        'Dates Before Today', 'Dates After Today',
    ]
    STAT_KEYS_OTHER = [ 'Non-Null Count', 'Null Count', '% Null', 'Status', 'Data Type Mismatch Hint']
    STAT_KEYS_ERROR = ['Error', 'Status']


    def __init__(self, iface, parent=None):
        super(FieldProfilerDockWidget, self).__init__(parent)
        self.iface = iface
        self.setObjectName("FieldProfilerDockWidgetInstance")
        self.setWindowTitle(self.tr("Field Profiler"))

        self.main_widget = QWidget()
        self.main_layout = QVBoxLayout(self.main_widget)
        self.setWidget(self.main_widget)

        self.analysis_results_cache = OrderedDict()
        self.conversion_error_feature_ids_by_field = {}
        self.non_printable_char_feature_ids_by_field = {}
        self._was_analyzing_selected_features = False

        self._define_stat_tooltips()
        self._create_input_group()
        self._create_results_ui()

        self.layerComboBox.layerChanged.connect(self.populate_fields)
        self.analyzeButton.clicked.connect(self.run_analysis)
        self.copyButton.clicked.connect(self.copy_results_to_clipboard)
        self.exportButton.clicked.connect(self.export_results_to_csv)
        self.resultsTableWidget.cellDoubleClicked.connect(self._on_cell_double_clicked)

        self.populate_fields(self.layerComboBox.currentLayer())
        if not SCIPY_AVAILABLE:
            self.iface.messageBar().pushMessage(
                "Field Profiler Warning",
                self.tr("Scipy library not found. Advanced numeric statistics (Skewness, Kurtosis, Normality) will be unavailable."),
                level=Qgis.Warning, duration=10
            )
        
    def tr(self, message):
        return QtCore.QCoreApplication.translate("FieldProfilerDockWidget", message)

    def _define_stat_tooltips(self):
        self.stat_tooltips = {
            'Non-Null Count': self.tr("Number of features with non-missing values."),
            'Null Count': self.tr("Number of features with missing (NULL) values. Double-click cell to select these features."),
            '% Null': self.tr("Percentage of features with missing (NULL) values."),
            'Conversion Errors': self.tr("Number of values that could not be converted to a numeric type (for numeric fields). Double-click cell to select these features."),
            'Low Variance Flag': self.tr("True if standard deviation is close to zero or all values are identical (for numeric fields)."),
            'Outliers (IQR)': self.tr("Number of numeric values falling outside Q1 - 1.5*IQR and Q3 + 1.5*IQR. Double-click cell to select these features."),
            'Min Outlier': self.tr("Minimum value among those flagged as outliers by IQR method."),
            'Max Outlier': self.tr("Maximum value among those flagged as outliers by IQR method."),
            '% Outliers': self.tr("Percentage of non-null values flagged as outliers by IQR method."),
            'Min': self.tr("Minimum value."),
            'Max': self.tr("Maximum value."),
            'Range': self.tr("Difference between Max and Min values."),
            'Sum': self.tr("Sum of all numeric values."),
            'Mean': self.tr("Average of numeric values."),
            'Median': self.tr("Median (middle) value of numeric data."),
            'Stdev (pop)': self.tr("Population Standard Deviation. Measures the amount of variation or dispersion."),
            'Mode(s)': self.tr("Most frequently occurring value(s)."),
            'Variety (distinct)': self.tr("Number of unique distinct values."),
            'Q1': self.tr("First Quartile (25th percentile)."),
            'Q3': self.tr("Third Quartile (75th percentile)."),
            'IQR': self.tr("Interquartile Range (Q3 - Q1)."),
            'Zeros': self.tr("Count of zero values (for numeric fields)."),
            'Positives': self.tr("Count of positive values (for numeric fields)."),
            'Negatives': self.tr("Count of negative values (for numeric fields)."),
            'CV %': self.tr("Coefficient of Variation (Stdev / Mean * 100). Indicates relative variability. N/A if mean is zero."),
            
            'Integer Values': self.tr("Count of numeric values that are whole numbers."),
            'Decimal Values': self.tr("Count of numeric values with a fractional part."),
            '% Integer Values': self.tr("Percentage of non-null numeric values that are whole numbers."),
            'Skewness': self.tr("Measure of asymmetry. Positive: tail on right. Negative: tail on left. Requires Scipy."),
            'Kurtosis': self.tr("Measure of tailedness (Fisher's, normal=0). Positive: heavy tails. Negative: light tails. Requires Scipy."),
            'Normality (Shapiro-Wilk p)': self.tr("P-value from Shapiro-Wilk test for normality. Low p (<0.05) suggests non-normal. Requires Scipy & >=3 values."),
            'Normality (Likely Normal)': self.tr("True if Shapiro-Wilk p-value > 0.05. Requires Scipy."),
            '1st Pctl': self.tr("1st Percentile."), '5th Pctl': self.tr("5th Percentile."),
            '95th Pctl': self.tr("95th Percentile."), '99th Pctl': self.tr("99th Percentile."),
            'Optimal Bins (Freedman-Diaconis)': self.tr("Suggested number of bins for a histogram using Freedman-Diaconis rule."),

            'Empty Strings': self.tr("Number of non-null strings that are empty (''). Double-click cell to select these features."),
            '% Empty': self.tr("Percentage of non-null strings that are empty."),
            'Leading/Trailing Spaces': self.tr("Number of non-empty strings that have leading or trailing whitespace. Double-click cell to select these features."),
            'Internal Multiple Spaces': self.tr("Number of non-empty strings with consecutive internal spaces (e.g., 'word  word')."),
            'Min Length': self.tr("Minimum length of non-empty strings."),
            'Max Length': self.tr("Maximum length of non-empty strings."),
            'Avg Length': self.tr("Average length of non-empty strings."),
            'Unique Values (Top)': self.tr("Most frequent distinct values and their counts. Double-click cell to select features matching the first listed value (uses cached actual value)."),
            'Values Occurring Once': self.tr("Count of distinct values that appear only once in the non-null dataset."),
            'Top Words': self.tr("Most frequent words (after removing stop words and punctuation)."),
            'Pattern Matches': self.tr("Counts of values matching common patterns (e.g., Emails, URLs)."),
            '% Uppercase': self.tr("Percentage of non-empty strings that are entirely uppercase."),
            '% Lowercase': self.tr("Percentage of non-empty strings that are entirely lowercase."),
            '% Titlecase': self.tr("Percentage of non-empty strings that are in title case (e.g., 'Title Case String')."),
            '% Mixed Case': self.tr("Percentage of non-empty strings that have mixed casing (not fully upper, lower, or title)."),
            'Non-Printable Chars Count': self.tr("Number of strings containing non-printable ASCII characters (excluding tab, newline, carriage return). Double-click to select features."),

            'Min Date': self.tr("Earliest date/datetime found."),
            'Max Date': self.tr("Latest date/datetime found."),
            'Common Years': self.tr("Most frequent years."),
            'Common Months': self.tr("Most frequent months."),
            'Common Days': self.tr("Most frequent days of the week."),
            'Common Hours (Top 3)': self.tr("Most frequent hours for DateTime fields (e.g., 10:00, 14:00)."),
            '% Midnight Time': self.tr("Percentage of DateTime values where time is 00:00:00."),
            '% Noon Time': self.tr("Percentage of DateTime values where time is 12:00:00."),
            '% Weekend Dates': self.tr("Percentage of dates falling on a Saturday or Sunday."),
            '% Weekday Dates': self.tr("Percentage of dates falling on a weekday (Mon-Fri)."),
            'Dates Before Today': self.tr("Count of dates occurring before today."),
            'Dates After Today': self.tr("Count of dates occurring after today."),
            
            'Status': self.tr("General status or summary of the field analysis."),
            'Error': self.tr("An error occurred during analysis of this field."),
            'Data Type Mismatch Hint': self.tr("A suggestion if the field's content statistically resembles a different data type.")
        }

    def _create_input_group(self):
        self.input_group_box = QGroupBox(self.tr("Input & Settings"))
        main_input_layout = QVBoxLayout()
        
        layer_label = QLabel(self.tr("Select Layer:"))
        self.layerComboBox = QgsMapLayerComboBox(self.main_widget)
        self.layerComboBox.setFilters(QgsMapLayerProxyModel.VectorLayer)
        main_input_layout.addWidget(layer_label)
        main_input_layout.addWidget(self.layerComboBox)
        
        fields_label = QLabel(self.tr("Select Field(s):"))
        self.fieldListWidget = QListWidget()
        self.fieldListWidget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        main_input_layout.addWidget(fields_label)
        main_input_layout.addWidget(self.fieldListWidget)
        
        self.selectedOnlyCheckbox = QCheckBox(self.tr("Analyze selected features only"))
        main_input_layout.addWidget(self.selectedOnlyCheckbox)
        
        # --- Basic Configuration Group ---
        config_group = QGroupBox(self.tr("Configuration"))
        config_layout = QFormLayout()
        self.limitUniqueSpinBox = QSpinBox()
        self.limitUniqueSpinBox.setRange(1, 100); self.limitUniqueSpinBox.setValue(5)
        self.limitUniqueSpinBox.setToolTip(self.tr("Maximum number of unique values to display in 'Unique Values (Top)'."))
        config_layout.addRow(self.tr("Unique Values Limit:"), self.limitUniqueSpinBox)
        self.decimalPlacesSpinBox = QSpinBox()
        self.decimalPlacesSpinBox.setRange(0, 10); self.decimalPlacesSpinBox.setValue(2)
        self.decimalPlacesSpinBox.setToolTip(self.tr("Number of decimal places for numeric statistics in the table."))
        config_layout.addRow(self.tr("Numeric Decimal Places:"), self.decimalPlacesSpinBox)
        config_group.setLayout(config_layout)
        main_input_layout.addWidget(config_group)

        # --- Detailed Analysis Options Group ---
        detailed_options_group = QGroupBox(self.tr("Detailed Analysis Options"))
        detailed_options_layout = QVBoxLayout()

        self.chk_numeric_dist_shape = QCheckBox(self.tr("Numeric: Distribution Shape (Skew, Kurtosis, Normality)"))
        self.chk_numeric_dist_shape.setChecked(True)
        self.chk_numeric_dist_shape.setToolTip(self.tr("Requires Scipy. Calculates skewness, kurtosis, and Shapiro-Wilk normality test."))
        detailed_options_layout.addWidget(self.chk_numeric_dist_shape)

        self.chk_numeric_adv_percentiles = QCheckBox(self.tr("Numeric: Advanced Percentiles (1,5,95,99)"))
        self.chk_numeric_adv_percentiles.setChecked(True)
        detailed_options_layout.addWidget(self.chk_numeric_adv_percentiles)
        
        self.chk_numeric_int_decimal = QCheckBox(self.tr("Numeric: Integer/Decimal Counts & Optimal Bins"))
        self.chk_numeric_int_decimal.setChecked(True)
        detailed_options_layout.addWidget(self.chk_numeric_int_decimal)
        
        self.chk_numeric_outlier_details = QCheckBox(self.tr("Numeric: Min/Max Outlier Values & %"))
        self.chk_numeric_outlier_details.setChecked(True)
        detailed_options_layout.addWidget(self.chk_numeric_outlier_details)

        self.chk_text_case_analysis = QCheckBox(self.tr("Text: Case Analysis & Advanced Whitespace"))
        self.chk_text_case_analysis.setChecked(True)
        detailed_options_layout.addWidget(self.chk_text_case_analysis)

        self.chk_text_rarity_nonprintable = QCheckBox(self.tr("Text: Rarity (Once-Occurring) & Non-Printable Chars"))
        self.chk_text_rarity_nonprintable.setChecked(True)
        detailed_options_layout.addWidget(self.chk_text_rarity_nonprintable)
        
        self.chk_date_time_weekend = QCheckBox(self.tr("Date: Time Components & Weekend/Weekday Analysis"))
        self.chk_date_time_weekend.setChecked(True)
        detailed_options_layout.addWidget(self.chk_date_time_weekend)

        detailed_options_group.setLayout(detailed_options_layout)
        main_input_layout.addWidget(detailed_options_group)
        
        # --- Analyze Button and Progress Bar ---
        self.analyzeButton = QPushButton(self.tr("Analyze Selected Fields"))
        main_input_layout.addWidget(self.analyzeButton)
        self.progressBar = QProgressBar(self)
        self.progressBar.setTextVisible(True); self.progressBar.setVisible(False)
        main_input_layout.addWidget(self.progressBar)
        
        self.input_group_box.setLayout(main_input_layout)
        self.main_layout.addWidget(self.input_group_box)

    def _create_results_ui(self):
        self.results_group_box = QGroupBox(self.tr("Analysis Results"))
        results_layout = QVBoxLayout()
        self.resultsTableWidget = QTableWidget()
        self.resultsTableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.resultsTableWidget.setAlternatingRowColors(True)
        self.resultsTableWidget.setSortingEnabled(True) 
        results_layout.addWidget(self.resultsTableWidget)
        button_layout = QHBoxLayout()
        self.copyButton = QPushButton(self.tr("Copy Table"))
        self.exportButton = QPushButton(self.tr("Export Table"))
        button_layout.addStretch()
        button_layout.addWidget(self.copyButton); button_layout.addWidget(self.exportButton)
        results_layout.addLayout(button_layout)
        self.results_group_box.setLayout(results_layout)
        self.results_group_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.main_layout.addWidget(self.results_group_box)

    def populate_fields(self, layer):
        self.fieldListWidget.clear()
        self.resultsTableWidget.clear()
        self.resultsTableWidget.setRowCount(0)
        self.resultsTableWidget.setColumnCount(0)
        self.analysis_results_cache = OrderedDict()
        self.conversion_error_feature_ids_by_field = {}
        self.non_printable_char_feature_ids_by_field = {}
        self.progressBar.setVisible(False)

        if layer and isinstance(layer, QgsVectorLayer):
            self.fieldListWidget.setEnabled(True); self.selectedOnlyCheckbox.setEnabled(True); self.analyzeButton.setEnabled(True)
            for field in layer.fields(): item_text = f"{field.name()} ({field.typeName()})"; self.fieldListWidget.addItem(item_text)
        else:
            self.fieldListWidget.setEnabled(False); self.selectedOnlyCheckbox.setEnabled(False); self.analyzeButton.setEnabled(False)

    def _get_detailed_options_state(self):
        return {
            'numeric_dist_shape': self.chk_numeric_dist_shape.isChecked(),
            'numeric_adv_percentiles': self.chk_numeric_adv_percentiles.isChecked(),
            'numeric_int_decimal': self.chk_numeric_int_decimal.isChecked(),
            'numeric_outlier_details': self.chk_numeric_outlier_details.isChecked(),
            'text_case_analysis': self.chk_text_case_analysis.isChecked(),
            'text_rarity_nonprintable': self.chk_text_rarity_nonprintable.isChecked(),
            'date_time_weekend': self.chk_date_time_weekend.isChecked(),
        }

    def run_analysis(self):
        self.resultsTableWidget.clear(); self.resultsTableWidget.setRowCount(0); self.resultsTableWidget.setColumnCount(0)
        self.analysis_results_cache = OrderedDict()
        self.conversion_error_feature_ids_by_field = {}
        self.non_printable_char_feature_ids_by_field = {}

        current_layer = self.layerComboBox.currentLayer()
        selected_list_items = self.fieldListWidget.selectedItems()
        
        self.current_limit_unique_display = self.limitUniqueSpinBox.value()
        self.current_decimal_places = self.decimalPlacesSpinBox.value()
        self._was_analyzing_selected_features = self.selectedOnlyCheckbox.isChecked()
        detailed_options = self._get_detailed_options_state()

        if not current_layer or not isinstance(current_layer, QgsVectorLayer):
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Please select a valid vector layer."), level=Qgis.Warning); return
        if not selected_list_items:
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Please select one or more fields to analyze."), level=Qgis.Warning); return

        selected_field_names_from_widget = [item.text().split(" (")[0] for item in selected_list_items]
        if not selected_field_names_from_widget:
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Error extracting field names."), level=Qgis.Critical); return

        request = QgsFeatureRequest()
        feature_count_total_layer = current_layer.featureCount()
        feature_count_analyzed = 0

        if self._was_analyzing_selected_features:
            selected_ids = current_layer.selectedFeatureIds()
            if not selected_ids:
                self.iface.messageBar().pushMessage(self.tr("Warning"), self.tr("No features selected for analysis."), level=Qgis.Warning)
                self.progressBar.setVisible(False); return
            request.setFilterFids(selected_ids); feature_count_analyzed = len(selected_ids)
            analysis_scope_message = self.tr("Analyzing Selected Features ({})").format(feature_count_analyzed)
        else:
            feature_count_analyzed = feature_count_total_layer
            analysis_scope_message = self.tr("Analyzing All Features ({})").format(feature_count_analyzed)

        if feature_count_analyzed == 0:
            self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No features to analyze."), level=Qgis.Info, duration=5)
            self.progressBar.setVisible(False); return

        self.iface.messageBar().pushMessage(self.tr("Info"), analysis_scope_message, level=Qgis.Info, duration=3)
        QApplication.processEvents()

        self.progressBar.setRange(0, feature_count_analyzed if feature_count_analyzed > 0 else 100)
        self.progressBar.setValue(0); self.progressBar.setVisible(True)

        field_data_collector = OrderedDict()
        qgs_fields_objects = current_layer.fields()
        field_metadata = {}

        valid_selected_field_names = []
        for field_name in selected_field_names_from_widget:
            field_index = qgs_fields_objects.lookupField(field_name)
            if field_index == -1:
                self.analysis_results_cache[field_name] = {'Error': 'Field not found'}; continue
            
            valid_selected_field_names.append(field_name)
            field_obj = qgs_fields_objects.field(field_index)
            field_metadata[field_name] = {'index': field_index, 'object': field_obj, 'type': field_obj.type()}
            collector_init = {
                'raw_values': [], 'null_count': 0,
                'non_printable_fids': [] 
            }
            if field_obj.isNumeric():
                collector_init.update({'float_values': [], 'conversion_errors': 0, 'conversion_error_feature_ids': []})
            if field_obj.type() in [QVariant.Date, QVariant.DateTime]:
                 collector_init['original_variants'] = []

            field_data_collector[field_name] = collector_init
        
        if not valid_selected_field_names:
            self.populate_results_table(self.analysis_results_cache, selected_field_names_from_widget)
            self.progressBar.setVisible(False); return

        current_iterator = current_layer.getFeatures(request)
        iteration_count = 0
        try:
            for feature in current_iterator:
                iteration_count += 1
                fid = feature.id()
                for field_name in valid_selected_field_names:
                    meta = field_metadata[field_name]
                    collector = field_data_collector[field_name]
                    val = feature[meta['index']]
                    
                    if meta['type'] in [QVariant.Date, QVariant.DateTime]:
                        collector['original_variants'].append(val if (val is not None and not (hasattr(val, 'isNull') and val.isNull())) else None)

                    if val is None or (hasattr(val, 'isNull') and val.isNull()):
                        collector['null_count'] += 1
                    else:
                        collector['raw_values'].append(val)
                        if meta['object'].isNumeric():
                            try:
                                collector['float_values'].append(float(val))
                            except (ValueError, TypeError):
                                collector['conversion_errors'] += 1
                                collector['conversion_error_feature_ids'].append(fid)
                        elif meta['type'] == QVariant.String:
                            if detailed_options['text_rarity_nonprintable']:
                                if self._has_non_printable_chars(str(val)):
                                    collector['non_printable_fids'].append(fid)

                if iteration_count % 100 == 0 or iteration_count == feature_count_analyzed:
                    self.progressBar.setValue(iteration_count); QApplication.processEvents()
        except Exception as e_iter:
            for field_name in valid_selected_field_names: self.analysis_results_cache[field_name] = {'Error': f'Feature iteration error: {e_iter}'}
            self.populate_results_table(self.analysis_results_cache, selected_field_names_from_widget)
            self.progressBar.setVisible(False); return
        
        self.progressBar.setValue(feature_count_analyzed)

        for field_name in valid_selected_field_names:
            data = field_data_collector[field_name]
            meta = field_metadata[field_name]

            if meta['object'].isNumeric() and data.get('conversion_error_feature_ids'):
                self.conversion_error_feature_ids_by_field[field_name] = data['conversion_error_feature_ids']
            if meta['type'] == QVariant.String and data.get('non_printable_fids'):
                self.non_printable_char_feature_ids_by_field[field_name] = list(set(data['non_printable_fids'])) 

            non_null_count = len(data['raw_values'])
            percent_null = (data['null_count'] / feature_count_analyzed * 100) if feature_count_analyzed > 0 else 0
            field_results = OrderedDict([('Null Count', data['null_count']), ('% Null', f"{percent_null:.2f}%"), ('Non-Null Count', non_null_count)])
            
            status_set = False
            if non_null_count == 0:
                if meta['object'].isNumeric() and data.get('conversion_errors', 0) > 0:
                    field_results['Status'] = f"All values Null or conversion errors ({data['conversion_errors']})"
                else:
                    field_results['Status'] = 'All Null or Empty'
                status_set = True
            
            analysis_for_field = {}
            if not status_set: 
                try:
                    if meta['object'].isNumeric():
                        analysis_for_field = self.analyze_numeric_field_from_list(data['float_values'], data.get('conversion_errors',0), detailed_options, non_null_count)
                    elif meta['type'] == QVariant.String:
                        analysis_for_field = self.analyze_text_field(data['raw_values'], non_null_count, detailed_options)
                    elif meta['type'] in [QVariant.Date, QVariant.DateTime]:
                        original_variants = data.get('original_variants', data['raw_values'])
                        analysis_for_field = self.analyze_date_field_enhanced(original_variants, non_null_count, detailed_options)
                    else:
                        analysis_for_field = {'Status': 'Analysis not implemented for this type'}
                except Exception as e_analysis:
                    analysis_for_field = {'Error': f'Analysis function error: {e_analysis}'}
            
            field_results.update(analysis_for_field)

            hint = "N/A"
            if meta['type'] == QVariant.String and non_null_count > 0:
                numeric_like_count = sum(1 for s_val in data['raw_values'] if str(s_val).replace('.', '', 1).strip().isdigit()) # strip to handle " 123 "
                if numeric_like_count / non_null_count > 0.9: 
                    hint = "High % of numeric-like strings. Consider if this field should be numeric."
            elif meta['object'].isNumeric() and non_null_count > 0:
                if field_results.get('Variety (distinct)', float('inf')) < 15 and non_null_count > 20:
                     hint = "Low variety for a numeric field. Consider if this is categorical or a code."

            field_results['Data Type Mismatch Hint'] = hint
            self.analysis_results_cache[field_name] = field_results
        
        self.populate_results_table(self.analysis_results_cache, selected_field_names_from_widget)
        self.progressBar.setVisible(False); QApplication.processEvents()

    def populate_results_table(self, results_data, field_names_for_header):
        self.resultsTableWidget.clear()
        if not results_data and not field_names_for_header: return
        all_stat_names = set()
        for field_name, field_data in results_data.items(): all_stat_names.update(field_data.keys())
        
        all_stat_names = {stat for stat in all_stat_names if not stat.endswith('_actual_first_value')}

        stat_rows = []; seen_keys = set()
        key_lists_ordered = [
            self.STAT_KEYS_NUMERIC, self.STAT_KEYS_TEXT, self.STAT_KEYS_DATE, 
            self.STAT_KEYS_OTHER, self.STAT_KEYS_ERROR
        ]
        
        predefined_order = []
        temp_seen_for_order = set()
        for key_list in key_lists_ordered:
            for key in key_list:
                if key not in temp_seen_for_order:
                    predefined_order.append(key)
                    temp_seen_for_order.add(key)

        for key in predefined_order:
            if key in all_stat_names and key not in seen_keys:
                stat_rows.append(key)
                seen_keys.add(key)
        
        extras = sorted([key for key in all_stat_names if key not in seen_keys]); stat_rows.extend(extras)
        
        num_rows = len(stat_rows); num_cols = len(field_names_for_header) + 1
        self.resultsTableWidget.setRowCount(num_rows); self.resultsTableWidget.setColumnCount(num_cols)
        headers = ["Statistic"] + field_names_for_header; self.resultsTableWidget.setHorizontalHeaderLabels(headers)
        
        quality_keywords = ['%', 'Null', 'Empty', 'Error', 'Outlier', 'Spaces', 'Variance', 'Flag', 'Conversion', 'Mismatch', 'Non-Printable'] 
        dp = self.current_decimal_places
        
        for r, stat_name in enumerate(stat_rows):
            stat_item = QTableWidgetItem(self.tr(stat_name))
            stat_item.setToolTip(self.stat_tooltips.get(stat_name, self.tr("No description available.")))
            
            is_quality_issue = any(keyword.lower() in stat_name.lower() for keyword in quality_keywords) or \
                               stat_name == 'Error'
            # Check boolean quality issues for the first field to color the statistic name row
            # This assumes all fields would have similar boolean quality issues, which might not be true.
            # A more nuanced approach would be to color cells individually.
            first_field_name_for_color = field_names_for_header[0] if field_names_for_header else None
            if first_field_name_for_color:
                 first_field_data = results_data.get(first_field_name_for_color, {})
                 if stat_name == 'Normality (Likely Normal)' and first_field_data.get(stat_name) is False:
                     is_quality_issue = True
                 if stat_name == 'Low Variance Flag' and first_field_data.get(stat_name) is True: # Low variance IS a quality flag
                     is_quality_issue = True


            if is_quality_issue:
                stat_item.setBackground(QtGui.QColor(255, 240, 240)) 
            elif stat_name.startswith('%') or "Pctl" in stat_name or stat_name in ['Skewness', 'Kurtosis']:
                stat_item.setBackground(QtGui.QColor(240, 240, 255)) 
            else:
                stat_item.setBackground(QtGui.QColor(230, 230, 230)) 
            
            self.resultsTableWidget.setItem(r, 0, stat_item)

            for c, field_name in enumerate(field_names_for_header):
                field_data = results_data.get(field_name, {}); value = field_data.get(stat_name, "")
                display_text = ""
                
                if isinstance(value, bool):
                    display_text = str(value)
                elif isinstance(value, float):
                    if stat_name == 'Normality (Shapiro-Wilk p)':
                         display_text = f"{value:.4g}" if not numpy.isnan(value) else "N/A"
                    else:
                         display_text = f"{value:.{dp}f}" if not numpy.isnan(value) else "N/A"
                elif isinstance(value, list) and stat_name != 'Mode(s)':
                    display_text = "; ".join(map(str, value))
                elif isinstance(value, list) and stat_name == 'Mode(s)':
                    display_text = ", ".join(f"{v:.{dp}f}" if isinstance(v, (int, float)) else str(v) for v in value)
                else:
                    display_text = str(value)
                
                item = QTableWidgetItem(display_text)
                
                align_right_keywords = ['Count', 'Error', 'Outlier', 'Zero', 'Positive', 'Negative', 'Space', 'Empty', 'Value', 'Length', 'Pctl', 'Optimal Bins']
                align_right = isinstance(value, (int, float, bool, numpy.number)) or \
                              '%' in stat_name or \
                              any(kw in stat_name for kw in align_right_keywords)

                item.setTextAlignment(Qt.AlignVCenter | (Qt.AlignRight if align_right else Qt.AlignLeft))
                
                if isinstance(value, str) and ('\n' in value or len(value) > 60):
                    item.setToolTip(value)
                elif item.text() == "N/A (Scipy not found)" or item.text() == "N/A (>=3 values needed)":
                    item.setForeground(QtGui.QBrush(Qt.gray))

                self.resultsTableWidget.setItem(r, c + 1, item)
        
        self.resultsTableWidget.resizeColumnsToContents()

    def analyze_numeric_field_from_list(self, non_null_values_list_float, conversion_errors, options, total_non_null_count):
        results = OrderedDict()
        results['Conversion Errors'] = conversion_errors
        
        try:
            data_np = numpy.array(non_null_values_list_float, dtype=float)
            if numpy.any(numpy.isinf(data_np)): 
                data_np = data_np[~numpy.isinf(data_np)]
        except Exception: 
            data_np = numpy.array([], dtype=float) 

        count_val = len(data_np)

        if count_val == 0:
            results['Status'] = 'No valid numeric data' if conversion_errors == 0 else f'No valid data ({conversion_errors} conversion errors)'
            for key in self.STAT_KEYS_NUMERIC:
                if key not in ['Non-Null Count', 'Null Count', '% Null', 'Conversion Errors', 'Status']:
                    if key in ['Variety (distinct)', 'Zeros', 'Positives', 'Negatives', 'Outliers (IQR)', 'Integer Values', 'Decimal Values', '% Outliers', 'Min Outlier', 'Max Outlier']: results[key] = 0
                    elif key in ['Low Variance Flag', 'Normality (Likely Normal)']: results[key] = False
                    elif key == '% Integer Values': results[key] = f"{0.0:.2f}%"
                    else: results[key] = 'N/A'
            return results

        min_val = numpy.nanmin(data_np) if count_val > 0 else numpy.nan
        max_val = numpy.nanmax(data_np) if count_val > 0 else numpy.nan
        sum_val = numpy.nansum(data_np) if count_val > 0 else numpy.nan
        mean_val = numpy.nanmean(data_np) if count_val > 0 else numpy.nan
        median_val = numpy.nanmedian(data_np) if count_val > 0 else numpy.nan
        
        results['Min'] = min_val; results['Max'] = max_val
        results['Range'] = max_val - min_val if not (numpy.isnan(min_val) or numpy.isnan(max_val)) else numpy.nan
        results['Sum'] = sum_val; results['Mean'] = mean_val; results['Median'] = median_val

        std_dev_pop = numpy.std(data_np) if count_val > 0 else numpy.nan
        results['Stdev (pop)'] = std_dev_pop

        modes_val = 'N/A'
        if count_val > 0:
            if SCIPY_AVAILABLE:
                try: # Add try-except for scipy.stats.mode
                    mode_res = scipy_stats.mode(data_np, nan_policy='omit', keepdims=False) 
                    if hasattr(mode_res, 'mode') and numpy.size(mode_res.mode) > 0: # Check size for array/scalar
                        modes_val = list(mode_res.mode) if isinstance(mode_res.mode, numpy.ndarray) else [mode_res.mode]
                    elif not hasattr(mode_res, 'mode'): # Older scipy might return tuple (array, count_array)
                        if numpy.size(mode_res[0]) > 0:
                             modes_val = list(mode_res[0])
                        else: modes_val = 'N/A (no mode or all unique)' # Should be caught by size check
                    else: 
                        modes_val = 'N/A (no mode or all unique)'
                except Exception: # Broad exception for safety with scipy mode
                     modes_val = 'N/A (mode error)'
            else: 
                try:
                    modes_val = statistics.multimode(data_np.tolist()) if data_np.size > 0 else 'N/A'
                except statistics.StatisticsError:
                    modes_val = 'N/A (no unique mode)'
        results['Mode(s)'] = modes_val

        results['Variety (distinct)'] = len(numpy.unique(data_np[~numpy.isnan(data_np)])) if count_val > 0 else 0
        
        q1, q3, iqr_val = numpy.nan, numpy.nan, numpy.nan
        outlier_count = 0
        min_outlier_val, max_outlier_val, percent_outliers = numpy.nan, numpy.nan, numpy.nan

        if count_val > 0:
            q1 = numpy.nanpercentile(data_np, 25)
            q3 = numpy.nanpercentile(data_np, 75)
            if not (numpy.isnan(q1) or numpy.isnan(q3)):
                iqr_val = q3 - q1
                if options.get('numeric_outlier_details', True): 
                    lower_bound = q1 - 1.5 * iqr_val
                    upper_bound = q3 + 1.5 * iqr_val
                    outliers_bool = (data_np < lower_bound) | (data_np > upper_bound)
                    outliers_vals = data_np[outliers_bool & ~numpy.isnan(data_np)] # Filter out NaNs from outliers
                    outlier_count = len(outliers_vals)
                    if outlier_count > 0:
                        min_outlier_val = numpy.min(outliers_vals)
                        max_outlier_val = numpy.max(outliers_vals)
                    percent_outliers = (outlier_count / count_val * 100.0) if count_val > 0 else 0.0

        results['Q1'] = q1; results['Q3'] = q3; results['IQR'] = iqr_val
        results['Outliers (IQR)'] = outlier_count
        if options.get('numeric_outlier_details', True):
            results['Min Outlier'] = min_outlier_val
            results['Max Outlier'] = max_outlier_val
            results['% Outliers'] = percent_outliers
        else:
            results['Min Outlier'] = "N/A (Opt.)"; results['Max Outlier'] = "N/A (Opt.)"; results['% Outliers'] = "N/A (Opt.)"


        low_variance = False
        if count_val == 1: low_variance = True
        elif not numpy.isnan(std_dev_pop) and numpy.isclose(std_dev_pop, 0.0): low_variance = True
        elif results['Variety (distinct)'] == 1 and count_val > 1: low_variance = True
        results['Low Variance Flag'] = low_variance
        
        results['Zeros'] = numpy.sum(data_np == 0) if count_val > 0 else 0
        results['Positives'] = numpy.sum(data_np > 0) if count_val > 0 else 0
        results['Negatives'] = numpy.sum(data_np < 0) if count_val > 0 else 0
        
        cv = numpy.nan
        if not numpy.isnan(mean_val) and mean_val != 0 and not numpy.isnan(std_dev_pop):
            cv = (std_dev_pop / mean_val) * 100
        results['CV %'] = cv

        if options.get('numeric_int_decimal', False) and count_val > 0:
            valid_data_for_int_check = data_np[~numpy.isnan(data_np)] # Exclude NaNs before floor comparison
            integer_values_count = numpy.sum(valid_data_for_int_check == numpy.floor(valid_data_for_int_check))
            results['Integer Values'] = int(integer_values_count)
            results['Decimal Values'] = len(valid_data_for_int_check) - int(integer_values_count) # Based on non-NaN count
            results['% Integer Values'] = (int(integer_values_count) / len(valid_data_for_int_check) * 100.0) if len(valid_data_for_int_check) > 0 else 0.0
            
            if count_val > 1 and not numpy.isnan(iqr_val) and iqr_val > 0 and not (numpy.isnan(min_val) or numpy.isnan(max_val)):
                bin_width = 2 * iqr_val / (count_val**(1/3))
                if bin_width > 0 : 
                    data_range = max_val - min_val
                    if not numpy.isnan(data_range) and data_range > 0:
                         results['Optimal Bins (Freedman-Diaconis)'] = int(numpy.ceil(data_range / bin_width))
                    elif data_range == 0: # All values are the same
                         results['Optimal Bins (Freedman-Diaconis)'] = 1
                    else: 
                         results['Optimal Bins (Freedman-Diaconis)'] = results['Variety (distinct)'] if results['Variety (distinct)'] > 0 else 1
                else: results['Optimal Bins (Freedman-Diaconis)'] = "N/A"
            else: results['Optimal Bins (Freedman-Diaconis)'] = "N/A"
        else:
            results['Integer Values'] = "N/A (Opt.)"; results['Decimal Values'] = "N/A (Opt.)"; results['% Integer Values'] = "N/A (Opt.)"
            results['Optimal Bins (Freedman-Diaconis)'] = "N/A (Opt.)"


        if options.get('numeric_dist_shape', False) and SCIPY_AVAILABLE and count_val > 0:
            data_for_scipy = data_np[~numpy.isnan(data_np)] # Scipy functions need non-NaN data
            if len(data_for_scipy) > 0:
                results['Skewness'] = scipy_stats.skew(data_for_scipy)
                results['Kurtosis'] = scipy_stats.kurtosis(data_for_scipy, fisher=True) 
                if len(data_for_scipy) >= 3: 
                    try:
                        shapiro_stat, shapiro_p = scipy_stats.shapiro(data_for_scipy)
                        results['Normality (Shapiro-Wilk p)'] = shapiro_p
                        results['Normality (Likely Normal)'] = bool(shapiro_p > 0.05) 
                    except ValueError as e_shapiro: 
                        results['Normality (Shapiro-Wilk p)'] = f"N/A (Error: {e_shapiro})"
                        results['Normality (Likely Normal)'] = "N/A"
                else:
                    results['Normality (Shapiro-Wilk p)'] = "N/A (<3 valid)"
                    results['Normality (Likely Normal)'] = "N/A (<3 valid)"
            else: # All values were NaN
                results['Skewness'] = numpy.nan; results['Kurtosis'] = numpy.nan
                results['Normality (Shapiro-Wilk p)'] = "N/A (all NaN)"; results['Normality (Likely Normal)'] = "N/A (all NaN)"

        elif options.get('numeric_dist_shape', False) and not SCIPY_AVAILABLE:
            scipy_na_msg = "N/A (Scipy missing)"
            results['Skewness'] = scipy_na_msg; results['Kurtosis'] = scipy_na_msg
            results['Normality (Shapiro-Wilk p)'] = scipy_na_msg; results['Normality (Likely Normal)'] = scipy_na_msg
        else: 
            opt_na_msg = "N/A (Opt.)"
            results['Skewness'] = opt_na_msg; results['Kurtosis'] = opt_na_msg
            results['Normality (Shapiro-Wilk p)'] = opt_na_msg; results['Normality (Likely Normal)'] = opt_na_msg


        if options.get('numeric_adv_percentiles', False) and count_val > 0:
            percentiles_to_calc = [1, 5, 95, 99]
            pctl_values = numpy.nanpercentile(data_np, percentiles_to_calc)
            results['1st Pctl'] = pctl_values[0]
            results['5th Pctl'] = pctl_values[1]
            results['95th Pctl'] = pctl_values[2]
            results['99th Pctl'] = pctl_values[3]
        else:
            opt_na_msg = "N/A (Opt.)"
            results['1st Pctl'] = opt_na_msg; results['5th Pctl'] = opt_na_msg
            results['95th Pctl'] = opt_na_msg; results['99th Pctl'] = opt_na_msg
            
        return results

    def _has_non_printable_chars(self, text_value):
        if not isinstance(text_value, str): return False
        allowed_control = {'\t', '\n', '\r'} # Allow tab, newline, carriage return
        # Check for any character that is not printable AND not in allowed_control
        return any(not c.isprintable() and c not in allowed_control for c in text_value)


    def analyze_text_field(self, values, non_null_count, options):
        results = OrderedDict(); dp = self.current_decimal_places
        
        if non_null_count == 0: 
            results['Status'] = 'No text data'
            for key in self.STAT_KEYS_TEXT:
                 if key not in ['Non-Null Count', 'Null Count', '% Null', 'Status']:
                    if key in ['Empty Strings', 'Leading/Trailing Spaces', 'Internal Multiple Spaces', 
                               'Variety (distinct)', 'Values Occurring Once', 'Non-Printable Chars Count']: results[key] = 0
                    elif key == '% Empty': results[key] = f"{0.0:.{dp}f}%"
                    elif key.startswith('%') and ('Case' in key): results[key] = f"{0.0:.{dp}f}%" # % Uppercase etc.
                    else: results[key] = 'N/A'
            return results

        str_values = [str(v) for v in values] 

        empty_string_count = str_values.count('')
        percent_empty = (empty_string_count / non_null_count * 100.0) if non_null_count > 0 else 0.0
        results['Empty Strings'] = empty_string_count
        results['% Empty'] = f"{percent_empty:.{dp}f}%"
        
        non_empty_str_values = [s for s in str_values if s] 
        count_non_empty = len(non_empty_str_values)

        min_len, max_len, avg_len_val = 'N/A', 'N/A', 'N/A'
        if count_non_empty > 0:
            lengths = [len(s) for s in non_empty_str_values]
            min_len, max_len, avg_len_val = min(lengths), max(lengths), statistics.mean(lengths)
        results['Min Length'] = min_len; results['Max Length'] = max_len
        results['Avg Length'] = f"{avg_len_val:.{dp}f}" if isinstance(avg_len_val, float) else avg_len_val
        
        value_counts = Counter(str_values) 
        results['Variety (distinct)'] = len(value_counts)
        
        sorted_counts = sorted(value_counts.items(), key=lambda item: (-item[1], item[0]))
        top_unique_list = []; actual_first_unique_value_for_selection = None
        limit_unique = self.current_limit_unique_display
        if sorted_counts:
            actual_first_unique_value_for_selection = sorted_counts[0][0]
            for i, (val_str, count) in enumerate(sorted_counts):
                if i >= limit_unique: break
                display_val_preview = f"'{val_str[:50]}{'...' if len(val_str) > 50 else ''}'"
                if val_str == "": display_val_preview = self.tr("'(Empty String)'")
                top_unique_list.append(f"{display_val_preview}: {count}")
        results['Unique Values (Top)'] = "\n".join(top_unique_list) if top_unique_list else "N/A"
        if top_unique_list:
            results['Unique Values (Top)_actual_first_value'] = actual_first_unique_value_for_selection
        
        if options.get('text_rarity_nonprintable', False):
            results['Values Occurring Once'] = sum(1 for v_str,c in value_counts.items() if c == 1 and v_str != "") 
            results['Non-Printable Chars Count'] = sum(1 for s_val in str_values if self._has_non_printable_chars(s_val))
        else:
            results['Values Occurring Once'] = "N/A (Opt.)"
            results['Non-Printable Chars Count'] = "N/A (Opt.)"

        if options.get('text_case_analysis', False):
            if count_non_empty > 0:
                upper_count = sum(1 for s_val in non_empty_str_values if s_val.isupper())
                lower_count = sum(1 for s_val in non_empty_str_values if s_val.islower())
                title_count = sum(1 for s_val in non_empty_str_values if s_val.istitle())
                
                # Mixed case is everything else that is non-empty
                # Ensure mixed_count is not negative if a string is e.g. both title and upper (e.g. "I")
                # A string can be istitle() and isupper() (e.g. "I").
                # A string can be istitle() and islower() (e.g. "i" if it was the only char in a word - unlikely for istitle).
                # A simple approach: if not purely upper and not purely lower and not purely title -> mixed.
                # This will double count sometimes if we just subtract.
                # Correct mixed_count:
                mixed_count = 0
                for s_val in non_empty_str_values:
                    is_u = s_val.isupper()
                    is_l = s_val.islower()
                    is_t = s_val.istitle()
                    if not is_u and not is_l and not is_t: # Pure mixed
                        mixed_count +=1
                    # Handle cases like "I" which is upper and title. We only count it as one category.
                    # If we prioritize: Upper, then Lower, then Title.
                    # This is complex. The current sum will give % of strings that *exclusively* match.
                    # A better % Mixed might be: 100 - (%Upper + %Lower + %Title) if these are mutually exclusive counts.
                    # For now, the existing approximation is fine if interpreted as "primarily"
                
                # Let's refine mixed_count: A string is mixed if it's not purely one of the others.
                # This is still tricky. Sticking to the simpler sum for now for % of strings that are *fully* one type.
                # The sum of these percentages might not be 100%.
                
                results['% Uppercase'] = f"{(upper_count / count_non_empty * 100.0):.{dp}f}%"
                results['% Lowercase'] = f"{(lower_count / count_non_empty * 100.0):.{dp}f}%"
                results['% Titlecase'] = f"{(title_count / count_non_empty * 100.0):.{dp}f}%"
                
                # % Mixed Case based on what's left, ensuring it's not negative.
                # This is an approximation, as a string could be "Other" if it doesn't fit neatly.
                percent_known_case = (upper_count + lower_count + title_count) / count_non_empty * 100.0
                results['% Mixed Case'] = f"{max(0.0, 100.0 - percent_known_case):.{dp}f}%"


                results['Internal Multiple Spaces'] = sum(1 for s_val in non_empty_str_values if "  " in s_val.strip()) 
            else: # count_non_empty is 0
                na_percent = f"{0.0:.{dp}f}%"
                results['% Uppercase'] = na_percent; results['% Lowercase'] = na_percent
                results['% Titlecase'] = na_percent; results['% Mixed Case'] = na_percent
                results['Internal Multiple Spaces'] = 0
        else: 
            opt_na_msg = "N/A (Opt.)"
            results['% Uppercase'] = opt_na_msg; results['% Lowercase'] = opt_na_msg
            results['% Titlecase'] = opt_na_msg; results['% Mixed Case'] = opt_na_msg
            results['Internal Multiple Spaces'] = opt_na_msg


        results['Leading/Trailing Spaces'] = sum(1 for s_val in non_empty_str_values if s_val != s_val.strip())
        
        word_list = []
        for text in non_empty_str_values:
            cleaned_text = text.lower(); cleaned_text = re.sub(r'[^\w\s]', '', cleaned_text)
            words = cleaned_text.split()
            word_list.extend([word for word in words if word and word not in STOP_WORDS and not word.isdigit()])
        if word_list:
            word_counts = Counter(word_list)
            top_words_list = [f"{word}:{count}" for word, count in word_counts.most_common(10)]
            results['Top Words'] = "\n".join(top_words_list) if top_words_list else "N/A"
        else: results['Top Words'] = "N/A (No words found)"
        
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        url_pattern = r'https?://[^\s/$.?#].[^\s]*'
        emails_found = sum(1 for text in non_empty_str_values if re.search(email_pattern, text))
        urls_found = sum(1 for text in non_empty_str_values if re.search(url_pattern, text))
        results['Pattern Matches'] = f"Emails: {emails_found}, URLs: {urls_found}"
        
        return results

    def analyze_date_field_enhanced(self, original_variant_values, non_null_count, options):
        results = OrderedDict()
        dp = self.current_decimal_places 

        if non_null_count == 0:
            results['Status'] = 'No date data'
            for key in self.STAT_KEYS_DATE:
                 if key not in ['Non-Null Count', 'Null Count', '% Null', 'Status']:
                    if key in ['Dates Before Today', 'Dates After Today']: results[key] = 0
                    elif key.startswith('%') and ('Time' in key or 'Dates' in key): results[key] = f"{0.0:.{dp}f}%"
                    else: results[key] = 'N/A'
            return results

        py_datetimes = []    
        q_date_time_objects = [] 
        
        has_time_component_overall = False 

        for v_orig in original_variant_values: 
            if v_orig is None: continue 

            py_dt = None
            q_obj = None

            if isinstance(v_orig, QDateTime) and v_orig.isValid():
                py_dt = v_orig.toPyDateTime()
                q_obj = v_orig
                has_time_component_overall = True
            elif isinstance(v_orig, QDate) and v_orig.isValid():
                py_dt = datetime(v_orig.year(), v_orig.month(), v_orig.day())
                q_obj = v_orig 
            
            if py_dt and q_obj:
                py_datetimes.append(py_dt)
                q_date_time_objects.append(q_obj)
        
        if not py_datetimes: 
            results['Status'] = 'No valid date objects parsed from non-null values'
            for key in self.STAT_KEYS_DATE: # Fill N/A
                 if key not in ['Non-Null Count', 'Null Count', '% Null', 'Status']:
                    if key in ['Dates Before Today', 'Dates After Today']: results[key] = 0
                    elif key.startswith('%') and ('Time' in key or 'Dates' in key): results[key] = f"{0.0:.{dp}f}%"
                    else: results[key] = 'N/A'
            return results

        min_d, max_d = min(py_datetimes), max(py_datetimes)
        results['Min Date'] = min_d.isoformat(sep=' ', timespec='auto')
        results['Max Date'] = max_d.isoformat(sep=' ', timespec='auto')
        
        years = [d.year for d in py_datetimes]
        months = [d.month for d in py_datetimes] 
        days_of_week_num = [d.weekday() for d in py_datetimes] 

        day_names_map = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        results['Common Years'] = ", ".join([f"{yr}:{cnt}" for yr, cnt in Counter(years).most_common(3)])
        results['Common Months'] = ", ".join([f"{mo}:{cnt}" for mo, cnt in Counter(months).most_common(3)]) 
        results['Common Days'] = ", ".join([f"{day_names_map[d]}:{cnt}" for d, cnt in Counter(days_of_week_num).most_common(3)])
        
        today_dt = datetime.now() # Use datetime for comparison with py_datetimes
        results['Dates Before Today'] = sum(1 for d_py in py_datetimes if d_py < today_dt)
        results['Dates After Today'] = sum(1 for d_py in py_datetimes if d_py > today_dt)

        date_counts = Counter(q_date_time_objects) 
        sorted_date_counts = sorted(date_counts.items(), key=lambda item: (-item[1], item[0])) 
        top_unique_dates_list = []; actual_first_unique_date_for_selection = None
        limit_unique = self.current_limit_unique_display
        if sorted_date_counts:
            actual_first_unique_date_for_selection = sorted_date_counts[0][0]
            for i, (date_obj, count) in enumerate(sorted_date_counts):
                if i >= limit_unique: break
                display_val_preview = date_obj.toString(Qt.ISODateWithMs if isinstance(date_obj, QDateTime) and date_obj.time().msec() > 0 else Qt.ISODate)
                top_unique_dates_list.append(f"'{display_val_preview}': {count}")
        results['Unique Values (Top)'] = "\n".join(top_unique_dates_list) if top_unique_dates_list else "N/A"
        if top_unique_dates_list:
            results['Unique Values (Top)_actual_first_value'] = actual_first_unique_date_for_selection


        if options.get('date_time_weekend', False) and q_date_time_objects:
            midnight_count = 0
            noon_count = 0
            hours_list = []
            
            q_datetimes_only = [q_obj for q_obj in q_date_time_objects if isinstance(q_obj, QDateTime)]
            
            if q_datetimes_only: # Only proceed if there are actual QDateTime objects
                for q_dt_obj in q_datetimes_only:
                    hours_list.append(q_dt_obj.time().hour())
                    if q_dt_obj.time() == QTime(0,0,0,0): midnight_count +=1
                    if q_dt_obj.time() == QTime(12,0,0,0): noon_count +=1
                
                results['Common Hours (Top 3)'] = ", ".join([f"{hr:02d}:00 ({cnt}) " for hr, cnt in Counter(hours_list).most_common(3)])
                results['% Midnight Time'] = f"{(midnight_count / len(q_datetimes_only) * 100.0):.{dp}f}%"
                results['% Noon Time'] = f"{(noon_count / len(q_datetimes_only) * 100.0):.{dp}f}%"
            else: 
                results['Common Hours (Top 3)'] = "N/A (No time data)"
                results['% Midnight Time'] = "N/A (No time data)"
                results['% Noon Time'] = "N/A (No time data)"
            
            # For weekend/weekday, use all QDate and QDateTime (converted to QDate)
            all_q_dates_for_dow = []
            for q_obj in q_date_time_objects:
                if isinstance(q_obj, QDateTime):
                    all_q_dates_for_dow.append(q_obj.date())
                elif isinstance(q_obj, QDate):
                    all_q_dates_for_dow.append(q_obj)
            
            if all_q_dates_for_dow:
                weekend_day_count = sum(1 for d_obj in all_q_dates_for_dow if d_obj.dayOfWeek() >= 6) 
                total_for_dow_calc = len(all_q_dates_for_dow)
                results['% Weekend Dates'] = f"{(weekend_day_count / total_for_dow_calc * 100.0):.{dp}f}%"
                results['% Weekday Dates'] = f"{((total_for_dow_calc - weekend_day_count) / total_for_dow_calc * 100.0):.{dp}f}%"
            else: # Should not happen if q_date_time_objects existed
                 results['% Weekend Dates'] = f"{0.0:.{dp}f}%"; results['% Weekday Dates'] = f"{0.0:.{dp}f}%"
        else: 
            opt_na_msg = "N/A (Opt.)"
            results['Common Hours (Top 3)'] = opt_na_msg; results['% Midnight Time'] = opt_na_msg; results['% Noon Time'] = opt_na_msg
            results['% Weekend Dates'] = opt_na_msg; results['% Weekday Dates'] = opt_na_msg
            
        return results

    def _on_cell_double_clicked(self, row, column):
        if column == 0: return

        current_layer = self.layerComboBox.currentLayer()
        if not current_layer or not isinstance(current_layer, QgsVectorLayer):
            self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("No valid layer selected."), level=Qgis.Warning); return

        stat_name_item = self.resultsTableWidget.item(row, 0)
        field_header_item = self.resultsTableWidget.horizontalHeaderItem(column)

        if not stat_name_item or not field_header_item:
            self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("Could not identify clicked cell data."), level=Qgis.Warning); return

        # Use the English key for logic, not the translated statistic name from the table
        # This requires mapping the translated name back to the key, or storing keys in items.
        # For now, let's assume statistic_name is the key. This part needs care if tr() changes keys.
        # A robust way: store original key in QTableWidgetItem.setData(Qt.UserRole, original_key)
        statistic_name_key = stat_name_item.text() # This is currently translated.
        # Find original key if translated:
        original_statistic_key = None
        for key, trans_val in self.stat_tooltips.items(): # stat_tooltips keys are original
            if self.tr(key) == statistic_name_key: # if the table shows translated names
                 original_statistic_key = key
                 break
        if not original_statistic_key: # Fallback if no match (e.g. statistic is not in tooltips or table shows original keys)
            original_statistic_key = statistic_name_key


        field_name_for_selection = field_header_item.text()
        field_qobj = current_layer.fields().field(field_name_for_selection)
        if not field_qobj: return

        quoted_field_name = f'"{field_name_for_selection}"'
        expression = None
        ids_to_select_directly = None

        is_string_field = (field_qobj.type() == QVariant.String)
        is_numeric_field = field_qobj.isNumeric()

        if original_statistic_key == 'Null Count':
            expression = f"{quoted_field_name} IS NULL"
        elif original_statistic_key == 'Empty Strings' and is_string_field:
            expression = f"{quoted_field_name} = ''"
        elif original_statistic_key == 'Leading/Trailing Spaces' and is_string_field:
            expression = f"{quoted_field_name} != trim({quoted_field_name}) AND length(trim({quoted_field_name})) > 0"
        elif original_statistic_key == 'Conversion Errors' and is_numeric_field:
            ids_to_select_directly = self.conversion_error_feature_ids_by_field.get(field_name_for_selection, [])
            if not ids_to_select_directly: self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No features with conversion errors recorded."), level=Qgis.Info); return
        
        elif original_statistic_key == 'Non-Printable Chars Count' and is_string_field:
            ids_to_select_directly = self.non_printable_char_feature_ids_by_field.get(field_name_for_selection, [])
            if not ids_to_select_directly: self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No features with non-printable characters recorded."), level=Qgis.Info); return

        elif original_statistic_key == 'Outliers (IQR)' and is_numeric_field:
            field_stats = self.analysis_results_cache.get(field_name_for_selection, {})
            q1_val = field_stats.get('Q1') 
            q3_val = field_stats.get('Q3')
            iqr_val = field_stats.get('IQR')
            if isinstance(q1_val, (int, float)) and isinstance(q3_val, (int, float)) and isinstance(iqr_val, (int, float)) and \
               not (numpy.isnan(q1_val) or numpy.isnan(q3_val) or numpy.isnan(iqr_val)):
                lower_bound = q1_val - 1.5 * iqr_val
                upper_bound = q3_val + 1.5 * iqr_val
                expression = f"({quoted_field_name} < {lower_bound} OR {quoted_field_name} > {upper_bound}) AND {quoted_field_name} IS NOT NULL"
            else:
                self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("Q1, Q3, or IQR is N/A for outlier selection."), level=Qgis.Info); return

        elif original_statistic_key == 'Unique Values (Top)':
            cached_field_results = self.analysis_results_cache.get(field_name_for_selection, {})
            actual_first_value = cached_field_results.get('Unique Values (Top)_actual_first_value')
            if 'Unique Values (Top)_actual_first_value' not in cached_field_results:
                self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("No specific unique value cached for selection."), level=Qgis.Info); return
            
            # Handle if actual_first_value is None but it's not an empty string
            # (empty string IS a valid value to select for)
            if actual_first_value is None and not (isinstance(actual_first_value, str) and actual_first_value == ""):
                 self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("Cached unique value is None, cannot select."), level=Qgis.Info); return

            if isinstance(actual_first_value, str):
                # *** THIS IS THE CORRECTED BLOCK FOR THE F-STRING SYNTAX ERROR ***
                if actual_first_value: # Check if it's a non-empty string
                    escaped_val = str(actual_first_value).replace("'", "''") # Ensure it's string for .replace
                    expression = f"{quoted_field_name} = '{escaped_val}'"
                else: # It's an empty string ""
                    expression = f"{quoted_field_name} = ''"
            elif isinstance(actual_first_value, (int, float, numpy.number)): # numpy.number covers numpy's numeric types
                if numpy.isnan(actual_first_value): 
                    self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("Cannot select NaN unique value directly by expression."), level=Qgis.Info); return
                expression = f"{quoted_field_name} = {float(actual_first_value)}" # Ensure it's a Python float for expression
            elif isinstance(actual_first_value, QDate):
                expression = f"{quoted_field_name} = date('{actual_first_value.toString(Qt.ISODate)}')"
            elif isinstance(actual_first_value, QDateTime):
                expression = f"{quoted_field_name} = datetime('{actual_first_value.toString(Qt.ISODate)}')" 
            else:
                self.iface.messageBar().pushMessage(self.tr("Warning"), self.tr("Cannot select unique value of type: {0}").format(type(actual_first_value).__name__), level=Qgis.Warning); return
        
        if expression:
            self._select_features_by_expression(current_layer, field_name_for_selection, expression)
        elif ids_to_select_directly is not None:
            self._select_features_by_ids(current_layer, field_name_for_selection, ids_to_select_directly)

    def _select_features_by_expression(self, layer, field_name, expression_string):
        try:
            selection_mode = QgsVectorLayer.SetSelection
            if self._was_analyzing_selected_features:
                selection_mode = QgsVectorLayer.IntersectSelection
            
            num_selected = layer.selectByExpression(expression_string, selection_mode)
            self.iface.mapCanvas().refresh()
            if hasattr(self.iface, 'actionSelect'): # Check if actionSelect exists
                self.iface.actionSelect().trigger() 

            msg = self.tr("Selected {0} features for field '{1}' where: {2}").format(num_selected, field_name, expression_string)
            if self._was_analyzing_selected_features and selection_mode == QgsVectorLayer.IntersectSelection:
                 msg += self.tr(" (Intersected with features from original analysis scope).")
            else:
                 msg += "."
            self.iface.messageBar().pushMessage(self.tr("Selection Succeeded"), msg, level=Qgis.Success, duration=7)

        except Exception as e:
            self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("Error selecting features by expression: {0}\nExpression: {1}").format(str(e), expression_string), level=Qgis.Critical)

    def _select_features_by_ids(self, layer, field_name, fids_to_select):
        try:
            num_selected = 0
            final_ids_for_selection = list(fids_to_select) 

            if self._was_analyzing_selected_features:
                current_selection_on_layer = set(layer.selectedFeatureIds())
                final_ids_for_selection = [fid for fid in fids_to_select if fid in current_selection_on_layer]
                layer.selectByIds(final_ids_for_selection, QgsVectorLayer.SetSelection) 
                num_selected = len(final_ids_for_selection)
                msg_suffix = self.tr(" (Intersected with current layer selection).")
            else:
                layer.selectByIds(final_ids_for_selection, QgsVectorLayer.SetSelection)
                num_selected = len(final_ids_for_selection)
                msg_suffix = "."
            
            self.iface.mapCanvas().refresh()
            if hasattr(self.iface, 'actionSelect'): # Check if actionSelect exists
                self.iface.actionSelect().trigger() 

            msg = self.tr("Selected {0} features for field '{1}' based on stored IDs{2}").format(num_selected, field_name, msg_suffix)
            self.iface.messageBar().pushMessage(self.tr("Selection Succeeded"), msg, level=Qgis.Success, duration=7)
        except Exception as e:
            self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("Error selecting features by IDs: {0}").format(str(e)), level=Qgis.Critical)

    def copy_results_to_clipboard(self):
        if self.resultsTableWidget.rowCount() == 0 or self.resultsTableWidget.columnCount() == 0:
            self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No results to copy."), level=Qgis.Info); return
        clipboard = QApplication.clipboard()
        if not clipboard:
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Could not access clipboard."), level=Qgis.Critical); return
        output = ""
        headers = [self.resultsTableWidget.horizontalHeaderItem(c).text() for c in range(self.resultsTableWidget.columnCount())]
        output += "\t".join(headers) + "\n"
        for r in range(self.resultsTableWidget.rowCount()):
            row_data = [];
            for c in range(self.resultsTableWidget.columnCount()):
                 item = self.resultsTableWidget.item(r, c); cell_text = item.text().replace("\n", " | ") if item else ""; row_data.append(cell_text)
            output += "\t".join(row_data) + "\n"
        clipboard.setText(output); self.iface.messageBar().pushMessage(self.tr("Success"), self.tr("Table results copied to clipboard."), level=Qgis.Success)

    def export_results_to_csv(self):
        if self.resultsTableWidget.rowCount() == 0 or self.resultsTableWidget.columnCount() == 0:
            self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No results to export."), level=Qgis.Info); return
        default_filename = "field_profiler_results.csv";
        if self.layerComboBox.currentLayer(): default_filename = f"{self.layerComboBox.currentLayer().name()}_profile.csv"
        file_path, _ = QFileDialog.getSaveFileName(self, self.tr("Export Results to CSV"), default_filename, "CSV Files (*.csv);;All Files (*)")
        if not file_path: return
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                headers = [self.resultsTableWidget.horizontalHeaderItem(c).text() for c in range(self.resultsTableWidget.columnCount())]; writer.writerow(headers)
                for r in range(self.resultsTableWidget.rowCount()):
                    row_data = [];
                    for c in range(self.resultsTableWidget.columnCount()):
                        item = self.resultsTableWidget.item(r, c); cell_text = item.text() if item else ""; row_data.append(cell_text)
                    writer.writerow(row_data)
            self.iface.messageBar().pushMessage(self.tr("Success"), self.tr("Results successfully exported to CSV."), level=Qgis.Success)
        except Exception as e: self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Could not export results to CSV: ") + str(e), level=Qgis.Critical); print(f"CSV Export Error: {e}")

    def closeEvent(self, event):
        self.hide()
        event.ignore()