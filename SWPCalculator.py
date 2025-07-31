import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                           QHBoxLayout, QGridLayout, QLabel, QLineEdit, QPushButton,
                           QComboBox, QTableWidget, QTableWidgetItem, QTabWidget,
                           QGroupBox, QMessageBox, QFileDialog, QHeaderView,
                           QScrollArea, QFrame, QSplitter, QCheckBox) # Added QCheckBox
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QFont, QPalette, QColor, QPixmap, QIcon
import pandas as pd
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np
from dateutil.relativedelta import relativedelta
import matplotlib
matplotlib.use('Qt5Agg')

class ModernButton(QPushButton):
    def __init__(self, text, color="#2196F3"):
        super().__init__(text)
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 11px;
            }}
            QPushButton:hover {{
                background-color: {self.darken_color(color)};
            }}
            QPushButton:pressed {{
                background-color: {self.darken_color(color, 0.3)};
            }}
        """)

    def darken_color(self, color, factor=0.2):
        # Simple color darkening
        if color == "#2196F3":
            return "#1976D2"
        elif color == "#4CAF50":
            return "#388E3C"
        elif color == "#FF9800":
            return "#F57C00"
        return color

class MatplotlibWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.figure = Figure(facecolor='white')
        self.canvas = FigureCanvas(self.figure)
        layout = QVBoxLayout()
        layout.addWidget(self.canvas)
        self.setLayout(layout)

class SWPCalculator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.results_data = []
        self.init_ui()
        self.set_default_values()
        self.apply_styles()

    def init_ui(self):
        self.setWindowTitle(" SWP Calculator")
        self.setGeometry(100, 100, 1400, 900)

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Main layout
        main_layout = QHBoxLayout(central_widget)

        # Create splitter for resizable panels
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)

        # Left panel - Input parameters
        self.create_input_panel(splitter)

        # Right panel - Results and charts
        self.create_results_panel(splitter)

        # Set splitter proportions
        splitter.setSizes([450, 950])

    def create_input_panel(self, parent):
        # Input panel widget
        input_widget = QWidget()
        input_layout = QVBoxLayout(input_widget)

        # Title
        title_label = QLabel("Investment Parameters")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        input_layout.addWidget(title_label)

        # Scroll area for inputs
        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        # Input group box
        input_group = QGroupBox("Core Parameters")
        input_group_layout = QGridLayout(input_group)

        # Input fields for core parameters
        self.input_fields = {}
        core_fields_config = [
            ("Initial Investment Amount (₹):", 'initial_amount', "line_edit"),
            ("Investment Date (DD/MM/YYYY):", 'investment_date', "line_edit"),
            ("SWP Amount (₹):", 'swp_amount', "line_edit"),
            ("SWP Start Date (DD/MM/YYYY):", 'swp_start_date', "line_edit"),
            ("SWP End Date (DD/MM/YYYY):", 'swp_end_date', "line_edit"),
            ("Expected Annual Return (%):", 'annual_return', "line_edit"),
            ("SWP Frequency:", 'frequency', "combo_box")
        ]

        for i, (label_text, field_name, field_type) in enumerate(core_fields_config):
            label = QLabel(label_text)
            label.setFont(QFont("Arial", 12)) # Larger font
            input_group_layout.addWidget(label, i, 0, Qt.AlignRight)

            if field_type == "line_edit":
                field = QLineEdit()
                field.setPlaceholderText("Enter value...")
            elif field_type == "combo_box":
                field = QComboBox()
                field.addItems(["Monthly", "Quarterly", "Half-Yearly", "Yearly"])

            field.setMinimumHeight(35) # Slightly larger height
            input_group_layout.addWidget(field, i, 1)
            self.input_fields[field_name] = field

        scroll_layout.addWidget(input_group)

        # Additional Parameters Checkbox
        self.additional_params_checkbox = QCheckBox("Show Additional Parameters")
        self.additional_params_checkbox.setFont(QFont("Arial", 11, QFont.Bold))
        self.additional_params_checkbox.setChecked(False) # Default to unchecked
        self.additional_params_checkbox.toggled.connect(self.toggle_additional_parameters)
        scroll_layout.addWidget(self.additional_params_checkbox)

        # Additional parameters group box
        self.additional_params_group = QGroupBox("Additional Parameters")
        self.additional_params_group_layout = QGridLayout(self.additional_params_group)
        self.additional_params_group.setVisible(False) # Hidden by default

        self.additional_input_fields = {} # Separate dictionary for additional fields
        additional_fields_config = [
            ("Expense Ratio (%):", 'expense_ratio', "line_edit"),
            ("Exit Load (%):", 'exit_load', "line_edit"),
            ("Inflation Rate (%):", 'inflation_rate', "line_edit"),
            ("Tax Rate on Gains (%):", 'tax_rate', "line_edit")
        ]

        for i, (label_text, field_name, field_type) in enumerate(additional_fields_config):
            label = QLabel(label_text)
            label.setFont(QFont("Arial", 12)) # Larger font
            self.additional_params_group_layout.addWidget(label, i, 0, Qt.AlignRight)

            field = QLineEdit()
            field.setPlaceholderText("Enter value...")
            field.setMinimumHeight(35) # Slightly larger height
            self.additional_params_group_layout.addWidget(field, i, 1)
            self.additional_input_fields[field_name] = field

        scroll_layout.addWidget(self.additional_params_group)

        # Buttons
        button_layout = QHBoxLayout()

        self.calc_button = ModernButton("Calculate SWP", "#2196F3")
        self.calc_button.clicked.connect(self.calculate_swp)
        button_layout.addWidget(self.calc_button)

        self.export_button = ModernButton("Export to Excel", "#4CAF50")
        self.export_button.clicked.connect(self.export_to_excel)
        button_layout.addWidget(self.export_button)

        self.reset_button = ModernButton("Reset", "#FF9800")
        self.reset_button.clicked.connect(self.reset_fields)
        button_layout.addWidget(self.reset_button)

        scroll_layout.addLayout(button_layout)
        scroll_layout.addStretch()

        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        input_layout.addWidget(scroll_area)

        parent.addWidget(input_widget)

    def toggle_additional_parameters(self, checked):
        self.additional_params_group.setVisible(checked)

    def create_results_panel(self, parent):
        # Results panel widget
        results_widget = QWidget()
        results_layout = QVBoxLayout(results_widget)

        # Title
        title_label = QLabel("Analysis Results")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        results_layout.addWidget(title_label)

        # Summary panel
        self.create_summary_panel(results_layout)

        # Tab widget for results
        self.tab_widget = QTabWidget()
        results_layout.addWidget(self.tab_widget)

        # Table tab
        self.create_table_tab()

        # Chart tab
        self.create_chart_tab()

        parent.addWidget(results_widget)

    def create_summary_panel(self, parent_layout):
        summary_group = QGroupBox("Summary")
        summary_layout = QGridLayout(summary_group)

        # Summary labels
        self.summary_labels = {}
        summary_fields = [
            ("Total Amount Withdrawn:", 'total_withdrawn'),
            ("Remaining Corpus:", 'remaining_corpus'),
            ("Months Sustainable:", 'months_sustainable'),
            ("Final Portfolio Value:", 'final_value')
        ]

        for i, (label_text, field_name) in enumerate(summary_fields):
            label = QLabel(label_text)
            label.setFont(QFont("Arial", 10, QFont.Bold))
            value_label = QLabel("₹0")
            value_label.setFont(QFont("Arial", 10, QFont.Bold))
            value_label.setStyleSheet("color: #2196F3;")

            row = i // 2
            col = (i % 2) * 2

            summary_layout.addWidget(label, row, col)
            summary_layout.addWidget(value_label, row, col + 1)

            self.summary_labels[field_name] = value_label

        parent_layout.addWidget(summary_group)

    def create_table_tab(self):
        # Table widget
        self.table_widget = QTableWidget()

        # Set up columns
        columns = ['Month', 'Date', 'Opening Balance', 'Growth', 'SWP Amount',
                  'Tax', 'Closing Balance', 'Real Value']
        self.table_widget.setColumnCount(len(columns))
        self.table_widget.setHorizontalHeaderLabels(columns)

        # Configure table appearance
        header = self.table_widget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        self.table_widget.setAlternatingRowColors(True)
        self.table_widget.setSelectionBehavior(QTableWidget.SelectRows)

        self.tab_widget.addTab(self.table_widget, "Monthly Analysis")

    def create_chart_tab(self):
        # Chart widget
        self.chart_widget = MatplotlibWidget()
        self.tab_widget.addTab(self.chart_widget, "Graphical Analysis")

    def set_default_values(self):
        """Set default values"""
        defaults = {
            'initial_amount': '10000000',  # 1 crore
            'investment_date': '01/06/2025',
            'swp_amount': '75000',
            'swp_start_date': '01/07/2025',
            'swp_end_date': '22/05/2035',
            'annual_return': '15.0',
            'frequency': 'Monthly'
        }

        additional_defaults = {
            'expense_ratio': '0.0',
            'exit_load': '0.0',
            'inflation_rate': '0.0',
            'tax_rate': '0.0'
        }

        for field_name, value in defaults.items():
            if isinstance(self.input_fields[field_name], QLineEdit):
                self.input_fields[field_name].setText(value)
            elif isinstance(self.input_fields[field_name], QComboBox):
                self.input_fields[field_name].setCurrentText(value)

        for field_name, value in additional_defaults.items():
            if field_name in self.additional_input_fields:
                self.additional_input_fields[field_name].setText(value)


    def get_input_values(self):
        """Get values from input fields"""
        values = {}
        for field_name, field_widget in self.input_fields.items():
            if isinstance(field_widget, QLineEdit):
                text = field_widget.text().strip()
                if field_name in ['initial_amount', 'swp_amount', 'annual_return']:
                    values[field_name] = float(text) if text else 0.0
                else:
                    values[field_name] = text
            elif isinstance(field_widget, QComboBox):
                values[field_name] = field_widget.currentText()

        # Get values from additional input fields
        for field_name, field_widget in self.additional_input_fields.items():
            text = field_widget.text().strip()
            values[field_name] = float(text) if text else 0.0 # These are always numeric

        return values

    def calculate_swp(self):
        try:
            # Get input values
            inputs = self.get_input_values()

            initial_amount = inputs['initial_amount']
            investment_date = datetime.strptime(inputs['investment_date'], "%d/%m/%Y")
            swp_amount = inputs['swp_amount']
            swp_start_date = datetime.strptime(inputs['swp_start_date'], "%d/%m/%Y")
            swp_end_date = datetime.strptime(inputs['swp_end_date'], "%d/%m/%Y")
            annual_return = inputs['annual_return'] / 100
            expense_ratio = inputs.get('expense_ratio', 0.0) / 100 # Use .get() with default
            exit_load = inputs.get('exit_load', 0.0) / 100 # Use .get() with default
            inflation_rate = inputs.get('inflation_rate', 0.0) / 100 # Use .get() with default
            tax_rate = inputs.get('tax_rate', 0.0) / 100 # Use .get() with default
            frequency = inputs['frequency']

            # Calculate frequency factor
            freq_map = {"Monthly": 12, "Quarterly": 4, "Half-Yearly": 2, "Yearly": 1}
            freq_factor = freq_map[frequency]

            # Calculate monthly return (net of expense ratio)
            monthly_return = ((1 + annual_return - expense_ratio) ** (1/12)) - 1
            monthly_inflation = ((1 + inflation_rate) ** (1/12)) - 1

            # Initialize variables
            current_balance = initial_amount
            current_date = swp_start_date
            results = []
            month_counter = 0
            total_withdrawn = 0
            initial_investment_remaining = initial_amount # For capital gain calculation

            # Calculate SWP month by month
            while current_date <= swp_end_date and current_balance > 0:
                month_counter += 1

                # Calculate opening balance
                opening_balance = current_balance

                # Apply monthly growth
                growth = opening_balance * monthly_return
                balance_after_growth = opening_balance + growth

                swp_withdrawal_this_month = 0
                tax_amount_this_month = 0
                exit_load_amount_this_month = 0

                # Calculate SWP withdrawal (adjust for frequency)
                if month_counter % (12 // freq_factor) == 0: # Corrected logic for first withdrawal and subsequent
                    # Calculate actual SWP amount for the period
                    swp_period_amount = swp_amount
                    
                    # Ensure we don't withdraw more than available balance
                    actual_swp_to_withdraw = min(swp_period_amount, balance_after_growth)

                    # Determine capital gain portion
                    # This is a simplified calculation and might need refinement for precise tax implications
                    # For a more accurate gain, you'd track unit cost
                    # Here, we assume gain is pro-rata from total capital vs. total investment
                    if initial_investment_remaining > 0:
                         capital_gain_ratio = (balance_after_growth - initial_investment_remaining) / balance_after_growth if balance_after_growth > 0 else 0
                         capital_gain = actual_swp_to_withdraw * capital_gain_ratio
                    else:
                         capital_gain = actual_swp_to_withdraw # If initial investment is depleted, all is gain


                    tax_amount_this_month = capital_gain * tax_rate

                    # Apply exit load if applicable (simplified: for first year from investment_date)
                    if (current_date - investment_date).days < 365:
                        exit_load_amount_this_month = actual_swp_to_withdraw * exit_load

                    final_withdrawal_deductions = actual_swp_to_withdraw + tax_amount_this_month + exit_load_amount_this_month

                    # Ensure we don't withdraw more than available balance (after accounting for tax/load)
                    if balance_after_growth >= final_withdrawal_deductions:
                        current_balance = balance_after_growth - final_withdrawal_deductions
                        swp_withdrawal_this_month = actual_swp_to_withdraw
                    else:
                        swp_withdrawal_this_month = balance_after_growth / (1 + tax_rate + exit_load) # Withdraw as much as possible net of charges
                        tax_amount_this_month = swp_withdrawal_this_month * tax_rate
                        exit_load_amount_this_month = swp_withdrawal_this_month * exit_load
                        current_balance = 0 # Portfolio depleted
                    
                    total_withdrawn += swp_withdrawal_this_month
                    initial_investment_remaining -= swp_withdrawal_this_month # Reduce initial investment part

                else:
                    # No SWP withdrawal this month, only growth
                    current_balance = balance_after_growth

                # Calculate real value (inflation adjusted)
                months_from_start = ((current_date.year - swp_start_date.year) * 12 +
                                   current_date.month - swp_start_date.month)
                
                # Handle division by zero for inflation adjustment if months_from_start is negative or zero
                if months_from_start >= 0:
                    real_value = current_balance / ((1 + monthly_inflation) ** months_from_start)
                else:
                    real_value = current_balance # No inflation adjustment if before start date

                # Store results
                results.append({
                    'Month': month_counter,
                    'Date': current_date.strftime("%d/%m/%Y"),
                    'Opening Balance': opening_balance,
                    'Growth': growth,
                    'SWP Amount': swp_withdrawal_this_month,
                    'Tax': tax_amount_this_month,
                    'Closing Balance': current_balance,
                    'Real Value': real_value
                })

                # Move to next month
                current_date += relativedelta(months=1)

                # Break if balance becomes too low and SWP is due
                if current_balance < swp_amount * 0.1 and month_counter % (12 // freq_factor) == 0:
                    break

            self.results_data = results
            self.update_display(results, total_withdrawn, current_balance, month_counter)

        except ValueError as e:
            QMessageBox.critical(self, "Error", f"Please check your input values: {str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

    def update_display(self, results, total_withdrawn, remaining_corpus, months_sustainable):
        # Update summary
        self.summary_labels['total_withdrawn'].setText(f"₹{total_withdrawn:,.2f}")
        self.summary_labels['remaining_corpus'].setText(f"₹{remaining_corpus:,.2f}")
        self.summary_labels['months_sustainable'].setText(f"{months_sustainable} months")
        self.summary_labels['final_value'].setText(f"₹{remaining_corpus:,.2f}")

        # Update table
        self.table_widget.setRowCount(len(results))

        for i, result in enumerate(results):
            items = [
                str(result['Month']),
                result['Date'],
                f"₹{result['Opening Balance']:,.2f}",
                f"₹{result['Growth']:,.2f}",
                f"₹{result['SWP Amount']:,.2f}",
                f"₹{result['Tax']:,.2f}",
                f"₹{result['Closing Balance']:,.2f}",
                f"₹{result['Real Value']:,.2f}"
            ]

            for j, item in enumerate(items):
                table_item = QTableWidgetItem(item)
                table_item.setTextAlignment(Qt.AlignCenter)
                self.table_widget.setItem(i, j, table_item)

        # Update chart
        self.update_chart(results)

    def update_chart(self, results):
        if not results:
            self.chart_widget.figure.clear()
            self.chart_widget.canvas.draw()
            return

        # Calculate total withdrawn from results
        total_withdrawn = sum(r['SWP Amount'] for r in results)

        # Clear existing chart
        self.chart_widget.figure.clear()

        # Create subplots with dynamic sizing/layout
        fig = self.chart_widget.figure
        gs = fig.add_gridspec(2, 2, wspace=0.3, hspace=0.4) # Adjust wspace and hspace
        
        ax00 = fig.add_subplot(gs[0, 0])
        ax01 = fig.add_subplot(gs[0, 1])
        ax10 = fig.add_subplot(gs[1, 0])
        ax11 = fig.add_subplot(gs[1, 1])

        fig.suptitle('SWP Analysis Charts', fontsize=16, fontweight='bold')

        months = [r['Month'] for r in results]
        closing_balance = [r['Closing Balance'] for r in results]
        real_value = [r['Real Value'] for r in results]
        swp_amounts = [r['SWP Amount'] for r in results]
        cumulative_withdrawn = np.cumsum(swp_amounts)

        # Chart 1: Portfolio Value Over Time
        ax00.plot(months, closing_balance, 'b-', linewidth=2, label='Nominal Value')
        ax00.plot(months, real_value, 'r--', linewidth=2, label='Real Value')
        ax00.set_title('Portfolio Value Over Time')
        ax00.set_xlabel('Months')
        ax00.set_ylabel('Amount (₹)')
        ax00.legend()
        ax00.grid(True, alpha=0.3)
        ax00.ticklabel_format(style='plain', axis='y')

        # Chart 2: Cumulative Withdrawal
        ax01.plot(months, cumulative_withdrawn, 'g-', linewidth=2)
        ax01.set_title('Cumulative Withdrawals')
        ax01.set_xlabel('Months')
        ax01.set_ylabel('Amount (₹)')
        ax01.grid(True, alpha=0.3)
        ax01.ticklabel_format(style='plain', axis='y')

        # Chart 3: Monthly SWP Amounts (Bar Chart)
        # Using a step for bars if there are too many months for clarity
        step = max(1, len(months) // 20)
        ax10.bar(months[::step], [swp_amounts[i] for i in range(0, len(swp_amounts), step)],
                       color='orange', alpha=0.7)
        ax10.set_title(f'Monthly SWP Withdrawals (Every {step} Months)')
        ax10.set_xlabel('Months')
        ax10.set_ylabel('SWP Amount (₹)')
        ax10.grid(True, alpha=0.3)
        ax10.ticklabel_format(style='plain', axis='y')

        # Chart 4: Portfolio Depletion Rate (Pie Chart - Final State) or Line Chart
        # Let's make this a pie chart showing initial vs. withdrawn vs. remaining if the period ends
        # If the portfolio is still active, maybe a line chart of depletion %
        
        final_balance = closing_balance[-1] if closing_balance else 0
        initial_investment = self.get_input_values()['initial_amount']
        
        if initial_investment > 0 and (total_withdrawn + final_balance) > 0:
            # Pie Chart: Distribution of Initial Investment
            labels = ['Total Withdrawn', 'Remaining Corpus']
            sizes = [total_withdrawn, final_balance]
            
            # Filter out zero values to prevent issues with pie chart
            non_zero_sizes = [s for s in sizes if s > 0]
            non_zero_labels = [labels[i] for i, s in enumerate(sizes) if s > 0]

            if non_zero_sizes: # Only plot if there's data to show
                ax11.pie(non_zero_sizes, labels=non_zero_labels, autopct='%1.1f%%', startangle=90, colors=['#FFC107', '#4CAF50'])
                ax11.set_title('Investment Outcome Distribution')
                ax11.axis('equal') # Equal aspect ratio ensures that pie is drawn as a circle.
            else:
                ax11.text(0.5, 0.5, "No data for pie chart", horizontalalignment='center', verticalalignment='center', transform=ax11.transAxes)
                ax11.set_title('Investment Outcome Distribution')
        else:
            # Fallback to Depletion Rate Line Chart if no meaningful pie chart
            if closing_balance and closing_balance[0] > 0:
                depletion_rate = [(closing_balance[0] - cb) / closing_balance[0] * 100
                                 for cb in closing_balance]
                ax11.plot(months, depletion_rate, 'm-', linewidth=2)
                ax11.set_title('Portfolio Depletion Rate')
                ax11.set_xlabel('Months')
                ax11.set_ylabel('Depletion %')
                ax11.grid(True, alpha=0.3)
            else:
                ax11.text(0.5, 0.5, "Insufficient data for chart", horizontalalignment='center', verticalalignment='center', transform=ax11.transAxes)
                ax11.set_title('Portfolio Depletion Rate')


        plt.tight_layout() # Adjust subplot parameters for a tight layout
        self.chart_widget.canvas.draw()

    def export_to_excel(self):
        if not self.results_data:
            QMessageBox.warning(self, "Warning", "Please calculate SWP first!")
            return

        try:
            # Get file path from user
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save SWP Analysis", "swp_analysis.xlsx",
                "Excel files (*.xlsx);;All files (*.*)"
            )

            if not file_path:
                return

            # Create DataFrame
            df = pd.DataFrame(self.results_data)

            # Get input values for summary
            inputs = self.get_input_values()

            # Create Excel file with multiple sheets
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Main analysis sheet
                df.to_excel(writer, sheet_name='SWP Analysis', index=False)
                
                # Auto-adjust column width for SWP Analysis sheet
                worksheet_analysis = writer.sheets['SWP Analysis']
                for column in worksheet_analysis.columns:
                    max_length = 0
                    column_name = column[0].column_letter # Get the column name (e.g., 'A', 'B')
                    for cell in column:
                        try: # Necessary to avoid error on empty cells
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2) # Add a little padding
                    worksheet_analysis.column_dimensions[column_name].width = adjusted_width

                # Summary sheet
                summary_data = {
                    'Parameter': ['Initial Investment', 'Total Withdrawn', 'Remaining Corpus',
                                'Months Sustainable', 'Expected Annual Return', 'SWP Amount',
                                'Expense Ratio', 'Exit Load', 'Inflation Rate', 'Tax Rate'],
                    'Value': [
                        f"₹{inputs['initial_amount']:,.2f}",
                        self.summary_labels['total_withdrawn'].text(),
                        self.summary_labels['remaining_corpus'].text(),
                        self.summary_labels['months_sustainable'].text(),
                        f"{inputs['annual_return']}%",
                        f"₹{inputs['swp_amount']:,.2f}",
                        f"{inputs.get('expense_ratio', 0.0)}%",
                        f"{inputs.get('exit_load', 0.0)}%",
                        f"{inputs.get('inflation_rate', 0.0)}%",
                        f"{inputs.get('tax_rate', 0.0)}%"
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)

                # Auto-adjust column width for Summary sheet
                worksheet_summary = writer.sheets['Summary']
                for column in worksheet_summary.columns:
                    max_length = 0
                    column_name = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet_summary.column_dimensions[column_name].width = adjusted_width


                # Input parameters sheet
                params_data = {
                    'Parameter': list(inputs.keys()),
                    'Value': list(inputs.values())
                }
                params_df = pd.DataFrame(params_data)
                params_df.to_excel(writer, sheet_name='Input Parameters', index=False)

                # Auto-adjust column width for Input Parameters sheet
                worksheet_params = writer.sheets['Input Parameters']
                for column in worksheet_params.columns:
                    max_length = 0
                    column_name = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet_params.column_dimensions[column_name].width = adjusted_width

            QMessageBox.information(self, "Success",
                                  f"SWP analysis exported successfully to:\n{file_path}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export to Excel: {str(e)}")

    def reset_fields(self):
        """Reset all fields to default values"""
        self.set_default_values()

        # Clear table
        self.table_widget.setRowCount(0)

        # Clear chart
        self.chart_widget.figure.clear()
        self.chart_widget.canvas.draw()

        # Reset summary
        for label in self.summary_labels.values():
            if "₹" in label.text():
                label.setText("₹0")
            else:
                label.setText("0")

        self.results_data = []
        self.additional_params_checkbox.setChecked(False) # Reset checkbox

    def apply_styles(self):
        """Apply modern styling to the application"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QLineEdit {
                border: 2px solid #ddd;
                border-radius: 5px;
                padding: 8px; /* Increased padding */
                font-size: 13px; /* Larger font */
            }
            QLineEdit:focus {
                border-color: #2196F3;
            }
            QComboBox {
                border: 2px solid #ddd;
                border-radius: 5px;
                padding: 8px; /* Increased padding */
                font-size: 13px; /* Larger font */
            }
            QComboBox:focus {
                border-color: #2196F3;
            }
            QLabel {
                font-size: 13px; /* Larger font for labels */
            }
            QCheckBox {
                font-size: 13px; /* Larger font for checkbox */
                padding: 5px 0;
            }
            QTableWidget {
                border: 1px solid #ddd;
                border-radius: 5px;
                background-color: white;
                alternate-background-color: #f9f9f9;
                font-size: 12px; /* Slightly larger table font */
            }
            QTableWidget::item {
                padding: 8px;
            }
            QTabWidget::pane {
                border: 1px solid #ddd;
                border-radius: 5px;
            }
            QTabBar::tab {
                background-color: #e1e1e1;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                font-size: 12px; /* Larger tab font */
            }
            QTabBar::tab:selected {
                background-color: #2196F3;
                color: white;
            }
        """)

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("SWP Calculator")

    app.setStyle("Fusion")  # Use Fusion style for a modern look
    app.setPalette(QPalette(QColor(240, 240, 240)))  # Set a light background color
    app.setWindowIcon(QIcon('logo.png'))

    calculator = SWPCalculator()
    calculator.show()

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()