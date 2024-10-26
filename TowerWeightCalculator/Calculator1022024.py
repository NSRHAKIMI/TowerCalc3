import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QComboBox, QPushButton, QVBoxLayout, QHBoxLayout, QGroupBox, \
    QGridLayout, QTableWidget, QTableWidgetItem, QHeaderView, QFileDialog, QMessageBox
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtCore import Qt

# Define tower data for options and weights
tower_data = {
    "NS5": {
        "BTB_options": {
            "BASIC TOWER BODY WHEN WITH E3 BODY EXTENSION": 7657.795000,
            "BASIC TOWER BODY WHEN WITH E6 to E12 BODY EXTENSIONS": 7318.671000,
            "BASIC TOWER BODY WHEN WITH -4M to +2M LEG EXTENSIONS": 7657.795000,
            "BASIC TOWER BODY WHEN WITH -5M LEG EXTENSION": 7658.179000,
        },
        "BE_values": {
            "E0": 0.0,
            "E3": 1398.814000,
            "E6": 2535.785000,
            "E9": 3185.689000,
            "E12": 4297.383000
        },
        "Leg_values": {
            "+0M": 471.061000,
            "-5M": 171.491000,
            "-4M": 236.135000,
            "-3M": 264.864000,
            "-2M": 342.284000,
            "-1M": 407.650000,
            "+1M": 535.807000,
            "+2M": 623.249000
        }
    },
    "LA5": {
        "BTB_options": {
            "BASIC TOWER BODY WHEN WITH -4M & -5M LEG EXTENSIONS": 11175.495000,
            "BASIC TOWER BODY WHEN WITH BE OR -3M,-2M,-1M,+0M,+1M AND +2M LEG EXTENSIONS": 11173.535000,
        },
        "BE_values": {
            "E0": 0.0,
            "E3": {
                "with -4M & -5M leg extensions": 2174.258000,
                "with -3M to +2M leg extensions": 2151.228000
            },
            "E6": {
                "with -4M & -5M leg extensions": 3596.067000,
                "with -3M to +2M leg extensions": 3571.704000
            },
            "E9": {
                "with -4M & -5M leg extensions": 4696.969000,
                "with -3M to +2M leg extensions": 4693.660000
            }
        },
        "Leg_values": {
            "-5M": 187.406000,
            "-4M": 266.396000,
            "-3M": 335.810000,
            "-2M": 416.244000,
            "-1M": 514.038000,
            "+0M": 588.899000,
            "+1M": 669.806000,
            "+2M": 765.779000
        }
    },
    "MA5": {
        "BTB_options": {
            "BASIC TOWER BODY WHEN WITH -4M & -5M LEG EXTENSIONS": 13726.899000,
            "BASIC TOWER BODY WHEN WITH BE OR -3M,-2M,-1M,+0M,+1M AND +2M LEG EXTENSIONS": 13724.584000,
        },
        "BE_values": {
            "E0": 0.0,
            "E3": {
                "with -4M & -5M leg extensions": 2228.388000,
                "with -3M to +2M leg extensions": 2228.100000
            },
            "E6": {
                "with -4M & -5M leg extensions": 3690.482000,
                "with -3M to +2M leg extensions": 3688.048000
            },
            "E9": {
                "with -4M & -5M leg extensions": 5240.190000,
                "with -3M to +2M leg extensions": 5237.598000
            }
        },
        "Leg_values": {
            "-5M": 219.517000,
            "-4M": 307.324000,
            "-3M": 357.019000,
            "-2M": 448.234000,
            "-1M": 576.856000,
            "+0M": 668.408000,
            "+1M": 757.754000,
            "+2M": 852.382000
        }
    },
    "HA5/DE5": {
        "BTB_options": {
            "BASIC TOWER BODY WHEN WITH ALL BODY EXTENSION": 14513.530000,
            "BASIC TOWER BODY WHEN WITH -4M & -5M LEG EXTENSIONS": 14515.305000,
            "BASIC TOWER BODY WHEN WITH -3M,-2M,-1M,+0M,+1M & +2M LEG EXTENSIONS": 14513.530000,
        },
        "BE_values": {
            "E0": 0.0,
            "E3": 2691.752000,
            "E6": 4344.253000,
            "E9": 5849.610000
        },
        "Leg_values": {
            "-5M": 322.898000,
            "-4M": 408.242000,
            "-3M": 516.451000,
            "-2M": 585.859000,
            "-1M": 705.206000,
            "+0M": 807.840000,
            "+1M": 936.229000,
            "+2M": 1044.828000
        }
    }
}


class TowerCalculator(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 800)
        self.setWindowTitle("Tower Weight Calculator")
        self.setStyleSheet("background-color: #EAF6F6;")

        main_layout = QVBoxLayout()
        self.setLayout(main_layout)
        main_font = QFont("Verdana", 10)
        label_style = "color: #333333; font-weight: bold;"

        top_section_layout = QGridLayout()
        top_section_widget = QWidget()
        top_section_widget.setLayout(top_section_layout)
        main_layout.addWidget(top_section_widget, stretch=1)

        # Tower Options Group
        label_group = QGroupBox("Tower Options")
        label_group.setStyleSheet("border: 1px solid #CCCCCC; padding: 10px; background-color: #F0F0F0;")
        label_layout = QVBoxLayout()
        label_group.setLayout(label_layout)

        # Selection Group
        input_group = QGroupBox("Selection")
        input_group.setStyleSheet("border: 1px solid #CCCCCC; padding: 10px; background-color: #FFFFFF;")
        input_layout = QVBoxLayout()
        input_group.setLayout(input_layout)

        # Tower Type
        tower_label = QLabel("Tower Type:")
        tower_label.setFont(main_font)
        tower_label.setStyleSheet(label_style)
        label_layout.addWidget(tower_label)

        self.tower_type = QComboBox()
        self.tower_type.addItems(tower_data.keys())
        self.tower_type.currentIndexChanged.connect(self.update_options)
        input_layout.addWidget(self.tower_type)

        # Body Extension ComboBox
        be_label = QLabel("Body Extension:")
        be_label.setFont(main_font)
        be_label.setStyleSheet(label_style)
        label_layout.addWidget(be_label)

        self.body_extension = QComboBox()
        input_layout.addWidget(self.body_extension)

        # Leg Extensions
        self.leg_extensions = []
        for i in range(4):
            leg_label = QLabel(f"Leg {i + 1} Extension:")
            leg_label.setFont(main_font)
            leg_label.setStyleSheet(label_style)
            label_layout.addWidget(leg_label)

            leg_ext = QComboBox()
            input_layout.addWidget(leg_ext)
            self.leg_extensions.append(leg_ext)

        top_section_layout.addWidget(label_group, 0, 0, 1, 1)
        top_section_layout.addWidget(input_group, 0, 1, 1, 2)

        # Buttons Layout
        button_layout = QHBoxLayout()
        self.calculate_button = QPushButton("Calculate")
        self.calculate_button.setFont(QFont("Verdana", 11, QFont.Bold))
        self.calculate_button.setStyleSheet(
            "background-color: #4CAF50; color: white; padding: 8px 15px; border-radius: 5px;")
        self.calculate_button.clicked.connect(self.calculate_and_add_row)

        self.remove_button = QPushButton("Remove")
        self.remove_button.setFont(QFont("Verdana", 11, QFont.Bold))
        self.remove_button.setStyleSheet(
            "background-color: #f44336; color: white; padding: 8px 15px; border-radius: 5px;")
        self.remove_button.clicked.connect(self.remove_selected_row)

        self.remove_all_button = QPushButton("Remove All")
        self.remove_all_button.setFont(QFont("Verdana", 11, QFont.Bold))
        self.remove_all_button.setStyleSheet(
            "background-color: #f44336; color: white; padding: 8px 15px; border-radius: 5px;")
        self.remove_all_button.clicked.connect(self.remove_all_rows)

        self.export_button = QPushButton("Export")
        self.export_button.setFont(QFont("Verdana", 11, QFont.Bold))
        self.export_button.setStyleSheet(
            "background-color: #2196F3; color: white; padding: 8px 15px; border-radius: 5px;")
        self.export_button.clicked.connect(self.export_to_excel)

        button_layout.addWidget(self.calculate_button)
        button_layout.addWidget(self.remove_button)
        button_layout.addWidget(self.remove_all_button)
        button_layout.addWidget(self.export_button)
        input_layout.addLayout(button_layout)

        # Table for displaying calculations
        self.table = QTableWidget()
        self.table.setColumnCount(11)  # Added BTB, so 11 columns now
        self.table.setHorizontalHeaderLabels([
            "SNO", "Tower_Type", "BTB", "BE", "Leg A", "Leg B", "Leg C", "Leg D", "BB+BE", "ALL_Legs", "Tower_Weight"
        ])
        self.table.setStyleSheet("background-color: #FFFFFF; border: 1px solid #CCCCCC;")
        self.table.verticalHeader().setVisible(False)

        main_layout.addWidget(self.table, stretch=2)
        self.update_options()
        self.showMaximized()

    def update_options(self):
        tower_type = self.tower_type.currentText()
        data = tower_data.get(tower_type, {})

        # Disconnect signals temporarily to prevent recursion
        self.body_extension.blockSignals(True)

        # Initialize BE values based on tower type
        self.update_body_extension_options()

        # Initialize Leg values
        for leg_ext in self.leg_extensions:
            leg_ext.clear()
            leg_ext.addItems(data["Leg_values"].keys())
            leg_ext.setCurrentText("+0M")

        # Reconnect signals
        self.body_extension.blockSignals(False)

    def update_body_extension_options(self):
        tower_type = self.tower_type.currentText()
        data = tower_data.get(tower_type, {})

        # Clear the BE ComboBox before updating
        self.body_extension.clear()

        if tower_type in ["LA5", "MA5"]:
            self.body_extension.addItems(["E0", "E3", "E6", "E9"])
        else:
            self.body_extension.addItems(data["BE_values"].keys())

        self.body_extension.setCurrentText("E0")

    def calculate_and_add_row(self):
        tower_type = self.tower_type.currentText()
        data = tower_data.get(tower_type, {})

        # Retrieve BE weights
        selected_be = self.body_extension.currentText()
        BE_weight = 0.0

        if selected_be == "E0":
            BE_weight = 0.0
        elif tower_type in ["LA5", "MA5"]:
            be_options = data["BE_values"][selected_be]
            if self.leg_extensions[0].currentText() in ["-4M", "-5M"]:
                BE_weight = be_options["with -4M & -5M leg extensions"] / 1000
            else:
                BE_weight = be_options["with -3M to +2M leg extensions"] / 1000
        elif tower_type == "HA5/DE5":  # Correct handling for HA5/DE5 towers
            BE_weight = data["BE_values"].get(selected_be, 0) / 1000
        else:
            BE_weight = data["BE_values"].get(selected_be, 0) / 1000

        # Automatically select BTB based on BE and Leg extensions
        if tower_type == "NS5":
            if selected_be == "E3":
                BTB_weight = data["BTB_options"]["BASIC TOWER BODY WHEN WITH E3 BODY EXTENSION"] / 1000
            elif any(ext == "-5M" for ext in [leg.currentText() for leg in self.leg_extensions]):
                BTB_weight = data["BTB_options"]["BASIC TOWER BODY WHEN WITH -5M LEG EXTENSION"] / 1000
            elif selected_be in ["E6", "E9", "E12"]:
                BTB_weight = data["BTB_options"]["BASIC TOWER BODY WHEN WITH E6 to E12 BODY EXTENSIONS"] / 1000
            else:
                BTB_weight = data["BTB_options"]["BASIC TOWER BODY WHEN WITH -4M to +2M LEG EXTENSIONS"] / 1000
        elif tower_type == "HA5/DE5":  # Fix for HA5/DE5 tower type logic
            if "-4M" in [leg.currentText() for leg in self.leg_extensions] or "-5M" in [leg.currentText() for leg in
                                                                                        self.leg_extensions]:
                BTB_weight = data["BTB_options"]["BASIC TOWER BODY WHEN WITH -4M & -5M LEG EXTENSIONS"] / 1000
            else:
                BTB_weight = data["BTB_options"][
                                 "BASIC TOWER BODY WHEN WITH -3M,-2M,-1M,+0M,+1M & +2M LEG EXTENSIONS"] / 1000
        else:
            if "-4M" in [leg.currentText() for leg in self.leg_extensions] or "-5M" in [leg.currentText() for leg in
                                                                                        self.leg_extensions]:
                BTB_weight = data["BTB_options"]["BASIC TOWER BODY WHEN WITH -4M & -5M LEG EXTENSIONS"] / 1000
            else:
                BTB_weight = data["BTB_options"][
                                 "BASIC TOWER BODY WHEN WITH BE OR -3M,-2M,-1M,+0M,+1M AND +2M LEG EXTENSIONS"] / 1000

        BB_BE_weight = BTB_weight + BE_weight

        # Retrieve Leg weights
        leg_weights = [data["Leg_values"].get(leg.currentText(), 0) / 1000 for leg in self.leg_extensions]
        ALL_Legs = sum(leg_weights)
        total_weight = BB_BE_weight + ALL_Legs

        tower_type_str = f"{tower_type}+{self.body_extension.currentText()}+({','.join([leg.currentText() for leg in self.leg_extensions])})"

        # Check for identical tower_weight between last two rows
        if self.table.rowCount() > 1:
            previous_weight = self.table.item(self.table.rowCount() - 2, 10).text()
            current_weight = self.table.item(self.table.rowCount() - 1, 10).text()
            if previous_weight == current_weight:
                reply = QMessageBox.question(self, "Identical Tower Weights",
                                             "The tower weights of the last two rows are identical. Do you want to add this row?",
                                             QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.No:
                    return  # Exit if user selects 'No'

        # Add row to table
        self.add_row_to_table(tower_type_str, BTB_weight, BE_weight, leg_weights, BB_BE_weight, ALL_Legs, total_weight)
        self.update_sno()

        # Auto-fit columns (SNO, Tower_Type, and Tower_Weight)
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Auto-fit SNO
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Auto-fit Tower_Type
        self.table.horizontalHeader().setSectionResizeMode(10, QHeaderView.ResizeToContents)  # Auto-fit Tower_Weight
        for i in range(2, 10):
            self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)  # Stretch remaining columns

    def add_row_to_table(self, tower_type_str, BTB_weight, BE_weight, leg_weights, BB_BE_weight, ALL_Legs,
                         total_weight):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        self.table.setItem(row_position, 0, QTableWidgetItem(str(row_position + 1)))
        self.table.setItem(row_position, 1, QTableWidgetItem(tower_type_str))
        self.table.setItem(row_position, 2, QTableWidgetItem(f"{BTB_weight:.4f}"))
        self.table.setItem(row_position, 3, QTableWidgetItem(f"{BE_weight:.4f}"))

        for i in range(4):
            self.table.setItem(row_position, 4 + i, QTableWidgetItem(f"{leg_weights[i]:.4f}"))

        # BB+BE and ALL_Legs in different color
        bb_be_item = QTableWidgetItem(f"{BB_BE_weight:.4f}")
        bb_be_item.setForeground(QColor(0, 0, 255))  # Blue color
        self.table.setItem(row_position, 8, bb_be_item)

        all_legs_item = QTableWidgetItem(f"{ALL_Legs:.4f}")
        all_legs_item.setForeground(QColor(0, 0, 255))  # Blue color
        self.table.setItem(row_position, 9, all_legs_item)

        # Tower Weight in bold red
        tower_weight_item = QTableWidgetItem(f"{total_weight:.4f}")
        tower_weight_item.setFont(QFont("Verdana", 10, QFont.Bold))
        tower_weight_item.setForeground(QColor(255, 0, 0))  # Red color
        self.table.setItem(row_position, 10, tower_weight_item)

        for col in range(11):
            self.table.item(row_position, col).setTextAlignment(Qt.AlignCenter)

        self.table.scrollToBottom()

    def remove_selected_row(self):
        selected_row = self.table.currentRow()
        if selected_row >= 0:
            self.table.removeRow(selected_row)
            self.update_sno()
        else:
            QMessageBox.warning(self, "No Row Selected", "Please select a row to remove.", QMessageBox.Ok)

    def remove_all_rows(self):
        self.table.setRowCount(0)
        self.update_sno()

    def update_sno(self):
        # Update serial numbers for each row
        for row in range(self.table.rowCount()):
            self.table.setItem(row, 0, QTableWidgetItem(str(row + 1)))

    def export_to_excel(self):
        try:
            # Open file dialog and get path for saving the file
            path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx)")

            if path:  # If a valid path is selected
                data = []
                # Collecting table data
                for row in range(self.table.rowCount()):
                    row_data = []
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        row_data.append(item.text() if item else "")
                    data.append(row_data)

                # Convert list of data to a Pandas DataFrame
                df = pd.DataFrame(data, columns=["SNO", "Tower_Type", "BTB", "BE", "Leg A", "Leg B", "Leg C", "Leg D",
                                                 "BB+BE", "ALL_Legs", "Tower_Weight"])

                # Try exporting the DataFrame to Excel
                df.to_excel(path, index=False)
                QMessageBox.information(self, "Success", "File saved successfully!")
            else:
                # No valid path was selected
                QMessageBox.warning(self, "No File Selected", "Please select a valid file path to save the Excel file.")

        except Exception as e:
            # Catch and display any errors that occur during the export process
            QMessageBox.critical(self, "Export Failed", f"An error occurred while saving the file: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    calculator = TowerCalculator()
    calculator.show()
    sys.exit(app.exec_())
