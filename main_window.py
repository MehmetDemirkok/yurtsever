from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QLineEdit, QDateEdit, 
    QTableWidget, QTableWidgetItem, QComboBox,
    QMessageBox, QHeaderView, QFileDialog, QDialog, QTextBrowser
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QDoubleValidator
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import quote_sheetname
from database import Database
from stay_model import StayModel
from helpers import validate_dates, format_currency
import pandas as pd
from datetime import datetime
from unidecode import unidecode

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = Database()
        self.stay_model = StayModel(self.db)
        self.current_sort_column = -1
        self.current_sort_order = Qt.AscendingOrder
        self.init_ui()
        
    def init_ui(self):
        """Initialize the user interface."""
        self.setWindowTitle('Otel Konaklama Puantaj Sistemi')
        self.setMinimumSize(1000, 600)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Create input form
        form_layout = QVBoxLayout()  # Changed to QVBoxLayout for better organization
        
        # First row - Name and Title
        name_title_layout = QHBoxLayout()
        
        # Guest name input
        self.guest_name_input = QLineEdit()
        self.guest_name_input.setPlaceholderText('Adı Soyadı')
        name_title_layout.addWidget(self.guest_name_input)
        
        # Guest title input
        self.guest_title_input = QLineEdit()
        self.guest_title_input.setPlaceholderText('Unvanı')
        name_title_layout.addWidget(self.guest_title_input)
        
        form_layout.addLayout(name_title_layout)
        
        # Second row - Country and City
        location_layout = QHBoxLayout()
        
        # Country input
        self.country_input = QLineEdit()
        self.country_input.setPlaceholderText('Ülke')
        location_layout.addWidget(self.country_input)
        
        # City input
        self.city_input = QLineEdit()
        self.city_input.setPlaceholderText('Şehir')
        location_layout.addWidget(self.city_input)
        
        form_layout.addLayout(location_layout)
        
        # Third row - Dates and Room Type
        dates_room_layout = QHBoxLayout()
        
        # Check-in date input
        self.check_in_date = QDateEdit()
        self.check_in_date.setDate(QDate.currentDate())
        self.check_in_date.setCalendarPopup(True)
        self.check_in_date.setDisplayFormat("dd.MM.yyyy")
        dates_room_layout.addWidget(QLabel("Giriş Tarihi:"))
        dates_room_layout.addWidget(self.check_in_date)
        
        # Check-out date input
        self.check_out_date = QDateEdit()
        self.check_out_date.setDate(QDate.currentDate().addDays(1))
        self.check_out_date.setCalendarPopup(True)
        self.check_out_date.setDisplayFormat("dd.MM.yyyy")
        dates_room_layout.addWidget(QLabel("Çıkış Tarihi:"))
        dates_room_layout.addWidget(self.check_out_date)
        
        # Room type input
        self.room_type_combo = QComboBox()
        self.room_type_combo.addItems(['Single Oda', 'Double Oda', 'Triple Oda', 'Suit Oda', 'Aile Odası'])
        dates_room_layout.addWidget(QLabel("Oda Tipi:"))
        dates_room_layout.addWidget(self.room_type_combo)
        
        form_layout.addLayout(dates_room_layout)
        
        # All action buttons and nightly rate input in one horizontal layout
        action_buttons_and_rate_layout = QHBoxLayout()
        
        # Nightly rate input
        self.nightly_rate_input = QLineEdit()
        self.nightly_rate_input.setPlaceholderText('Gecelik Ücret')
        self.nightly_rate_input.setValidator(QDoubleValidator(0.00, 999999.99, 2))
        action_buttons_and_rate_layout.addWidget(self.nightly_rate_input)
        
        # Add button
        self.add_button = QPushButton('Ekle')
        self.add_button.setObjectName('add_button')
        self.add_button.clicked.connect(self.add_stay)
        action_buttons_and_rate_layout.addWidget(self.add_button)
        
        # Edit selected button
        self.edit_selected_button = QPushButton('Seçili Kaydı Düzenle')
        self.edit_selected_button.clicked.connect(self.edit_selected_stay)
        action_buttons_and_rate_layout.addWidget(self.edit_selected_button)

        # Delete selected button
        self.delete_selected_button = QPushButton('Seçili Kaydı Sil')
        self.delete_selected_button.clicked.connect(self.delete_selected_stay)
        action_buttons_and_rate_layout.addWidget(self.delete_selected_button)

        form_layout.addLayout(action_buttons_and_rate_layout)
        
        # Connect returnPressed signal to add_button's clicked signal
        self.guest_name_input.returnPressed.connect(self.add_button.click)
        self.guest_title_input.returnPressed.connect(self.add_button.click)
        self.country_input.returnPressed.connect(self.add_button.click)
        self.city_input.returnPressed.connect(self.add_button.click)
        self.nightly_rate_input.returnPressed.connect(self.add_button.click)
        
        main_layout.addLayout(form_layout)
        
        # Create table
        self.table = QTableWidget()
        self.table.setColumnCount(10)  # Changed column count from 9 to 10
        self.table.setHorizontalHeaderLabels([
            'ID', 'Adı Soyadı', 'Unvan', 'Ülke', 'Şehir',
            'Giriş Tarihi', 'Çıkış Tarihi', 'Oda Tipi', 'Gecelik Ücret', 'Toplam Ücret'
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.horizontalHeader().sectionClicked.connect(self.sort_data)
        main_layout.addWidget(self.table)
        
        # Create filter controls
        filter_layout = QHBoxLayout()
        
        self.filter_guest_name_input = QLineEdit()
        self.filter_guest_name_input.setPlaceholderText('Misafir Adı Soyadı Filtrele')
        self.filter_guest_name_input.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.filter_guest_name_input)
        
        self.filter_country_input = QLineEdit()
        self.filter_country_input.setPlaceholderText('Ülke Filtrele')
        self.filter_country_input.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.filter_country_input)
        
        self.filter_city_input = QLineEdit()
        self.filter_city_input.setPlaceholderText('Şehir Filtrele')
        self.filter_city_input.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.filter_city_input)

        self.filter_room_type_combo = QComboBox()
        self.filter_room_type_combo.addItems(['Tümü', 'Single Oda', 'Double Oda', 'Triple Oda', 'Suit Oda', 'Aile Odası'])
        self.filter_room_type_combo.currentIndexChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.filter_room_type_combo)

        main_layout.addLayout(filter_layout)
        
        # Create report and import controls
        control_layout = QHBoxLayout()
        
        # Statistics Report button
        stats_button = QPushButton('Puantaj Raporu')
        stats_button.clicked.connect(self.generate_puantaj_report)
        control_layout.addWidget(stats_button)
        
        # Import Excel button
        import_button = QPushButton('Excel\'den İçe Aktar')
        import_button.clicked.connect(self.import_excel)
        control_layout.addWidget(import_button)
        
        # Export Excel button
        export_button = QPushButton('Excel\'e Aktar')
        export_button.clicked.connect(self.export_excel)
        control_layout.addWidget(export_button)
        
        # Download Template button
        template_button = QPushButton('Excel Şablonu İndir')
        template_button.clicked.connect(self.download_excel_template)
        control_layout.addWidget(template_button)
        
        main_layout.addLayout(control_layout)
        
        # Load initial data
        self.load_stays()
        
    def add_stay(self):
        """Add a new stay record."""
        guest_name = self.guest_name_input.text().strip()
        guest_title = self.guest_title_input.text().strip()
        country = self.country_input.text().strip()
        city = self.city_input.text().strip()
        check_in = self.check_in_date.date().toString('yyyy-MM-dd')
        check_out = self.check_out_date.date().toString('yyyy-MM-dd')
        room_type = self.room_type_combo.currentText()
        
        print(f"DEBUG: check_in: {check_in}, check_out: {check_out}")

        try:
            nightly_rate = float(self.nightly_rate_input.text().strip())
        except ValueError:
            QMessageBox.warning(self, 'Hata', 'Geçerli bir gecelik ücret giriniz.')
            return
        
        if not all([guest_name, guest_title, country, city]):
            QMessageBox.warning(self, 'Hata', 'Tüm alanları doldurunuz.')
            return
        
        if not validate_dates(check_in, check_out):
            QMessageBox.warning(self, 'Hata', 'Geçerli giriş ve çıkış tarihleri giriniz.')
            return
        
        try:
            self.stay_model.create_stay(
                guest_name=guest_name,
                guest_title=guest_title,
                country=country,
                city=city,
                check_in_date=check_in,
                check_out_date=check_out,
                room_type=room_type,
                nightly_rate=nightly_rate
            )
            self.load_stays()
            self.clear_inputs()
        except Exception as e:
            print(f"DEBUG: Error adding stay: {e}")
            QMessageBox.critical(self, 'Hata', f'Kayıt eklenirken hata oluştu: {str(e)}')
    
    def load_stays(self, guest_name: str = None, country: str = None, city: str = None, room_type: str = None, sort_column_index: int = None, sort_order: Qt.SortOrder = Qt.AscendingOrder):
        """Load all stays into the table with optional filtering and sorting."""
        try:
            # Get header labels for sorting
            if sort_column_index is None:
                sort_column_index = self.current_sort_column
            
            column_name = None
            if sort_column_index != -1:
                column_name = self.table.horizontalHeaderItem(sort_column_index).text()
            
            sort_order_str = "ASC" if sort_order == Qt.AscendingOrder else "DESC"

            stays = self.stay_model.get_all_stays(guest_name, country, city, room_type, column_name, sort_order_str)
            self.table.setRowCount(len(stays))
            
            for row, stay in enumerate(stays):
                self.table.setItem(row, 0, QTableWidgetItem(str(stay['id'])))
                self.table.setItem(row, 1, QTableWidgetItem(stay['guest_name']))
                self.table.setItem(row, 2, QTableWidgetItem(stay['guest_title']))
                self.table.setItem(row, 3, QTableWidgetItem(stay['country']))
                self.table.setItem(row, 4, QTableWidgetItem(stay['city']))
                self.table.setItem(row, 5, QTableWidgetItem(stay['check_in_date']))
                self.table.setItem(row, 6, QTableWidgetItem(stay['check_out_date']))
                self.table.setItem(row, 7, QTableWidgetItem(stay['room_type']))
                self.table.setItem(row, 8, QTableWidgetItem(format_currency(stay['nightly_rate'])))
                self.table.setItem(row, 9, QTableWidgetItem(format_currency(stay['total_amount'])))
                
            self.table.sortItems(sort_column_index if sort_column_index != -1 else 0, sort_order) # Apply visual sort

        except Exception as e:
            QMessageBox.critical(self, 'Hata', f'Veriler yüklenirken hata oluştu: {str(e)}')
    
    def clear_inputs(self):
        """Clear all input fields."""
        self.guest_name_input.clear()
        self.guest_title_input.clear()
        self.country_input.clear()
        self.city_input.clear()
        self.nightly_rate_input.clear()
        self.check_in_date.setDate(QDate.currentDate())
        self.check_out_date.setDate(QDate.currentDate().addDays(1))
        self.room_type_combo.setCurrentIndex(0)
    
    def edit_stay(self, row):
        """Populate input fields with data from selected stay for editing."""
        try:
            stay_id = int(self.table.item(row, 0).text())
            stay = self.stay_model.get_stay_by_id(stay_id)
            
            if stay:
                # Store the original clicked handler of add_button
                self._original_clicked_handler = self.add_button.clicked
                
                # Disconnect the old handler
                try:
                    self.add_button.clicked.disconnect(self.add_stay)
                except TypeError: # If not connected to add_stay, it's connected to save_edit. Disconnect that.
                    try:
                        self.add_button.clicked.disconnect(lambda: self.save_edit(stay_id))
                    except TypeError: # If not connected to save_edit, it's not connected at all
                        pass
                
                self.add_button.setText('Kaydet')
                # Connect the add button to save_edit with the current stay_id
                self.add_button.clicked.connect(lambda: self.save_edit(stay_id))

                self.guest_name_input.setText(stay['guest_name'])
                self.guest_title_input.setText(stay['guest_title'])
                self.country_input.setText(stay['country'])
                self.city_input.setText(stay['city'])
                self.check_in_date.setDate(QDate.fromString(stay['check_in_date'], 'yyyy-MM-dd'))
                self.check_out_date.setDate(QDate.fromString(stay['check_out_date'], 'yyyy-MM-dd'))
                self.room_type_combo.setCurrentText(stay['room_type'])
                self.nightly_rate_input.setText(str(stay['nightly_rate']))
                
        except Exception as e:
            QMessageBox.critical(self, 'Hata', f'Kayıt düzenlenirken hata oluştu: {str(e)}')
    
    def edit_selected_stay(self):
        """Edit the currently selected stay record."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Hata', 'Lütfen düzenlemek için bir kayıt seçin.')
            return
        row = selected_rows[0].row()
        self.edit_stay(row)

    def delete_stay(self, row):
        """Delete the stay record at the specified row."""
        stay_id = int(self.table.item(row, 0).text())
        reply = QMessageBox.question(self, 'Onay', 'Bu kaydı silmek istediğinizden emin misiniz?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                if self.stay_model.delete_stay(stay_id):
                    self.load_stays()
                else:
                    QMessageBox.warning(self, 'Hata', 'Kayıt silinemedi.')
            except Exception as e:
                QMessageBox.critical(self, 'Hata', f'Kayıt silinirken hata oluştu: {str(e)}')

    def delete_selected_stay(self):
        """Delete the currently selected stay record."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Hata', 'Lütfen silmek için bir kayıt seçin.')
            return
        row = selected_rows[0].row()
        self.delete_stay(row)
    
    def save_edit(self, stay_id):
        """Save changes to an edited stay record."""
        guest_name = self.guest_name_input.text().strip()
        guest_title = self.guest_title_input.text().strip()
        country = self.country_input.text().strip()
        city = self.city_input.text().strip()
        check_in_date = self.check_in_date.date().toString('yyyy-MM-dd')
        check_out_date = self.check_out_date.date().toString('yyyy-MM-dd')
        room_type = self.room_type_combo.currentText()
        
        try:
            nightly_rate = float(self.nightly_rate_input.text().strip())
        except ValueError:
            QMessageBox.warning(self, 'Hata', 'Geçerli bir gecelik ücret giriniz.')
            return
        
        if not all([guest_name, guest_title, country, city]):
            QMessageBox.warning(self, 'Hata', 'Tüm alanları doldurunuz.')
            return
        
        if not validate_dates(check_in_date, check_out_date):
            QMessageBox.warning(self, 'Hata', 'Geçerli giriş ve çıkış tarihleri giriniz.')
            return

        try:
            self.stay_model.update_stay(
                stay_id=stay_id,
                guest_name=guest_name,
                guest_title=guest_title,
                country=country,
                city=city,
                check_in_date=check_in_date,
                check_out_date=check_out_date,
                room_type=room_type,
                nightly_rate=nightly_rate
            )
            self.load_stays()
            self.clear_inputs()
            QMessageBox.information(self, 'Başarılı', 'Kayıt başarıyla güncellendi.')
            
        except Exception as e:
            QMessageBox.critical(self, 'Hata', f'Kayıt güncellenirken hata oluştu: {str(e)}')
        finally:
            # Restore the add_button to its original state
            self.add_button.setText('Ekle')
            try:
                self.add_button.clicked.disconnect()
            except TypeError:
                pass
            self.add_button.clicked.connect(self._original_clicked_handler)
            self._original_clicked_handler = None
    
    def import_excel(self):
        """Import data from Excel file."""
        try:
            # Open file dialog
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Excel Dosyası Seç",
                "",
                "Excel Dosyaları (*.xlsx *.xls)"
            )
            
            if not file_path:
                return
            
            # Read Excel file
            df = pd.read_excel(file_path)
            
            # Normalize column names by stripping whitespace and converting to a consistent format
            def normalize_column_name(col_name):
                return unidecode(col_name.strip().lower().replace(" ", "_"))

            df.columns = [normalize_column_name(col) for col in df.columns]

            # Validate required columns (normalized)
            required_columns_map = {
                normalize_column_name('Adı Soyadı'): 'Adı Soyadı',
                normalize_column_name('Unvan'): 'Unvan',
                normalize_column_name('Ülke'): 'Ülke',
                normalize_column_name('Şehir'): 'Şehir',
                normalize_column_name('Giriş Tarihi'): 'Giriş Tarihi',
                normalize_column_name('Çıkış Tarihi'): 'Çıkış Tarihi',
                normalize_column_name('Oda Tipi'): 'Oda Tipi',
                normalize_column_name('Gecelik Ücret'): 'Gecelik Ücret'
            }
            
            missing_normalized_columns = [col for col in required_columns_map.keys() if col not in df.columns]
            if missing_normalized_columns:
                # Map back to original names for the error message
                missing_original_names = [required_columns_map[col] for col in missing_normalized_columns]
                QMessageBox.warning(
                    self,
                    'Hata',
                    f'Excel dosyasında eksik sütunlar var: {", ".join(missing_original_names)}'
                )
                return
            
            # Process each row
            success_count = 0
            error_count = 0
            skipped_rows_details = []
            
            for index, row in df.iterrows():
                row_num = index + 2  # Excel row numbers start from 1, and header is row 1, so data starts from row 2
                try:
                    # Convert dates to string format and handle potential errors
                    try:
                        check_in = pd.to_datetime(row[normalize_column_name('Giriş Tarihi')]).strftime('%Y-%m-%d')
                    except ValueError:
                        skipped_rows_details.append(f"Satır {row_num}: Geçersiz Giriş Tarihi formatı ('{row[normalize_column_name('Giriş Tarihi')]}').")
                        error_count += 1
                        continue
                    
                    try:
                        check_out = pd.to_datetime(row[normalize_column_name('Çıkış Tarihi')]).strftime('%Y-%m-%d')
                    except ValueError:
                        skipped_rows_details.append(f"Satır {row_num}: Geçersiz Çıkış Tarihi formatı ('{row[normalize_column_name('Çıkış Tarihi')]}').")
                        error_count += 1
                        continue

                    # Validate dates for logical consistency (check_out after check_in)
                    if not validate_dates(check_in, check_out):
                        skipped_rows_details.append(f"Satır {row_num}: Çıkış Tarihi, Giriş Tarihinden önce olamaz.")
                        error_count += 1
                        continue

                    # Convert nightly rate and handle potential errors
                    try:
                        nightly_rate = float(str(row[normalize_column_name('Gecelik Ücret')]).replace(',', '.')) # Handle comma as decimal separator
                    except ValueError:
                        skipped_rows_details.append(f"Satır {row_num}: Geçersiz Gecelik Ücret formatı ('{row[normalize_column_name('Gecelik Ücret')]}').")
                        error_count += 1
                        continue
                    
                    # Ensure required text fields are not empty after stripping
                    guest_name = str(row[normalize_column_name('Adı Soyadı')]).strip()
                    guest_title = str(row[normalize_column_name('Unvan')]).strip()
                    country = str(row[normalize_column_name('Ülke')]).strip()
                    city = str(row[normalize_column_name('Şehir')]).strip()
                    room_type = str(row[normalize_column_name('Oda Tipi')]).strip()

                    if not all([guest_name, guest_title, country, city, room_type]):
                        missing_fields = []
                        if not guest_name: missing_fields.append('Adı Soyadı')
                        if not guest_title: missing_fields.append('Unvan')
                        if not country: missing_fields.append('Ülke')
                        if not city: missing_fields.append('Şehir')
                        if not room_type: missing_fields.append('Oda Tipi')
                        skipped_rows_details.append(f"Satır {row_num}: Eksik veya boş alanlar: {', '.join(missing_fields)}.")
                        error_count += 1
                        continue

                    # Create stay record
                    self.stay_model.create_stay(
                        guest_name=guest_name,
                        guest_title=guest_title,
                        country=country,
                        city=city,
                        check_in_date=check_in,
                        check_out_date=check_out,
                        room_type=room_type,
                        nightly_rate=nightly_rate
                    )
                    success_count += 1
                except Exception as e:
                    skipped_rows_details.append(f"Satır {row_num}: Beklenmeyen hata: {e}")
                    error_count += 1
            
            # Reload table
            self.load_stays()
            
            # Show result message
            result_message = f'Başarıyla içe aktarılan: {success_count}\nToplam Hatalı Kayıt: {error_count}'
            if skipped_rows_details:
                result_message += "\n\nAtlanan satırlar ve nedenleri:\n" + "\n".join(skipped_rows_details[:10]) # Show up to 10 details
                if len(skipped_rows_details) > 10:
                    result_message += f"\n... ve {len(skipped_rows_details) - 10} adet daha."
            
            QMessageBox.information(
                self,
                'İçe Aktarma Tamamlandı',
                result_message
            )
            
        except Exception as e:
            QMessageBox.critical(
                self,
                'Hata',
                f'Excel dosyası içe aktarılırken genel bir hata oluştu: {str(e)}'
            )
    
    def download_excel_template(self):
        """Create and download Excel template file."""
        try:
            # Create a DataFrame with example data
            example_data = {
                'Adı Soyadı': ['Ahmet Yılmaz', 'Ayşe Demir'],
                'Unvan': ['Bay', 'Bayan'],
                'Ülke': ['Türkiye', 'Türkiye'],
                'Şehir': ['İstanbul', 'Ankara'],
                'Giriş Tarihi': ['2024-03-20', '2024-03-21'],
                'Çıkış Tarihi': ['2024-03-25', '2024-03-23'],
                'Oda Tipi': ['Single Oda', 'Double Oda'],
                'Gecelik Ücret': [500, 750]
            }
            df = pd.DataFrame(example_data)
            
            # Get save file path
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Excel Şablonunu Kaydet",
                "misafir_kayit_sablonu.xlsx",
                "Excel Dosyaları (*.xlsx)"
            )
            
            if not file_path:
                return
            
            # Add .xlsx extension if not present
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            
            # Create Excel writer
            writer = pd.ExcelWriter(file_path, engine='openpyxl')
            
            # Get workbook
            workbook = writer.book
            
            # Create room types sheet first
            room_types = ['Single Oda', 'Double Oda', 'Triple Oda', 'Suit Oda', 'Aile Odası']
            room_types_sheet = workbook.create_sheet(title='Oda Tipleri')
            
            # Add room types to the sheet
            for i, room_type in enumerate(room_types, 1):
                room_types_sheet.cell(row=i, column=1, value=room_type)
            
            # Create named range for room types
            room_types_range = f"'{room_types_sheet.title}'!$A$1:$A${len(room_types)}"
            workbook.defined_names.add(DefinedName('room_types', attr_text=room_types_range))
            
            # Write data to main sheet
            df.to_excel(writer, index=False, sheet_name='Misafir Kayıtları')
            
            # Get the main worksheet
            worksheet = writer.sheets['Misafir Kayıtları']
            
            # Add data validation for room types
            dv = DataValidation(type="list", formula1=f"=room_types", allow_blank=True)
            worksheet.add_data_validation(dv)
            dv.add('G2:G1000')  # Apply to column G (Oda Tipi) for a generous range
            
            # Set column widths for better readability
            worksheet.column_dimensions['A'].width = 15
            worksheet.column_dimensions['B'].width = 15
            worksheet.column_dimensions['C'].width = 15
            worksheet.column_dimensions['D'].width = 15
            worksheet.column_dimensions['E'].width = 18
            worksheet.column_dimensions['F'].width = 18
            worksheet.column_dimensions['G'].width = 15
            worksheet.column_dimensions['H'].width = 15

            # Add instructions
            instructions = [
                'Excel Şablonu Kullanım Talimatları:',
                '',
                '1. Tüm sütunları doldurunuz:',
                '   - Adı Soyadı: Misafirin tam adı',
                '   - Unvan: Bay/Bayan',
                '   - Ülke: Misafirin ülkesi',
                '   - Şehir: Misafirin şehri',
                '   - Giriş Tarihi: YYYY-MM-DD formatında veya Excel tarih formatında',
                '   - Çıkış Tarihi: YYYY-MM-DD formatında veya Excel tarih formatında',
                '   - Oda Tipi: Dropdown menüden seçiniz',
                '   - Gecelik Ücret: Sayısal değer (örnek: 150.00)',
                '',
                '2. Örnek kayıtları silip kendi kayıtlarınızı ekleyebilirsiniz.',
                '3. Tarihleri doğru formatta girdiğinizden emin olun.',
                '4. Oda tipini sağdaki dropdown menüden seçiniz.',
                '5. Gecelik ücreti geçerli bir sayı olarak giriniz.',
                '',
                'Not: Bu şablonu doldurduktan sonra "Excel\'den İçe Aktar" butonu ile verileri sisteme aktarabilirsiniz.'
            ]
            
            # Create instructions sheet
            instructions_sheet = workbook.create_sheet(title='Kullanım Talimatları')
            for i, instruction in enumerate(instructions, 1):
                instructions_sheet.cell(row=i, column=1, value=instruction)

            # Hide the room types sheet
            room_types_sheet.sheet_state = 'hidden'
            
            # Set active sheet to Misafir Kayıtları
            workbook.active = workbook['Misafir Kayıtları']
            
            # Save the workbook
            writer.close()
            
            QMessageBox.information(
                self,
                'Başarılı',
                f'Excel şablonu başarıyla oluşturuldu:\n{file_path}'
            )
            
        except Exception as e:
            QMessageBox.critical(
                self,
                'Hata',
                f'Excel şablonu oluşturulurken hata oluştu: {str(e)}'
            )
    
    def closeEvent(self, event):
        """Handle window close event."""
        self.db.close()
        event.accept()

    def apply_filters(self):
        """Apply filters to the table by reloading stays with current filter values."""
        guest_name = self.filter_guest_name_input.text().strip()
        country = self.filter_country_input.text().strip()
        city = self.filter_city_input.text().strip()
        room_type = self.filter_room_type_combo.currentText()
        
        self.load_stays(guest_name, country, city, room_type, self.current_sort_column, self.current_sort_order)

    def sort_data(self, column_index: int):
        """Sort the table data based on the clicked column.
        
        Args:
            column_index (int): The index of the column that was clicked.
        """
        if column_index == self.current_sort_column:
            self.current_sort_order = Qt.DescendingOrder if self.current_sort_order == Qt.AscendingOrder else Qt.AscendingOrder
        else:
            self.current_sort_column = column_index
            self.current_sort_order = Qt.AscendingOrder
        
        self.apply_filters() # Reapply filters and sorting 

    def export_excel(self):
        """Export data from the table to an Excel file."""
        try:
            # Get table data
            column_headers = []
            for col in range(self.table.columnCount()):
                column_headers.append(self.table.horizontalHeaderItem(col).text())
            
            data = []
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    if item is not None:
                        row_data.append(item.text())
                    else:
                        row_data.append("") # Handle empty cells
                data.append(row_data)
            
            df = pd.DataFrame(data, columns=column_headers)
            
            # Get save file path
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Excel Dosyası Kaydet",
                "konaklama_kayitlari.xlsx",
                "Excel Dosyaları (*.xlsx)"
            )
            
            if not file_path:
                return
            
            # Add .xlsx extension if not present
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            
            # Write data to Excel
            writer = pd.ExcelWriter(file_path, engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name='Konaklama Kayıtları')
            writer.close()
            
            QMessageBox.information(
                self,
                'Başarılı',
                f"""Veriler başarıyla Excel'e aktarıldı:
{file_path}"""
            )
            
        except Exception as e:
            QMessageBox.critical(
                self,
                'Hata',
                f"""Excel'e aktarılırken hata oluştu: {str(e)}"""
            )

    def generate_puantaj_report(self):
        """Generate detailed stay report in Excel format."""
        try:
            # Get report data
            report_data = self.stay_model.get_detailed_stay_report()
            
            # Get save file path
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Puantaj Raporunu Kaydet",
                "konaklama_puantaji.xlsx",
                "Excel Dosyaları (*.xlsx)"
            )
            
            if not file_path:
                return
            
            # Add .xlsx extension if not present
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            
            # Create Excel writer
            writer = pd.ExcelWriter(file_path, engine='openpyxl')
            
            # Create main report sheet
            df = pd.DataFrame(report_data)
            
            # Rename columns for better readability
            column_names = {
                'guest_name': 'Misafir Adı',
                'guest_title': 'Unvan',
                'country': 'Ülke',
                'city': 'Şehir',
                'check_in_date': 'Giriş Tarihi',
                'check_out_date': 'Çıkış Tarihi',
                'room_type': 'Oda Tipi',
                'nightly_rate': 'Gecelik Ücret',
                'total_amount': 'Toplam Ücret',
                'stay_type': 'Konaklama Tipi',
                'nights': 'Gece Sayısı'
            }
            df = df.rename(columns=column_names)
            
            # Write main report
            df.to_excel(writer, sheet_name='Konaklama Puantajı', index=False)
            
            # Get workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Konaklama Puantajı']
            
            # Create summary sheet
            summary_sheet = workbook.create_sheet(title='Özet Rapor')
            
            # Calculate summary statistics
            total_stays = len(df)
            total_revenue = df['Toplam Ücret'].sum()
            avg_nights = df['Gece Sayısı'].mean()
            
            # Room type statistics
            room_stats = df.groupby('Oda Tipi').agg({
                'Misafir Adı': 'count',
                'Toplam Ücret': 'sum',
                'Gece Sayısı': 'sum'
            }).reset_index()
            
            # Stay type statistics
            stay_type_stats = df.groupby('Konaklama Tipi').agg({
                'Misafir Adı': 'count',
                'Toplam Ücret': 'sum',
                'Gece Sayısı': 'sum'
            }).reset_index()
            
            # Write summary statistics
            summary_sheet['A1'] = 'KONAKLAMA PUANTAJI ÖZET RAPORU'
            summary_sheet['A2'] = f'Rapor Tarihi: {datetime.now().strftime("%d.%m.%Y")}'
            summary_sheet['A4'] = 'GENEL İSTATİSTİKLER'
            summary_sheet['A5'] = 'Toplam Konaklama Sayısı:'
            summary_sheet['B5'] = total_stays
            summary_sheet['A6'] = 'Toplam Gelir:'
            summary_sheet['B6'] = total_revenue
            summary_sheet['A7'] = 'Ortalama Konaklama Süresi:'
            summary_sheet['B7'] = f'{avg_nights:.1f} gece'
            
            # Write room type statistics
            summary_sheet['A9'] = 'ODA TİPİ BAZINDA İSTATİSTİKLER'
            summary_sheet['A10'] = 'Oda Tipi'
            summary_sheet['B10'] = 'Konaklama Sayısı'
            summary_sheet['C10'] = 'Toplam Gelir'
            summary_sheet['D10'] = 'Toplam Gece'
            
            for i, row in enumerate(room_stats.itertuples(), 11):
                summary_sheet[f'A{i}'] = row._1
                summary_sheet[f'B{i}'] = row._2
                summary_sheet[f'C{i}'] = row._3
                summary_sheet[f'D{i}'] = row._4
            
            # Write stay type statistics
            start_row = len(room_stats) + 13
            summary_sheet[f'A{start_row}'] = 'KONAKLAMA TİPİ BAZINDA İSTATİSTİKLER'
            summary_sheet[f'A{start_row+1}'] = 'Konaklama Tipi'
            summary_sheet[f'B{start_row+1}'] = 'Konaklama Sayısı'
            summary_sheet[f'C{start_row+1}'] = 'Toplam Gelir'
            summary_sheet[f'D{start_row+1}'] = 'Toplam Gece'
            
            for i, row in enumerate(stay_type_stats.itertuples(), start_row+2):
                summary_sheet[f'A{i}'] = row._1
                summary_sheet[f'B{i}'] = row._2
                summary_sheet[f'C{i}'] = row._3
                summary_sheet[f'D{i}'] = row._4
            
            # Format main report sheet
            for column in worksheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
            
            # Format summary sheet
            for column in summary_sheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                summary_sheet.column_dimensions[column[0].column_letter].width = adjusted_width
            
            # Save the workbook
            writer.close()
            
            QMessageBox.information(
                self,
                'Başarılı',
                f'Puantaj raporu başarıyla oluşturuldu:\n{file_path}'
            )
            
        except Exception as e:
            QMessageBox.critical(
                self,
                'Hata',
                f'Puantaj raporu oluşturulurken hata oluştu: {str(e)}'
            ) 