from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QLineEdit, QDateEdit, 
    QTableWidget, QTableWidgetItem, QComboBox,
    QMessageBox, QHeaderView, QFileDialog
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QDoubleValidator
from database import Database
from stay_model import StayModel
from report_utils import StayReport
from helpers import validate_dates, format_currency
import pandas as pd
from datetime import datetime

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = Database()
        self.stay_model = StayModel(self.db)
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
        main_layout.addWidget(self.table)
        
        # Create report and import controls
        control_layout = QHBoxLayout()
        
        # Period selection
        self.period_combo = QComboBox()
        self.period_combo.addItems(['Haftalık', 'Aylık', 'Yıllık'])
        control_layout.addWidget(self.period_combo)
        
        # Report type selection
        self.report_type_combo = QComboBox()
        self.report_type_combo.addItems(['Konaklama Raporu'])
        control_layout.addWidget(self.report_type_combo)
        
        # Generate report button
        report_button = QPushButton('Rapor Oluştur')
        report_button.clicked.connect(self.generate_report)
        control_layout.addWidget(report_button)
        
        # Import Excel button
        import_button = QPushButton('Excel\'den İçe Aktar')
        import_button.clicked.connect(self.import_excel)
        control_layout.addWidget(import_button)
        
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
    
    def load_stays(self):
        """Load all stays into the table."""
        try:
            stays = self.stay_model.get_all_stays()
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
                
                # Removed Edit and Delete buttons from table cells
        except Exception as e:
            QMessageBox.critical(self, 'Hata', f'Veriler yüklenirken hata oluştu: {str(e)}')
    
    def generate_report(self):
        """Generate and display report based on selected period and type."""
        try:
            period_map = {'Haftalık': 'week', 'Aylık': 'month', 'Yıllık': 'year'}
            period = period_map[self.period_combo.currentText()]
            
            stays = self.stay_model.get_all_stays()
            report = StayReport(stays)
            
            # Simplified report generation as only 'Konaklama Raporu' remains
            report_df = report.get_period_report(period)
            
            if report_df.empty:
                QMessageBox.information(self, 'Bilgi', 'Seçilen dönem için veri bulunamadı.')
                return
            
            # Update table with report data
            self.table.setColumnCount(len(report_df.columns))
            self.table.setHorizontalHeaderLabels(report_df.columns)
            self.table.setRowCount(len(report_df))
            
            for row in range(len(report_df)):
                for col in range(len(report_df.columns)):
                    value = str(report_df.iloc[row, col])
                    self.table.setItem(row, col, QTableWidgetItem(value))
                    
        except Exception as e:
            QMessageBox.critical(self, 'Hata', f'Rapor oluşturulurken hata oluştu: {str(e)}')
    
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
            
            # Validate required columns
            required_columns = [
                'Adı Soyadı', 'Unvan', 'Ülke', 'Şehir',
                'Giriş Tarihi', 'Çıkış Tarihi', 'Oda Tipi', 'Gecelik Ücret'
            ]
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                QMessageBox.warning(
                    self,
                    'Hata',
                    f'Excel dosyasında eksik sütunlar var: {", ".join(missing_columns)}'
                )
                return
            
            # Process each row
            success_count = 0
            error_count = 0
            
            for _, row in df.iterrows():
                try:
                    # Convert dates to string format
                    check_in = pd.to_datetime(row['Giriş Tarihi']).strftime('%Y-%m-%d')
                    check_out = pd.to_datetime(row['Çıkış Tarihi']).strftime('%Y-%m-%d')
                    
                    # Create stay record
                    self.stay_model.create_stay(
                        guest_name=str(row['Adı Soyadı']),
                        guest_title=str(row['Unvan']),
                        country=str(row['Ülke']),
                        city=str(row['Şehir']),
                        check_in_date=check_in,
                        check_out_date=check_out,
                        room_type=str(row['Oda Tipi']),
                        nightly_rate=float(row['Gecelik Ücret'])
                    )
                    success_count += 1
                except Exception as e:
                    print(f"Error importing row: {e}")
                    error_count += 1
            
            # Reload table
            self.load_stays()
            
            # Show result message
            QMessageBox.information(
                self,
                'İçe Aktarma Tamamlandı',
                f'Başarıyla içe aktarılan: {success_count}\nHatalı kayıt: {error_count}'
            )
            
        except Exception as e:
            QMessageBox.critical(
                self,
                'Hata',
                f'Excel dosyası içe aktarılırken hata oluştu: {str(e)}'
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
                'Oda Tipi': ['Standart Oda', 'Deluxe Oda'],
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
            
            # Write data to Excel
            df.to_excel(writer, index=False, sheet_name='Misafir Kayıtları')
            
            # Get workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Misafir Kayıtları']
            
            # Add data validation for room types
            room_types = ['Standart Oda', 'Deluxe Oda', 'Suit Oda', 'Aile Odası']
            for row in range(3, 100):  # Apply to rows 3-100
                cell = worksheet.cell(row=row, column=7)  # Column G (Oda Tipi)
                cell.value = f'=INDIRECT("room_types")'
            
            # Add named range for room types
            room_types_sheet = workbook.create_sheet('Oda Tipleri')
            for i, room_type in enumerate(room_types, 1):
                room_types_sheet.cell(row=i, column=1, value=room_type)
            
            # Add data validation
            from openpyxl.worksheet.datavalidation import DataValidation
            dv = DataValidation(type="list", formula1="=Oda Tipleri!$A$1:$A$4", allow_blank=True)
            worksheet.add_data_validation(dv)
            dv.add(f'G3:G100')  # Apply to column G (Oda Tipi)
            
            # Add instructions
            instructions = [
                'Excel Şablonu Kullanım Talimatları:',
                '',
                '1. Tüm sütunları doldurunuz:',
                '   - Adı Soyadı: Misafirin tam adı',
                '   - Unvan: Bay/Bayan',
                '   - Ülke: Misafirin ülkesi',
                '   - Şehir: Misafirin şehri',
                '   - Giriş Tarihi: YYYY-MM-DD formatında',
                '   - Çıkış Tarihi: YYYY-MM-DD formatında',
                '   - Oda Tipi: Dropdown menüden seçiniz',
                '   - Gecelik Ücret: Sayısal değer',
                '',
                '2. Örnek kayıtları silip kendi kayıtlarınızı ekleyebilirsiniz.',
                '3. Tarihleri Excel tarih formatında girebilirsiniz.',
                '4. Oda tipini dropdown menüden seçiniz.',
                '5. Gecelik ücreti sayısal değer olarak giriniz.',
                '',
                'Not: Bu şablonu doldurduktan sonra "Excel\'den İçe Aktar" butonu ile verileri sisteme aktarabilirsiniz.'
            ]
            
            # Create instructions sheet
            instructions_sheet = workbook.create_sheet('Kullanım Talimatları')
            for i, instruction in enumerate(instructions, 1):
                instructions_sheet.cell(row=i, column=1, value=instruction)
            
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