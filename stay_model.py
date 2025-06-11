from datetime import datetime
from typing import List, Optional, Dict, Any
from database import Database
from PyQt5.QtCore import pyqtSignal, QObject

class StayModel(QObject):
    report_generated = pyqtSignal()

    def __init__(self, db: Database):
        """Initialize stay model with database connection.
        
        Args:
            db (Database): Database instance
        """
        super().__init__()
        self.db = db
    
    def create_stay(self, guest_name: str, guest_title: str, country: str, city: str,
                   check_in_date: str, check_out_date: str, room_type: str,
                   nightly_rate: float) -> int:
        """Create a new stay record.
        
        Args:
            guest_name (str): Name of the guest
            guest_title (str): Title of the guest
            country (str): Country of the guest
            city (str): City of the guest
            check_in_date (str): Check-in date (YYYY-MM-DD)
            check_out_date (str): Check-out date (YYYY-MM-DD)
            room_type (str): Type of the room
            nightly_rate (float): Nightly rate for the stay
            
        Returns:
            int: ID of the created stay record
        """
        try:
            check_in = datetime.strptime(check_in_date, "%Y-%m-%d")
            check_out = datetime.strptime(check_out_date, "%Y-%m-%d")
            nights = (check_out - check_in).days
            total_amount = nights * nightly_rate
            
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            query = '''
                INSERT INTO stays (
                    guest_name, guest_title, country, city,
                    check_in_date, check_out_date, room_type,
                    nightly_rate, total_amount, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            '''
            
            self.db.cursor.execute(query, (
                guest_name, guest_title, country, city,
                check_in_date, check_out_date, room_type,
                nightly_rate, total_amount, now, now
            ))
            self.db.conn.commit()
            return self.db.cursor.lastrowid
        except Exception as e:
            print(f"Error creating stay: {e}")
            raise
    
    def get_all_stays(self, guest_name: str = None, country: str = None, city: str = None, room_type: str = None, sort_column: str = None, sort_order: str = "DESC") -> List[Dict[str, Any]]:
        """Get all stay records, with optional filtering and sorting.
        
        Args:
            guest_name (str): Optional filter for guest name.
            country (str): Optional filter for country.
            city (str): Optional filter for city.
            room_type (str): Optional filter for room type.
            sort_column (str): Column to sort by (e.g., 'guest_name', 'check_in_date').
            sort_order (str): Sort order ('ASC' or 'DESC'). Defaults to 'DESC'.

        Returns:
            List[Dict[str, Any]]: List of stay records
        """
        try:
            query = 'SELECT id, guest_name, guest_title, country, city, check_in_date, check_out_date, room_type, nightly_rate, total_amount, created_at, updated_at FROM stays'
            conditions = []
            params = []

            if guest_name:
                conditions.append("guest_name LIKE ?")
                params.append(f'%{guest_name}%')
            if country:
                conditions.append("country LIKE ?")
                params.append(f'%{country}%')
            if city:
                conditions.append("city LIKE ?")
                params.append(f'%{city}%')
            if room_type and room_type != 'Tümü':
                conditions.append("room_type = ?")
                params.append(room_type)
            
            if conditions:
                query += " WHERE " + " AND ".join(conditions)
            
            # Add sorting
            if sort_column:
                # Map display names to actual column names in the database
                column_map = {
                    'ID': 'id',
                    'Adı Soyadı': 'guest_name',
                    'Unvan': 'guest_title',
                    'Ülke': 'country',
                    'Şehir': 'city',
                    'Giriş Tarihi': 'check_in_date',
                    'Çıkış Tarihi': 'check_out_date',
                    'Oda Tipi': 'room_type',
                    'Gecelik Ücret': 'nightly_rate',
                    'Toplam Ücret': 'total_amount'
                }
                db_column = column_map.get(sort_column, 'check_in_date') # Default to check_in_date
                query += f" ORDER BY {db_column} {sort_order}"
            else:
                query += " ORDER BY check_in_date DESC" # Default sort

            self.db.cursor.execute(query, tuple(params))
            columns = [description[0] for description in self.db.cursor.description]
            return [dict(zip(columns, row)) for row in self.db.cursor.fetchall()]
        except Exception as e:
            print(f"Error fetching stays: {e}")
            raise
    
    def get_stay_by_id(self, stay_id: int) -> Optional[Dict[str, Any]]:
        """Get a specific stay record by ID.
        
        Args:
            stay_id (int): ID of the stay record
            
        Returns:
            Optional[Dict[str, Any]]: Stay record if found, None otherwise
        """
        try:
            self.db.cursor.execute('SELECT id, guest_name, guest_title, country, city, check_in_date, check_out_date, room_type, nightly_rate, total_amount, created_at, updated_at FROM stays WHERE id = ?', (stay_id,))
            row = self.db.cursor.fetchone()
            if row:
                columns = [description[0] for description in self.db.cursor.description]
                return dict(zip(columns, row))
            return None
        except Exception as e:
            print(f"Error fetching stay: {e}")
            raise
    
    def update_stay(self, stay_id: int, **kwargs) -> bool:
        """Update a stay record.
        
        Args:
            stay_id (int): ID of the stay record to update
            **kwargs: Fields to update and their new values
            
        Returns:
            bool: True if update successful, False otherwise
        """
        try:
            if not kwargs:
                return False
                
            update_fields = []
            values = []
            
            for key, value in kwargs.items():
                if key in ['guest_name', 'guest_title', 'country', 'city',
                          'check_in_date', 'check_out_date', 'room_type',
                          'nightly_rate']:
                    update_fields.append(f"{key} = ?")
                    values.append(value)
            
            if not update_fields:
                return False
                
            # Recalculate total amount if dates or rate changed
            if any(key in kwargs for key in ['check_in_date', 'check_out_date', 'nightly_rate']):
                stay = self.get_stay_by_id(stay_id)
                if stay:
                    check_in = datetime.strptime(
                        kwargs.get('check_in_date', stay['check_in_date']), 
                        "%Y-%m-%d"
                    )
                    check_out = datetime.strptime(
                        kwargs.get('check_out_date', stay['check_out_date']), 
                        "%Y-%m-%d"
                    )
                    nights = (check_out - check_in).days
                    nightly_rate = kwargs.get('nightly_rate', stay['nightly_rate'])
                    total_amount = nights * nightly_rate
                    
                    update_fields.append("total_amount = ?")
                    values.append(total_amount)
            
            update_fields.append("updated_at = ?")
            values.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            values.append(stay_id)
            
            query = f'''
                UPDATE stays 
                SET {', '.join(update_fields)}
                WHERE id = ?
            '''
            
            self.db.cursor.execute(query, values)
            self.db.conn.commit()
            return True
        except Exception as e:
            print(f"Error updating stay: {e}")
            raise
    
    def delete_stay(self, stay_id: int) -> bool:
        """Delete a stay record.
        
        Args:
            stay_id (int): ID of the stay record to delete
            
        Returns:
            bool: True if deletion successful, False otherwise
        """
        try:
            self.db.cursor.execute('DELETE FROM stays WHERE id = ?', (stay_id,))
            self.db.conn.commit()
            return self.db.cursor.rowcount > 0
        except Exception as e:
            print(f"Error deleting stay: {e}")
            raise
    
    def get_stay_statistics(self) -> Dict[str, Any]:
        """Get statistics about stays including room type counts and total amounts.
        
        Returns:
            Dict[str, Any]: Dictionary containing stay statistics
        """
        try:
            # Get room type counts
            self.db.cursor.execute('''
                SELECT room_type, COUNT(*) as count, SUM(total_amount) as total_amount
                FROM stays
                GROUP BY room_type
            ''')
            room_stats = self.db.cursor.fetchall()
            
            # Get total stays and amount
            self.db.cursor.execute('''
                SELECT COUNT(*) as total_stays, SUM(total_amount) as total_amount
                FROM stays
            ''')
            total_stats = self.db.cursor.fetchone()
            
            # Format the statistics
            stats = {
                'room_types': {
                    row[0]: {
                        'count': row[1],
                        'total_amount': row[2]
                    } for row in room_stats
                },
                'total_stays': total_stats[0],
                'total_amount': total_stats[1]
            }
            
            return stats
        except Exception as e:
            print(f"Error getting stay statistics: {e}")
            raise
    
    def get_detailed_stay_report(self, start_date: str = None, end_date: str = None) -> List[Dict[str, Any]]:
        """Get detailed stay report data for Excel export with comprehensive statistics.
        
        Args:
            start_date (str): Optional start date for filtering (YYYY-MM-DD)
            end_date (str): Optional end date for filtering (YYYY-MM-DD)
            
        Returns:
            List[Dict[str, Any]]: List of dictionaries containing stay report data
        """
        try:
            # Base query with date filtering
            query = '''
                WITH stay_stats AS (
                    SELECT 
                        guest_name,
                        guest_title,
                        country,
                        city,
                        check_in_date,
                        check_out_date,
                        room_type,
                        nightly_rate,
                        total_amount,
                        CASE 
                            WHEN julianday(check_out_date) - julianday(check_in_date) = 1 THEN 'Single'
                            WHEN julianday(check_out_date) - julianday(check_in_date) = 2 THEN 'Double'
                            WHEN julianday(check_out_date) - julianday(check_in_date) = 3 THEN 'Triple'
                            ELSE 'Multiple'
                        END as stay_type,
                        julianday(check_out_date) - julianday(check_in_date) as nights
                    FROM stays
                    WHERE 1=1
                '''
            
            params = []
            if start_date:
                query += " AND check_in_date >= ?"
                params.append(start_date)
            if end_date:
                query += " AND check_out_date <= ?"
                params.append(end_date)
            
            # Add summary statistics
            query += '''
                ),
                summary_stats AS (
                    SELECT
                        COUNT(DISTINCT guest_name) as total_guests,
                        COUNT(*) as total_stays,
                        SUM(nights) as total_nights,
                        SUM(total_amount) as total_revenue
                    FROM stay_stats
                )
                SELECT 
                    s.*,
                    ss.total_guests,
                    ss.total_stays,
                    ss.total_nights,
                    ss.total_revenue
                FROM stay_stats s
                CROSS JOIN summary_stats ss
                ORDER BY s.check_in_date DESC
            '''
            
            self.db.cursor.execute(query, tuple(params))
            stays = self.db.cursor.fetchall()
            
            # Format the data with additional statistics
            report_data = []
            for stay in stays:
                report_data.append({
                    'Misafir Adı': stay[0],
                    'Unvan': stay[1],
                    'Ülke': stay[2],
                    'Şehir': stay[3],
                    'Giriş Tarihi': stay[4],
                    'Çıkış Tarihi': stay[5],
                    'Oda Tipi': stay[6],
                    'Gecelik Ücret': stay[7],
                    'Toplam Ücret': stay[8],
                    'Konaklama Tipi': stay[9],
                    'Konaklama Süresi (Gün)': stay[10],
                    'Toplam Misafir Sayısı': stay[11],
                    'Toplam Konaklama Sayısı': stay[12],
                    'Toplam Konaklama Günü': stay[13],
                    'Toplam Gelir': stay[14]
                })
            
            # Emit signal when report is generated
            self.report_generated.emit()
            return report_data
            
        except Exception as e:
            print(f"Error generating detailed stay report: {e}")
            raise 