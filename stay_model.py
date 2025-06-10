from datetime import datetime
from typing import List, Optional, Dict, Any
from database import Database

class StayModel:
    def __init__(self, db: Database):
        """Initialize stay model with database connection.
        
        Args:
            db (Database): Database instance
        """
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
    
    def get_all_stays(self) -> List[Dict[str, Any]]:
        """Get all stay records.
        
        Returns:
            List[Dict[str, Any]]: List of stay records
        """
        try:
            self.db.cursor.execute('SELECT id, guest_name, check_in_date, check_out_date, nightly_rate, total_amount, created_at, updated_at FROM stays ORDER BY check_in_date DESC')
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
            self.db.cursor.execute('SELECT id, guest_name, check_in_date, check_out_date, nightly_rate, total_amount, created_at, updated_at FROM stays WHERE id = ?', (stay_id,))
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