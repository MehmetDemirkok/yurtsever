import sqlite3
from datetime import datetime
from pathlib import Path

class Database:
    def __init__(self, db_path: str = "hotel_stays.db"):
        """Initialize database connection and create tables if they don't exist.
        
        Args:
            db_path (str): Path to SQLite database file
        """
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self.connect()
        self.create_tables()
    
    def connect(self):
        """Establish connection to SQLite database."""
        try:
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()
        except sqlite3.Error as e:
            print(f"Database connection error: {e}")
            raise
    
    def create_tables(self):
        """Create necessary database tables if they don't exist."""
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS stays (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                guest_name TEXT NOT NULL,
                guest_title TEXT NOT NULL,
                country TEXT NOT NULL,
                city TEXT NOT NULL,
                check_in_date TEXT NOT NULL,
                check_out_date TEXT NOT NULL,
                room_type TEXT NOT NULL,
                nightly_rate REAL NOT NULL,
                total_amount REAL NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        ''')
        self.conn.commit()

        # Add new columns if they don't exist
        self._add_column_if_not_exists('stays', 'company_name', 'TEXT NOT NULL DEFAULT \'\'')
        self._add_column_if_not_exists('stays', 'hotel_name', 'TEXT NOT NULL DEFAULT \'\'')

        # Migrate column name from nightly_rate to hotel_purchase_price
        self._migrate_nightly_rate_to_hotel_purchase_price()

        # Add new columns for hotel sale price and total sale amount if they don't exist
        self._add_column_if_not_exists('stays', 'hotel_sale_price', 'REAL NOT NULL DEFAULT 0.0')
        self._add_column_if_not_exists('stays', 'total_sale_amount', 'REAL NOT NULL DEFAULT 0.0')

        # Migrate column name from total_amount to hotel_purchase_total_amount
        self._migrate_total_amount_to_hotel_purchase_total_amount()
    
    def _add_column_if_not_exists(self, table_name: str, column_name: str, column_definition: str):
        self.cursor.execute(f"PRAGMA table_info({table_name})")
        columns = [info[1] for info in self.cursor.fetchall()]
        if column_name not in columns:
            self.cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_definition}")
            self.conn.commit()

    def _column_exists(self, table_name: str, column_name: str) -> bool:
        self.cursor.execute(f"PRAGMA table_info({table_name})")
        columns = [info[1] for info in self.cursor.fetchall()]
        return column_name in columns

    def _migrate_nightly_rate_to_hotel_purchase_price(self):
        if self._column_exists('stays', 'nightly_rate') and not self._column_exists('stays', 'hotel_purchase_price'):
            print("Migrating 'nightly_rate' to 'hotel_purchase_price'...")
            try:
                self.cursor.execute("ALTER TABLE stays RENAME TO old_stays")
                self.conn.commit()

                self.cursor.execute('''
                    CREATE TABLE stays (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        guest_name TEXT NOT NULL,
                        guest_title TEXT NOT NULL,
                        company_name TEXT NOT NULL DEFAULT '',
                        country TEXT NOT NULL,
                        city TEXT NOT NULL,
                        check_in_date TEXT NOT NULL,
                        check_out_date TEXT NOT NULL,
                        room_type TEXT NOT NULL,
                        hotel_purchase_price REAL NOT NULL DEFAULT 0.0,
                        hotel_purchase_total_amount REAL NOT NULL DEFAULT 0.0,
                        hotel_name TEXT NOT NULL DEFAULT '',
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL,
                        hotel_sale_price REAL NOT NULL DEFAULT 0.0,
                        total_sale_amount REAL NOT NULL DEFAULT 0.0
                    )
                ''')
                self.conn.commit()

                self.cursor.execute('''
                    INSERT INTO stays (
                        id, guest_name, guest_title, company_name, country, city,
                        check_in_date, check_out_date, room_type, hotel_purchase_price,
                        hotel_purchase_total_amount, hotel_name, created_at, updated_at, hotel_sale_price, total_sale_amount
                    ) SELECT
                        id, guest_name, guest_title, company_name, country, city,
                        check_in_date, check_out_date, room_type, nightly_rate,
                        total_amount, hotel_name, created_at, updated_at, 0.0, 0.0  -- Default values for new columns
                    FROM old_stays
                ''')
                self.conn.commit()

                self.cursor.execute("DROP TABLE old_stays")
                self.conn.commit()
                print("Migration complete.")
            except sqlite3.Error as e:
                print(f"Database migration error: {e}")
                # Revert if migration fails to prevent data loss
                if self._table_exists('old_stays') and not self._table_exists('stays'):
                    self.cursor.execute("ALTER TABLE old_stays RENAME TO stays")
                    self.conn.commit()
                raise
        
    def _migrate_total_amount_to_hotel_purchase_total_amount(self):
        if self._column_exists('stays', 'total_amount') and not self._column_exists('stays', 'hotel_purchase_total_amount'):
            print("Migrating 'total_amount' to 'hotel_purchase_total_amount'...")
            try:
                self.cursor.execute("ALTER TABLE stays RENAME TO old_stays")
                self.conn.commit()

                self.cursor.execute('''
                    CREATE TABLE stays (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        guest_name TEXT NOT NULL,
                        guest_title TEXT NOT NULL,
                        company_name TEXT NOT NULL DEFAULT '',
                        country TEXT NOT NULL,
                        city TEXT NOT NULL,
                        check_in_date TEXT NOT NULL,
                        check_out_date TEXT NOT NULL,
                        room_type TEXT NOT NULL,
                        hotel_purchase_price REAL NOT NULL DEFAULT 0.0,
                        hotel_purchase_total_amount REAL NOT NULL DEFAULT 0.0,
                        hotel_name TEXT NOT NULL DEFAULT '',
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL,
                        hotel_sale_price REAL NOT NULL DEFAULT 0.0,
                        total_sale_amount REAL NOT NULL DEFAULT 0.0
                    )
                ''')
                self.conn.commit()

                self.cursor.execute('''
                    INSERT INTO stays (
                        id, guest_name, guest_title, company_name, country, city,
                        check_in_date, check_out_date, room_type, hotel_purchase_price,
                        hotel_purchase_total_amount, hotel_name, created_at, updated_at, hotel_sale_price, total_sale_amount
                    ) SELECT
                        id, guest_name, guest_title, company_name, country, city,
                        check_in_date, check_out_date, room_type, hotel_purchase_price,
                        total_amount, hotel_name, created_at, updated_at, 0.0, 0.0
                    FROM old_stays
                ''')
                self.conn.commit()

                self.cursor.execute("DROP TABLE old_stays")
                self.conn.commit()
                print("Migration complete.")
            except sqlite3.Error as e:
                print(f"Database migration error: {e}")
                if self._table_exists('old_stays') and not self._table_exists('stays'):
                    self.cursor.execute("ALTER TABLE old_stays RENAME TO stays")
                    self.conn.commit()
                raise

    def _table_exists(self, table_name: str) -> bool:
        self.cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
        return self.cursor.fetchone() is not None

    def close(self):
        """Close database connection."""
        if self.conn:
            self.conn.close() 