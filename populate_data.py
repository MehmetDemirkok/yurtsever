from datetime import datetime, timedelta
from faker import Faker
import random
from database import Database
from stay_model import StayModel

def populate_demo_data(num_records: int = 100):
    """Populate the database with demo stay records.

    Args:
        num_records (int): The number of demo records to create.
    """
    fake = Faker('tr_TR')  # Use Turkish locale for more realistic names/places
    db = Database()
    stay_model = StayModel(db)

    print(f"Generating {num_records} demo stay records...")

    for i in range(num_records):
        guest_name = fake.name()
        guest_title = random.choice(['Bay', 'Bayan', 'Çocuk'])
        country = fake.country()
        city = fake.city()

        check_in_date = fake.date_between(start_date='-2y', end_date='today')
        check_out_date = check_in_date + timedelta(days=random.randint(1, 10))
        
        room_type = random.choice(['Single Oda', 'Double Oda', 'Triple Oda', 'Suit Oda', 'Aile Odası'])
        nightly_rate = round(random.uniform(50.0, 500.0), 2)

        try:
            stay_model.create_stay(
                guest_name=guest_name,
                guest_title=guest_title,
                country=country,
                city=city,
                check_in_date=check_in_date.strftime("%Y-%m-%d"),
                check_out_date=check_out_date.strftime("%Y-%m-%d"),
                room_type=room_type,
                nightly_rate=nightly_rate
            )
            if (i + 1) % 10 == 0:
                print(f"{i + 1} records created.")
        except Exception as e:
            print(f"Error creating demo stay record {i+1}: {e}")
            db.conn.rollback() # Rollback in case of error
            
    db.close()
    print("Demo data population complete.")

if __name__ == "__main__":
    populate_demo_data(100) 