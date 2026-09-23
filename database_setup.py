import mysql.connector # type: ignore

def create_db():
    try:
        # CONNECT TO MYSQL 
        db = mysql.connector.connect(
            host="localhost",
            user="root",
            password="@ritesh"  
        )
        cursor = db.cursor()
        
        # Create Database and Table
        cursor.execute("CREATE DATABASE IF NOT EXISTS traffic_ai")
        cursor.execute("USE traffic_ai")
        
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS violations (
            id INT AUTO_INCREMENT PRIMARY KEY,
            violation_type VARCHAR(50),
            plate_number VARCHAR(20),
            confidence FLOAT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
            location VARCHAR(100) DEFAULT 'Main Gate Camera 1',
            evidence_image VARCHAR(255) DEFAULT NULL
        )
        """)
        print("✅ Database and Table Created Successfully!")
        db.close()
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    create_db()