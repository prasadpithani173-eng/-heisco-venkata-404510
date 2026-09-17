import os
import sqlite3
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "users.db")

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initializes the SQLite users.db and ensures the default admin exists."""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL,
            display_name TEXT DEFAULT '',
            employee_id TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()

    try:
        cursor.execute("ALTER TABLE users ADD COLUMN display_name TEXT DEFAULT ''")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN employee_id TEXT DEFAULT ''")
    except Exception:
        pass
    conn.commit()

    # ALWAYS do this:
    # DELETE FROM users WHERE username='admin'
    # INSERT new admin with VENKATA NAGA PRASAD, 404510, and generate_password_hash('heisco123')
    cursor.execute("DELETE FROM users WHERE username='admin'")
    hashed_pw = generate_password_hash("heisco123")
    cursor.execute(
        "INSERT INTO users (username, password, role, display_name, employee_id) VALUES (?, ?, ?, ?, ?)",
        ("admin", hashed_pw, "admin", "VENKATA NAGA PRASAD", "404510")
    )
    conn.commit()
    print("Admin re-created with VENKATA NAGA PRASAD (404510) and heisco123")

    # Ensure demo role accounts exist for easy testing of role-based access
    demo_accounts = [
        ("officer", "officer123", "officer", "HSE Officer", "404511"),
        ("supervisor", "super123", "supervisor", "HSE Supervisor", "404512"),
        ("engineer", "eng123", "engineer", "HSE Engineer", "404513")
    ]
    for u, p, r, dn, eid in demo_accounts:
        hp = generate_password_hash(p)
        cursor.execute("SELECT * FROM users WHERE username = ?", (u,))
        if not cursor.fetchone():
            cursor.execute(
                "INSERT INTO users (username, password, role, display_name, employee_id) VALUES (?, ?, ?, ?, ?)",
                (u, hp, r, dn, eid)
            )
        else:
            cursor.execute(
                "UPDATE users SET password = ?, role = ?, display_name = ?, employee_id = ? WHERE username = ?",
                (hp, r, dn, eid, u)
            )
    conn.commit()
    conn.close()

def get_user_by_username(username: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE username = ?", (username.strip(),))
    user = cursor.fetchone()
    conn.close()
    return user

def verify_user(username: str, password: str):
    """
    Verifies user credentials.
    Supports both hashed passwords and legacy plain text comparison (e.g. heisco123).
    """
    user = get_user_by_username(username)
    if not user:
        return None
    
    stored_pw = user["password"]
    is_valid = False
    
    # Check hashed or fallback plain text
    try:
        if check_password_hash(stored_pw, password):
            is_valid = True
    except Exception:
        pass

    if not is_valid and stored_pw == password:
        is_valid = True

    if is_valid:
        return user
    return None

def update_user_password(username: str, new_password: str) -> bool:
    """Updates the user's password in users.db."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM users WHERE username = ?", (username.strip(),))
    user = cursor.fetchone()
    if not user:
        conn.close()
        return False
    
    hashed_pw = generate_password_hash(new_password)
    cursor.execute("UPDATE users SET password = ? WHERE username = ?", (hashed_pw, username.strip()))
    conn.commit()
    conn.close()
    return True

def create_new_user(username: str, password: str, role: str) -> tuple[bool, str]:
    """Creates a new user account with specified role."""
    valid_roles = ["admin", "officer", "supervisor", "engineer"]
    role = role.strip().lower()
    if role not in valid_roles:
        return False, f"Invalid role. Must be one of: {', '.join(valid_roles)}"
    
    username = username.strip()
    if not username or len(username) < 3:
        return False, "Username must be at least 3 characters."
    if not password or len(password) < 4:
        return False, "Password must be at least 4 characters."

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM users WHERE username = ?", (username,))
    if cursor.fetchone():
        conn.close()
        return False, f"Username '{username}' already exists."

    hashed_pw = generate_password_hash(password)
    cursor.execute(
        "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
        (username, hashed_pw, role)
    )
    conn.commit()
    conn.close()
    return True, f"User '{username}' created successfully as '{role}'."

def get_all_users():
    """Retrieves all registered users."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, role, display_name, employee_id, created_at FROM users ORDER BY id ASC")
    users = cursor.fetchall()
    conn.close()
    return users

def delete_user_by_id(user_id: int) -> bool:
    conn = get_db_connection()
    cursor = conn.cursor()
    # Prevent deleting the main admin
    cursor.execute("SELECT username FROM users WHERE id = ?", (user_id,))
    row = cursor.fetchone()
    if row and row["username"] == "admin":
        conn.close()
        return False
    cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()
    return True
