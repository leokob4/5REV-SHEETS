import os
from flask import Flask, send_from_directory, request, jsonify
import bcrypt
import openpyxl

# --- Flask App Setup ---
# Corrected static_folder: This assumes web_server.py is in the root of 5REV-SHEETS
# and the 'js' folder is a direct subdirectory of 5REV-SHEETS.
app = Flask(__name__, static_folder='js')

# --- File Paths Configuration (Duplicate from gui.py for backend context) ---
USER_SHEETS_DIR = "user_sheets"
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx")

# Ensure user_sheets directory exists
os.makedirs(USER_SHEETS_DIR, exist_ok=True)

# --- Sheet Helpers (Adapted from gui.py) ---
# Note: For a production web application, direct Excel file manipulation
# is highly discouraged due to concurrency issues. A proper database (e.g., PostgreSQL, SQLite)
# and an ORM (e.g., SQLAlchemy) would be much more robust.
# This implementation is for demonstration purposes to integrate with your existing Excel data.

def load_users_from_excel_backend():
    """Loads user data from the database Excel file for backend use."""
    users = {}
    try:
        if not os.path.exists(DB_EXCEL_PATH):
            return {} # Return empty if file doesn't exist yet

        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        # Check if 'users' sheet exists, create if not (for first run setup)
        if 'users' not in wb.sheetnames:
            return {}

        users_sheet = wb["users"]
        for row in users_sheet.iter_rows(min_row=2): # Skip header row
            if len(row) >= 4 and all(cell.value is not None for cell in row[:4]):
                users[row[1].value] = { # username as key
                    "id": row[0].value,
                    "username": row[1].value,
                    "password_hash": row[2].value,
                    "role": row[3].value
                }
    except Exception as e:
        print(f"Error loading users in backend: {e}")
    return users

def register_user_backend(username, password, role="user"):
    """Registers a new user into the database Excel file for backend use."""
    try:
        # Check if DB_EXCEL_PATH exists. If not, create a new workbook.
        if not os.path.exists(DB_EXCEL_PATH):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "users"
            ws.append(["id", "username", "password_hash", "role"]) # Add headers
            # Also create an 'access' sheet if it doesn't exist
            access_ws = wb.create_sheet("access")
            access_ws.append(["role", "allowed_tools"])
            access_ws.append(["user", "mod1,mod2,mes_pcp"]) # Default user permissions
            access_ws.append(["admin", "all"]) # Default admin permissions
        else:
            wb = openpyxl.load_workbook(DB_EXCEL_PATH)

        if 'users' not in wb.sheetnames:
            ws = wb.create_sheet("users")
            ws.append(["id", "username", "password_hash", "role"]) # Add headers
        else:
            ws = wb["users"]

        # Check for existing username
        for row in ws.iter_rows(min_row=2):
            if row[1].value == username:
                raise ValueError("Username already exists.")

        # Determine next ID based on current max row
        next_id = ws.max_row if ws.max_row > 1 else 1 # If only headers, start at 1
        password_hash = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        ws.append([next_id, username, password_hash, role])
        wb.save(DB_EXCEL_PATH)
        return True
    except Exception as e:
        print(f"Error registering user in backend: {e}")
        return False

# --- Routes for serving HTML files ---
@app.route('/')
@app.route('/login') # Added route for /login
def serve_login():
    """Serves the login.html file when accessing the root URL or /login."""
    return send_from_directory(app.static_folder, 'login.html')

@app.route('/dashboard.html')
def serve_dashboard():
    """Serves the dashboard.html file."""
    return send_from_directory(app.static_folder, 'dashboard.html')

# --- API Endpoints ---
@app.route('/api/login', methods=['POST'])
def api_login():
    """Handles user login requests from the web frontend."""
    data = request.json
    username = data.get('username')
    password = data.get('password')

    if not username or not password:
        return jsonify({"message": "Username and password are required."}), 400

    users = load_users_from_excel_backend()
    user = users.get(username)

    if user and bcrypt.checkpw(password.encode('utf-8'), user["password_hash"].encode('utf-8')):
        # In a real app, generate and return a JWT or session token here
        return jsonify({"message": "Login successful!", "token": "fake-jwt-token-123"}), 200
    else:
        return jsonify({"message": "Invalid username or password."}), 401

@app.route('/api/register', methods=['POST'])
def api_register():
    """Handles user registration requests from the web frontend."""
    data = request.json
    username = data.get('username')
    password = data.get('password')

    if not username or not password:
        return jsonify({"message": "Username and password are required."}), 400

    users = load_users_from_excel_backend()
    if username in users:
        return jsonify({"message": "Username already exists."}), 409 # Conflict

    if register_user_backend(username, password):
        return jsonify({"message": f"User '{username}' registered successfully."}), 201 # Created
    else:
        return jsonify({"message": "An error occurred during registration."}), 500

# --- Main entry point for running the Flask app ---
if __name__ == '__main__':
    # Print the path Flask is serving static files from for debugging
    print(f"Flask app is serving static files from: {app.static_folder}")

    # Initialize the users database if it doesn't exist for the first run
    if not os.path.exists(DB_EXCEL_PATH):
        print(f"Creating initial database at {DB_EXCEL_PATH}...")
        try:
            wb = openpyxl.Workbook()
            ws_users = wb.active
            ws_users.title = "users"
            ws_users.append(["id", "username", "password_hash", "role"])

            ws_access = wb.create_sheet("access")
            ws_access.append(["role", "allowed_tools"])
            ws_access.append(["user", "mod1,mod2,mes_pcp"])
            ws_access.append(["admin", "all"])

            wb.save(DB_EXCEL_PATH)
            print("Database created successfully with default users and access rules.")
        except Exception as e:
            print(f"Failed to create initial database: {e}")

    # Run the Flask app
    # host='0.0.0.0' makes the server accessible from other devices on the network.
    # debug=True allows for automatic reloading on code changes and provides a debugger.
    app.run(host='0.0.0.0', port=8000, debug=True)
