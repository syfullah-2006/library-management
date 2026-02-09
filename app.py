from flask import Flask, render_template, request, redirect, url_for, session, flash
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = 'super_secret_key'

DB_FILE = 'library_data.xlsx'

# --- Helper Functions ---

def init_db():
    """Creates the Excel file if it doesn't exist."""
    if not os.path.exists(DB_FILE):
        df = pd.DataFrame(columns=['BookID', 'Title', 'Author', 'Status', 'IssuedTo'])
        # FIX: Added engine='openpyxl'
        df.to_excel(DB_FILE, index=False, engine='openpyxl')

def read_db():
    """Reads the Excel file and forces correct data types."""
    if os.path.exists(DB_FILE):
        # FIX 1: Explicitly use openpyxl engine
        df = pd.read_excel(DB_FILE, engine='openpyxl')
        
        # FIX 2: Force 'IssuedTo' to be Object (Text), avoiding the float64 error
        # We also force BookID to be string so '101' doesn't become 101.0
        df['IssuedTo'] = df['IssuedTo'].astype(object)
        df['BookID'] = df['BookID'].astype(str)
        
        # Fill empty cells with 'None' string so they aren't treated as math errors
        df['IssuedTo'] = df['IssuedTo'].fillna('None')
        
        return df
    return pd.DataFrame(columns=['BookID', 'Title', 'Author', 'Status', 'IssuedTo'])

def save_db(df):
    """Saves the DataFrame back to Excel."""
    # FIX: Added engine='openpyxl'
    df.to_excel(DB_FILE, index=False, engine='openpyxl')

# Initialize DB on start
init_db()

# --- Routes ---

@app.route('/', methods=['GET', 'POST'])
def index():
    df = read_db()
    results = None
    
    if request.method == 'POST':
        search_query = request.form.get('search', '').lower()
        # Convert columns to string before searching to avoid errors
        results = df[
            df['Title'].astype(str).str.lower().str.contains(search_query) |
            df['Author'].astype(str).str.lower().str.contains(search_query) |
            df['BookID'].astype(str).str.contains(search_query)
        ].to_dict('records')
    
    if results is None:
        results = df.to_dict('records')

    return render_template('index.html', books=results)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        if username == 'admin' and password == 'pass123':
            session['logged_in'] = True
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid Credentials!', 'danger')
            
    return render_template('login.html')

@app.route('/dashboard')
def dashboard():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    df = read_db()
    books = df.to_dict('records')
    return render_template('dashboard.html', books=books)

@app.route('/add_book', methods=['POST'])
def add_book():
    if not session.get('logged_in'): return redirect(url_for('login'))
    
    df = read_db()
    new_book = {
        'BookID': str(request.form['book_id']),
        'Title': request.form['title'],
        'Author': request.form['author'],
        'Status': 'Available',
        'IssuedTo': 'None'
    }
    
    new_df = pd.DataFrame([new_book])
    # Ensure the new dataframe also has the right types
    new_df['IssuedTo'] = new_df['IssuedTo'].astype(object)
    
    df = pd.concat([df, new_df], ignore_index=True)
    save_db(df)
    
    flash('Book added successfully!', 'success')
    return redirect(url_for('dashboard'))

@app.route('/issue_book', methods=['POST'])
def issue_book():
    if not session.get('logged_in'): return redirect(url_for('login'))
    
    book_id = str(request.form['book_id'])
    student_name = request.form['student_name']
    
    df = read_db()
    
    # Use string comparison
    mask = df['BookID'] == book_id
    
    if mask.any():
        idx = df.index[mask][0]
        if df.at[idx, 'Status'] == 'Available':
            df.at[idx, 'Status'] = 'Issued'
            # This line caused your error before, but now 'IssuedTo' is text-compatible
            df.at[idx, 'IssuedTo'] = student_name 
            save_db(df)
            flash(f'Book issued to {student_name}', 'success')
        else:
            flash('Book is already issued!', 'warning')
    else:
        flash('Book ID not found.', 'danger')
        
    return redirect(url_for('dashboard'))

@app.route('/return_book', methods=['POST'])
def return_book():
    if not session.get('logged_in'): return redirect(url_for('login'))
    
    book_id = str(request.form['book_id'])
    df = read_db()
    mask = df['BookID'] == book_id
    
    if mask.any():
        idx = df.index[mask][0]
        df.at[idx, 'Status'] = 'Available'
        df.at[idx, 'IssuedTo'] = 'None'
        save_db(df)
        flash('Book returned successfully.', 'success')
    else:
        flash('Book ID not found.', 'danger')

    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)

    