# LISQuant - Financial Analysis Application

A Flask web application for financial analysis, allowing users to analyze stock tickers and view historical analyses.

## Prerequisites

*   Python 3.7+
*   PostgreSQL server installed and running.
*   Access to a Bloomberg Terminal and `blpapi` Python library installed (for the actual financial data population via `Inputs_Cur.py`).

## Setup Instructions

1.  **Clone the Repository:**
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Create a Python Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows: venv\Scripts\activate
    ```

3.  **Install Python Dependencies:**
    A `requirements.txt` file is not yet present in this repository. You will need to manually install the necessary packages. Key dependencies include:
    *   `Flask`
    *   `Flask-Login`
    *   `psycopg2-binary` (or `psycopg2` if you have build tools)
    *   `openpyxl`
    *   `blpapi` (requires separate installation from Bloomberg)
    *   `Werkzeug` (usually installed with Flask)
    
    Example:
    ```bash
    pip install Flask Flask-Login psycopg2-binary openpyxl Werkzeug
    # blpapi needs to be installed following Bloomberg's documentation
    ```
    *(It is recommended to generate a `requirements.txt` file for easier dependency management.)*

4.  **Set Environment Variables:**
    The application uses environment variables to configure the database connection. Create a `.env` file (and ensure it's in your `.gitignore`) or set these in your shell environment:

    *   `DB_NAME`: The name of your database (defaults to `lisquant_db` in the init script).
    *   `DB_USER`: Your PostgreSQL username (e.g., `postgres`).
    *   `DB_PASSWORD`: **Required.** Your PostgreSQL password for the specified user.
    *   `DB_HOST`: The database host (defaults to `localhost`).
    *   `DB_PORT`: The database port (defaults to `5432`).
    *   `FLASK_APP`: Set to `app.py`.
    *   `FLASK_ENV`: Set to `development` for debug mode (optional).

    Example for bash:
    ```bash
    export DB_PASSWORD="your_secure_password"
    export FLASK_APP=app.py
    export FLASK_ENV=development
    ```

5.  **Initialize the Database:**
    Once the environment variables are set (especially `DB_PASSWORD`), run the database initialization script:
    ```bash
    python initialize_database.py
    ```
    This will create the database (if it doesn't exist) and the necessary tables (`users`, `analyses`).

6.  **Provide the Excel Template:**
    The financial analysis feature (`Inputs_Cur.py`) requires an Excel template named `LIS_Valuation_Empty.xlsx` to be present in the root of the repository. You will need to provide this file. A dummy version was created during development, which should be replaced.

7.  **Run the Flask Application:**
    ```bash
    flask run
    ```
    The application should now be running, typically at `http://127.0.0.1:5000/`.

## Project Structure

*   `app.py`: Main Flask application file.
*   `initialize_database.py`: Script to set up the database schema.
*   `Inputs_Cur.py`: Contains logic for fetching financial data (requires Bloomberg).
*   `templates/`: HTML templates for the web interface.
*   `static/`: CSS and JavaScript files.
*   `LIS_Valuation_Empty.xlsx`: (User-provided) Template for financial models.
*   `init_db.sql`: Original SQL script for database setup (now superseded by `initialize_database.py` but kept for reference).
