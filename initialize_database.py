import os
import psycopg2
from psycopg2 import sql
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT # Needed for CREATE DATABASE

def get_db_config():
    config = {
        'dbname': os.environ.get('DB_NAME', 'lisquant_db'),
        'user': os.environ.get('DB_USER', 'postgres'),
        'password': os.environ.get('DB_PASSWORD'), # No default, script should fail if not set and required
        'host': os.environ.get('DB_HOST', 'localhost'),
        'port': os.environ.get('DB_PORT', '5432')
    }
    if not config['password']:
        raise ValueError("DB_PASSWORD environment variable is not set. Please set it before running.")
    return config

def main():
    db_config = get_db_config()
    db_to_create = db_config.pop('dbname') # Remove dbname for initial connection, store it

    print(f"Attempting to connect to PostgreSQL server at {db_config['host']}:{db_config['port']} as user {db_config['user']}...")
    
    conn = None
    try:
        # Connect to the default 'postgres' database (or any existing database) to create the new one
        conn_params_default = db_config.copy()
        conn_params_default['dbname'] = 'postgres' # Or 'template1'
        
        conn = psycopg2.connect(**conn_params_default)
        conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT) # Needed for CREATE DATABASE
        cur = conn.cursor()
        
        print(f"Checking if database '{db_to_create}' exists...")
        cur.execute(sql.SQL("SELECT 1 FROM pg_database WHERE datname = %s"), (db_to_create,))
        exists = cur.fetchone()
        
        if not exists:
            print(f"Database '{db_to_create}' does not exist. Creating it...")
            cur.execute(sql.SQL("CREATE DATABASE {}").format(sql.Identifier(db_to_create)))
            print(f"Database '{db_to_create}' created successfully.")
        else:
            print(f"Database '{db_to_create}' already exists.")
            
        cur.close()
        conn.close()

        # Now connect to the newly created (or existing) database
        print(f"Connecting to database '{db_to_create}'...")
        conn_params_target = db_config.copy()
        conn_params_target['dbname'] = db_to_create
        conn = psycopg2.connect(**conn_params_target)
        cur = conn.cursor()

        print("Creating table 'users' if it doesn't exist...")
        cur.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username VARCHAR(50) UNIQUE NOT NULL,
                password TEXT NOT NULL
            );
        ''')
        print("'users' table created or already exists.")

        print("Creating table 'analyses' if it doesn't exist...")
        cur.execute('''
            CREATE TABLE IF NOT EXISTS analyses (
                id SERIAL PRIMARY KEY,
                user_id INTEGER REFERENCES users(id),
                ticker VARCHAR(20) NOT NULL,
                filename VARCHAR(255) NOT NULL,
                created_at TIMESTAMP NOT NULL
            );
        ''')
        print("'analyses' table created or already exists.")

        conn.commit()
        print("Database schema initialized successfully.")

    except psycopg2.Error as e:
        print(f"An error occurred: {e}")
    finally:
        if conn:
            cur.close()
            conn.close()
            print("Database connection closed.")

if __name__ == '__main__':
    main()
