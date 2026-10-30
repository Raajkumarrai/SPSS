import mysql.connector
from mysql.connector import Error
from config import config

class DatabaseManager:
    def __init__(self):
        self.config = config.get_database_config()
    
    def get_connection(self):
        """Get MySQL database connection"""
        return self._get_mysql_connection()
    
    def _get_mysql_connection(self):
        """Get MySQL connection"""
        try:
            connection = mysql.connector.connect(**self.config)
            connection.autocommit = True
            connection.get_warnings = False
            return connection
        except Error as e:
            print(f"Error connecting to MySQL: {e}")
            raise
    
    def execute_query(self, query, params=None):
        """Execute a query and return results"""
        connection = None
        cursor = None
        
        try:
            # Get a fresh connection for each query
            connection = self.get_connection()
            
            # Use regular cursor to avoid issues
            cursor = connection.cursor()
            
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            
            # Check if query returns results (SELECT, SHOW, DESCRIBE, etc.)
            query_upper = query.strip().upper()
            returns_results = (
                query_upper.startswith('SELECT') or 
                query_upper.startswith('SHOW') or 
                query_upper.startswith('DESCRIBE') or
                query_upper.startswith('DESC') or
                query_upper.startswith('EXPLAIN')
            )
            
            if returns_results and cursor.description:
                results = cursor.fetchall()
                
                # Get column names
                column_names = [desc[0] for desc in cursor.description]
                # Convert tuples to dictionaries
                dict_results = []
                for row in results:
                    dict_row = {}
                    for i, value in enumerate(row):
                        dict_row[column_names[i]] = value
                    dict_results.append(dict_row)
                return dict_results
            else:
                connection.commit()
                return cursor.lastrowid
        except Exception as e:
            if connection:
                connection.rollback()
            print(f"Error executing query: {e}")
            raise
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()
    
    def execute_many(self, query, params_list):
        """Execute multiple queries with different parameters"""
        connection = self.get_connection()
        cursor = None
        
        try:
            cursor = connection.cursor()
            cursor.executemany(query, params_list)
            connection.commit()
            return cursor.rowcount
        except Exception as e:
            if connection:
                connection.rollback()
            print(f"Error executing multiple queries: {e}")
            raise
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()
    
    def table_exists(self, table_name):
        """Check if a table exists"""
        query = "SHOW TABLES LIKE %s"
        result = self.execute_query(query, (table_name,))
        return len(result) > 0

# Global database manager instance
db_manager = DatabaseManager()

def get_db():
    """Get database connection (compatibility function)"""
    return db_manager.get_connection()

def execute_query(query, params=None):
    """Execute a query (compatibility function)"""
    return db_manager.execute_query(query, params) 