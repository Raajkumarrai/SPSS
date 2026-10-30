import os

class Config:
    SECRET_KEY = 'your-secret-key-change-this-in-production'
    
    # MySQL Configuration (Only Database)
    MYSQL_HOST = 'localhost'
    MYSQL_USER = 'root'
    MYSQL_PASSWORD = ''  # Update with your MySQL root password
    MYSQL_DATABASE = 'spps_db'
    MYSQL_PORT = 3306
    
    @staticmethod
    def get_database_config():
        return {
            'host': Config.MYSQL_HOST,
            'user': Config.MYSQL_USER,
            'password': Config.MYSQL_PASSWORD,
            'database': Config.MYSQL_DATABASE,
            'port': Config.MYSQL_PORT,
            'charset': 'utf8mb4',
            'autocommit': True,
            'pool_reset_session': True,
            'use_unicode': True,
            'get_warnings': False
        }

class DevelopmentConfig(Config):
    DEBUG = True

class ProductionConfig(Config):
    DEBUG = False

# Choose configuration based on environment
config = ProductionConfig()  # Change to DevelopmentConfig for development 