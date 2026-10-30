#!/usr/bin/env python3
"""
SPPS - Student Performance Prediction System
Run script for the application
"""

import os
import sys

# Add the app directory to the Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'app'))

from app import app, create_default_admin

if __name__ == '__main__':
    # Create default admin user
    create_default_admin()
    
    # Run the application
    print("Starting SPPS Application...")
    print("Default admin credentials: admin@spps.com / admin123")
    print("Access the application at: http://localhost:5000")
    
    app.run(debug=True, host='0.0.0.0', port=5000) 