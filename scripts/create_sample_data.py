"""
Create sample data for testing
"""
import pandas as pd
import os
from pathlib import Path

def create_sample_excel():
    """Create sample Excel file"""
    data = {
        'ID': ['OBS001', 'OBS002', 'OBS003', 'OBS004'],
        'Observation': [
            'Login button not responding',
            'Page loads slowly',
            'Mobile menu broken',
            'API returns 500 error'
        ],
        'POC': ['login_error', 'performance', 'mobile_menu', 'api_error'],
        'Severity': ['High', 'Medium', 'High', 'Critical'],
        'Category': ['UI', 'Performance', 'Mobile', 'API']
    }
    
    df = pd.DataFrame(data)
    
    # Ensure folder exists
    Path("inputs/excel").mkdir(parents=True, exist_ok=True)
    
    # Save
    output = "inputs/excel/sample_observations.xlsx"
    df.to_excel(output, index=False)
    print(f"✅ Created: {output}")

def create_sample_pocs():
    """Create placeholder POC files"""
    pocs = ['login_error', 'performance', 'mobile_menu', 'api_error']
    
    # Ensure folder exists
    Path("inputs/pocs").mkdir(parents=True, exist_ok=True)
    
    for poc in pocs:
        file_path = f"inputs/pocs/{poc}.txt"  # Using .txt as placeholder
        with open(file_path, 'w') as f:
            f.write(f"Placeholder for {poc} screenshot")
    
    print(f"✅ Created {len(pocs)} placeholder POC files in inputs/pocs/")

if __name__ == "__main__":
    print("Creating sample data...")
    create_sample_excel()
    create_sample_pocs()
    print("\nDone! You can now run: python main.py")