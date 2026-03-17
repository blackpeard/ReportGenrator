"""
Utility functions
"""
import os
from pathlib import Path

def ensure_folder(folder_path):
    """Create folder if it doesn't exist"""
    Path(folder_path).mkdir(parents=True, exist_ok=True)
    return folder_path

def get_output_filename(template_type):
    """Generate output filename"""
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{template_type}_report_{timestamp}.docx"