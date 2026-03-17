"""
POC Finder - Matches POC references with actual files
Supports both reference-based and folder-based matching
"""
from pathlib import Path
import glob
import os
import re

class POCFinder:
    def __init__(self):
        self.poc_files = {}
        self.poc_folder = None
        self.folder_structure = {}  # New: Store folder-based POCs
        
    def scan_folder(self, folder_path):
        """Scan folder and index all image files"""
        self.poc_folder = Path(folder_path)
        self.poc_files = {}
        self.folder_structure = {}
        
        image_extensions = ['*.png', '*.jpg', '*.jpeg', '*.gif', '*.bmp']
        
        for ext in image_extensions:
            for file in self.poc_folder.rglob(ext):
                # Original indexing for reference-based matching
                name = file.stem.lower()
                full_name = file.name.lower()
                
                self.poc_files[name] = str(file)
                self.poc_files[full_name] = str(file)
                
                # Clean name indexing
                clean_name = re.sub(r'[^a-zA-Z0-9]', '', name)
                if clean_name != name:
                    self.poc_files[clean_name] = str(file)
                
                # NEW: Index by folder structure (vulnerability name)
                parent_folder = file.parent.name
                if parent_folder not in self.folder_structure:
                    self.folder_structure[parent_folder] = []
                
                self.folder_structure[parent_folder].append(str(file))
        
        # Sort images in each folder naturally
        for folder in self.folder_structure:
            self.folder_structure[folder].sort(
                key=lambda x: [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', x)]
            )
        
        return len(self.poc_files)
    
    def find_poc(self, reference):
        """Find POC file matching the reference (original method)"""
        if not reference or not self.poc_files:
            return None
            
        ref = str(reference).strip().lower()
        
        # Direct match
        if ref in self.poc_files:
            return self.poc_files[ref]
        
        # Try without extension
        ref_without_ext = Path(ref).stem
        if ref_without_ext in self.poc_files:
            return self.poc_files[ref_without_ext]
        
        # Try partial match
        for key, path in self.poc_files.items():
            if ref in key or key in ref:
                return path
        
        return None
    
    # NEW METHOD 1: Get POCs by vulnerability title (folder name)
    def get_pocs_by_vulnerability(self, vulnerability_title):
        """
        Get all POC images for a vulnerability based on folder name
        
        Args:
            vulnerability_title: Title of the vulnerability from Excel
            
        Returns:
            List of image paths sorted naturally, or empty list if not found
        """
        if not vulnerability_title or not self.folder_structure:
            return []
        
        # Clean the title to match folder name
        clean_title = re.sub(r'[\\/*?:"<>|]', '', vulnerability_title).strip()
        
        # Try exact match first
        if clean_title in self.folder_structure:
            return self.folder_structure[clean_title]
        
        # Try case-insensitive match
        for folder in self.folder_structure:
            if folder.lower() == clean_title.lower():
                return self.folder_structure[folder]
        
        # Try partial match (if folder name contains title or vice versa)
        for folder, images in self.folder_structure.items():
            if (clean_title.lower() in folder.lower()) or (folder.lower() in clean_title.lower()):
                return images
        
        return []
    
    # NEW METHOD 2: Check if vulnerability has POCs
    def has_pocs(self, vulnerability_title):
        """Check if a vulnerability has any POC images"""
        return len(self.get_pocs_by_vulnerability(vulnerability_title)) > 0
    
    # NEW METHOD 3: Get folder names (for debugging/listing)
    def get_all_vulnerability_folders(self):
        """Get list of all folder names that have POCs"""
        return list(self.folder_structure.keys())