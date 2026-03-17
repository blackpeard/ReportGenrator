"""
Dynamic Excel Reader - Works with ANY Excel file structure
Automatically detects sheets and columns
"""
import pandas as pd
from pathlib import Path
import re
import os

class ExcelReader:
    def __init__(self, config=None):
        self.config = config or {}
        self.excel_file = None
        self.sheets = {}
        
    def load(self, excel_path):
        """Load Excel file and get all sheet names"""
        self.excel_file = excel_path
        xl = pd.ExcelFile(excel_path)
        self.sheets = {name: None for name in xl.sheet_names}
        print(f"  ✓ Loaded Excel with sheets: {list(self.sheets.keys())}")
        return self.sheets.keys()
    
    def read_index_sheet(self):
        """Read Index/Summary sheet dynamically"""
        # Try common sheet names
        index_sheet_names = ["Index", "Summary", "Overview", "Dashboard", "Cover"]
        
        for sheet_name in index_sheet_names:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
                print(f"  ✓ Found index sheet: '{sheet_name}'")
                return df
            except:
                continue
        
        # If no index sheet found, return empty dataframe
        print("  ⚠ No index sheet found, using empty data")
        return pd.DataFrame()
    
    def read_scope_sheet(self):
        """Read Scope sheet dynamically"""
        scope_sheet_names = ["Scope", "Scopes", "In Scope", "Scope of Work"]
        
        for sheet_name in scope_sheet_names:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
                print(f"  ✓ Found scope sheet: '{sheet_name}'")
                return df
            except:
                continue
        
        return pd.DataFrame()
    
    def read_limitation_sheet(self):
        """Read Limitation sheet dynamically"""
        limitation_sheet_names = ["Limitation", "Limitations", "Out of Scope", "Exclusions"]
        
        for sheet_name in limitation_sheet_names:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
                print(f"  ✓ Found limitation sheet: '{sheet_name}'")
                return df
            except:
                continue
        
        return pd.DataFrame()
    
    def read_observations(self):
        """
        Read observations from ANY sheet structure
        Automatically detects columns based on content patterns
        """
        # Try common observation sheet names
        obs_sheet_names = ["Observations", "Issues", "Findings", "Vulnerabilities", "Bugs", "Results"]
        
        df = None
        used_sheet = None
        
        for sheet_name in obs_sheet_names:
            try:
                df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                used_sheet = sheet_name
                print(f"  ✓ Found observations sheet: '{sheet_name}'")
                break
            except:
                continue
        
        if df is None:
            # If no named sheet found, use first sheet
            df = pd.read_excel(self.excel_file, sheet_name=0)
            used_sheet = "first sheet"
            print(f"  ✓ Using first sheet for observations")
        
        # Clean column names
        original_columns = df.columns.tolist()
        df.columns = [str(col).strip() for col in df.columns]
        
        # Detect and map columns dynamically
        observations = self._map_observations_dynamically(df)
        
        print(f"  ✓ Found {len(observations)} observations")
        return observations
    
    def _map_observations_dynamically(self, df):
        """
        Dynamically map columns based on content patterns
        No hardcoded column names!
        """
        observations = []
        
        # Get all column names (lowercase for matching)
        columns = {col: str(col).lower().strip() for col in df.columns}
        
        # Detect column types based on header names and content
        col_mapping = self._detect_column_types(df, columns)
        
        # Convert each row to observation dict
        for idx, row in df.iterrows():
            obs = {}
            
            # Map each detected column
            for obs_field, excel_col in col_mapping.items():
                if excel_col and excel_col in df.columns:
                    value = row.get(excel_col, "")
                    obs[obs_field] = "" if pd.isna(value) else str(value)
                else:
                    obs[obs_field] = ""
            
            # Add row index for reference
            obs['_row_index'] = idx
            
            # Only add if there's at least some content
            if obs.get('description') or obs.get('title') or obs.get('finding'):
                observations.append(obs)
        
        return observations
    
    def _detect_column_types(self, df, columns):
        """
        Intelligently detect what each column represents
        Based on header names AND content patterns
        """
        mapping = {
            'sr_no': None,
            'severity': None,
            'title': None,
            'description': None,
            'affected_url': None,
            'cve': None,
            'poc': None,
            'recommendation': None,
            'impact': None,
            'status': None
        }
        
        # Common patterns for each field
        patterns = {
            'sr_no': ['sr', 's.no', 'sno', 'serial', 'no.', 'id', '#'],
            'severity': ['severity', 'risk', 'criticality', 'priority', 'level'],
            'title': ['title', 'name', 'vulnerability', 'finding', 'issue', 'heading'],
            'description': ['description', 'desc', 'details', 'detail', 'observation'],
            'affected_url': ['url', 'affected', 'endpoint', 'path', 'location', 'page'],
            'cve': ['cve', 'cwe', 'vulnerability id', 'id', 'reference'],
            'poc': ['poc', 'proof', 'screenshot', 'evidence', 'image', 'attachment'],
            'recommendation': ['recommendation', 'fix', 'remediation', 'solution', 'mitigation'],
            'impact': ['impact', 'effect', 'consequence'],
            'status': ['status', 'state']
        }
        
        # First pass: match by column name
        for col_name, col_lower in columns.items():
            for field, field_patterns in patterns.items():
                if any(pattern in col_lower for pattern in field_patterns):
                    if mapping[field] is None:  # Don't override if already set
                        mapping[field] = col_name
                        break
        
        # Second pass: check content for columns that weren't matched
        for field, field_patterns in patterns.items():
            if mapping[field] is None:
                # Look at first few rows to detect content
                for col_name in df.columns[:10]:  # Check first 10 columns
                    sample_values = df[col_name].dropna().astype(str).head(3).tolist()
                    sample_text = ' '.join(sample_values).lower()
                    
                    # Check if content matches field patterns
                    if field == 'severity' and any(word in sample_text for word in ['high', 'medium', 'low', 'critical']):
                        mapping[field] = col_name
                        break
                    elif field == 'cve' and any('cve-' in sample_text or 'cwe-' in sample_text for text in sample_values):
                        mapping[field] = col_name
                        break
                    elif field == 'url' and any('http' in text or 'https' in text or 'www.' in text for text in sample_values):
                        mapping[field] = col_name
                        break
        
        return mapping
    
    def get_cell_value(self, sheet, row, col, default=""):
        """Safely get cell value from any sheet"""
        try:
            if sheet.empty:
                return default
            val = sheet.iloc[row, col]
            return default if pd.isna(val) else str(val)
        except:
            return default

    def extract_report_name(self, index_df):
        """Extract report name from index sheet"""
        if index_df.empty:
            return os.path.splitext(os.path.basename(self.excel_file))[0]
        
        # Try common positions where report name might be
        possible_positions = [(5,1), (4,1), (3,1), (2,1), (1,1), (0,0)]
        
        for row, col in possible_positions:
            try:
                val = index_df.iloc[row, col]
                if pd.notna(val) and str(val).strip():
                    return str(val).strip()
            except:
                continue
        
        return os.path.splitext(os.path.basename(self.excel_file))[0]
    
    def extract_summary_stats(self, index_df):
        """Extract summary statistics from index sheet"""
        stats = {
            'high': '0',
            'medium': '0', 
            'low': '0',
            'total': '0',
            'initial_date': '',
            'app_url': ''
        }
        
        if index_df.empty:
            return stats
        
        # Try to find numbers in the sheet that look like severity counts
        for row in range(min(30, len(index_df))):
            for col in range(min(10, len(index_df.columns))):
                try:
                    val = str(index_df.iloc[row, col]).lower()
                    if 'high' in val or 'critical' in val:
                        # Check next cell for the number
                        if col + 1 < len(index_df.columns):
                            next_val = index_df.iloc[row, col + 1]
                            if pd.notna(next_val) and str(next_val).strip().isdigit():
                                stats['high'] = str(next_val).strip()
                    
                    elif 'medium' in val or 'moderate' in val:
                        if col + 1 < len(index_df.columns):
                            next_val = index_df.iloc[row, col + 1]
                            if pd.notna(next_val) and str(next_val).strip().isdigit():
                                stats['medium'] = str(next_val).strip()
                    
                    elif 'low' in val or 'minor' in val:
                        if col + 1 < len(index_df.columns):
                            next_val = index_df.iloc[row, col + 1]
                            if pd.notna(next_val) and str(next_val).strip().isdigit():
                                stats['low'] = str(next_val).strip()
                    
                    elif 'total' in val or 'overall' in val:
                        if col + 1 < len(index_df.columns):
                            next_val = index_df.iloc[row, col + 1]
                            if pd.notna(next_val) and str(next_val).strip().isdigit():
                                stats['total'] = str(next_val).strip()
                except:
                    continue
        
        return stats