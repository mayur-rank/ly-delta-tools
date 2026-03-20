try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("\n[!] ERROR: Missing dependencies (pandas/numpy).")
    print("[!] Please run: venv\\Scripts\\python.exe -m pip install -r requirements.txt\n")
    exit(1)

import os
import glob
from datetime import datetime

class PremiumPatternAnalyzer:
    def __init__(self, folder_path, threshold_percentile=10):
        self.folder_path = folder_path
        self.threshold_percentile = threshold_percentile
        self.master_df = None

    def load_data(self):
        """Recursively scan folder for Excel files and sheets."""
        all_data = []
        excel_files = glob.glob(os.path.join(self.folder_path, "**/*.xlsx"), recursive=True)
        
        print(f"Found {len(excel_files)} Excel files. Loading sheets...")
        
        for file in excel_files:
            try:
                # Load all sheets in the file
                xl = pd.ExcelFile(file)
                for sheet_name in xl.sheet_names:
                    # In your format, metadata is Row 1, Header is Row 3 (skip index 0,1)
                    df = pd.read_excel(file, sheet_name=sheet_name, skiprows=2)
                    
                    # If not found at Row 3 (skiprows 2), try without skip (Row 1)
                    if "DateTime" not in df.columns:
                        df = pd.read_excel(file, sheet_name=sheet_name)
                    
                    # Basic validation: must have DateTime and Difference
                    if "DateTime" in df.columns and "Difference" in df.columns:
                        print(f"DEBUG: Successfully loaded sheet '{sheet_name}' from {os.path.basename(file)}")
                        print(f"DEBUG: Columns found: {list(df.columns)}")
                        print(f"DEBUG: Sample DateTime: {df['DateTime'].iloc[0] if not df.empty else 'EMPTY'}")
                        
                        df['Sheet'] = sheet_name
                        df['File'] = os.path.basename(file)
                        
                        # Use flexible time parsing
                        # We try to force it to just the time component
                        def parse_to_time(val):
                            if pd.isna(val): return None
                            if isinstance(val, (datetime, pd.Timestamp)):
                                return val.time()
                            try:
                                # Try as string
                                return pd.to_datetime(str(val)).time()
                            except:
                                try:
                                    # Try special format
                                    return pd.to_datetime(str(val), format='%I:%M:%S %p').time()
                                except:
                                    return None

                        df['TimeObj'] = df['DateTime'].apply(parse_to_time)
                        
                        # Filter out rows where TimeObj couldn't be parsed
                        invalid_times = df['TimeObj'].isna().sum()
                        if invalid_times > 0:
                            print(f"DEBUG: WARNING! {invalid_times} rows had unparseable Time in {sheet_name}. Sample: {df['DateTime'].iloc[0]}")
                            df = df[df['TimeObj'].notna()]

                        # Add Weekday if we can infer date from sheet name (YYYY-MM-DD)
                        try:
                            date_obj = datetime.strptime(sheet_name, '%Y-%m-%d')
                            df['Weekday'] = date_obj.strftime('%A')
                        except:
                            df['Weekday'] = "Unknown"
                            
                        all_data.append(df)
                    else:
                        print(f"DEBUG: Sheet '{sheet_name}' missed required columns. Found: {list(df.columns)}")
            except Exception as e:
                print(f"Error loading {file}: {e}")
                
        if all_data:
            self.master_df = pd.concat(all_data, ignore_index=True)
            print(f"Loaded {len(self.master_df)} total rows.")
            return True
        return False

    def find_patterns(self):
        """Identify significant drops and group them by time/weekday."""
        if self.master_df is None or self.master_df.empty:
            print("DEBUG: master_df is empty, cannot find patterns.")
            return None, 0
            
        # 1. Dynamic Thresholding (Outlier detection)
        # We look for the "Difference" values that are in the bottom X percentile
        valid_diffs = self.master_df[self.master_df['Difference'].notnull()]['Difference']
        if valid_diffs.empty:
            print("DEBUG: No valid 'Difference' data found.")
            return None, 0
            
        threshold = np.percentile(valid_diffs, self.threshold_percentile)
        
        # Robustness: If many values are 0, threshold might be 0. 
        # We want ONLY those that are actually dropping significantly.
        # If threshold is >= 0, we take the bottom 5% of NEGATIVE values instead.
        if threshold >= 0:
            negative_diffs = valid_diffs[valid_diffs < 0]
            if not negative_diffs.empty:
                threshold = np.percentile(negative_diffs, 50) # Take median of drops
                print(f"DEBUG: Threshold was >= 0, adjusted to median of drops: {threshold:.4f}")

        print(f"DEBUG: Data Range {valid_diffs.min()} to {valid_diffs.max()}")
        print(f"DEBUG: Calculated final threshold: {threshold:.4f}")
        
        # 2. Filter for "Events"
        events = self.master_df[self.master_df['Difference'] <= threshold].copy()
        print(f"DEBUG: Found {len(events)} event rows below threshold.")
        
        if events.empty:
            return None, threshold

        # 3. Grouping Logic
        # Group by Time only to find "Daily Patterns" (e.g., every day at 3:25 PM)
        patterns = events.groupby(['TimeObj']).agg({
            'Difference': ['count', 'mean', 'min'],
            'Premium': 'mean',
            'Weekday': lambda x: ", ".join(sorted(list(set(x)))) # Show which days it happened
        }).reset_index()
        
        # Flatten columns
        patterns.columns = ['Time', 'Occurrences', 'Avg_Drop', 'Max_Drop', 'Avg_Premium', 'Weekdays']
        
        # Sort by Occurrences (Frequency) AND then by Max_Drop (Magnitude)
        patterns = patterns.sort_values(by=['Occurrences', 'Max_Drop'], ascending=[False, True])
        
        print(f"DEBUG: Grouped into {len(patterns)} unique time patterns.")
        return patterns, threshold

    def generate_report(self, output_file="Pattern_Report.xlsx"):
        """Export patterns and summary stats to Excel."""
        patterns, threshold = self.find_patterns()
        if patterns is None:
            print("No patterns found.")
            return
            
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Sheet 1: Top Patterns
            patterns.to_excel(writer, sheet_name='Top_Patterns', index=False)
            
            # Sheet 2: Weekday Analysis 
            # (We expand the comma-separated 'Weekdays' back into individual rows to count)
            try:
                wb_expanded = patterns.assign(Weekdays=patterns['Weekdays'].str.split(', ')).explode('Weekdays')
                weekday_summary = wb_expanded.groupby('Weekdays')['Occurrences'].sum().sort_values(ascending=False)
                weekday_summary.to_excel(writer, sheet_name='Weekday_Analysis')
            except Exception as e:
                print(f"DEBUG: Weekday summary failed: {e}")
            
            # Sheet 3: Time Clustering
            time_summary = patterns.groupby('Time')['Occurrences'].sum().sort_values(ascending=False)
            time_summary.to_excel(writer, sheet_name='Time_Clustering')
            
        print(f"Report saved to {output_file}")
        return patterns

if __name__ == "__main__":
    # Example usage (can be pointed to v_test_data)
    analyzer = PremiumPatternAnalyzer(folder_path="v_test_data")
    if analyzer.load_data():
        p, t = analyzer.find_patterns()
        print("\nTOP PATTERNS IDENTIFIED:")
        print(p.head(20))
        analyzer.generate_report()
