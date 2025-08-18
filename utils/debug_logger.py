import logging
import json
from datetime import datetime
import traceback
import zipfile
import os

# Set up detailed logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

class DebugLogger:
    def __init__(self, name="TeleiosDebug"):
        self.logger = logging.getLogger(name)
        self.extraction_trace = []
        
    def reset_trace(self):
        """Reset the extraction trace for a new operation"""
        self.extraction_trace = []
        
    def log_file_info(self, file_obj):
        """Log detailed file information"""
        info = {
            "timestamp": datetime.now().isoformat(),
            "filename": getattr(file_obj, 'filename', 'Unknown'),
            "file_size": None,
            "is_zip_file": False,
            "zip_test": None,
            "file_type": None
        }
        
        try:
            # Try to get file size
            file_obj.seek(0, 2)  # Seek to end
            info["file_size"] = file_obj.tell()
            file_obj.seek(0)  # Reset to beginning
            
            # Test if it's a valid zip file (Excel files are zip archives)
            try:
                with zipfile.ZipFile(file_obj, 'r') as zf:
                    info["is_zip_file"] = True
                    info["zip_contents"] = zf.namelist()[:10]  # First 10 entries
                file_obj.seek(0)  # Reset after zip test
            except zipfile.BadZipFile as e:
                info["zip_test"] = f"BadZipFile: {str(e)}"
                # Check first bytes to determine file type
                file_obj.seek(0)
                first_bytes = file_obj.read(8)
                file_obj.seek(0)
                
                # Check for Excel file signatures
                if first_bytes[:4] == b'PK\x03\x04':
                    info["file_type"] = "Looks like ZIP/XLSX but corrupted"
                elif first_bytes[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                    info["file_type"] = "Old Excel format (XLS)"
                else:
                    info["file_type"] = f"Unknown format. First bytes: {first_bytes.hex()}"
                    
        except Exception as e:
            info["error"] = str(e)
            info["traceback"] = traceback.format_exc()
            
        self.logger.info(f"File info: {json.dumps(info, indent=2)}")
        self.extraction_trace.append({"step": "file_info", "data": info})
        return info
        
    def log_sheet_detection(self, workbook, sheet_names):
        """Log sheet detection process"""
        detection_info = {
            "timestamp": datetime.now().isoformat(),
            "all_sheets": sheet_names,
            "county_trend_candidates": [],
            "selected_sheet": None
        }
        
        for sheet_name in sheet_names:
            if 'trend' in sheet_name.lower() and 'county' in sheet_name.lower():
                detection_info["county_trend_candidates"].append(sheet_name)
                if not detection_info["selected_sheet"]:
                    detection_info["selected_sheet"] = sheet_name
                    
        self.logger.info(f"Sheet detection: {json.dumps(detection_info, indent=2)}")
        self.extraction_trace.append({"step": "sheet_detection", "data": detection_info})
        return detection_info
        
    def log_row_extraction(self, row_num, year_val, medicare_val, action, reason=""):
        """Log each row extraction decision"""
        row_info = {
            "row": row_num,
            "year_value": year_val,
            "medicare_value": medicare_val,
            "action": action,  # "extract", "skip", "stop"
            "reason": reason
        }
        
        self.logger.debug(f"Row {row_num}: {action} - {reason}")
        self.extraction_trace.append({"step": f"row_{row_num}", "data": row_info})
        return row_info
        
    def log_extraction_summary(self, county_name, total_rows_checked, rows_extracted, stop_reason):
        """Log extraction summary"""
        summary = {
            "timestamp": datetime.now().isoformat(),
            "county": county_name,
            "total_rows_checked": total_rows_checked,
            "rows_extracted": rows_extracted,
            "stop_reason": stop_reason
        }
        
        self.logger.info(f"Extraction summary: {json.dumps(summary, indent=2)}")
        self.extraction_trace.append({"step": "summary", "data": summary})
        return summary
        
    def get_full_trace(self):
        """Get the complete extraction trace"""
        return {
            "trace": self.extraction_trace,
            "trace_count": len(self.extraction_trace)
        }
        
    def save_trace_to_file(self, filename):
        """Save trace to a JSON file for analysis"""
        filepath = f"debug_traces/{filename}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        os.makedirs("debug_traces", exist_ok=True)
        
        with open(filepath, 'w') as f:
            json.dump(self.get_full_trace(), f, indent=2)
            
        self.logger.info(f"Trace saved to {filepath}")
        return filepath

# Global logger instance
debug_logger = DebugLogger()