"""
API endpoints for hodiny application
Provides structured REST API for better data handling and performance
"""

from flask import Blueprint, request, jsonify, g
from datetime import datetime
import logging
from typing import Dict, Any, Optional, List

# Configure logger
logger = logging.getLogger(__name__)

# Create API Blueprint
api_bp = Blueprint('api', __name__, url_prefix='/api/v1')


class APIResponse:
    """Standard API response format"""
    
    @staticmethod
    def success(data: Any = None, message: str = "Success", status_code: int = 200) -> tuple:
        """Return successful API response"""
        response = {
            "success": True,
            "message": message,
            "data": data,
            "timestamp": datetime.now().isoformat()
        }
        return jsonify(response), status_code
    
    @staticmethod
    def error(message: str, error_code: str = "GENERAL_ERROR", status_code: int = 400, details: Optional[Dict] = None) -> tuple:
        """Return error API response"""
        response = {
            "success": False,
            "error": {
                "message": message,
                "code": error_code,
                "details": details or {}
            },
            "timestamp": datetime.now().isoformat()
        }
        return jsonify(response), status_code


def validate_required_fields(data: Dict, required_fields: List[str]) -> Optional[tuple]:
    """Validate required fields in request data"""
    missing_fields = [field for field in required_fields if not data.get(field)]
    if missing_fields:
        return APIResponse.error(
            f"Missing required fields: {', '.join(missing_fields)}",
            "MISSING_FIELDS",
            400,
            {"missing_fields": missing_fields}
        )
    return None


def validate_date_format(date_string: str) -> bool:
    """Validate date format YYYY-MM-DD"""
    try:
        datetime.strptime(date_string, "%Y-%m-%d")
        return True
    except ValueError:
        return False


def validate_time_format(time_string: str) -> bool:
    """Validate time format HH:MM"""
    try:
        datetime.strptime(time_string, "%H:%M")
        return True
    except ValueError:
        return False


@api_bp.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    try:
        # Basic health checks
        health_status = {
            "status": "healthy",
            "version": "1.0.0",
            "timestamp": datetime.now().isoformat(),
            "services": {
                "database": "operational",  # Could be expanded to check actual DB
                "excel_manager": "operational",
                "employee_manager": "operational"
            }
        }
        
        return APIResponse.success(health_status, "Service is healthy")
    
    except Exception as e:
        logger.error("Health check failed: %s", e, exc_info=True)
        return APIResponse.error("Service health check failed", "HEALTH_CHECK_ERROR", 503)


@api_bp.route('/employees', methods=['GET'])
def get_employees():
    """Get all employees"""
    try:
        employees = g.employee_manager.get_all_employees()
        selected_employees = g.employee_manager.get_vybrani_zamestnanci()
        
        employee_data = []
        for emp in employees:
            employee_data.append({
                "name": emp,
                "selected": emp in selected_employees
            })
        
        return APIResponse.success(employee_data, f"Retrieved {len(employee_data)} employees")
    
    except Exception as e:
        logger.error("Error retrieving employees: %s", e, exc_info=True)
        return APIResponse.error("Failed to retrieve employees", "EMPLOYEE_RETRIEVAL_ERROR", 500)


@api_bp.route('/employees/selected', methods=['GET', 'POST'])
def manage_selected_employees():
    """Get or update selected employees"""
    if request.method == 'GET':
        try:
            selected = g.employee_manager.get_vybrani_zamestnanci()
            return APIResponse.success(selected, f"Retrieved {len(selected)} selected employees")
        
        except Exception as e:
            logger.error("Error retrieving selected employees: %s", e, exc_info=True)
            return APIResponse.error("Failed to retrieve selected employees", "SELECTED_EMPLOYEES_ERROR", 500)
    
    elif request.method == 'POST':
        try:
            data = request.get_json()
            if not data:
                return APIResponse.error("No data provided", "NO_DATA", 400)
            
            employees = data.get('employees', [])
            if not isinstance(employees, list):
                return APIResponse.error("Employees must be a list", "INVALID_DATA_TYPE", 400)
            
            # Update selected employees
            g.employee_manager.set_vybrani_zamestnanci(employees)
            
            return APIResponse.success(employees, f"Updated selected employees: {len(employees)} selected")
        
        except Exception as e:
            logger.error("Error updating selected employees: %s", e, exc_info=True)
            return APIResponse.error("Failed to update selected employees", "UPDATE_EMPLOYEES_ERROR", 500)


@api_bp.route('/time-entry', methods=['POST'])
def create_time_entry():
    """Create a new time entry with enhanced validation"""
    try:
        data = request.get_json()
        if not data:
            return APIResponse.error("No data provided", "NO_DATA", 400)
        
        # Validate required fields
        required_fields = ['date']
        validation_error = validate_required_fields(data, required_fields)
        if validation_error:
            return validation_error
        
        # Extract and validate data
        date = data.get('date')
        start_time = data.get('start_time')
        end_time = data.get('end_time')
        lunch_duration = data.get('lunch_duration', '1.0')
        is_free_day = data.get('is_free_day', False)
        notes = data.get('notes', '')
        
        # Validate date format
        if not validate_date_format(date):
            return APIResponse.error("Invalid date format. Use YYYY-MM-DD", "INVALID_DATE_FORMAT", 400)
        
        # Get selected employees
        selected_employees = g.employee_manager.get_vybrani_zamestnanci()
        if not selected_employees:
            return APIResponse.error("No employees selected", "NO_EMPLOYEES_SELECTED", 400)
        
        # Validate time fields for work days
        if not is_free_day:
            if not start_time or not end_time:
                return APIResponse.error("Start time and end time required for work days", "MISSING_TIME", 400)
            
            if not validate_time_format(start_time) or not validate_time_format(end_time):
                return APIResponse.error("Invalid time format. Use HH:MM", "INVALID_TIME_FORMAT", 400)
            
            # Validate lunch duration
            try:
                lunch_float = float(lunch_duration)
                if lunch_float < 0 or lunch_float > 8:
                    return APIResponse.error("Lunch duration must be between 0 and 8 hours", "INVALID_LUNCH_DURATION", 400)
            except ValueError:
                return APIResponse.error("Invalid lunch duration format", "INVALID_LUNCH_FORMAT", 400)
        
        # Process the time entry
        if is_free_day:
            g.excel_manager.ulozit_pracovni_dobu(
                date, "00:00", "00:00", "0", selected_employees
            )
            g.hodiny2025_manager.zapis_pracovni_doby(
                date, "00:00", "00:00", "0", len(selected_employees)
            )
            message = f"Free day recorded for {date} ({len(selected_employees)} employees)"
        else:
            g.excel_manager.ulozit_pracovni_dobu(
                date, start_time, end_time, lunch_duration, selected_employees
            )
            g.hodiny2025_manager.zapis_pracovni_doby(
                date, start_time, end_time, lunch_duration, len(selected_employees)
            )
            message = f"Work time recorded for {date} ({len(selected_employees)} employees)"
        
        # Prepare response data
        response_data = {
            "date": date,
            "start_time": start_time if not is_free_day else None,
            "end_time": end_time if not is_free_day else None,
            "lunch_duration": lunch_duration if not is_free_day else None,
            "is_free_day": is_free_day,
            "employees_count": len(selected_employees),
            "notes": notes
        }
        
        return APIResponse.success(response_data, message)
    
    except Exception as e:
        logger.error("Error creating time entry: %s", e, exc_info=True)
        return APIResponse.error("Failed to create time entry", "TIME_ENTRY_ERROR", 500)


@api_bp.route('/time-entries', methods=['GET'])
def get_time_entries():
    """Get time entries with optional filtering"""
    try:
        # Get query parameters
        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')
        employee = request.args.get('employee')
        week_number = request.args.get('week')
        
        # Validate date parameters
        if start_date and not validate_date_format(start_date):
            return APIResponse.error("Invalid start_date format. Use YYYY-MM-DD", "INVALID_DATE_FORMAT", 400)
        
        if end_date and not validate_date_format(end_date):
            return APIResponse.error("Invalid end_date format. Use YYYY-MM-DD", "INVALID_DATE_FORMAT", 400)
        
        # Get data based on filters
        if week_number:
            try:
                week_num = int(week_number)
                week_data = g.excel_manager.get_current_week_data(week_num)
                if week_data:
                    return APIResponse.success(week_data, f"Retrieved data for week {week_num}")
                else:
                    return APIResponse.success([], f"No data found for week {week_num}")
            except ValueError:
                return APIResponse.error("Invalid week number", "INVALID_WEEK_NUMBER", 400)
        
        # For now, return current week data as default
        current_week_data = g.excel_manager.get_current_week_data()
        
        return APIResponse.success(current_week_data or [], "Retrieved current week time entries")
    
    except Exception as e:
        logger.error("Error retrieving time entries: %s", e, exc_info=True)
        return APIResponse.error("Failed to retrieve time entries", "TIME_ENTRIES_ERROR", 500)


@api_bp.route('/excel/status', methods=['GET'])
def get_excel_status():
    """Get Excel file status and information"""
    try:
        # Get Excel status information
        active_filename = g.excel_manager.get_active_filename() if hasattr(g.excel_manager, 'get_active_filename') else "Unknown"
        excel_exists = g.excel_manager.file_exists() if hasattr(g.excel_manager, 'file_exists') else False
        
        status_data = {
            "active_filename": active_filename,
            "excel_exists": excel_exists,
            "timestamp": datetime.now().isoformat()
        }
        
        return APIResponse.success(status_data, "Excel status retrieved")
    
    except Exception as e:
        logger.error("Error getting Excel status: %s", e, exc_info=True)
        return APIResponse.error("Failed to get Excel status", "EXCEL_STATUS_ERROR", 500)


@api_bp.route('/settings', methods=['GET', 'POST'])
def manage_settings():
    """Get or update application settings"""
    if request.method == 'GET':
        try:
            # Get current settings from session or defaults
            from flask import session
            settings = session.get('settings', {
                'start_time': '07:00',
                'end_time': '18:00',
                'lunch_duration': 1.0,
                'theme': 'light'
            })
            
            return APIResponse.success(settings, "Settings retrieved")
        
        except Exception as e:
            logger.error("Error retrieving settings: %s", e, exc_info=True)
            return APIResponse.error("Failed to retrieve settings", "SETTINGS_ERROR", 500)
    
    elif request.method == 'POST':
        try:
            data = request.get_json()
            if not data:
                return APIResponse.error("No data provided", "NO_DATA", 400)
            
            # Validate settings
            from flask import session
            current_settings = session.get('settings', {})
            
            # Update settings
            for key, value in data.items():
                if key in ['start_time', 'end_time'] and not validate_time_format(value):
                    return APIResponse.error(f"Invalid time format for {key}", "INVALID_TIME_FORMAT", 400)
                current_settings[key] = value
            
            session['settings'] = current_settings
            
            return APIResponse.success(current_settings, "Settings updated successfully")
        
        except Exception as e:
            logger.error("Error updating settings: %s", e, exc_info=True)
            return APIResponse.error("Failed to update settings", "SETTINGS_UPDATE_ERROR", 500)


# Error handlers for the API blueprint
@api_bp.errorhandler(404)
def api_not_found(error):
    """Handle 404 errors for API endpoints"""
    return APIResponse.error("API endpoint not found", "ENDPOINT_NOT_FOUND", 404)


@api_bp.errorhandler(405)
def api_method_not_allowed(error):
    """Handle 405 errors for API endpoints"""
    return APIResponse.error("Method not allowed for this endpoint", "METHOD_NOT_ALLOWED", 405)


@api_bp.errorhandler(500)
def api_internal_error(error):
    """Handle 500 errors for API endpoints"""
    logger.error("Internal API error: %s", error, exc_info=True)
    return APIResponse.error("Internal server error", "INTERNAL_ERROR", 500)