"""
API endpoints for hodiny application
Provides structured REST API for better data handling and performance
"""

import logging
from datetime import datetime
from typing import Any, Dict, List, Optional

from flask import Blueprint, g, jsonify, request, session

from performance_optimizations import invalidate_employee_stats_cache, invalidate_user_settings_cache
from services.api_service import (
    create_time_entry as create_time_entry_payload,
    filter_time_entries_by_employee,
    get_excel_status as get_excel_status_payload,
    get_settings,
    get_time_entries as get_time_entries_payload,
    serialize_employees,
    update_selected_employees,
    update_settings,
)

# Configure logger
logger = logging.getLogger(__name__)

# Create API Blueprint
api_bp = Blueprint("api", __name__, url_prefix="/api/v1")


class APIResponse:
    """Standard API response format"""

    @staticmethod
    def success(data: Any = None, message: str = "Success", status_code: int = 200) -> tuple:
        """Return successful API response"""
        response = {"success": True, "message": message, "data": data, "timestamp": datetime.now().isoformat()}
        return jsonify(response), status_code

    @staticmethod
    def error(
        message: str, error_code: str = "GENERAL_ERROR", status_code: int = 400, details: Optional[Dict] = None
    ) -> tuple:
        """Return error API response"""
        response = {
            "success": False,
            "error": {"message": message, "code": error_code, "details": details or {}},
            "timestamp": datetime.now().isoformat(),
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
            {"missing_fields": missing_fields},
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


@api_bp.route("/health", methods=["GET"])
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
                "employee_manager": "operational",
            },
        }

        return APIResponse.success(health_status, "Service is healthy")

    except Exception as e:
        logger.error("Health check failed: %s", e, exc_info=True)
        return APIResponse.error("Service health check failed", "HEALTH_CHECK_ERROR", 503)


@api_bp.route("/employees", methods=["GET"])
def get_employees():
    """Get all employees"""
    try:
        employee_data = serialize_employees(g.employee_manager)
        return APIResponse.success(employee_data, f"Retrieved {len(employee_data)} employees")

    except Exception as e:
        logger.error("Error retrieving employees: %s", e, exc_info=True)
        return APIResponse.error("Failed to retrieve employees", "EMPLOYEE_RETRIEVAL_ERROR", 500)


@api_bp.route("/employees/selected", methods=["GET", "POST"])
def manage_selected_employees():
    """Get or update selected employees"""
    if request.method == "GET":
        try:
            selected = g.employee_manager.get_vybrani_zamestnanci()
            return APIResponse.success(selected, f"Retrieved {len(selected)} selected employees")

        except Exception as e:
            logger.error("Error retrieving selected employees: %s", e, exc_info=True)
            return APIResponse.error("Failed to retrieve selected employees", "SELECTED_EMPLOYEES_ERROR", 500)

    elif request.method == "POST":
        try:
            data = request.get_json()
            if not data:
                return APIResponse.error("No data provided", "NO_DATA", 400)

            employees = data.get("employees", [])
            if not isinstance(employees, list):
                return APIResponse.error("Employees must be a list", "INVALID_DATA_TYPE", 400)

            selected = update_selected_employees(g.employee_manager, employees)
            invalidate_employee_stats_cache()
            return APIResponse.success(selected, f"Updated selected employees: {len(selected)} selected")

        except Exception as e:
            logger.error("Error updating selected employees: %s", e, exc_info=True)
            return APIResponse.error("Failed to update selected employees", "UPDATE_EMPLOYEES_ERROR", 500)


@api_bp.route("/time-entry", methods=["POST"])
def create_time_entry():
    """Create a new time entry with enhanced validation"""
    try:
        data = request.get_json()
        if not data:
            return APIResponse.error("No data provided", "NO_DATA", 400)

        # Validate required fields
        required_fields = ["date"]
        validation_error = validate_required_fields(data, required_fields)
        if validation_error:
            return validation_error

        # Extract and validate data
        date = data.get("date")
        start_time = data.get("start_time")
        end_time = data.get("end_time")
        lunch_duration = data.get("lunch_duration", "1.0")
        is_free_day = data.get("is_free_day", False)

        # Validate date format
        if not validate_date_format(date):
            return APIResponse.error("Invalid date format. Use YYYY-MM-DD", "INVALID_DATE_FORMAT", 400)

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
                    return APIResponse.error(
                        "Lunch duration must be between 0 and 8 hours", "INVALID_LUNCH_DURATION", 400
                    )
            except ValueError:
                return APIResponse.error("Invalid lunch duration format", "INVALID_LUNCH_FORMAT", 400)

        response_data = create_time_entry_payload(data, g.employee_manager, g.excel_manager, g.hodiny2025_manager)
        return APIResponse.success(response_data, response_data["message"])

    except ValueError as e:
        if str(e) == "No employees selected":
            return APIResponse.error("No employees selected", "NO_EMPLOYEES_SELECTED", 400)
        return APIResponse.error(str(e), "TIME_ENTRY_VALIDATION_ERROR", 400)

    except Exception as e:
        logger.error("Error creating time entry: %s", e, exc_info=True)
        return APIResponse.error("Failed to create time entry", "TIME_ENTRY_ERROR", 500)


@api_bp.route("/time-entries", methods=["GET"])
def get_time_entries():
    """Get time entries with optional filtering"""
    try:
        # Get query parameters
        start_date = request.args.get("start_date")
        end_date = request.args.get("end_date")
        employee_filter = request.args.get("employee")
        week_number = request.args.get("week")

        # Validate date parameters
        if start_date and not validate_date_format(start_date):
            return APIResponse.error("Invalid start_date format. Use YYYY-MM-DD", "INVALID_DATE_FORMAT", 400)

        if end_date and not validate_date_format(end_date):
            return APIResponse.error("Invalid end_date format. Use YYYY-MM-DD", "INVALID_DATE_FORMAT", 400)

        # Get data based on filters
        if week_number:
            try:
                week_num = int(week_number)
                week_data = get_time_entries_payload(g.excel_manager, week_num)
                if week_data:
                    week_data = filter_time_entries_by_employee(week_data, employee_filter)
                    return APIResponse.success(week_data, f"Retrieved data for week {week_num}")
                else:
                    return APIResponse.success([], f"No data found for week {week_num}")
            except ValueError:
                return APIResponse.error("Invalid week number", "INVALID_WEEK_NUMBER", 400)

        # For now, return current week data as default
        current_week_data = get_time_entries_payload(g.excel_manager)
        current_week_data = filter_time_entries_by_employee(current_week_data, employee_filter)

        return APIResponse.success(current_week_data or [], "Retrieved current week time entries")

    except Exception as e:
        logger.error("Error retrieving time entries: %s", e, exc_info=True)
        return APIResponse.error("Failed to retrieve time entries", "TIME_ENTRIES_ERROR", 500)


@api_bp.route("/excel/status", methods=["GET"])
def get_excel_status():
    """Get Excel file status and information"""
    try:
        status_data = get_excel_status_payload()
        return APIResponse.success(status_data, "Excel status retrieved")

    except Exception as e:
        logger.error("Error getting Excel status: %s", e, exc_info=True)
        return APIResponse.error("Failed to get Excel status", "EXCEL_STATUS_ERROR", 500)


@api_bp.route("/settings", methods=["GET", "POST"])
def manage_settings():
    """Get or update application settings"""
    if request.method == "GET":
        try:
            settings = get_settings()
            session["settings"] = settings
            return APIResponse.success(settings, "Settings retrieved")

        except Exception as e:
            logger.error("Error retrieving settings: %s", e, exc_info=True)
            return APIResponse.error("Failed to retrieve settings", "SETTINGS_ERROR", 500)

    elif request.method == "POST":
        try:
            data = request.get_json()
            if not data:
                return APIResponse.error("No data provided", "NO_DATA", 400)

            current_settings = session.get("settings", {})

            for key, value in data.items():
                if key in ["start_time", "end_time"] and not validate_time_format(value):
                    return APIResponse.error(f"Invalid time format for {key}", "INVALID_TIME_FORMAT", 400)
            updated_settings = update_settings(current_settings, data)
            session["settings"] = updated_settings
            invalidate_user_settings_cache()
            return APIResponse.success(updated_settings, "Settings updated successfully")

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
