"""
Performance optimizations for hodiny application
Implements caching, database optimizations, and general performance improvements
"""

import logging
import time
from datetime import datetime, timedelta
from functools import lru_cache, wraps
from typing import Any, Dict, Optional

from flask import g, session

logger = logging.getLogger(__name__)


class SimpleCache:
    """Simple in-memory cache with TTL support"""

    def __init__(self, default_ttl: int = 300):  # 5 minutes default
        self.cache = {}
        self.default_ttl = default_ttl

    def get(self, key: str) -> Optional[Any]:
        """Get value from cache if not expired"""
        if key in self.cache:
            value, expiry = self.cache[key]
            if datetime.now() < expiry:
                logger.debug(f"Cache hit for key: {key}")
                return value
            else:
                logger.debug(f"Cache expired for key: {key}")
                del self.cache[key]

        logger.debug(f"Cache miss for key: {key}")
        return None

    def set(self, key: str, value: Any, ttl: Optional[int] = None) -> None:
        """Set value in cache with TTL"""
        ttl = ttl or self.default_ttl
        expiry = datetime.now() + timedelta(seconds=ttl)
        self.cache[key] = (value, expiry)
        logger.debug(f"Cache set for key: {key}, TTL: {ttl}s")

    def delete(self, key: str) -> None:
        """Delete key from cache"""
        if key in self.cache:
            del self.cache[key]
            logger.debug(f"Cache deleted for key: {key}")

    def clear(self) -> None:
        """Clear all cache"""
        self.cache.clear()
        logger.debug("Cache cleared")

    def cleanup_expired(self) -> None:
        """Remove expired entries"""
        now = datetime.now()
        expired_keys = [key for key, (_, expiry) in self.cache.items() if now >= expiry]

        for key in expired_keys:
            del self.cache[key]

        if expired_keys:
            logger.debug(f"Cleaned up {len(expired_keys)} expired cache entries")


# Global cache instance
app_cache = SimpleCache()


def cache_result(ttl: int = 300, key_func=None):
    """Decorator to cache function results"""

    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            # Generate cache key
            if key_func:
                cache_key = key_func(*args, **kwargs)
            else:
                cache_key = f"{func.__name__}:{hash(str(args) + str(kwargs))}"

            # Try to get from cache
            result = app_cache.get(cache_key)
            if result is not None:
                return result

            # Execute function and cache result
            start_time = time.time()
            result = func(*args, **kwargs)
            execution_time = time.time() - start_time

            app_cache.set(cache_key, result, ttl)
            logger.debug(f"Function {func.__name__} executed in {execution_time:.3f}s and cached")

            return result

        return wrapper

    return decorator


def timing_decorator(func):
    """Decorator to log function execution time"""

    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        execution_time = time.time() - start_time

        if execution_time > 1.0:  # Log slow operations
            logger.warning(f"Slow operation: {func.__name__} took {execution_time:.3f}s")
        else:
            logger.debug(f"Operation: {func.__name__} took {execution_time:.3f}s")

        return result

    return wrapper


@lru_cache(maxsize=100)
def get_week_number_cached(date_str: str) -> int:
    """Cached version of week number calculation"""
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.isocalendar()[1]
    except ValueError:
        return datetime.now().isocalendar()[1]


@cache_result(ttl=60)  # Cache for 1 minute
def get_employee_stats():
    """Get cached employee statistics"""
    try:
        all_employees = g.employee_manager.get_all_employees()
        selected_employees = g.employee_manager.get_vybrani_zamestnanci()

        return {
            "total_employees": len(all_employees),
            "selected_employees": len(selected_employees),
            "selection_percentage": (len(selected_employees) / len(all_employees) * 100) if all_employees else 0,
        }
    except Exception as e:
        logger.error(f"Error getting employee stats: {e}")
        return {"total_employees": 0, "selected_employees": 0, "selection_percentage": 0}


@cache_result(ttl=300)  # Cache for 5 minutes
def get_excel_file_info():
    """Get cached Excel file information"""
    try:
        return {
            "exists": g.excel_manager.file_exists() if hasattr(g.excel_manager, "file_exists") else False,
            "filename": (
                g.excel_manager.get_active_filename() if hasattr(g.excel_manager, "get_active_filename") else "Unknown"
            ),
            "last_checked": datetime.now().isoformat(),
        }
    except Exception as e:
        logger.error(f"Error getting Excel file info: {e}")
        return {"exists": False, "filename": "Unknown", "last_checked": datetime.now().isoformat()}


def optimize_session_data():
    """Optimize session data by removing unnecessary keys"""
    if "settings" not in session:
        return

    # Remove old or unnecessary session data
    current_time = datetime.now()

    # Clean up temporary form data older than 1 hour
    if "temp_form_data" in session:
        temp_data = session.get("temp_form_data", {})
        if "timestamp" in temp_data:
            try:
                data_time = datetime.fromisoformat(temp_data["timestamp"])
                if current_time - data_time > timedelta(hours=1):
                    del session["temp_form_data"]
                    logger.debug("Cleaned up old temporary form data from session")
            except (ValueError, TypeError):
                del session["temp_form_data"]


class PerformanceMonitor:
    """Monitor and log performance metrics"""

    def __init__(self):
        self.request_times = []
        self.slow_requests = []

    def record_request(self, endpoint: str, duration: float):
        """Record request timing"""
        self.request_times.append({"endpoint": endpoint, "duration": duration, "timestamp": datetime.now()})

        # Keep only last 100 requests
        if len(self.request_times) > 100:
            self.request_times = self.request_times[-100:]

        # Track slow requests
        if duration > 2.0:  # Slow request threshold
            self.slow_requests.append({"endpoint": endpoint, "duration": duration, "timestamp": datetime.now()})
            logger.warning(f"Slow request detected: {endpoint} took {duration:.3f}s")

    def get_performance_stats(self) -> Dict[str, Any]:
        """Get performance statistics"""
        if not self.request_times:
            return {"avg_response_time": 0, "slow_requests": 0, "total_requests": 0}

        avg_time = sum(r["duration"] for r in self.request_times) / len(self.request_times)
        slow_count = len([r for r in self.request_times if r["duration"] > 2.0])

        return {
            "avg_response_time": round(avg_time, 3),
            "slow_requests": slow_count,
            "total_requests": len(self.request_times),
            "slow_request_percentage": round((slow_count / len(self.request_times)) * 100, 2),
        }


# Global performance monitor
perf_monitor = PerformanceMonitor()


def optimize_excel_operations():
    """Optimize Excel file operations"""
    try:
        # Only check Excel file existence once per request
        if not hasattr(g, "_excel_exists_cache"):
            g._excel_exists_cache = g.excel_manager.file_exists() if hasattr(g.excel_manager, "file_exists") else False

        return g._excel_exists_cache
    except Exception as e:
        logger.error(f"Error optimizing Excel operations: {e}")
        return False


def batch_employee_operations(operations: list):
    """Batch multiple employee operations for better performance"""
    try:
        results = []
        for operation in operations:
            func, args, kwargs = operation
            result = func(*args, **kwargs)
            results.append(result)

        logger.debug(f"Batched {len(operations)} employee operations")
        return results
    except Exception as e:
        logger.error(f"Error in batch employee operations: {e}")
        return []


def cleanup_old_data():
    """Cleanup old data and optimize performance"""
    try:
        # Cleanup expired cache entries
        app_cache.cleanup_expired()

        # Optimize session data
        optimize_session_data()

        # Clean up performance monitor data older than 1 hour
        cutoff_time = datetime.now() - timedelta(hours=1)
        perf_monitor.request_times = [r for r in perf_monitor.request_times if r["timestamp"] > cutoff_time]
        perf_monitor.slow_requests = [r for r in perf_monitor.slow_requests if r["timestamp"] > cutoff_time]

        logger.debug("Completed data cleanup and optimization")

    except Exception as e:
        logger.error(f"Error during data cleanup: {e}")


def get_system_performance_info() -> Dict[str, Any]:
    """Get system performance information"""
    try:
        import psutil

        memory_info = psutil.virtual_memory()
        cpu_percent = psutil.cpu_percent(interval=1)

        return {
            "memory_usage_percent": memory_info.percent,
            "memory_available_mb": memory_info.available // (1024 * 1024),
            "cpu_usage_percent": cpu_percent,
            "cache_size": len(app_cache.cache),
            "performance_stats": perf_monitor.get_performance_stats(),
        }
    except ImportError:
        # psutil not available
        return {
            "memory_usage_percent": "N/A",
            "memory_available_mb": "N/A",
            "cpu_usage_percent": "N/A",
            "cache_size": len(app_cache.cache),
            "performance_stats": perf_monitor.get_performance_stats(),
        }
    except Exception as e:
        logger.error(f"Error getting system performance info: {e}")
        return {
            "error": str(e),
            "cache_size": len(app_cache.cache),
            "performance_stats": perf_monitor.get_performance_stats(),
        }


# Helper functions for common optimizations
def memoize_user_settings(user_id: str = "default"):
    """Memoize user settings to avoid repeated file reads"""
    cache_key = f"user_settings:{user_id}"

    settings = app_cache.get(cache_key)
    if settings is None:
        # Load settings from file
        from app import load_settings_from_file

        settings = load_settings_from_file()
        app_cache.set(cache_key, settings, ttl=600)  # Cache for 10 minutes

    return settings


def invalidate_user_settings_cache(user_id: str = "default"):
    """Invalidate user settings cache when settings change"""
    cache_key = f"user_settings:{user_id}"
    app_cache.delete(cache_key)
    logger.debug(f"Invalidated settings cache for user: {user_id}")


# Startup optimization
def initialize_performance_optimizations():
    """Initialize performance optimizations on app startup"""
    logger.info("Initializing performance optimizations...")

    # Pre-warm cache with common operations
    try:
        get_employee_stats()
        get_excel_file_info()
        logger.info("Performance optimizations initialized successfully")
    except Exception as e:
        logger.error(f"Error initializing performance optimizations: {e}")
