"""
NITDA BYOD System - Configuration Management Module
Provides centralized configuration with validation and fallbacks
"""
import os
from typing import Optional, Dict, Any
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


class Config:
    """Configuration class with validation and defaults"""

    # =========================================================================
    # SUPABASE CONFIGURATION
    # =========================================================================
    SUPABASE_URL: str = os.getenv("SUPABASE_URL", "").strip()
    SUPABASE_KEY: str = os.getenv("SUPABASE_KEY", "").strip()
    
    # Validate Supabase credentials
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise ValueError(
            "CRITICAL: SUPABASE_URL and SUPABASE_KEY must be set in .env file. "
            "These are required to connect to the database."
        )

    # =========================================================================
    # EMAIL CONFIGURATION
    # =========================================================================
    SMTP_SERVER: str = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    SMTP_PORT: int = int(os.getenv("SMTP_PORT", "587"))
    SENDER_EMAIL: str = os.getenv("SENDER_EMAIL", "").strip()
    SENDER_PASSWORD: str = os.getenv("SENDER_PASSWORD", "").strip()
    
    # Validate email configuration
    if not SENDER_EMAIL:
        raise ValueError("SENDER_EMAIL must be set in .env file")
    if not SENDER_PASSWORD:
        raise ValueError("SENDER_PASSWORD must be set in .env file (use Google App Password)")

    # =========================================================================
    # SERVER CONFIGURATION
    # =========================================================================
    APPROVAL_SERVER_URL: str = os.getenv("APPROVAL_SERVER_URL", "http://localhost:5000")
    APPROVAL_SERVER_HOST: str = os.getenv("APPROVAL_SERVER_HOST", "0.0.0.0")
    APPROVAL_SERVER_PORT: int = int(os.getenv("APPROVAL_SERVER_PORT", "5000"))
    APPROVAL_SERVER_DEBUG: bool = os.getenv("APPROVAL_SERVER_DEBUG", "False").lower() == "true"

    # =========================================================================
    # AUTOMATION CONFIGURATION
    # =========================================================================
    CHECK_INTERVAL: int = int(os.getenv("CHECK_INTERVAL", "30"))  # seconds
    INSPECTION_LEAD_TIME: int = int(os.getenv("INSPECTION_LEAD_TIME", "2"))  # days
    MAX_RETRIES: int = int(os.getenv("MAX_RETRIES", "3"))
    RETRY_DELAY: int = int(os.getenv("RETRY_DELAY", "5"))  # seconds

    # =========================================================================
    # LOGGING CONFIGURATION
    # =========================================================================
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO")
    LOG_FILE: str = os.getenv("LOG_FILE", "logs/byod_system.log")
    LOG_MAX_SIZE: int = int(os.getenv("LOG_MAX_SIZE", "10485760"))  # 10MB
    LOG_BACKUP_COUNT: int = int(os.getenv("LOG_BACKUP_COUNT", "5"))

    # =========================================================================
    # SECURITY CONFIGURATION
    # =========================================================================
    ENABLE_RLS: bool = os.getenv("ENABLE_RLS", "True").lower() == "true"
    ALLOWED_IPS: list = os.getenv("ALLOWED_IPS", "").split(",") if os.getenv("ALLOWED_IPS") else []
    SESSION_TIMEOUT: int = int(os.getenv("SESSION_TIMEOUT", "3600"))  # seconds

    # =========================================================================
    # QR CODE CONFIGURATION
    # =========================================================================
    QR_CODE_VERSION: int = int(os.getenv("QR_CODE_VERSION", "1"))
    QR_CODE_BOX_SIZE: int = int(os.getenv("QR_CODE_BOX_SIZE", "10"))
    QR_CODE_BORDER: int = int(os.getenv("QR_CODE_BORDER", "4"))
    QR_STORAGE_PATH: str = os.getenv("QR_STORAGE_PATH", "qr_codes/")

    # =========================================================================
    # ORGANIZATION CONFIGURATION
    # =========================================================================
    ORG_NAME: str = os.getenv("ORG_NAME", "NITDA")
    IT_DEPARTMENT_EMAIL: str = os.getenv("IT_DEPARTMENT_EMAIL", "").strip()
    SUPPORT_EMAIL: str = os.getenv("SUPPORT_EMAIL", "").strip()

    # =========================================================================
    # OPTIONAL FEATURES
    # =========================================================================
    ENABLE_GOOGLE_FORMS_SYNC: bool = os.getenv("ENABLE_GOOGLE_FORMS_SYNC", "False").lower() == "true"
    ENABLE_AUDIT_LOGGING: bool = os.getenv("ENABLE_AUDIT_LOGGING", "True").lower() == "true"
    ENABLE_EMAIL_NOTIFICATIONS: bool = os.getenv("ENABLE_EMAIL_NOTIFICATIONS", "True").lower() == "true"

    @classmethod
    def to_dict(cls) -> Dict[str, Any]:
        """Return configuration as dictionary (excluding sensitive data)"""
        return {
            "supabase_url": cls.SUPABASE_URL[:50] + "..." if cls.SUPABASE_URL else "NOT SET",
            "smtp_server": cls.SMTP_SERVER,
            "smtp_port": cls.SMTP_PORT,
            "sender_email": cls.SENDER_EMAIL,
            "approval_server_url": cls.APPROVAL_SERVER_URL,
            "check_interval": cls.CHECK_INTERVAL,
            "log_level": cls.LOG_LEVEL,
            "org_name": cls.ORG_NAME,
        }

    @classmethod
    def validate(cls) -> tuple[bool, list[str]]:
        """Validate all configuration parameters. Returns (is_valid, error_list)"""
        errors = []

        # Check Supabase
        if not cls.SUPABASE_URL.startswith("https://"):
            errors.append("SUPABASE_URL must start with https://")
        if len(cls.SUPABASE_KEY) < 50:
            errors.append("SUPABASE_KEY appears to be invalid (too short)")

        # Check email
        if "@" not in cls.SENDER_EMAIL:
            errors.append("SENDER_EMAIL must be a valid email address")
        if cls.SMTP_PORT not in [587, 465, 25]:
            errors.append("SMTP_PORT should be 25, 465, or 587")

        # Check intervals
        if cls.CHECK_INTERVAL < 10:
            errors.append("CHECK_INTERVAL should be at least 10 seconds")
        if cls.INSPECTION_LEAD_TIME < 1:
            errors.append("INSPECTION_LEAD_TIME must be at least 1 day")

        return len(errors) == 0, errors


# Initialize config at import time
def initialize_config() -> None:
    """Initialize and validate configuration on startup"""
    is_valid, errors = Config.validate()
    if not is_valid:
        print("CONFIGURATION WARNINGS:")
        for error in errors:
            print(f"   - {error}")


# Auto-initialize on import
initialize_config()
