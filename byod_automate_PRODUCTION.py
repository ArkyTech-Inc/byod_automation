"""
NITDA BYOD Automation Daemon - Production Ready
Background process that runs automation workflows at regular intervals
"""
import os
import sys
import time
import logging
import signal
from pathlib import Path
from datetime import datetime
from logging.handlers import RotatingFileHandler

from dotenv import load_dotenv

# Load environment variables
load_dotenv()

from config import Config
from byod_automation_PRODUCTION import BYODAutomation, DatabaseConnectionError

# ============================================================================
# LOGGING SETUP
# ============================================================================

def setup_logging():
    """Configure logging with file rotation"""
    # Create logs directory if it doesn't exist
    log_dir = Path(Config.LOG_FILE).parent
    log_dir.mkdir(parents=True, exist_ok=True)

    # Create logger
    logger = logging.getLogger('byod_daemon')
    logger.setLevel(getattr(logging, Config.LOG_LEVEL))

    # File handler with rotation
    file_handler = RotatingFileHandler(
        Config.LOG_FILE,
        maxBytes=Config.LOG_MAX_SIZE,
        backupCount=Config.LOG_BACKUP_COUNT
    )
    file_handler.setLevel(getattr(logging, Config.LOG_LEVEL))

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(getattr(logging, Config.LOG_LEVEL))

    # Formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # Add handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


logger = setup_logging()


# ============================================================================
# BYOD DAEMON CLASS
# ============================================================================

class BYODDaemon:
    """Background daemon for BYOD automation"""

    def __init__(self, check_interval: int = 30):
        """Initialize daemon with check interval"""
        self.check_interval = check_interval
        self.running = False
        self.automation = None
        
        # Set up signal handlers for graceful shutdown
        signal.signal(signal.SIGTERM, self._signal_handler)
        signal.signal(signal.SIGINT, self._signal_handler)

    def _signal_handler(self, signum, frame):
        """Handle shutdown signals gracefully"""
        signal_name = 'SIGINT' if signum == signal.SIGINT else 'SIGTERM'
        logger.info(f"Received {signal_name} signal. Shutting down gracefully...")
        self.running = False
        sys.exit(0)

    def initialize(self) -> bool:
        """Initialize the daemon and automation engine"""
        try:
            logger.info("🚀 Initializing BYOD Automation Daemon...")
            
            # Validate configuration
            is_valid, errors = Config.validate()
            if not is_valid:
                logger.error("❌ Configuration validation failed:")
                for error in errors:
                    logger.error(f"   - {error}")
                return False
            
            logger.info(f"✅ Configuration validated successfully")
            logger.info(f"   Log Level: {Config.LOG_LEVEL}")
            logger.info(f"   Check Interval: {self.check_interval}s")
            logger.info(f"   Max Retries: {Config.MAX_RETRIES}")
            
            # Initialize automation engine
            try:
                self.automation = BYODAutomation()
                logger.info("✅ Automation engine initialized")
                return True
            except DatabaseConnectionError as e:
                logger.error(f"❌ Failed to initialize automation engine: {e}")
                return False
                
        except Exception as e:
            logger.error(f"❌ Daemon initialization failed: {e}", exc_info=True)
            return False

    def run(self):
        """Main daemon loop"""
        if not self.initialize():
            logger.error("❌ Failed to initialize daemon. Exiting.")
            sys.exit(1)

        logger.info("\n" + "=" * 70)
        logger.info("🚀 BYOD AUTOMATION DAEMON STARTED")
        logger.info("=" * 70)
        logger.info(f"Start Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"Check Interval: {self.check_interval} seconds")
        logger.info(f"Log File: {Config.LOG_FILE}")
        logger.info("Press Ctrl+C to stop gracefully")
        logger.info("=" * 70 + "\n")

        self.running = True
        cycle_count = 0
        consecutive_failures = 0
        max_consecutive_failures = 5

        while self.running:
            try:
                cycle_count += 1
                cycle_start = time.time()

                logger.debug(f"[Cycle {cycle_count}] Starting automation run...")
                
                # Run automation
                try:
                    self.automation.run_automation()
                    consecutive_failures = 0  # Reset failure counter on success
                except Exception as e:
                    consecutive_failures += 1
                    logger.error(f"❌ Automation cycle failed (attempt {consecutive_failures}): {e}", exc_info=True)
                    
                    # Exit if too many consecutive failures
                    if consecutive_failures >= max_consecutive_failures:
                        logger.critical(
                            f"❌ CRITICAL: {consecutive_failures} consecutive failures. "
                            f"Shutting down daemon."
                        )
                        self.running = False
                        break

                # Calculate actual run time and sleep for remainder
                cycle_time = time.time() - cycle_start
                sleep_time = max(0, self.check_interval - cycle_time)

                if sleep_time > 0:
                    logger.debug(f"[Cycle {cycle_count}] Sleeping for {sleep_time:.1f}s...")
                    time.sleep(sleep_time)
                else:
                    logger.warning(
                        f"[Cycle {cycle_count}] Automation took {cycle_time:.1f}s "
                        f"(exceeds check_interval of {self.check_interval}s)"
                    )

            except KeyboardInterrupt:
                logger.info("\n🛑 Keyboard interrupt received. Shutting down...")
                self.running = False
                break
            except Exception as e:
                consecutive_failures += 1
                logger.error(f"❌ Unexpected error in daemon loop: {e}", exc_info=True)
                
                if consecutive_failures >= max_consecutive_failures:
                    logger.critical(f"❌ Too many errors. Shutting down.")
                    self.running = False
                    break
                
                # Wait before retrying
                logger.info(f"⏳ Waiting {Config.RETRY_DELAY}s before retry...")
                time.sleep(Config.RETRY_DELAY)

        # Shutdown
        logger.info("\n" + "=" * 70)
        logger.info("🛑 BYOD AUTOMATION DAEMON STOPPED")
        logger.info("=" * 70)
        logger.info(f"End Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"Total Cycles: {cycle_count}")
        logger.info("=" * 70 + "\n")
        
        sys.exit(0)


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """Main entry point"""
    try:
        # Get check interval from environment or config
        check_interval = Config.CHECK_INTERVAL
        
        # Validate interval
        if check_interval < 10:
            logger.warning(f"⚠️ CHECK_INTERVAL is very low ({check_interval}s). Minimum recommended is 30s.")
        if check_interval > 3600:
            logger.warning(f"⚠️ CHECK_INTERVAL is very high ({check_interval}s). Consider lowering for responsiveness.")

        # Create and run daemon
        daemon = BYODDaemon(check_interval=check_interval)
        daemon.run()

    except Exception as e:
        logger.critical(f"❌ Fatal error: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
