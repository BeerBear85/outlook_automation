"""
Configuration Management Module

Loads configuration from multiple sources with priority:
1. Environment variables (highest priority)
2. config.json file
3. Default values (lowest priority)

@author: Generated for outlook_automation repository (Python Graph implementation)
"""

import os
import json
from pathlib import Path
from typing import Dict, Any, Optional, List


class Config:
    """
    Configuration manager for Outlook Graph automation.

    Loads settings from environment variables, config file, or defaults.
    """

    # Default configuration values
    DEFAULTS = {
        "tenant_id": "common",
        "client_id": None,  # Must be provided by user
        "scopes": ["Calendars.ReadWrite", "Mail.ReadWrite", "User.Read"],
        "timezone": "Europe/Copenhagen",
        "api_version": "v1.0",
        "base_url": "https://graph.microsoft.com"
    }

    def __init__(self, config_file: Optional[Path] = None):
        """
        Initialize configuration.

        Args:
            config_file: Path to config.json file. If None, looks in standard locations.
        """
        self._config = {}
        self._load_config(config_file)

    def _load_config(self, config_file: Optional[Path] = None):
        """
        Load configuration from all sources.

        Priority: Environment variables > config.json > defaults
        """
        # Start with defaults
        self._config = self.DEFAULTS.copy()

        # Try to load from config file
        if config_file is None:
            # Look in standard locations
            config_file = self._find_config_file()

        if config_file and config_file.exists():
            try:
                with open(config_file, 'r') as f:
                    file_config = json.load(f)
                    self._config.update(file_config)
                print(f"✓ Loaded configuration from: {config_file}")
            except Exception as e:
                print(f"⚠ Warning: Failed to load config file {config_file}: {e}")

        # Override with environment variables (highest priority)
        self._load_from_env()

    def _find_config_file(self) -> Optional[Path]:
        """
        Find config.json in standard locations.

        Searches:
        1. Current directory
        2. Script directory (scripts_using_python/)
        3. User home directory (~/.outlook_automation/)

        Returns:
            Path to config file if found, None otherwise
        """
        search_paths = [
            Path.cwd() / "config.json",
            Path(__file__).parent.parent.parent / "config.json",
            Path.home() / ".outlook_automation" / "config.json"
        ]

        for path in search_paths:
            if path.exists():
                return path

        return None

    def _load_from_env(self):
        """Load configuration from environment variables."""
        env_mappings = {
            "GRAPH_TENANT_ID": "tenant_id",
            "GRAPH_CLIENT_ID": "client_id",
            "GRAPH_SCOPES": "scopes",
            "TIMEZONE": "timezone"
        }

        for env_var, config_key in env_mappings.items():
            value = os.environ.get(env_var)
            if value:
                # Special handling for scopes (comma-separated string)
                if config_key == "scopes":
                    value = [s.strip() for s in value.split(",")]
                self._config[config_key] = value

    def get(self, key: str, default: Any = None) -> Any:
        """
        Get configuration value.

        Args:
            key: Configuration key
            default: Default value if key not found

        Returns:
            Configuration value
        """
        return self._config.get(key, default)

    @property
    def tenant_id(self) -> str:
        """Get Azure AD tenant ID."""
        return self._config.get("tenant_id", "common")

    @property
    def client_id(self) -> Optional[str]:
        """Get Azure AD application (client) ID."""
        client_id = self._config.get("client_id")
        if not client_id:
            raise ValueError(
                "Client ID not configured. Please set GRAPH_CLIENT_ID environment variable "
                "or add 'client_id' to config.json"
            )
        return client_id

    @property
    def scopes(self) -> List[str]:
        """Get Microsoft Graph permission scopes."""
        return self._config.get("scopes", self.DEFAULTS["scopes"])

    @property
    def timezone(self) -> str:
        """Get timezone for date/time operations."""
        return self._config.get("timezone", "Europe/Copenhagen")

    @property
    def api_version(self) -> str:
        """Get Microsoft Graph API version."""
        return self._config.get("api_version", "v1.0")

    @property
    def base_url(self) -> str:
        """Get Microsoft Graph base URL."""
        return self._config.get("base_url", "https://graph.microsoft.com")

    @property
    def graph_endpoint(self) -> str:
        """Get full Microsoft Graph endpoint URL."""
        return f"{self.base_url}/{self.api_version}"

    def to_dict(self) -> Dict[str, Any]:
        """
        Get configuration as dictionary.

        Returns:
            Configuration dictionary
        """
        return self._config.copy()

    def validate(self) -> bool:
        """
        Validate that required configuration is present.

        Returns:
            True if valid, raises ValueError otherwise
        """
        if not self.client_id:
            raise ValueError("client_id is required")

        if not self.scopes:
            raise ValueError("scopes cannot be empty")

        return True

    def __repr__(self) -> str:
        """String representation of configuration."""
        # Don't include sensitive data
        safe_config = self._config.copy()
        if "client_id" in safe_config and safe_config["client_id"]:
            safe_config["client_id"] = safe_config["client_id"][:8] + "..."
        return f"Config({safe_config})"


# Global config instance (lazy-loaded)
_global_config: Optional[Config] = None


def get_config(config_file: Optional[Path] = None) -> Config:
    """
    Get global configuration instance.

    Args:
        config_file: Optional path to config file

    Returns:
        Config instance
    """
    global _global_config
    if _global_config is None:
        _global_config = Config(config_file)
    return _global_config


def reload_config(config_file: Optional[Path] = None) -> Config:
    """
    Reload global configuration from sources.

    Args:
        config_file: Optional path to config file

    Returns:
        New Config instance
    """
    global _global_config
    _global_config = Config(config_file)
    return _global_config
