"""
Unit tests for config module

Tests configuration loading from multiple sources with priority.

@author: Generated for outlook_automation repository
"""

import pytest
import os
import json
import tempfile
from pathlib import Path
import sys

# Add src to path
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))

from outlook_graph.config import Config


class TestConfigLoading:
    """Test configuration loading from various sources."""

    def test_config_defaults(self):
        """Test that default configuration values are set."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create empty config file to prevent loading from standard locations
            config_file = Path(tmpdir) / "config.json"
            config_file.write_text("{}")

            config = Config(config_file)

            assert config.tenant_id == "common"
            assert config.api_version == "v1.0"
            assert config.timezone == "Europe/Copenhagen"
            assert "Calendars.ReadWrite" in config.scopes

    def test_config_from_file(self):
        """Test loading configuration from JSON file."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_data = {
                "tenant_id": "test-tenant-id",
                "client_id": "test-client-id",
                "timezone": "America/New_York",
                "scopes": ["Calendars.Read", "User.Read"]
            }
            config_file.write_text(json.dumps(config_data))

            config = Config(config_file)

            assert config.tenant_id == "test-tenant-id"
            assert config.client_id == "test-client-id"
            assert config.timezone == "America/New_York"
            assert config.scopes == ["Calendars.Read", "User.Read"]

    def test_config_from_env_overrides_file(self):
        """Test that environment variables override file config."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_data = {
                "tenant_id": "file-tenant",
                "client_id": "file-client-id"
            }
            config_file.write_text(json.dumps(config_data))

            # Set environment variables
            os.environ["GRAPH_TENANT_ID"] = "env-tenant"
            os.environ["GRAPH_CLIENT_ID"] = "env-client-id"

            try:
                config = Config(config_file)

                # Environment variables should override file
                assert config.tenant_id == "env-tenant"
                assert config.client_id == "env-client-id"

            finally:
                # Clean up environment variables
                del os.environ["GRAPH_TENANT_ID"]
                del os.environ["GRAPH_CLIENT_ID"]

    def test_config_scopes_from_env(self):
        """Test loading comma-separated scopes from environment variable."""
        os.environ["GRAPH_SCOPES"] = "Scope1,Scope2,Scope3"

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                config_file = Path(tmpdir) / "config.json"
                config_file.write_text("{}")

                config = Config(config_file)

                assert len(config.scopes) == 3
                assert "Scope1" in config.scopes
                assert "Scope2" in config.scopes
                assert "Scope3" in config.scopes

        finally:
            del os.environ["GRAPH_SCOPES"]

    def test_config_missing_client_id_raises_error(self):
        """Test that missing client_id raises ValueError."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_file.write_text('{"tenant_id": "test-tenant"}')

            config = Config(config_file)

            with pytest.raises(ValueError, match="Client ID not configured"):
                _ = config.client_id

    def test_config_validate_success(self):
        """Test config validation with valid configuration."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_data = {
                "client_id": "test-client-id",
                "scopes": ["Calendars.ReadWrite"]
            }
            config_file.write_text(json.dumps(config_data))

            config = Config(config_file)

            # Should not raise
            assert config.validate() is True

    def test_config_validate_missing_client_id(self):
        """Test config validation fails with missing client_id."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_file.write_text("{}")

            config = Config(config_file)

            with pytest.raises(ValueError, match="Client ID not configured"):
                config.validate()

    def test_config_graph_endpoint(self):
        """Test that graph_endpoint combines base_url and api_version."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_file.write_text("{}")

            config = Config(config_file)

            assert config.graph_endpoint == "https://graph.microsoft.com/v1.0"

    def test_config_to_dict(self):
        """Test converting config to dictionary."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_data = {
                "tenant_id": "test-tenant",
                "client_id": "test-client-id"
            }
            config_file.write_text(json.dumps(config_data))

            config = Config(config_file)
            config_dict = config.to_dict()

            assert isinstance(config_dict, dict)
            assert config_dict["tenant_id"] == "test-tenant"
            assert config_dict["client_id"] == "test-client-id"

    def test_config_get_method(self):
        """Test get() method with default value."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_file.write_text('{"custom_key": "custom_value"}')

            config = Config(config_file)

            assert config.get("custom_key") == "custom_value"
            assert config.get("nonexistent", "default") == "default"
            assert config.get("nonexistent") is None


class TestConfigEdgeCases:
    """Test edge cases and error handling."""

    def test_config_invalid_json(self):
        """Test that invalid JSON in config file is handled gracefully."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_file.write_text("{ invalid json }")

            # Should fall back to defaults instead of crashing
            config = Config(config_file)

            assert config.tenant_id == "common"  # Default value

    def test_config_nonexistent_file(self):
        """Test loading when config file doesn't exist."""
        nonexistent_file = Path("/nonexistent/path/config.json")

        # Should use defaults without crashing
        config = Config(nonexistent_file)

        assert config.tenant_id == "common"
        assert config.timezone == "Europe/Copenhagen"

    def test_config_empty_scopes(self):
        """Test that empty scopes list uses defaults."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_file.write_text('{"scopes": []}')

            config = Config(config_file)

            # Empty scopes should be preserved (not replaced with defaults)
            assert config.scopes == []

    def test_config_repr(self):
        """Test string representation of Config."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_file.write_text('{"client_id": "very-long-client-id-12345"}')

            config = Config(config_file)
            repr_str = repr(config)

            assert "Config" in repr_str
            # Should truncate client_id for security
            assert "very-lon..." in repr_str or "..." in repr_str


class TestConfigIntegration:
    """Integration tests for configuration management."""

    def test_config_priority_env_over_file_over_defaults(self):
        """Test full priority chain: env > file > defaults."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create config file with some values
            config_file = Path(tmpdir) / "config.json"
            config_data = {
                "tenant_id": "file-tenant",
                "client_id": "file-client",
                "timezone": "America/New_York"
            }
            config_file.write_text(json.dumps(config_data))

            # Set some (but not all) env vars
            os.environ["GRAPH_TENANT_ID"] = "env-tenant"

            try:
                config = Config(config_file)

                # tenant_id from env (highest priority)
                assert config.tenant_id == "env-tenant"

                # client_id from file (env not set)
                assert config.client_id == "file-client"

                # timezone from file (env not set)
                assert config.timezone == "America/New_York"

                # api_version from defaults (not in file or env)
                assert config.api_version == "v1.0"

            finally:
                del os.environ["GRAPH_TENANT_ID"]


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
