"""
Configuration loader for M365 email population.

Handles loading and validation of:
- mailboxes.yaml - User mailbox configuration
- sites.json - SharePoint sites configuration  
- environments.json - Tenant/environment configuration
"""

import json
import os
from pathlib import Path
from typing import Dict, List, Optional, Any

# Try to import yaml, provide helpful error if not installed
try:
    import yaml
except ImportError:
    yaml = None

# Configuration paths
SCRIPT_DIR = Path(__file__).parent.parent.resolve()
PROJECT_DIR = SCRIPT_DIR.parent
CONFIG_DIR = PROJECT_DIR / "config"

MAILBOXES_FILE = CONFIG_DIR / "mailboxes.yaml"
SITES_FILE = CONFIG_DIR / "sites.json"
ENVIRONMENTS_FILE = CONFIG_DIR / "environments.json"


def check_yaml_installed() -> bool:
    """Check if PyYAML is installed."""
    return yaml is not None


def load_mailbox_config(config_path: Optional[Path] = None) -> Dict[str, Any]:
    """
    Load and validate mailbox configuration from YAML.
    
    Args:
        config_path: Optional path to mailboxes.yaml. Uses default if not provided.
        
    Returns:
        Dictionary containing mailbox configuration.
        
    Raises:
        FileNotFoundError: If config file doesn't exist.
        ValueError: If config is invalid.
        ImportError: If PyYAML is not installed.
    """
    if not check_yaml_installed():
        raise ImportError(
            "PyYAML is required for email population. "
            "Install it with: pip install pyyaml"
        )
    
    path = config_path or MAILBOXES_FILE
    
    if not path.exists():
        raise FileNotFoundError(
            f"Mailbox configuration not found: {path}\n"
            f"Please create {path.name} with your mailbox definitions."
        )
    
    with open(path, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    
    # Validate configuration
    validate_config(config)
    
    # Apply defaults
    config = apply_defaults(config)
    
    return config


def load_sites_config(sites_path: Optional[Path] = None) -> Dict[str, Any]:
    """
    Load SharePoint sites configuration.
    
    Args:
        sites_path: Optional path to sites.json. Uses default if not provided.
        
    Returns:
        Dictionary containing sites configuration.
    """
    path = sites_path or SITES_FILE
    
    if not path.exists():
        return {"sites": []}
    
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_environment_config(env_path: Optional[Path] = None) -> Dict[str, Any]:
    """
    Load environment/tenant configuration.
    
    Args:
        env_path: Optional path to environments.json. Uses default if not provided.
        
    Returns:
        Dictionary containing environment configuration.
    """
    path = env_path or ENVIRONMENTS_FILE
    
    if not path.exists():
        return {"environments": [], "default_environment": None}
    
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def get_environment(env_config: Dict[str, Any], env_name: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """
    Get a specific environment configuration.
    
    Args:
        env_config: Full environment configuration.
        env_name: Name of environment to get. Uses default if not provided.
        
    Returns:
        Environment configuration dictionary or None if not found.
    """
    environments = env_config.get("environments", [])
    
    if not environments:
        return None
    
    # Use specified name or default
    target_name = env_name or env_config.get("default_environment")
    
    if not target_name:
        # Return first environment if no default specified
        return environments[0] if environments else None
    
    # Find matching environment
    for env in environments:
        if env.get("name") == target_name:
            return env
    
    return None


def get_tenant_domain(env: Dict[str, Any]) -> Optional[str]:
    """
    Get the M365 tenant domain from environment config.
    
    Args:
        env: Environment configuration dictionary.
        
    Returns:
        Tenant domain (e.g., 'contoso.onmicrosoft.com') or None.
    """
    m365_config = env.get("m365", {})
    tenant_name = m365_config.get("tenant_name")
    
    if tenant_name:
        # Handle both formats: 'contoso' and 'contoso.onmicrosoft.com'
        if ".onmicrosoft.com" in tenant_name:
            return tenant_name
        return f"{tenant_name}.onmicrosoft.com"
    
    return None


def get_tenant_id(env: Dict[str, Any]) -> Optional[str]:
    """
    Get the Azure tenant ID from environment config.
    
    Args:
        env: Environment configuration dictionary.
        
    Returns:
        Tenant ID or None.
    """
    azure_config = env.get("azure", {})
    return azure_config.get("tenant_id")


def validate_config(config: Dict[str, Any]) -> None:
    """
    Validate mailbox configuration structure.
    
    Args:
        config: Configuration dictionary to validate.
        
    Raises:
        ValueError: If configuration is invalid.
    """
    if not config:
        raise ValueError("Configuration is empty")
    
    # Check required top-level keys
    if "settings" not in config:
        raise ValueError("Missing required 'settings' section in configuration")
    
    if "users" not in config:
        raise ValueError("Missing required 'users' section in configuration")
    
    users = config.get("users", [])
    if not users:
        raise ValueError("No users defined in configuration")
    
    # Validate each user
    for i, user in enumerate(users):
        if "upn" not in user:
            raise ValueError(f"User at index {i} is missing required 'upn' field")
        
        if "role" not in user:
            raise ValueError(f"User '{user.get('upn', f'index {i}')}' is missing required 'role' field")
        
        # Validate UPN format
        upn = user["upn"]
        if "@" not in upn:
            raise ValueError(f"Invalid UPN format for user: {upn}")
    
    # Validate email distribution if present
    settings = config.get("settings", {})
    distribution = settings.get("email_distribution", {})
    
    if distribution:
        total = sum(distribution.values())
        if total != 100:
            raise ValueError(
                f"Email distribution percentages must sum to 100, got {total}"
            )


def apply_defaults(config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Apply default values to configuration.
    
    Args:
        config: Configuration dictionary.
        
    Returns:
        Configuration with defaults applied.
    """
    defaults = {
        "settings": {
            "default_email_count": 50,
            "date_range_months": 12,
            "include_sensitivity_labels": True,
            "email_distribution": {
                "newsletters": 20,
                "links": 20,
                "attachments": 20,
                "organisational": 20,
                "interdepartmental": 20,
            },
            "threading": {
                "enabled": True,
                "single_email_percentage": 60,
                "reply_chain_percentage": 25,
                "forward_chain_percentage": 10,
                "reply_all_percentage": 5,
                "max_thread_depth": 5,
            },
            "sender_distribution": {
                "internal_users": 60,
                "internal_system": 20,
                "external": 20,
            },
            "date_settings": {
                "business_hours_percentage": 85,
                "weekday_percentage": 90,
                "recent_bias": True,
            },
        },
        "volume_mappings": {
            "high": 100,
            "medium": 50,
            "low": 25,
        },
    }
    
    # Deep merge defaults with config
    result = _deep_merge(defaults, config)
    
    return result


def _deep_merge(base: Dict, override: Dict) -> Dict:
    """
    Deep merge two dictionaries.
    
    Args:
        base: Base dictionary with defaults.
        override: Dictionary with overrides.
        
    Returns:
        Merged dictionary.
    """
    result = base.copy()
    
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    
    return result


def get_user_email_count(user: Dict[str, Any], settings: Dict[str, Any]) -> int:
    """
    Calculate email count for a user based on volume setting.
    
    Args:
        user: User configuration dictionary.
        settings: Global settings dictionary.
        
    Returns:
        Number of emails to generate for this user.
    """
    volume = user.get("email_volume", "medium")
    
    # If volume is already a number, use it directly
    if isinstance(volume, int):
        return max(1, min(volume, 500))  # Clamp between 1 and 500
    
    # Look up volume mapping
    volume_mappings = settings.get("volume_mappings", {
        "high": 100,
        "medium": 50,
        "low": 25,
    })
    
    return volume_mappings.get(str(volume).lower(), 50)


def get_user_department(user: Dict[str, Any]) -> str:
    """
    Get the department for a user, inferring from role if not specified.
    
    Args:
        user: User configuration dictionary.
        
    Returns:
        Department name.
    """
    # Return explicit department if set
    if "department" in user:
        return user["department"]
    
    # Try to infer from role
    role = user.get("role", "").lower()
    
    role_department_map = {
        "ceo": "Executive Leadership",
        "cfo": "Executive Leadership",
        "coo": "Executive Leadership",
        "cto": "Executive Leadership",
        "chief": "Executive Leadership",
        "head of it": "IT Department",
        "it manager": "IT Department",
        "it director": "IT Department",
        "system admin": "IT Department",
        "hr": "Human Resources",
        "human resources": "Human Resources",
        "recruitment": "Human Resources",
        "finance": "Finance Department",
        "accountant": "Finance Department",
        "financial": "Finance Department",
        "marketing": "Marketing Department",
        "sales": "Sales Department",
        "legal": "Legal & Compliance",
        "compliance": "Legal & Compliance",
        "operations": "Operations Department",
        "customer service": "Customer Service",
        "support": "Customer Service",
        "training": "Training & Development",
        "project manager": "Project Management Office",
        "pmo": "Project Management Office",
        "research": "Research & Development",
        "r&d": "Research & Development",
    }
    
    for keyword, department in role_department_map.items():
        if keyword in role:
            return department
    
    return "General"


def get_department_site(department: str, config: Dict[str, Any]) -> Optional[str]:
    """
    Get the SharePoint site name for a department.
    
    Args:
        department: Department name.
        config: Full mailbox configuration.
        
    Returns:
        SharePoint site name or None.
    """
    departments = config.get("departments", {})
    dept_config = departments.get(department, {})
    return dept_config.get("sharepoint_site")


def get_all_users(config: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Get all users from configuration with computed fields.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        List of user dictionaries with computed fields.
    """
    users = config.get("users", [])
    settings = config.get("settings", {})
    
    result = []
    for user in users:
        # Create a copy with computed fields
        user_copy = user.copy()
        user_copy["department"] = get_user_department(user)
        user_copy["email_count"] = get_user_email_count(user, settings)
        user_copy["display_name"] = _get_display_name(user)
        result.append(user_copy)
    
    return result


def _get_display_name(user: Dict[str, Any]) -> str:
    """
    Get display name for a user from UPN.
    
    Args:
        user: User configuration dictionary.
        
    Returns:
        Display name.
    """
    upn = user.get("upn", "")
    
    # Extract name part before @
    name_part = upn.split("@")[0]
    
    # Split by common separators and capitalize
    parts = name_part.replace(".", " ").replace("_", " ").replace("-", " ").split()
    
    return " ".join(part.capitalize() for part in parts)
