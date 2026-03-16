"""
Configuration loader for M365 email population.

Handles loading and validation of:
- mailboxes.yaml - User mailbox configuration
- sites.json - SharePoint sites configuration
- environments.json - Tenant/environment configuration
- Azure AD auto-discovery settings
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
AZURE_AD_CACHE_FILE = CONFIG_DIR / ".azure_ad_cache.json"


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
        # Azure AD auto-discovery defaults
        "azure_ad": {
            "enabled": False,
            "users": {
                "max_users": None,  # No limit by default
                "include_departments": None,  # All departments
                "exclude_departments": None,
                "exclude_external_users": False,
                "exclude_service_accounts": True,
                "validate_mailbox_exists": True,
            },
            "groups": {
                "enabled": True,
                "include_m365_groups": True,
                "include_security_groups": True,
                "include_distribution_lists": True,
                "exclude_patterns": None,
                "min_members": 2,
                "max_groups": None,
            },
            "cache": {
                "enabled": True,
                "ttl_minutes": 60,
                "cache_file": str(AZURE_AD_CACHE_FILE),
            },
        },
        # CC/BCC configuration defaults
        "cc_bcc": {
            "cc": {
                "enabled": True,
                "probability": 0.3,
                "max_recipients": 5,
            },
            "bcc": {
                "enabled": True,
                "probability": 0.1,
                "max_recipients": 3,
            },
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
        "claims": "Claims Department",
        "adjuster": "Claims Department",
        "claims adjuster": "Claims Department",
        "claims analyst": "Claims Department",
        "claims manager": "Claims Department",
        "claims examiner": "Claims Department",
        "loss adjuster": "Claims Department",
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


# =============================================================================
# AZURE AD CONFIGURATION HELPERS
# =============================================================================

def is_azure_ad_enabled(config: Dict[str, Any]) -> bool:
    """
    Check if Azure AD auto-discovery is enabled.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        True if Azure AD discovery is enabled.
    """
    return config.get("azure_ad", {}).get("enabled", False)


def get_azure_ad_config(config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Get Azure AD configuration section.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        Azure AD configuration dictionary.
    """
    return config.get("azure_ad", {})


def get_azure_ad_user_config(config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Get Azure AD user discovery configuration.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        User discovery configuration dictionary.
    """
    return config.get("azure_ad", {}).get("users", {})


def get_azure_ad_group_config(config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Get Azure AD group discovery configuration.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        Group discovery configuration dictionary.
    """
    return config.get("azure_ad", {}).get("groups", {})


def get_azure_ad_cache_config(config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Get Azure AD cache configuration.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        Cache configuration dictionary.
    """
    return config.get("azure_ad", {}).get("cache", {})


def get_cc_bcc_config(config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Get CC/BCC configuration.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        CC/BCC configuration dictionary.
    """
    return config.get("cc_bcc", {})


def get_azure_ad_cache_path(config: Dict[str, Any]) -> Path:
    """
    Get the path to the Azure AD cache file.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        Path to cache file.
    """
    cache_config = get_azure_ad_cache_config(config)
    cache_file = cache_config.get("cache_file", str(AZURE_AD_CACHE_FILE))
    return Path(cache_file)


def is_azure_ad_cache_enabled(config: Dict[str, Any]) -> bool:
    """
    Check if Azure AD caching is enabled.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        True if caching is enabled.
    """
    return get_azure_ad_cache_config(config).get("enabled", True)


def get_azure_ad_cache_ttl(config: Dict[str, Any]) -> int:
    """
    Get Azure AD cache TTL in minutes.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        Cache TTL in minutes.
    """
    return get_azure_ad_cache_config(config).get("ttl_minutes", 60)


def get_mailbox_users_from_config(config: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Get users that have mailboxes from YAML configuration.
    
    This returns users from the 'users' section that are expected
    to have mailboxes (for email population).
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        List of user dictionaries with mailboxes.
    """
    return get_all_users(config)


def merge_yaml_and_azure_ad_users(
    yaml_users: List[Dict[str, Any]],
    azure_ad_users: List[Dict[str, Any]],
    prefer_azure_ad: bool = True
) -> List[Dict[str, Any]]:
    """
    Merge users from YAML configuration and Azure AD discovery.
    
    Args:
        yaml_users: Users from YAML configuration.
        azure_ad_users: Users from Azure AD discovery.
        prefer_azure_ad: If True, Azure AD data takes precedence for duplicates.
        
    Returns:
        Merged list of users.
    """
    user_map: Dict[str, Dict[str, Any]] = {}
    
    # Add users based on preference
    if prefer_azure_ad:
        # Add YAML users first
        for user in yaml_users:
            upn = user.get("upn", "").lower()
            if upn:
                user_map[upn] = user
        
        # Override with Azure AD users
        for user in azure_ad_users:
            upn = user.get("upn", user.get("email", "")).lower()
            if upn:
                user_map[upn] = user
    else:
        # Add Azure AD users first
        for user in azure_ad_users:
            upn = user.get("upn", user.get("email", "")).lower()
            if upn:
                user_map[upn] = user
        
        # Override with YAML users
        for user in yaml_users:
            upn = user.get("upn", "").lower()
            if upn:
                user_map[upn] = user
    
    return list(user_map.values())


def validate_azure_ad_config(config: Dict[str, Any]) -> List[str]:
    """
    Validate Azure AD configuration and return any warnings.
    
    Args:
        config: Full mailbox configuration.
        
    Returns:
        List of warning messages (empty if valid).
    """
    warnings: List[str] = []
    azure_ad = config.get("azure_ad", {})
    
    if not azure_ad.get("enabled", False):
        return warnings
    
    # Check user configuration
    user_config = azure_ad.get("users", {})
    max_users = user_config.get("max_users")
    if max_users is not None and max_users < 1:
        warnings.append("azure_ad.users.max_users should be at least 1 or None for unlimited")
    
    # Check group configuration
    group_config = azure_ad.get("groups", {})
    min_members = group_config.get("min_members", 2)
    if min_members < 1:
        warnings.append("azure_ad.groups.min_members should be at least 1")
    
    # Check cache configuration
    cache_config = azure_ad.get("cache", {})
    ttl = cache_config.get("ttl_minutes", 60)
    if ttl < 1:
        warnings.append("azure_ad.cache.ttl_minutes should be at least 1")
    
    # Check CC/BCC configuration
    cc_bcc = config.get("cc_bcc", {})
    cc_config = cc_bcc.get("cc", {})
    bcc_config = cc_bcc.get("bcc", {})
    
    cc_prob = cc_config.get("probability", 0.3)
    if not 0 <= cc_prob <= 1:
        warnings.append("cc_bcc.cc.probability should be between 0 and 1")
    
    bcc_prob = bcc_config.get("probability", 0.1)
    if not 0 <= bcc_prob <= 1:
        warnings.append("cc_bcc.bcc.probability should be between 0 and 1")
    
    cc_max = cc_config.get("max_recipients", 5)
    if cc_max < 1:
        warnings.append("cc_bcc.cc.max_recipients should be at least 1")
    
    bcc_max = bcc_config.get("max_recipients", 3)
    if bcc_max < 1:
        warnings.append("cc_bcc.bcc.max_recipients should be at least 1")
    
    return warnings
