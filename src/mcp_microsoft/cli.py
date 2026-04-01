"""
Interactive CLI for managing Microsoft 365 profiles.

Run directly:
    python -m mcp_microsoft.cli
    # or via the installed script:
    mcp-microsoft-setup
"""

from __future__ import annotations

import sys

from mcp_microsoft.profiles import ProfileManager


def _input(prompt: str, default: str = "") -> str:
    """Prompt with optional default value."""
    if default:
        raw = input(f"{prompt} [{default}]: ").strip()
        return raw or default
    while True:
        raw = input(f"{prompt}: ").strip()
        if raw:
            return raw
        print("  This field is required.")


def cmd_add(pm: ProfileManager) -> None:
    """Interactive profile creation."""
    print("\n--- Add Microsoft 365 Profile ---\n")
    name = _input("Profile name (e.g. personal, work)")
    client_id = _input("Azure App Registration Client ID")
    tenant_id = _input("Tenant ID", default="common")

    existing = pm.profiles
    set_default = not existing
    if existing:
        resp = input("Set as default profile? [y/N]: ").strip().lower()
        set_default = resp in ("y", "yes")

    try:
        cfg = pm.add_profile(
            name=name,
            client_id=client_id,
            tenant_id=tenant_id,
            set_as_default=set_default,
        )
    except ValueError as exc:
        print(f"\nError: {exc}")
        return

    print(f"\nProfile '{cfg.name}' created.")
    print(f"  Client ID: {cfg.client_id}")
    print(f"  Tenant:    {cfg.tenant_id}")
    print(f"  Cache:     {cfg.cache_path}")
    if set_default:
        print("  Set as default profile.")

    resp = input("\nAuthenticate now? [Y/n]: ").strip().lower()
    if resp not in ("n", "no"):
        cmd_auth(pm, cfg.name)


def cmd_auth(pm: ProfileManager, profile_name: str | None = None) -> None:
    """Trigger interactive auth for a profile."""
    if profile_name is None:
        profiles = pm.profiles
        if not profiles:
            print("No profiles configured. Add one first.")
            return
        if len(profiles) == 1:
            profile_name = next(iter(profiles))
        else:
            print("\nAvailable profiles:")
            for name in sorted(profiles):
                tag = " (default)" if name == pm.default_profile_name else ""
                print(f"  - {name}{tag}")
            profile_name = _input("\nProfile to authenticate")

    print(f"\nAuthenticating '{profile_name}'... A browser window will open.")
    try:
        pm.get_token(profile_name)
        print(f"Profile '{profile_name}' authenticated successfully.")
    except (ValueError, RuntimeError) as exc:
        print(f"Authentication failed: {exc}")


def cmd_list(pm: ProfileManager) -> None:
    """List all profiles."""
    profiles = pm.profiles
    if not profiles:
        print("No profiles configured.")
        return

    print("\n--- Microsoft 365 Profiles ---\n")
    for name, cfg in sorted(profiles.items()):
        default_tag = " (default)" if name == pm.default_profile_name else ""
        masked_cid = f"{cfg.client_id[:8]}..." if len(cfg.client_id) > 8 else cfg.client_id
        try:
            authed = "yes" if pm.is_authenticated(name) else "no"
        except Exception:
            authed = "error"
        print(f"  {name}{default_tag}")
        print(f"    Client ID:     {masked_cid}")
        print(f"    Tenant:        {cfg.tenant_id}")
        print(f"    Authenticated: {authed}")
        print()


def cmd_remove(pm: ProfileManager) -> None:
    """Remove a profile."""
    profiles = pm.profiles
    if not profiles:
        print("No profiles configured.")
        return

    print("\nAvailable profiles:")
    for name in sorted(profiles):
        print(f"  - {name}")

    name = _input("\nProfile to remove")
    resp = input(f"Remove '{name}' and delete its token cache? [y/N]: ").strip().lower()
    if resp not in ("y", "yes"):
        print("Cancelled.")
        return

    try:
        pm.remove_profile(name)
        print(f"Profile '{name}' removed.")
    except ValueError as exc:
        print(f"Error: {exc}")


def cmd_set_default(pm: ProfileManager) -> None:
    """Change the default profile."""
    profiles = pm.profiles
    if not profiles:
        print("No profiles configured.")
        return

    print("\nAvailable profiles:")
    for name in sorted(profiles):
        tag = " (default)" if name == pm.default_profile_name else ""
        print(f"  - {name}{tag}")

    name = _input("\nNew default profile")
    try:
        pm.set_default(name)
        print(f"Default profile changed to '{name}'.")
    except ValueError as exc:
        print(f"Error: {exc}")


COMMANDS = {
    "add": ("Add a new profile", cmd_add),
    "auth": ("Authenticate a profile", lambda pm: cmd_auth(pm)),
    "list": ("List all profiles", cmd_list),
    "remove": ("Remove a profile", cmd_remove),
    "default": ("Change the default profile", cmd_set_default),
}


def main() -> None:
    """CLI entry point."""
    pm = ProfileManager.get()

    # If a subcommand was passed directly
    if len(sys.argv) > 1:
        cmd_name = sys.argv[1].lower()
        if cmd_name in COMMANDS:
            COMMANDS[cmd_name][1](pm)
            return
        print(f"Unknown command: {cmd_name}")
        print(f"Available: {', '.join(COMMANDS)}")
        sys.exit(1)

    # Interactive menu
    print("=== mcp-microsoft profile setup ===\n")
    for key, (desc, _) in COMMANDS.items():
        print(f"  {key:10s} {desc}")

    print()
    cmd_name = _input("Command").lower()
    if cmd_name in COMMANDS:
        COMMANDS[cmd_name][1](pm)
    else:
        print(f"Unknown command: {cmd_name}")
        sys.exit(1)


if __name__ == "__main__":
    main()
