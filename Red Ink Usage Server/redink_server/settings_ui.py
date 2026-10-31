"""Strict, atomic server.toml editor and encryption of operational secrets.

Never treat an editable TOML boolean as sufficient authorization for passwordless
administration. The encryption key lives separately in secrets.json and both
files must be protected by operating-system permissions.
"""
from __future__ import annotations

import json
import os
import re
import secrets
import threading
from dataclasses import replace
import tomllib
from pathlib import Path

from cryptography.fernet import Fernet, InvalidToken

from .config import ConfigError, _KNOWN, DEFAULT_TOML, load_settings, private_write, token_hash

_LOCK = threading.RLock()


def encrypt_secret(credentials: dict, value: str) -> str:
    return Fernet(credentials['configuration_encryption_key'].encode('ascii')).encrypt(value.encode('utf-8')).decode('ascii')


def decrypt_secret(credentials: dict, value: str) -> str:
    try:
        return Fernet(credentials['configuration_encryption_key'].encode('ascii')).decrypt(value.encode('ascii')).decode('utf-8')
    except (KeyError, ValueError, InvalidToken, UnicodeError) as exc:
        raise ConfigError('Protected configuration value is missing, invalid or has been modified.') from exc


def is_passwordless(settings) -> bool:
    token = settings.passwordless_mode
    if not token:
        return False
    return decrypt_secret(settings.credentials, token) == 'localhost-only-v1'


def value_toml(value):
    if type(value) is bool:
        return 'true' if value else 'false'
    if type(value) is int:
        return str(value)
    if isinstance(value, str):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, list) and all(isinstance(item, str) for item in value):
        return '[' + ', '.join(value_toml(item) for item in value) + ']'
    raise ConfigError('Unsupported value type in configuration editor.')


def replace_values(original: str, updates: dict) -> str:
    """Retain existing comments and section order; change only named values."""
    lines = original.splitlines(keepends=True)
    done = set()
    current = None
    output = []
    for line in lines:
        header = re.match(r'^\s*\[([A-Za-z0-9_]+)\]\s*(?:#.*)?$', line.rstrip('\r\n'))
        if header:
            if current:
                for key, value in updates.get(current, {}).items():
                    if (current, key) not in done:
                        output.append(f'{key} = {value_toml(value)}\n')
                        done.add((current, key))
            current = header.group(1)
            output.append(line)
            continue
        match = re.match(r'^(\s*)([a-z_]+)(\s*=\s*)(.*)$', line)
        if current in updates and match and match.group(2) in updates[current]:
            key = match.group(2)
            if (current, key) in done:
                raise ConfigError('Duplicate configuration key: ' + current + '.' + key)
            output.append(f'{match.group(1)}{key}{match.group(3)}{value_toml(updates[current][key])}\n')
            done.add((current, key))
        else:
            output.append(line)
    if current:
        for key, value in updates.get(current, {}).items():
            if (current, key) not in done:
                output.append(f'{key} = {value_toml(value)}\n')
                done.add((current, key))
    for section, items in updates.items():
        if any((section, key) not in done for key in items):
            output.append(f'\n[{section}]\n')
            for key, value in items.items():
                if (section, key) not in done:
                    output.append(f'{key} = {value_toml(value)}\n')
    result = ''.join(output)
    tomllib.loads(result)
    return result


def configuration_document(settings) -> dict:
    values = tomllib.loads((settings.data_dir / 'server.toml').read_text(encoding='utf-8'))
    defaults = tomllib.loads(DEFAULT_TOML)
    # Older installations did not have every section/key. Present all keys,
    # while preserving the existing TOML text until explicitly edited.
    values = {section: {**defaults.get(section, {}), **values.get(section, {})}
              for section in _KNOWN}
    # Security property is editable only via a separate authenticated action.
    auth = values.setdefault('authentication', {})
    auth['passwordless_mode'] = bool(is_passwordless(settings))
    return {
        'sections': values,
        'restart_required_on_save': False,
        'restart_only_sections': ['server', 'tls', 'collection'],
        'data_dir': str(settings.data_dir),
        'published_directory': str(settings.config_directory),
    }


def update_configuration(settings, updates: dict):
    if not isinstance(updates, dict) or not updates or set(updates) - set(_KNOWN):
        raise ConfigError('Invalid configuration sections.')
    for section, values in updates.items():
        if not isinstance(values, dict) or set(values) - _KNOWN[section]:
            raise ConfigError('Unknown or noneditable configuration option: ' + section)
        if section == 'authentication':
            raise ConfigError('Use the separate authenticated sign-in policy control.')
    path = settings.data_dir / 'server.toml'
    with _LOCK:
        old = path.read_text(encoding='utf-8')
        new = replace_values(old, updates)
        temporary = path.parent / ('.settings-' + secrets.token_hex(8) + '.tmp')
        private_write(temporary, new)
        try:
            os.replace(temporary, path)
            try:
                load_settings(settings.data_dir)
            except Exception:
                rollback = path.parent / ('.settings-rollback-' + secrets.token_hex(8) + '.tmp')
                private_write(rollback, old)
                os.replace(rollback, path)
                raise
        finally:
            temporary.unlink(missing_ok=True)
    restart_required = bool(set(updates) & {'server', 'tls', 'collection'})
    return {'saved': True, 'restart_required': restart_required}


def update_passwordless(settings, enabled: bool):
    if type(enabled) is not bool:
        raise ConfigError('Passwordless setting must be true or false.')
    if enabled and not settings.credentials.get('configuration_encryption_key'):
        raise ConfigError('The encryption key has not been initialized; restart after upgrading.')
    value = encrypt_secret(settings.credentials, 'localhost-only-v1') if enabled else ''
    with _LOCK:
        path = settings.data_dir / 'server.toml'
        old = path.read_text(encoding='utf-8')
        new = replace_values(old, {'authentication': {'passwordless_mode': value}})
        tmp = path.parent / ('.auth-' + secrets.token_hex(8) + '.tmp')
        private_write(tmp, new)
        try:
            os.replace(tmp, path)
            try:
                return load_settings(settings.data_dir)
            except Exception:
                restore = path.parent / ('.auth-restore-' + secrets.token_hex(8) + '.tmp')
                private_write(restore, old)
                os.replace(restore, path)
                raise
        finally:
            tmp.unlink(missing_ok=True)


def is_loopback_request(address: str | None) -> bool:
    import ipaddress
    try:
        return ipaddress.ip_address(address or '').is_loopback
    except ValueError:
        return False


def rotate_token(settings, kind: str) -> str:
    """Durable, atomic immediate rotation with live collector/config hash update.

    Caller returns the new token only on successful fsync + replacement.
    """
    if kind not in ('ingestion', 'config'):
        raise ConfigError('Unsupported token role.')
    with _LOCK:
        new_token = secrets.token_urlsafe(32)
        updated = dict(settings.credentials)
        updated[kind + '_token_hash'] = token_hash(new_token)
        updated[kind + '_token_ciphertext'] = encrypt_secret(updated, new_token)
        temp = settings.data_dir / ('.secrets-rotation-' + secrets.token_hex(8) + '.tmp')
        try:
            private_write(temp, json.dumps(updated, indent=2) + '\n')
            os.replace(temp, settings.data_dir / 'secrets.json')
            settings.credentials.clear()
            settings.credentials.update(updated)
        finally:
            temp.unlink(missing_ok=True)
        return new_token


LIVE_FIELDS = {
    'imports': {'counter_paths': 'counter_paths', 'log_paths': 'log_paths',
                'recursive': 'recursive', 'scan_seconds': 'scan_seconds',
                'maximum_file_bytes': 'maximum_file_bytes',
                'maximum_files_per_scan': 'maximum_files_per_scan'},
    'publishing': {'source_directory': 'publishing_source_directory'},
    'billing': {'unit_price': 'unit_price', 'currency': 'currency'},
    'retention': {'cutoff_date': 'cutoff_date', 'keep_months': 'keep_months',
                  'auto_cleanup': 'auto_cleanup'},
}


def apply_live_settings(current, validated, updates: dict):
    """Only rebind reload-safe fields; other fields need a process restart.

    The returned Settings is immutable, so the receiver, Store and scanner can
    be switched to the same consistent snapshot without rewriting each field.
    """
    fields = {field: getattr(validated, field)
              for section, values in updates.items()
              for field in (LIVE_FIELDS.get(section, {}).get(key) for key in values)
              if field is not None}
    return (replace(current, **fields) if fields else current), sorted(fields)
