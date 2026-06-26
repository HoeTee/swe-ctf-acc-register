#!/usr/bin/env python3
import argparse
import json
import os
import re
import sys
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import httpx
import yaml
from openpyxl import load_workbook


DEFAULT_CONFIG = Path("config.yaml")


@dataclass
class RegistrationRow:
    row: int
    name: str
    phone: str
    unit: str
    department: str
    username: str
    email: str
    password: str


def load_config(path: Path) -> dict[str, Any]:
    config = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    config["_config_path"] = str(path.resolve())
    config["_base_dir"] = str(path.resolve().parent)
    return config


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return text[:-2] if text.endswith(".0") else text


def normalize_phone(value: Any) -> str:
    return re.sub(r"\D+", "", normalize_text(value))


def normalize_name(value: Any) -> str:
    return re.sub(r"\s+", "", normalize_text(value))


def resolve_path(path_value: str, base_dir: Path) -> Path:
    path = Path(path_value)
    if path.is_absolute():
        return path
    return base_dir / path


def find_header(rows: list[tuple[Any, ...]], required: list[str]) -> tuple[int, list[str]]:
    for index, row in enumerate(rows[:30]):
        headers = [normalize_text(value) for value in row]
        if all(any(req == h or req in h for h in headers) for req in required):
            return index, headers
    raise RuntimeError(f"cannot find header row with columns: {required}")


def column_index(headers: list[str], name: str) -> int:
    if name in headers:
        return headers.index(name)
    matches = [index for index, value in enumerate(headers) if name in value]
    if len(matches) == 1:
        return matches[0]
    raise RuntimeError(f"cannot locate column {name!r}; headers={headers!r}")


def load_registration(config: dict[str, Any]) -> tuple[list[RegistrationRow], list[dict[str, Any]]]:
    base_dir = Path(config["_base_dir"])
    reg_cfg = config["registration"]
    account_cfg = config.get("account", {})
    xlsx_path = resolve_path(reg_cfg["xlsx_path"], base_dir)
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[reg_cfg.get("sheet_name")] if reg_cfg.get("sheet_name") else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    required = [
        reg_cfg["name_column"],
        reg_cfg["unit_column"],
        reg_cfg["department_column"],
        reg_cfg["phone_column"],
    ]
    header_idx, headers = find_header(rows, required)
    name_idx = column_index(headers, reg_cfg["name_column"])
    unit_idx = column_index(headers, reg_cfg["unit_column"])
    dept_idx = column_index(headers, reg_cfg["department_column"])
    phone_idx = column_index(headers, reg_cfg["phone_column"])

    email_domain = str(account_cfg.get("email_domain", "zjrcu.com")).strip().lstrip("@")
    accounts: list[RegistrationRow] = []
    warnings: list[dict[str, Any]] = []
    seen_usernames: dict[str, int] = {}

    for row_number, row in enumerate(rows[header_idx + 1 :], start=header_idx + 2):
        name = normalize_name(row[name_idx] if name_idx < len(row) else None)
        phone = normalize_phone(row[phone_idx] if phone_idx < len(row) else None)
        unit = normalize_text(row[unit_idx] if unit_idx < len(row) else None)
        department = normalize_text(row[dept_idx] if dept_idx < len(row) else None)
        if not name and not phone:
            continue
        if not name:
            warnings.append({"type": "missing_name", "row": row_number, "phone": phone})
            continue
        if len(phone) < 4:
            warnings.append({"type": "invalid_phone", "row": row_number, "name": name, "phone": phone})
            continue
        username = f"{phone}@{email_domain}"
        if username in seen_usernames:
            warnings.append(
                {
                    "type": "duplicate_username",
                    "row": row_number,
                    "first_row": seen_usernames[username],
                    "username": username,
                }
            )
            continue
        seen_usernames[username] = row_number
        accounts.append(
            RegistrationRow(
                row=row_number,
                name=name,
                phone=phone,
                unit=unit,
                department=department,
                username=username,
                email=username,
                password=phone[-4:],
            )
        )

    return accounts, warnings


class GZCTFClient:
    def __init__(self, cluster: dict[str, Any]):
        self.cluster = cluster
        self.name = str(cluster["name"])
        self.base_url = str(cluster["base_url"]).rstrip("/")
        self.client = httpx.Client(base_url=self.base_url, timeout=float(cluster.get("timeout_seconds", 15)))

    def close(self) -> None:
        self.client.close()

    def admin_credentials(self) -> tuple[str, str]:
        username = self.cluster.get("admin_username")
        password = self.cluster.get("admin_password")
        if self.cluster.get("admin_username_env"):
            username = os.environ.get(str(self.cluster["admin_username_env"]))
        if self.cluster.get("admin_password_env"):
            password = os.environ.get(str(self.cluster["admin_password_env"]))
        if not username or not password:
            raise RuntimeError(
                f"missing admin credentials for {self.name}; set admin_username/password or env vars"
            )
        return str(username), str(password)

    def login_admin(self) -> None:
        username, password = self.admin_credentials()
        response = self.client.post("/api/account/login", json={"userName": username, "password": password})
        response.raise_for_status()

    def get_platform_config(self) -> dict[str, Any]:
        response = self.client.get("/api/admin/config")
        response.raise_for_status()
        payload = response.json()
        if not isinstance(payload, dict):
            raise RuntimeError(f"unexpected /api/admin/config response: {payload!r}")
        return payload

    def update_platform_config(self, config_payload: dict[str, Any], method: str = "auto") -> str:
        methods = ["put", "patch", "post"] if method == "auto" else [method.lower()]
        last_error = None
        for candidate in methods:
            request = getattr(self.client, candidate, None)
            if request is None:
                continue
            response = request("/api/admin/config", json=config_payload)
            if response.status_code < 400:
                return candidate.upper()
            last_error = f"{candidate.upper()} /api/admin/config -> HTTP {response.status_code}: {response.text[:500]}"
            if response.status_code not in {404, 405, 415}:
                response.raise_for_status()
        raise RuntimeError(last_error or "no config update method tried")

    def register_account(self, account: RegistrationRow) -> dict[str, Any]:
        payload = {
            "challenge": None,
            "userName": account.username,
            "password": account.password,
            "email": account.email,
        }
        response = self.client.post("/api/account/register", json=payload)
        text = response.text
        if response.status_code in {200, 201, 204}:
            return {"status": "created", "http_status": response.status_code}
        lowered = text.lower()
        if response.status_code in {400, 409} and any(
            marker in lowered for marker in ["duplicate", "already", "exist", "taken"]
        ):
            return {"status": "already_exists", "http_status": response.status_code, "message": text[:500]}
        return {"status": "failed", "http_status": response.status_code, "message": text[:1000]}


def account_to_public_dict(account: RegistrationRow, include_password: bool = False) -> dict[str, Any]:
    value = {
        "row": account.row,
        "name": account.name,
        "phone": account.phone,
        "unit": account.unit,
        "department": account.department,
        "username": account.username,
        "email": account.email,
    }
    if include_password:
        value["password"] = account.password
    return value


def build_enabled_registration_config(original: dict[str, Any], updates: dict[str, Any]) -> dict[str, Any]:
    payload = deepcopy(original)
    for key, value in updates.items():
        payload[key] = value
    return payload


def run_cluster(
    cluster: dict[str, Any],
    accounts: list[RegistrationRow],
    registration_settings: dict[str, Any],
    dry_run: bool,
    limit: int | None,
    keep_register_open: bool,
) -> dict[str, Any]:
    selected_accounts = accounts[:limit] if limit else accounts
    result = {
        "cluster": cluster["name"],
        "base_url": cluster["base_url"],
        "dry_run": dry_run,
        "total_planned": len(selected_accounts),
        "created": 0,
        "already_exists": 0,
        "failed": 0,
        "items": [],
        "config_update_method": None,
        "config_restored": False,
    }

    if dry_run:
        result["items"] = [
            {**account_to_public_dict(account, include_password=True), "status": "planned"}
            for account in selected_accounts
        ]
        return result

    client = GZCTFClient(cluster)
    original_config = None
    try:
        client.login_admin()
        original_config = client.get_platform_config()
        enabled_config = build_enabled_registration_config(original_config, registration_settings)
        result["config_update_method"] = client.update_platform_config(
            enabled_config,
            method=str(cluster.get("config_update_method", "auto")),
        )
        for account in selected_accounts:
            register_result = client.register_account(account)
            status = register_result["status"]
            if status == "created":
                result["created"] += 1
            elif status == "already_exists":
                result["already_exists"] += 1
            else:
                result["failed"] += 1
            result["items"].append(
                {
                    **account_to_public_dict(account, include_password=True),
                    **register_result,
                }
            )
    finally:
        if original_config is not None and not keep_register_open:
            try:
                client.update_platform_config(
                    original_config,
                    method=str(cluster.get("config_update_method", "auto")),
                )
                result["config_restored"] = True
            except Exception as exc:
                result["config_restore_error"] = str(exc)
        client.close()

    return result


def main() -> None:
    parser = argparse.ArgumentParser(description="Batch-register GZCTF accounts from the AI marathon registration Excel.")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    parser.add_argument("--dry-run", action="store_true", help="Generate accounts without calling GZCTF APIs.")
    parser.add_argument("--check-excel", action="store_true", help="Only read Excel and print generated account summary.")
    parser.add_argument("--limit", type=int, help="Only process the first N accounts.")
    parser.add_argument("--keep-register-open", action="store_true", help="Do not restore the original platform config.")
    parser.add_argument("--print-passwords", action="store_true", help="Print generated passwords in --check-excel output.")
    args = parser.parse_args()

    config = load_config(args.config)
    accounts, warnings = load_registration(config)
    output = {
        "source": config["registration"]["xlsx_path"],
        "total_accounts": len(accounts),
        "warnings": warnings,
    }

    if args.check_excel:
        output["sample"] = [
            account_to_public_dict(account, include_password=args.print_passwords)
            for account in accounts[:20]
        ]
        print(json.dumps(output, ensure_ascii=False, indent=2))
        sys.exit(0 if not warnings else 1)

    registration_settings = config.get("registration_settings") or {}
    clusters = config.get("gzctf", {}).get("clusters") or []
    if not clusters:
        raise SystemExit("config.gzctf.clusters is empty")

    output["clusters"] = []
    for cluster in clusters:
        output["clusters"].append(
            run_cluster(
                cluster=cluster,
                accounts=accounts,
                registration_settings=registration_settings,
                dry_run=args.dry_run,
                limit=args.limit,
                keep_register_open=args.keep_register_open,
            )
        )

    result_file = config.get("output", {}).get("result_file")
    if result_file:
        path = resolve_path(result_file, Path(config["_base_dir"]))
        path.write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
