#!/usr/bin/env python3
"""
Bulk-insert DHIS2 users from an Excel spreadsheet via direct SQL.
Targets the DHIS2 PostgreSQL schema (tested against 2.36–2.41).

Usage:
    # Generate a SQL file for review:
    python create_users.py --xlsx LGA_EMIS_Logins_26_states.xlsx

    # Generate and execute directly against the database:
    python create_users.py --xlsx LGA_EMIS_Logins_26_states.xlsx \
        --execute "host=localhost dbname=dhis2 user=dhis password=..."
"""

import argparse
import json
import random
import string
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

import bcrypt
import openpyxl

_UID_FIRST = string.ascii_letters
_UID_REST = string.ascii_letters + string.digits


def generate_uid() -> str:
    """Generate a DHIS2-compatible 11-character UID."""
    return random.choice(_UID_FIRST) + "".join(random.choices(_UID_REST, k=10))


def hash_password(plaintext: str) -> str:
    return bcrypt.hashpw(plaintext.encode(), bcrypt.gensalt(rounds=10)).decode()


def parse_id_list(value) -> list[str]:
    if not value:
        return []
    try:
        return [item["id"] for item in json.loads(str(value))]
    except (json.JSONDecodeError, KeyError, TypeError, ValueError):
        return []


def q(s: str) -> str:
    """Escape and single-quote a SQL string literal."""
    return "'" + str(s).replace("'", "''") + "'"


def read_users(path: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = list(rows[0])
    return [
        dict(zip(headers, row))
        for row in rows[1:]
        if any(row)
    ]


def build_statements(user: dict) -> list[str]:
    """Return ordered SQL statements to insert one user and all FK relationships."""
    uid = generate_uid()
    username = (user.get("username") or "").strip()
    firstname = (user.get("firstName") or "").strip()
    surname = (user.get("surname") or "").strip()
    password_hash = hash_password((user.get("password") or "").strip())
    user_uuid = str(uuid.uuid4())

    role_ids = parse_id_list(user.get("userRoles"))
    ou_ids = parse_id_list(user.get("organisationUnits"))
    dv_ids = parse_id_list(user.get("dataViewOrganisationUnits"))
    tei_ids = parse_id_list(user.get("teiSearchOrganisationUnits"))
    ug_ids = parse_id_list(user.get("userGroups"))

    stmts = [f"-- {username}"]

    # Core user record — userinfoid has no default, must use hibernate_sequence
    stmts.append(
        f"INSERT INTO userinfo "
        f"(userinfoid, uid, uuid, created, lastupdated, username, password, "
        f"firstname, surname, disabled, externalauth, selfregistered, invitation, passwordlastupdated) "
        f"VALUES (nextval('hibernate_sequence'), {q(uid)}, {q(user_uuid)}::uuid, NOW(), NOW(), "
        f"{q(username)}, {q(password_hash)}, "
        f"{q(firstname)}, {q(surname)}, "
        f"FALSE, FALSE, FALSE, FALSE, NOW()) "
        f"ON CONFLICT (username) DO NOTHING;"
    )

    # User roles  (column is 'userid', not 'userinfoid')
    for role_uid in role_ids:
        stmts.append(
            f"INSERT INTO userrolemembers (userid, userroleid) "
            f"SELECT ui.userinfoid, ur.userroleid "
            f"FROM userinfo ui JOIN userrole ur ON ur.uid = {q(role_uid)} "
            f"WHERE ui.uid = {q(uid)} "
            f"ON CONFLICT DO NOTHING;"
        )

    # Data-capture org units
    for ou_uid in ou_ids:
        stmts.append(
            f"INSERT INTO usermembership (userinfoid, organisationunitid) "
            f"SELECT ui.userinfoid, ou.organisationunitid "
            f"FROM userinfo ui JOIN organisationunit ou ON ou.uid = {q(ou_uid)} "
            f"WHERE ui.uid = {q(uid)} "
            f"ON CONFLICT DO NOTHING;"
        )

    # Data-view org units
    for ou_uid in dv_ids:
        stmts.append(
            f"INSERT INTO userdatavieworgunits (userinfoid, organisationunitid) "
            f"SELECT ui.userinfoid, ou.organisationunitid "
            f"FROM userinfo ui JOIN organisationunit ou ON ou.uid = {q(ou_uid)} "
            f"WHERE ui.uid = {q(uid)} "
            f"ON CONFLICT DO NOTHING;"
        )

    # TEI search org units
    for ou_uid in tei_ids:
        stmts.append(
            f"INSERT INTO userteisearchorgunits (userinfoid, organisationunitid) "
            f"SELECT ui.userinfoid, ou.organisationunitid "
            f"FROM userinfo ui JOIN organisationunit ou ON ou.uid = {q(ou_uid)} "
            f"WHERE ui.uid = {q(uid)} "
            f"ON CONFLICT DO NOTHING;"
        )

    # User groups  (column is 'userid', not 'memberid')
    for ug_uid in ug_ids:
        stmts.append(
            f"INSERT INTO usergroupmembers (usergroupid, userid) "
            f"SELECT ug.usergroupid, ui.userinfoid "
            f"FROM usergroup ug JOIN userinfo ui ON ui.uid = {q(uid)} "
            f"WHERE ug.uid = {q(ug_uid)} "
            f"ON CONFLICT DO NOTHING;"
        )

    return stmts


def generate_sql_file(users: list[dict]) -> tuple[str, list[list[str]]]:
    """Return (sql_string, per_user_statements) for file output and direct execution."""
    now = datetime.now(timezone.utc).isoformat()
    all_user_stmts: list[list[str]] = []

    header = "\n".join([
        "-- DHIS2 bulk user creation",
        f"-- Generated : {now}",
        f"-- Users     : {len(users)}",
        f"-- Schema    : DHIS2 2.36-2.41 (PostgreSQL)",
        "",
        "BEGIN;",
        "",
    ])

    body_lines = []
    for i, user in enumerate(users, 1):
        username = user.get("username") or ""
        print(f"  [{i:>3}/{len(users)}] hashing {username}", file=sys.stderr)
        stmts = build_statements(user)
        all_user_stmts.append(stmts)
        body_lines.extend(stmts)
        body_lines.append("")

    return header + "\n".join(body_lines) + "\nCOMMIT;\n", all_user_stmts


def execute(all_user_stmts: list[list[str]], dsn: str) -> None:
    import psycopg2

    print("Connecting to database...", file=sys.stderr)
    conn = psycopg2.connect(dsn)
    total = len(all_user_stmts)
    try:
        with conn:
            with conn.cursor() as cur:
                for i, stmts in enumerate(all_user_stmts, 1):
                    username = stmts[0].lstrip("- ").strip()
                    print(f"  [{i:>3}/{total}] inserting {username}", file=sys.stderr)
                    for stmt in stmts:
                        if stmt.startswith("--") or not stmt.strip():
                            continue
                        cur.execute(stmt)
        print(f"Done. {total} users committed.", file=sys.stderr)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Bulk-insert DHIS2 users from Excel via SQL"
    )
    parser.add_argument(
        "--xlsx",
        default="LGA_EMIS_Logins_26_states.xlsx",
        help="Path to the source spreadsheet",
    )
    parser.add_argument(
        "--output",
        default="dhis2_users.sql",
        help="SQL file to write (always written, default: dhis2_users.sql)",
    )
    parser.add_argument(
        "--execute",
        metavar="DSN",
        help=(
            "Also execute directly against PostgreSQL. "
            "Example: 'host=localhost dbname=dhis2 user=dhis password=dhis'"
        ),
    )
    args = parser.parse_args()

    print(f"Reading {args.xlsx}...", file=sys.stderr)
    users = read_users(args.xlsx)
    print(f"Found {len(users)} users. Hashing passwords...", file=sys.stderr)

    sql, all_user_stmts = generate_sql_file(users)

    out = Path(args.output)
    out.write_text(sql, encoding="utf-8")
    print(f"\nSQL written → {out} ({out.stat().st_size:,} bytes)", file=sys.stderr)

    if args.execute:
        execute(all_user_stmts, args.execute)


if __name__ == "__main__":
    main()
