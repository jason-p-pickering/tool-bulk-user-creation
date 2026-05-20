# DHIS2 Bulk User Creation

Inserts users from an Excel spreadsheet directly into a DHIS2 PostgreSQL database via SQL, bypassing the HTTP API.

Targets the DHIS2 PostgreSQL schema (2.36–2.41). Passwords are hashed with bcrypt.

## Setup

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## Usage

**Generate a SQL file for review (default):**
```bash
.venv/bin/python create_users.py --xlsx LGA_EMIS_Logins_26_states.xlsx
```
Writes `dhis2_users.sql`. Inspect it before applying.

**Generate and execute directly against the database:**
```bash
.venv/bin/python create_users.py \
    --xlsx LGA_EMIS_Logins_26_states.xlsx \
    --execute "host=<host> dbname=<db> user=<user> password=<password>"
```

**Options:**

| Flag | Default | Description |
|---|---|---|
| `--xlsx` | `LGA_EMIS_Logins_26_states.xlsx` | Source spreadsheet |
| `--output` | `dhis2_users.sql` | SQL file to write |
| `--execute` | *(unset)* | PostgreSQL DSN — if set, executes the SQL after writing the file |

## Spreadsheet format

| Column | Description |
|---|---|
| `firstName` | User first name |
| `surname` | User surname |
| `username` | DHIS2 username (must be unique) |
| `password` | Plaintext password (will be bcrypt-hashed) |
| `userRoles` | JSON array of `{"id": "<uid>"}` objects |
| `organisationUnits` | Data-capture org units |
| `dataViewOrganisationUnits` | Data-view org units |
| `teiSearchOrganisationUnits` | TEI search org units |
| `userGroups` | JSON array of `{"id": "<uid>"}` objects |

## Notes

- Re-running is safe — every insert uses `ON CONFLICT DO NOTHING`.
- `userinfoid` is drawn from `hibernate_sequence`, keeping Hibernate's ID counter consistent.
- The SQL file is always written before any database connection is attempted.
