import json
import os
from pathlib import Path
from datetime import datetime, timedelta

_ENV_PATH = Path(__file__).with_name(".env")
_VISIBLE_TASK_RUNS_WHERE = "(COALESCE(success_count, 0) > 0 OR COALESCE(no_graph_count, 0) > 0)"


def _env_file_values():
    out = {}
    if not _ENV_PATH.is_file():
        return out
    for line in _ENV_PATH.read_text(encoding="utf-8", errors="replace").splitlines():
        s = line.strip()
        if not s or s.startswith("#") or "=" not in s:
            continue
        k, v = s.split("=", 1)
        out[k.strip()] = v.strip().strip('"').strip("'")
    return out


def _cfg_value(key, default=""):
    val = os.environ.get(key)
    if val not in (None, ""):
        return val
    return _env_file_values().get(key, default)


def _cfg():
    return {
        "host": (_cfg_value("WEBUI_MYSQL_HOST", "") or "").strip(),
        "port": int((_cfg_value("WEBUI_MYSQL_PORT", "3306") or "3306").strip() or "3306"),
        "user": (_cfg_value("WEBUI_MYSQL_USER", "") or "").strip(),
        "password": _cfg_value("WEBUI_MYSQL_PASSWORD", "") or "",
        "database": (_cfg_value("WEBUI_MYSQL_DATABASE", "") or "").strip(),
        "charset": (_cfg_value("WEBUI_MYSQL_CHARSET", "utf8mb4") or "utf8mb4").strip() or "utf8mb4",
    }


def is_configured():
    cfg = _cfg()
    return bool(cfg["host"] and cfg["user"] and cfg["database"])


def _connect():
    if not is_configured():
        raise RuntimeError("MySQL 未配置")
    try:
        import pymysql
        from pymysql.cursors import DictCursor
    except Exception as exc:
        raise RuntimeError("缺少 PyMySQL 依赖") from exc
    cfg = _cfg()
    return pymysql.connect(
        host=cfg["host"],
        port=cfg["port"],
        user=cfg["user"],
        password=cfg["password"],
        database=cfg["database"],
        charset=cfg["charset"],
        autocommit=True,
        cursorclass=DictCursor,
    )


def ensure_schema():
    if not is_configured():
        return False
    sql_runs = """
    CREATE TABLE IF NOT EXISTS task_runs (
      id BIGINT PRIMARY KEY AUTO_INCREMENT,
      run_id VARCHAR(64) NOT NULL UNIQUE,
      script_id VARCHAR(128) NOT NULL,
      script_title VARCHAR(255) NOT NULL DEFAULT '',
      args_json LONGTEXT NULL,
      status VARCHAR(32) NOT NULL DEFAULT 'running',
      started_at DATETIME NOT NULL,
      ended_at DATETIME NULL,
      duration_seconds DECIMAL(12,2) NULL,
      success_count INT NOT NULL DEFAULT 0,
      fail_count INT NOT NULL DEFAULT 0,
      no_graph_count INT NOT NULL DEFAULT 0,
      total_count INT NOT NULL DEFAULT 0,
      avg_success_duration_seconds DECIMAL(12,2) NULL,
      log_file_path TEXT NULL,
      created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
      INDEX idx_task_runs_started_at (started_at),
      INDEX idx_task_runs_status (status)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """
    sql_accounts = """
    CREATE TABLE IF NOT EXISTS task_accounts (
      id BIGINT PRIMARY KEY AUTO_INCREMENT,
      task_run_id BIGINT NULL,
      email VARCHAR(255) NOT NULL,
      password VARCHAR(255) NOT NULL DEFAULT '',
      client_id TEXT NULL,
      refresh_token LONGTEXT NULL,
      generated_at DATETIME NOT NULL,
      register_ip VARCHAR(64) NULL,
      register_region VARCHAR(255) NULL,
      status VARCHAR(32) NOT NULL DEFAULT 'imported',
      source VARCHAR(32) NOT NULL DEFAULT 'import',
      created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
      INDEX idx_task_accounts_task_run_id (task_run_id),
      INDEX idx_task_accounts_generated_at (generated_at),
      INDEX idx_task_accounts_status (status),
      INDEX idx_task_accounts_email (email(191))
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """
    sql_fk = """
    ALTER TABLE task_accounts
      ADD CONSTRAINT fk_task_accounts_task_run
      FOREIGN KEY (task_run_id) REFERENCES task_runs(id)
      ON DELETE SET NULL;
    """
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(sql_runs)
            cur.execute(sql_accounts)
            try:
                cur.execute(sql_fk)
            except Exception:
                pass
    return True


def _to_dt(value):
    if isinstance(value, datetime):
        return value
    if not value:
        return datetime.now()
    if isinstance(value, str):
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
        except Exception:
            return datetime.now()
    return datetime.now()


def create_task_run(run_id, script_id, script_title, args, log_file_path, started_at=None):
    ensure_schema()
    payload = json.dumps(args or {}, ensure_ascii=False, sort_keys=True)
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO task_runs
                (run_id, script_id, script_title, args_json, status, started_at, log_file_path)
                VALUES (%s, %s, %s, %s, 'running', %s, %s)
                """,
                (run_id, script_id, script_title or "", payload, _to_dt(started_at), log_file_path or ""),
            )
            return cur.lastrowid


def finish_task_run(
    run_id,
    status,
    ended_at=None,
    duration_seconds=None,
    success_count=0,
    fail_count=0,
    no_graph_count=0,
    total_count=0,
    avg_success_duration_seconds=None,
):
    ensure_schema()
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE task_runs
                   SET status=%s,
                       ended_at=%s,
                       duration_seconds=%s,
                       success_count=%s,
                       fail_count=%s,
                       no_graph_count=%s,
                       total_count=%s,
                       avg_success_duration_seconds=%s
                 WHERE run_id=%s
                """,
                (
                    status,
                    _to_dt(ended_at),
                    duration_seconds,
                    int(success_count or 0),
                    int(fail_count or 0),
                    int(no_graph_count or 0),
                    int(total_count or 0),
                    avg_success_duration_seconds,
                    run_id,
                ),
            )


def _lookup_task_run_id(run_id):
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM task_runs WHERE run_id=%s LIMIT 1", (run_id,))
            row = cur.fetchone()
            return (row or {}).get("id")


def add_task_account(
    run_id,
    email,
    password,
    client_id="",
    refresh_token="",
    generated_at=None,
    register_ip="",
    register_region="",
    status="ok",
    source="runtime",
):
    ensure_schema()
    task_run_id = _lookup_task_run_id(run_id) if run_id else None
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO task_accounts
                (task_run_id, email, password, client_id, refresh_token, generated_at,
                 register_ip, register_region, status, source)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    task_run_id,
                    email,
                    password or "",
                    client_id or "",
                    refresh_token or "",
                    _to_dt(generated_at),
                    register_ip or "",
                    register_region or "",
                    status or "ok",
                    source or "runtime",
                ),
            )


def import_accounts(rows):
    ensure_schema()
    saved = 0
    with _connect() as conn:
        with conn.cursor() as cur:
            for row in rows or []:
                cur.execute(
                    """
                    INSERT INTO task_accounts
                    (task_run_id, email, password, client_id, refresh_token, generated_at,
                     register_ip, register_region, status, source)
                    VALUES (NULL, %s, %s, %s, %s, %s, %s, %s, 'imported', 'import')
                    """,
                    (
                        row.get("email") or "",
                        row.get("password") or "",
                        row.get("client_id") or "",
                        row.get("refresh_token") or "",
                        _to_dt(row.get("generated_at")),
                        row.get("register_ip") or "",
                        row.get("register_region") or "",
                    ),
                )
                saved += 1
    return saved


def list_task_runs(limit=None, page=1, page_size=20):
    ensure_schema()
    page = max(1, int(page or 1))
    if limit:
        page_size = max(1, int(limit))
    else:
        page_size = int(page_size or 20)
        if page_size not in (20, 50, 100):
            page_size = 20
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(f"SELECT COUNT(*) AS total FROM task_runs WHERE {_VISIBLE_TASK_RUNS_WHERE}")
            total = int((cur.fetchone() or {}).get("total") or 0)
            cur.execute(
                f"""
                SELECT id, run_id, script_id, script_title, args_json, status,
                       started_at, ended_at, duration_seconds,
                       success_count, fail_count, no_graph_count, total_count,
                       avg_success_duration_seconds, log_file_path
                  FROM task_runs
                 WHERE {_VISIBLE_TASK_RUNS_WHERE}
              ORDER BY id DESC
                 LIMIT %s OFFSET %s
                """,
                (page_size, (page - 1) * page_size),
            )
            return {
                "items": cur.fetchall() or [],
                "total": total,
                "page": page,
                "page_size": page_size,
            }


def get_task_run(task_id):
    ensure_schema()
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id, run_id, script_id, script_title, args_json, status,
                       started_at, ended_at, duration_seconds,
                       success_count, fail_count, no_graph_count, total_count,
                       avg_success_duration_seconds, log_file_path
                  FROM task_runs
                 WHERE id=%s
                 LIMIT 1
                """,
                (int(task_id),),
            )
            return cur.fetchone()


def delete_task_run(task_id):
    ensure_schema()
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM task_runs WHERE id=%s", (int(task_id),))
            return cur.rowcount > 0


def list_accounts(
    limit=None,
    task_run_id=None,
    email="",
    status="",
    ids=None,
    page=1,
    page_size=20,
    sort_by="generated_at",
    sort_dir="desc",
):
    ensure_schema()
    where = []
    params = []
    if task_run_id:
        where.append("task_run_id=%s")
        params.append(int(task_run_id))
    if email:
        where.append("LOWER(email) LIKE %s")
        params.append(f"%{str(email).strip().lower()}%")
    if status:
        where.append("status=%s")
        params.append(str(status).strip())
    if ids:
        picked = [int(x) for x in (ids or []) if str(x).strip()]
        if picked:
            where.append("id IN (" + ",".join(["%s"] * len(picked)) + ")")
            params.extend(picked)
    safe_sort_by = "created_at" if str(sort_by or "").strip().lower() != "generated_at" else "generated_at"
    safe_sort_dir = "ASC" if str(sort_dir or "").strip().lower() == "asc" else "DESC"
    page = max(1, int(page or 1))
    if limit:
        page_size = max(1, int(limit))
    else:
        page_size = int(page_size or 20)
        if page_size not in (20, 50, 100):
            page_size = 20
    sql = """
        SELECT id, task_run_id, email, password, client_id, refresh_token,
               generated_at, register_ip, register_region, status, source, created_at
          FROM task_accounts
    """
    if where:
        sql += " WHERE " + " AND ".join(where)
    count_sql = "SELECT COUNT(*) AS total FROM task_accounts"
    if where:
        count_sql += " WHERE " + " AND ".join(where)
    sql += f" ORDER BY {safe_sort_by} {safe_sort_dir}, id DESC LIMIT %s OFFSET %s"
    count_params = tuple(params)
    params.extend([page_size, (page - 1) * page_size])
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(count_sql, count_params)
            total = int((cur.fetchone() or {}).get("total") or 0)
            cur.execute(sql, tuple(params))
            return {
                "items": cur.fetchall() or [],
                "total": total,
                "page": page,
                "page_size": page_size,
                "sort_by": safe_sort_by,
                "sort_dir": safe_sort_dir.lower(),
            }


def get_account(account_id):
    ensure_schema()
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id, task_run_id, email, password, client_id, refresh_token,
                       generated_at, register_ip, register_region, status, source, created_at
                  FROM task_accounts
                 WHERE id=%s
                 LIMIT 1
                """,
                (int(account_id),),
            )
            return cur.fetchone()


def update_account(account_id, fields):
    ensure_schema()
    allowed = ("email", "password", "client_id", "refresh_token")
    updates = []
    params = []
    for key in allowed:
        if key in (fields or {}):
            updates.append(f"{key}=%s")
            params.append(str((fields or {}).get(key) or "").strip())
    if not updates:
        return get_account(account_id)
    params.append(int(account_id))
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                f"UPDATE task_accounts SET {', '.join(updates)} WHERE id=%s",
                tuple(params),
            )
    return get_account(account_id)


def get_overview(days=7):
    ensure_schema()
    days = max(1, int(days or 7))
    start_day = (datetime.now() - timedelta(days=days - 1)).date()
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                f"""
                SELECT COUNT(*) AS total_tasks,
                       COALESCE(SUM(CASE WHEN status='running' THEN 1 ELSE 0 END), 0) AS running_tasks,
                       COALESCE(SUM(success_count + no_graph_count), 0) AS total_success_accounts,
                       COALESCE(AVG(duration_seconds), 0) AS avg_task_duration_seconds,
                       COALESCE(AVG(NULLIF(avg_success_duration_seconds, 0)), 0) AS avg_success_duration_seconds
                  FROM task_runs
                 WHERE {_VISIBLE_TASK_RUNS_WHERE}
                """
            )
            cards = cur.fetchone() or {}
            cur.execute(
                f"""
                SELECT DATE(started_at) AS day,
                       COUNT(*) AS task_count,
                       COALESCE(SUM(success_count + no_graph_count), 0) AS success_count
                  FROM task_runs
                 WHERE started_at >= %s
                   AND {_VISIBLE_TASK_RUNS_WHERE}
              GROUP BY DATE(started_at)
              ORDER BY day ASC
                """,
                (start_day,),
            )
            rows = cur.fetchall() or []
            cur.execute(
                """
                SELECT COALESCE(NULLIF(register_region, ''), 'UNKNOWN') AS region,
                       COUNT(*) AS total
                  FROM task_accounts
                 WHERE source='runtime'
              GROUP BY COALESCE(NULLIF(register_region, ''), 'UNKNOWN')
              ORDER BY total DESC
                 LIMIT 8
                """
            )
            top_regions = cur.fetchall() or []
    trend_map = {str(r["day"]): r for r in rows}
    task_trend = []
    success_trend = []
    for i in range(days):
        day = str(start_day + timedelta(days=i))
        item = trend_map.get(day, {})
        task_trend.append({"day": day, "value": int(item.get("task_count") or 0)})
        success_trend.append({"day": day, "value": int(item.get("success_count") or 0)})
    return {
        "cards": {
            "total_tasks": int(cards.get("total_tasks") or 0),
            "running_tasks": int(cards.get("running_tasks") or 0),
            "total_success_accounts": int(cards.get("total_success_accounts") or 0),
            "avg_task_duration_seconds": float(cards.get("avg_task_duration_seconds") or 0),
            "avg_success_duration_seconds": float(cards.get("avg_success_duration_seconds") or 0),
        },
        "task_trend": task_trend,
        "success_trend": success_trend,
        "top_regions": [
            {"label": row.get("region") or "UNKNOWN", "value": int(row.get("total") or 0)}
            for row in top_regions
        ],
    }
