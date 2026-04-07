import csv
from copy import deepcopy
import io
import json
import re
from datetime import date, datetime
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook, load_workbook
from sqlalchemy import func


BASE_DIR = Path(__file__).resolve().parent
ALLOWED_COLUMN_TYPES = {"text", "number", "money", "date", "datetime", "boolean"}
DEFAULT_STATE = {
    "left_panel_width": 280,
    "right_panel_visible": True,
    "errors_visible": True,
    "errors_height": 180,
    "last_template_id": None,
    "last_loaded_file": "",
}
EMPTY_DATASET = {
    "file_name": "",
    "header_row_number": None,
    "columns": [],
    "rows": [],
    "warnings": [],
    "errors": [],
    "partial": False,
    "missing_columns": [],
    "row_count": 0,
    "template": None,
    "next_row_id": 1,
    "can_undo_parse": False,
}
CURRENT_DATASET = deepcopy(EMPTY_DATASET)
LAST_PARSE_SNAPSHOT = None
SUPPORTED_SETTING_KEYS = ("output_folder", "instructions_file")
APP_SESSION_STARTED_AT = datetime.utcnow().replace(microsecond=0)

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///rbx_fakt_manag.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


class Template(db.Model):
    __tablename__ = "templates"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False, unique=True)
    created_at = db.Column(db.DateTime, server_default=func.now(), nullable=False)
    columns = db.relationship(
        "TemplateColumn",
        backref="template",
        cascade="all, delete-orphan",
        order_by="TemplateColumn.position",
        lazy=True,
    )


class TemplateColumn(db.Model):
    __tablename__ = "template_columns"

    id = db.Column(db.Integer, primary_key=True)
    template_id = db.Column(db.Integer, db.ForeignKey("templates.id"), nullable=False)
    column_name = db.Column(db.String(120), nullable=False)
    column_type = db.Column(db.String(40), nullable=False, default="text")
    position = db.Column(db.Integer, nullable=False, default=0)


class AppSetting(db.Model):
    __tablename__ = "app_settings"

    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(120), nullable=False, unique=True)
    value = db.Column(db.Text, nullable=False, default="")


class AppState(db.Model):
    __tablename__ = "app_state"

    id = db.Column(db.Integer, primary_key=True)
    state_key = db.Column(db.String(120), nullable=False, unique=True)
    state_value = db.Column(db.Text, nullable=False, default="")


class EventLog(db.Model):
    __tablename__ = "event_log"

    id = db.Column(db.Integer, primary_key=True)
    event_type = db.Column(db.String(80), nullable=False)
    event_description = db.Column(db.Text, nullable=False, default="")
    created_at = db.Column(db.DateTime, server_default=func.now(), nullable=False)


def serialize_template(template):
    return {
        "id": template.id,
        "name": template.name,
        "created_at": template.created_at.isoformat() if template.created_at else None,
        "columns": [
            {
                "id": column.id,
                "name": column.column_name,
                "type": column.column_type,
                "position": column.position,
            }
            for column in template.columns
        ],
    }


def get_settings_map():
    settings = {key: "" for key in SUPPORTED_SETTING_KEYS}
    for item in AppSetting.query.order_by(AppSetting.key).all():
        if item.key in settings:
            settings[item.key] = item.value
    return settings


def get_state_map():
    state = dict(DEFAULT_STATE)
    for item in AppState.query.all():
        try:
            state[item.state_key] = json.loads(item.state_value)
        except json.JSONDecodeError:
            state[item.state_key] = item.state_value
    return state


def upsert_setting(key, value):
    item = AppSetting.query.filter_by(key=key).first()
    if item is None:
        item = AppSetting(key=key, value=value)
        db.session.add(item)
    else:
        item.value = value


def upsert_state(key, value):
    item = AppState.query.filter_by(state_key=key).first()
    serialized = json.dumps(value, ensure_ascii=False)
    if item is None:
        item = AppState(state_key=key, state_value=serialized)
        db.session.add(item)
    else:
        item.state_value = serialized


def log_event(event_type, description):
    db.session.add(EventLog(event_type=event_type, event_description=description))


def infer_event_level(event_type):
    if event_type.endswith("_failed") or "error" in event_type:
        return "error"
    if event_type.endswith("_partial") or "warning" in event_type:
        return "warning"
    return "info"


def serialize_log_entry(item):
    return {
        "id": item.id,
        "event_level": infer_event_level(item.event_type),
        "event_type": item.event_type,
        "event_description": item.event_description,
        "created_at": item.created_at.isoformat() if item.created_at else None,
    }


def build_logs_payload(limit=80, current_limit=20, history_limit=30):
    entries = EventLog.query.order_by(EventLog.created_at.desc(), EventLog.id.desc()).limit(limit).all()
    current_session = []
    history = []

    for entry in entries:
        target = current_session if entry.created_at and entry.created_at >= APP_SESSION_STARTED_AT else history
        if target is current_session and len(current_session) >= current_limit:
            continue
        if target is history and len(history) >= history_limit:
            continue
        target.append(serialize_log_entry(entry))

        if len(current_session) >= current_limit and len(history) >= history_limit:
            break

    return {
        "session_started_at": APP_SESSION_STARTED_AT.isoformat(),
        "current_session": current_session,
        "history": history,
    }


def read_instructions_file(file_path):
    path = Path(file_path)
    if not path.exists() or not path.is_file():
        raise ValueError("Файл инструкций недоступен. Проверьте путь в настройках.")

    try:
        content = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("Файл инструкций должен быть текстовым в UTF-8.") from exc

    if "\x00" in content:
        raise ValueError("Файл инструкций должен быть plain text.")

    return {
        "file_name": path.name,
        "file_path": str(path),
        "content": content,
    }


def log_dataset_quality(dataset, source_event):
    warning_count = len(dataset["warnings"])
    error_count = len(dataset["errors"])

    if warning_count:
        log_event(
            "system_warning",
            f"{source_event}: получено предупреждений {warning_count}.",
        )

    if error_count:
        log_event(
            "typing_error",
            f"{source_event}: найдено ошибок типизации {error_count}.",
        )


def serialize_dataset():
    payload = deepcopy(CURRENT_DATASET)
    payload.pop("next_row_id", None)
    return payload


def build_dataset_errors(rows, columns):
    errors = []
    for row in rows:
        for column in columns:
            value = row["values"].get(column["name"], "")
            reason = validate_value(value, column["type"])
            if reason:
                errors.append(
                    {
                        "row_number": row["row_number"],
                        "source_row_number": row["source_row_number"],
                        "column_name": column["name"],
                        "column_type": column["type"],
                        "value": value,
                        "reason": reason,
                    }
                )
    return errors


def rebuild_row_numbers(rows):
    for index, row in enumerate(rows, start=1):
        row["row_number"] = index


def split_parse_fragments(value, delimiter):
    return [fragment.strip() for fragment in str(value or "").split(delimiter) if fragment.strip()]


def build_export_path(output_dir):
    base_name = datetime.now().strftime("data_%d%m%y_%H%M")
    candidate = output_dir / f"{base_name}.xlsx"
    suffix = 2

    while candidate.exists():
        candidate = output_dir / f"{base_name}_{suffix}.xlsx"
        suffix += 1

    return candidate


def normalize_header_name(raw_name, index):
    normalized = (raw_name or "").strip()
    if normalized:
        return normalized
    return f"Столбец_{index + 1}"


def uniquify_headers(headers):
    seen = {}
    result = []
    warnings = []

    for index, raw_name in enumerate(headers):
        base_name = normalize_header_name(raw_name, index)
        count = seen.get(base_name.lower(), 0) + 1
        seen[base_name.lower()] = count
        if count == 1:
            unique_name = base_name
        else:
            unique_name = f"{base_name}_{count}"
            warnings.append(
                f"Дублирующийся столбец '{base_name}' переименован в '{unique_name}'."
            )
        if not (raw_name or "").strip():
            warnings.append(f"Пустой заголовок столбца заменён на '{unique_name}'.")
        result.append(unique_name)

    return result, warnings


def parse_decimal(value):
    normalized = str(value).strip()
    if not normalized:
        return None

    if not re.fullmatch(r"[+-]?(?:\d+|\d{1,3}(?: \d{3})+)(?:\.\d+)?", normalized):
        raise ValueError("invalid decimal")

    return float(normalized.replace(" ", ""))


def normalize_cell_value(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y %H:%M")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    return str(value).strip()


def load_csv_rows(file_storage):
    try:
        decoded = file_storage.read().decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("CSV-файл должен быть в кодировке UTF-8 with BOM.") from exc

    return [list(row) for row in csv.reader(io.StringIO(decoded), delimiter=";")]


def load_xlsx_rows(file_storage):
    try:
        workbook = load_workbook(filename=io.BytesIO(file_storage.read()), read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError("Не удалось прочитать XLSX-файл.") from exc

    worksheet = workbook.active
    try:
        return [
            [normalize_cell_value(cell) for cell in row]
            for row in worksheet.iter_rows(values_only=True)
        ]
    finally:
        workbook.close()


def parse_boolean(value):
    normalized = str(value).strip().lower()
    if normalized in {"true", "yes", "да", "1", "истина"}:
        return True
    if normalized in {"false", "no", "нет", "0", "ложь"}:
        return False
    raise ValueError("invalid boolean")


def validate_value(value, column_type):
    raw_value = (value or "").strip()
    if column_type == "text":
        return None

    try:
        if column_type in {"number", "money"}:
            if not raw_value:
                return "пустое значение трактуется как 0"
            parse_decimal(raw_value)
            return None
        if column_type == "boolean":
            if not raw_value:
                return "пустое значение трактуется как Нет"
            parse_boolean(raw_value)
            return None
        if column_type == "date":
            if not raw_value:
                return "пустое значение трактуется как 01.01.1900"
            for pattern in ("%Y-%m-%d", "%d.%m.%Y"):
                try:
                    datetime.strptime(raw_value, pattern)
                    return None
                except ValueError:
                    continue
            return "ожидалась дата"
        if column_type == "datetime":
            if not raw_value:
                return "ожидались дата и время"
            for pattern in (
                "%d.%m.%Y %H:%M",
                "%Y-%m-%d %H:%M",
                "%Y-%m-%dT%H:%M:%S",
            ):
                try:
                    datetime.strptime(raw_value, pattern)
                    return None
                except ValueError:
                    continue
            return "ожидались дата и время"
    except ValueError:
        if column_type in {"number", "money"}:
            return "ожидалось числовое значение"
        if column_type == "boolean":
            return "ожидалось логическое значение"

    return None


def build_dataset(file_storage, template):
    if file_storage is None:
        raise ValueError("Файл не выбран.")

    file_name = Path(file_storage.filename or "").name
    if not file_name:
        raise ValueError("Файл не выбран.")
    suffix = Path(file_name).suffix.lower()
    if suffix == ".csv":
        source_rows = load_csv_rows(file_storage)
    elif suffix == ".xlsx":
        source_rows = load_xlsx_rows(file_storage)
    else:
        raise ValueError("Поддерживаются только CSV- и XLSX-файлы.")

    header = None
    header_row_number = None
    warnings = []
    rows = []

    for row_number, row in enumerate(source_rows, start=1):
        if header is None and not any((cell or "").strip() for cell in row):
            continue
        if header is None:
            header_row_number = row_number
            header, header_warnings = uniquify_headers(row)
            warnings.extend(header_warnings)
            continue

        values = list(row)
        if len(values) < len(header):
            values.extend([""] * (len(header) - len(values)))
        elif len(values) > len(header):
            warnings.append(
                f"Строка {row_number}: лишние значения после столбца '{header[-1]}' были отброшены."
            )
            values = values[: len(header)]

        rows.append(
            {
                "row_id": len(rows) + 1,
                "row_number": len(rows) + 1,
                "source_row_number": row_number,
                "values": {column_name: values[index] for index, column_name in enumerate(header)},
            }
        )

    if header is None:
        raise ValueError("В файле не найден заголовок.")

    template_columns = template.columns if template else []
    template_types = {column.column_name.lower(): column.column_type for column in template_columns}
    expected_columns = [column.column_name for column in template_columns]
    missing_columns = [name for name in expected_columns if name.lower() not in {item.lower() for item in header}]
    partial = bool(missing_columns)
    if missing_columns:
        warnings.append(
            "Частичная загрузка: отсутствуют ожидаемые столбцы шаблона: " + ", ".join(missing_columns)
        )

    known_template_columns = {column.column_name.lower() for column in template_columns}
    for column_name in header:
        if template and column_name.lower() not in known_template_columns:
            warnings.append(
                f"Новый столбец '{column_name}' не найден в шаблоне и получил тип text по умолчанию."
            )

    columns = []
    for position, column_name in enumerate(header):
        column_type = template_types.get(column_name.lower(), "text")
        columns.append(
            {
                "name": column_name,
                "type": column_type,
                "position": position,
                "from_template": column_name.lower() in template_types,
            }
        )

    errors = build_dataset_errors(rows, columns)

    return {
        "file_name": file_name,
        "header_row_number": header_row_number,
        "columns": columns,
        "rows": rows,
        "warnings": warnings,
        "errors": errors,
        "partial": partial,
        "missing_columns": missing_columns,
        "row_count": len(rows),
        "template": {"id": template.id, "name": template.name} if template else None,
        "next_row_id": len(rows) + 1,
        "can_undo_parse": False,
    }


def validate_settings(payload):
    errors = {}
    folder_fields = {
        "output_folder": "Папка результата",
    }
    file_fields = {
        "instructions_file": "Файл инструкций",
    }

    for key, label in folder_fields.items():
        value = (payload.get(key) or "").strip()
        if value and (not Path(value).exists() or not Path(value).is_dir()):
            errors[key] = f"{label} должна существовать и быть папкой."

    for key, label in file_fields.items():
        value = (payload.get(key) or "").strip()
        if value and (not Path(value).exists() or not Path(value).is_file()):
            errors[key] = f"{label} должен существовать и быть файлом."

    return errors


def normalize_columns(raw_columns):
    columns = []
    errors = []
    seen_names = set()

    for index, item in enumerate(raw_columns or []):
        name = (item.get("name") or "").strip()
        column_type = (item.get("type") or "text").strip().lower()
        if not name:
            errors.append(f"Столбец #{index + 1}: укажите название.")
            continue
        if name.lower() in seen_names:
            errors.append(f"Столбец '{name}' указан несколько раз.")
            continue
        if column_type not in ALLOWED_COLUMN_TYPES:
            errors.append(f"Столбец '{name}': тип '{column_type}' не поддерживается.")
            continue
        seen_names.add(name.lower())
        columns.append({"name": name, "type": column_type, "position": len(columns)})

    return columns, errors


@app.route("/")
def home():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/styles.css")
def styles():
    return send_from_directory(BASE_DIR, "styles.css")


@app.route("/app.js")
def script():
    return send_from_directory(BASE_DIR, "app.js")


@app.get("/api/bootstrap")
def bootstrap():
    return jsonify(
        {
            "settings": get_settings_map(),
            "templates": [serialize_template(item) for item in Template.query.order_by(Template.name).all()],
            "state": get_state_map(),
            "dataset": serialize_dataset(),
            "logs": build_logs_payload(),
        }
    )


@app.get("/api/settings")
def get_settings():
    return jsonify(get_settings_map())


@app.put("/api/settings")
def save_settings():
    payload = request.get_json(silent=True) or {}
    errors = validate_settings(payload)
    if errors:
        log_event("settings_failed", "Ошибка валидации настроек приложения.")
        db.session.commit()
        return jsonify({"ok": False, "errors": errors}), 400

    for key in SUPPORTED_SETTING_KEYS:
        upsert_setting(key, (payload.get(key) or "").strip())

    log_event("settings_saved", "Настройки приложения обновлены.")
    db.session.commit()
    return jsonify({"ok": True, "settings": get_settings_map()})


@app.get("/api/templates")
def list_templates():
    templates = Template.query.order_by(Template.name).all()
    return jsonify([serialize_template(item) for item in templates])


@app.post("/api/templates")
def create_template():
    payload = request.get_json(silent=True) or {}
    name = (payload.get("name") or "").strip()
    columns, column_errors = normalize_columns(payload.get("columns") or [])

    errors = []
    if not name:
        errors.append("Укажите имя шаблона.")
    if Template.query.filter(func.lower(Template.name) == name.lower()).first():
        errors.append("Шаблон с таким именем уже существует.")
    errors.extend(column_errors)
    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    template = Template(name=name)
    db.session.add(template)
    db.session.flush()

    for column in columns:
        db.session.add(
            TemplateColumn(
                template_id=template.id,
                column_name=column["name"],
                column_type=column["type"],
                position=column["position"],
            )
        )

    log_event("template_created", f"Создан шаблон '{name}'.")
    db.session.commit()
    return jsonify({"ok": True, "template": serialize_template(template)}), 201


@app.put("/api/templates/<int:template_id>")
def update_template(template_id):
    template = Template.query.get_or_404(template_id)
    payload = request.get_json(silent=True) or {}
    name = (payload.get("name") or "").strip()
    columns, column_errors = normalize_columns(payload.get("columns") or [])

    errors = []
    if not name:
        errors.append("Укажите имя шаблона.")

    duplicate = Template.query.filter(func.lower(Template.name) == name.lower(), Template.id != template_id).first()
    if duplicate:
        errors.append("Шаблон с таким именем уже существует.")

    errors.extend(column_errors)
    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    template.name = name
    TemplateColumn.query.filter_by(template_id=template.id).delete()
    for column in columns:
        db.session.add(
            TemplateColumn(
                template_id=template.id,
                column_name=column["name"],
                column_type=column["type"],
                position=column["position"],
            )
        )

    log_event("template_updated", f"Обновлён шаблон '{name}'.")
    db.session.commit()
    return jsonify({"ok": True, "template": serialize_template(template)})


@app.delete("/api/templates/<int:template_id>")
def delete_template(template_id):
    template = Template.query.get_or_404(template_id)
    template_name = template.name
    db.session.delete(template)

    state = get_state_map()
    if state.get("last_template_id") == template_id:
        upsert_state("last_template_id", None)

    log_event("template_deleted", f"Удалён шаблон '{template_name}'.")
    db.session.commit()
    return jsonify({"ok": True})


@app.get("/api/state")
def get_state():
    return jsonify(get_state_map())


@app.put("/api/state")
def save_state():
    payload = request.get_json(silent=True) or {}
    for key in DEFAULT_STATE:
        if key in payload:
            upsert_state(key, payload[key])

    db.session.commit()
    return jsonify({"ok": True, "state": get_state_map()})


@app.get("/api/logs")
def get_logs():
    return jsonify(build_logs_payload())


@app.delete("/api/logs/history")
def clear_logs_history():
    deleted_count = EventLog.query.filter(EventLog.created_at < APP_SESSION_STARTED_AT).delete(
        synchronize_session=False
    )
    log_event("journal_history_cleared", f"История журнала очищена: удалено {deleted_count} записей.")
    db.session.commit()
    return jsonify({"ok": True, "deleted_count": deleted_count, **build_logs_payload()})


@app.get("/api/instructions")
def get_instructions():
    settings = get_settings_map()
    file_path = (settings.get("instructions_file") or "").strip()
    if not file_path:
        return jsonify({"ok": False, "error": "Файл инструкций не задан в настройках."}), 400

    try:
        payload = read_instructions_file(file_path)
    except ValueError as exc:
        log_event("instructions_failed", str(exc))
        db.session.commit()
        return jsonify({"ok": False, "error": str(exc)}), 400

    return jsonify({"ok": True, **payload})


@app.get("/api/data/current")
def get_current_data():
    return jsonify(serialize_dataset())


@app.post("/api/data/upload")
def upload_data():
    global LAST_PARSE_SNAPSHOT

    file_storage = request.files.get("file")
    template_id = request.form.get("template_id", type=int)
    template = Template.query.get(template_id) if template_id else None

    try:
        dataset = build_dataset(file_storage, template)
    except ValueError as exc:
        log_event("upload_failed", str(exc))
        db.session.commit()
        return jsonify({"ok": False, "error": str(exc)}), 400

    CURRENT_DATASET.clear()
    CURRENT_DATASET.update(dataset)
    LAST_PARSE_SNAPSHOT = None

    error_count = len(dataset["errors"])
    warning_count = len(dataset["warnings"])
    if dataset["partial"]:
        event_type = "upload_partial"
        description = (
            f"Файл '{dataset['file_name']}' загружен частично: "
            f"{dataset['row_count']} строк, {warning_count} предупреждений, {error_count} ошибок."
        )
    else:
        event_type = "upload_success"
        description = (
            f"Файл '{dataset['file_name']}' загружен: "
            f"{dataset['row_count']} строк, {warning_count} предупреждений, {error_count} ошибок."
        )

    log_event(event_type, description)
    log_dataset_quality(dataset, f"Загрузка файла '{dataset['file_name']}'")
    db.session.commit()

    return jsonify({"ok": True, "dataset": dataset})


@app.post("/api/data/parse")
def parse_data():
    global LAST_PARSE_SNAPSHOT

    if not CURRENT_DATASET["rows"]:
        return jsonify({"ok": False, "error": "Нет данных для парсинга."}), 400

    payload = request.get_json(silent=True) or {}
    column_name = (payload.get("column_name") or "").strip()
    delimiter = str(payload.get("delimiter") or "")
    row_ids = payload.get("row_ids") or []

    if not column_name:
        return jsonify({"ok": False, "error": "Выберите столбец для парсинга."}), 400
    if not delimiter:
        return jsonify({"ok": False, "error": "Укажите разделитель парсинга."}), 400
    if column_name not in {column["name"] for column in CURRENT_DATASET["columns"]}:
        return jsonify({"ok": False, "error": "Выбранный столбец не найден в рабочем наборе."}), 400

    try:
        affected_row_ids = {int(item) for item in row_ids}
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "Передан некорректный набор строк для парсинга."}), 400

    if not affected_row_ids:
        return jsonify({"ok": False, "error": "Нет строк после фильтрации для парсинга."}), 400

    LAST_PARSE_SNAPSHOT = deepcopy(CURRENT_DATASET)

    deleted_rows = 0
    created_rows = 0
    next_row_id = CURRENT_DATASET.get("next_row_id", 1)
    updated_rows = []

    for row in CURRENT_DATASET["rows"]:
        if row.get("row_id") not in affected_row_ids:
            updated_rows.append(row)
            continue

        deleted_rows += 1
        fragments = split_parse_fragments(row["values"].get(column_name, ""), delimiter)
        if not fragments:
            continue

        for fragment in fragments:
            new_values = dict(row["values"])
            new_values[column_name] = fragment
            updated_rows.append(
                {
                    "row_id": next_row_id,
                    "row_number": 0,
                    "source_row_number": row["source_row_number"],
                    "values": new_values,
                }
            )
            next_row_id += 1
            created_rows += 1

    rebuild_row_numbers(updated_rows)
    CURRENT_DATASET["rows"] = updated_rows
    CURRENT_DATASET["row_count"] = len(updated_rows)
    CURRENT_DATASET["errors"] = build_dataset_errors(updated_rows, CURRENT_DATASET["columns"])
    CURRENT_DATASET["next_row_id"] = next_row_id
    CURRENT_DATASET["can_undo_parse"] = True

    log_event(
        "parse_applied",
        f"Парсинг по столбцу '{column_name}' выполнен: удалено {deleted_rows}, создано {created_rows}.",
    )
    log_dataset_quality(serialize_dataset(), f"Парсинг по столбцу '{column_name}'")
    db.session.commit()

    return jsonify(
        {
            "ok": True,
            "dataset": serialize_dataset(),
            "summary": {
                "deleted_rows": deleted_rows,
                "created_rows": created_rows,
            },
        }
    )


@app.post("/api/data/parse/undo")
def undo_parse():
    global LAST_PARSE_SNAPSHOT

    if LAST_PARSE_SNAPSHOT is None:
        return jsonify({"ok": False, "error": "Нет доступного шага отката парсинга."}), 400

    restored_dataset = deepcopy(LAST_PARSE_SNAPSHOT)
    restored_dataset["can_undo_parse"] = False

    CURRENT_DATASET.clear()
    CURRENT_DATASET.update(restored_dataset)
    LAST_PARSE_SNAPSHOT = None

    log_event("parse_undo", "Выполнен откат последнего парсинга.")
    db.session.commit()

    return jsonify({"ok": True, "dataset": serialize_dataset()})


@app.post("/api/data/export")
def export_data():
    if not CURRENT_DATASET["columns"]:
        return jsonify({"ok": False, "error": "Нет данных для сохранения."}), 400

    payload = request.get_json(silent=True) or {}
    raw_columns = payload.get("columns") or []
    raw_rows = payload.get("rows") or []

    if not isinstance(raw_columns, list) or not isinstance(raw_rows, list):
        return jsonify({"ok": False, "error": "Переданы некорректные данные для сохранения."}), 400

    columns = []
    seen_columns = set()
    for item in raw_columns:
        name = (item.get("name") or "").strip()
        if not name:
            return jsonify({"ok": False, "error": "Список столбцов для сохранения содержит пустое имя."}), 400
        if name.lower() in seen_columns:
            return jsonify({"ok": False, "error": f"Столбец '{name}' передан для сохранения несколько раз."}), 400
        seen_columns.add(name.lower())
        columns.append(name)

    if not columns:
        return jsonify({"ok": False, "error": "Нет видимых столбцов для сохранения."}), 400

    settings = get_settings_map()
    output_folder = (settings.get("output_folder") or "").strip()
    if not output_folder:
        message = "Не задана папка результата в настройках."
        log_event("export_failed", message)
        db.session.commit()
        return jsonify({"ok": False, "error": message}), 400

    output_dir = Path(output_folder)
    if not output_dir.exists() or not output_dir.is_dir():
        message = "Папка результата недоступна. Проверьте настройки приложения."
        log_event("export_failed", message)
        db.session.commit()
        return jsonify({"ok": False, "error": message}), 400

    target_path = build_export_path(output_dir)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "data"
    worksheet.append(columns)

    try:
        for row in raw_rows:
            worksheet.append([str((row or {}).get(column_name, "")) for column_name in columns])
        workbook.save(target_path)
    except Exception as exc:
        message = f"Не удалось сохранить XLSX: {exc}"
        log_event("export_failed", message)
        db.session.commit()
        return jsonify({"ok": False, "error": message}), 500

    log_event(
        "export_success",
        f"Результат сохранён в файл '{target_path.name}': {len(raw_rows)} строк, {len(columns)} столбцов.",
    )
    db.session.commit()

    return jsonify(
        {
            "ok": True,
            "file_name": target_path.name,
            "file_path": str(target_path),
            "row_count": len(raw_rows),
            "column_count": len(columns),
        }
    )


with app.app_context():
    db.create_all()


if __name__ == "__main__":
    app.run(debug=True)