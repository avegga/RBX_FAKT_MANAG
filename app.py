import json
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory
from flask_sqlalchemy import SQLAlchemy
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
    return {item.key: item.value for item in AppSetting.query.order_by(AppSetting.key).all()}


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


def validate_settings(payload):
    errors = {}
    folder_fields = {
        "source_folder": "Папка источника",
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
        return jsonify({"ok": False, "errors": errors}), 400

    for key in ("source_folder", "output_folder", "instructions_file"):
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
    logs = EventLog.query.order_by(EventLog.created_at.desc(), EventLog.id.desc()).limit(50).all()
    return jsonify(
        [
            {
                "id": item.id,
                "event_type": item.event_type,
                "event_description": item.event_description,
                "created_at": item.created_at.isoformat() if item.created_at else None,
            }
            for item in logs
        ]
    )


with app.app_context():
    db.create_all()


if __name__ == "__main__":
    app.run(debug=True)