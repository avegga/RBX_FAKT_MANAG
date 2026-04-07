# Project Structure

## UI
- Frontend: HTML, CSS, JavaScript
- Navigation: Main tabs (Analysis, Instructions, Journal, Settings, Reserve)

## API
- Framework: Flask
- Endpoints:
  - `/api/templates` (CRUD operations for templates)
  - `/api/settings` (Manage application settings)
  - `/api/logs` (Access event logs)

## Storage
- Database: SQLite
- Tables:
  - `templates`
  - `template_columns`
  - `app_settings`
  - `app_state`
  - `event_log`

## Processing
- Data validation
- CRUD operations
- State management