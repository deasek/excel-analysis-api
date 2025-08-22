# Excel Analysis API

Django REST API for uploading Excel files and calculating column summaries.

## Quick Start

### Prerequisites

- Docker and Docker Compose installed

### Setup

1. **Clone and configure:**
```bash
git clone <your-repo>
cd excel-analysis-api
cp .env.example .env
```

2. **Start the API:**
```bash
docker compose up --build -d
```

3. **Run database migrations:**
```bash
docker compose exec web python manage.py migrate
```

**API ready at:** http://localhost:8000

## Usage

### Upload and Analyze Excel File

```bash
curl -X POST http://localhost:8000/api/analyze/ \
  -F 'file=file.xlsx' \
  -F 'columns=price' \
  -F 'columns=quantity'
```

### Response

```json
{
  "file": "file.xlsx",
  "summary": [
    {"column": "price", "sum": 1012.0, "avg": 651.55},
    {"column": "quantity", "sum": 150.0, "avg": 25.0}
  ]
}
```

## API Documentation
- **Swagger UI:** http://localhost:8000/api/docs/

## Testing

### Run all tests:
```bash
docker compose exec web python manage.py test
```

## Future changes, ideas, notes:
- May be worth to introduce background processing with Celery for more complex processing of data (exports, integrations)
- Doesn't use database (may be worth to store processing results or cache them)
- May be extended with extra processing (like counting values, extra validations etc)
- Extra work work required to support example excel file provided in task description (multi sheet, no explicit header with columns etc)
- Current setup isn't proper for production deployment (uses django app server). May be worth to add application server like gunicorn and update Dockerfile to use it (with runserver used only in docker-compose).