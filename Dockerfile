# Use official Python runtime as base image
FROM python:3.11-slim

# Set working directory in container
WORKDIR /app

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV FLASK_APP=webapp.py
ENV FLASK_ENV=production

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first to leverage Docker cache
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create necessary directories
RUN mkdir -p app/static/images app/static/css app/static/js app/templates
RUN mkdir -p server/logs
RUN mkdir -p app/flask_session

# Copy template files to all locations
COPY MERQ_TIMESHEET_ETH-CAL_TEMPLATE.xlsx ./
COPY MERQ_TIMESHEET_ETH-CAL_TEMPLATE.xlsx ./app/
COPY MERQ_TIMESHEET_ETH-CAL_TEMPLATE.xlsx ./src/
COPY MERQ_TIMESHEET_ETH-CAL_TEMPLATE.xlsx ./server/

# Copy merq.png to static directories
COPY merq.png ./app/static/images/
COPY merq.png ./src/
COPY merq.png ./server/

# Copy database file
COPY merq_timesheet_db.sqlite ./
COPY merq_timesheet_db.sqlite ./app/
COPY merq_timesheet_db.sqlite ./src/

# Expose port
EXPOSE 5000

# Create a non-root user and switch to it
RUN useradd -m -u 1000 merquser
RUN chown -R merquser:merquser /app
USER merquser

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:5000/ || exit 1

# Command to run the application
CMD ["python", "app/webapp.py"]