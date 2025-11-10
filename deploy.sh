#!/bin/bash

# Build and start the containers
echo "Building and starting MERQ Timesheet containers..."
docker-compose down
docker-compose build --no-cache
docker-compose up -d

echo "Waiting for services to start..."
sleep 30

# Check if services are healthy
echo "Checking service health..."
curl -f http://localhost:5000/health

if [ $? -eq 0 ]; then
    echo "✅ MERQ Timesheet is running successfully!"
    echo "🌐 Web application: http://localhost:5000"
    echo "📧 SMTP service: Ready"
else
    echo "❌ Service health check failed"
    docker-compose logs merq-timesheet-web
fi