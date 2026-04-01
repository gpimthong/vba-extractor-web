# Use a slim Alpine-based Python image
FROM python:3.11-alpine

# Install build dependencies for pandas/openpyxl
RUN apk add --no-cache gcc musl-dev g++ libffi-dev

WORKDIR /app

# Install python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Expose port and run
EXPOSE 5000
CMD ["python", "app.py"]