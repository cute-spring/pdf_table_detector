#!/bin/bash
# This script creates a basic Python project structure for a PDF table detector.

# Define the project root folder name
PROJECT_NAME="pdf_table_detector"

# Create directories
mkdir -p "${PROJECT_NAME}/pdf_table_detector"   # Main package directory
mkdir -p "${PROJECT_NAME}/tests"                # Test directory
mkdir -p "${PROJECT_NAME}/data"                 # Data directory

# Create Python package files
touch "${PROJECT_NAME}/pdf_table_detector/__init__.py"
touch "${PROJECT_NAME}/pdf_table_detector/core.py"
touch "${PROJECT_NAME}/pdf_table_detector/pdf_parser.py"
touch "${PROJECT_NAME}/pdf_table_detector/table_detector.py"
touch "${PROJECT_NAME}/pdf_table_detector/utils.py"

# Create test file
touch "${PROJECT_NAME}/tests/test_core.py"

# Create additional project files
touch "${PROJECT_NAME}/.gitignore"
touch "${PROJECT_NAME}/requirements.txt"
touch "${PROJECT_NAME}/README.md"

# Create a placeholder for a sample PDF file (this will be an empty file)
touch "${PROJECT_NAME}/data/sample.pdf"

echo "Project structure for '${PROJECT_NAME}' has been created successfully."