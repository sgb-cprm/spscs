#!/bin/bash

source venv/bin/activate
gunicorn --workers=${NUM_WORKERS} -b 0.0.0.0:8000 app:app
