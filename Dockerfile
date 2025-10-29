FROM python:3.9-bookworm 

WORKDIR /usr/src/app
COPY . .
RUN apt-get update -y && apt-get install -y build-essential
RUN python -m pip install --no-cache-dir -r requirements.txt
RUN chmod -R 775 downloads ternary_plots uploads
EXPOSE 8000
CMD ["gunicorn", "--workers=2", "-b 0.0.0.0:8000", "app:app"]
