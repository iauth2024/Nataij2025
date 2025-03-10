"""
Django settings for nataij project.
"""

import os
from pathlib import Path
import dj_database_url  # ✅ Required for PostgreSQL on Render

BASE_DIR = Path(__file__).resolve().parent.parent

# ✅ Keep secret key hidden in production (use environment variable)
SECRET_KEY = os.getenv('SECRET_KEY', 'django-insecure-2au_4jgf7*mxddvcblo(75=b(4ob7tw@i86m2*@!ifs_3k3owi')

# ✅ Debug mode should be False in production
DEBUG = os.getenv('DEBUG', 'False') == 'True'

# ✅ Allow local and production domains
ALLOWED_HOSTS = os.getenv('ALLOWED_HOSTS', 'nataij2025.onrender.com','localhost','127.0.0.1').split(',')

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'results',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'nataij.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [BASE_DIR / "results" / "templates"],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'nataij.wsgi.application'

# ✅ PostgreSQL for Render, SQLite for local development
DATABASES = {
    'default': dj_database_url.config(
        default=f"sqlite:///{BASE_DIR / 'db.sqlite3'}", conn_max_age=600
    )
}

AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator'},
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator'},
    {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator'},
    {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator'},
]

LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'UTC'
USE_I18N = True
USE_TZ = True

# ✅ Static files settings (Render + local)
STATIC_URL = '/static/'
STATICFILES_DIRS = [BASE_DIR / "results" / "static"]
STATIC_ROOT = BASE_DIR / "staticfiles"

# ✅ Media files (for Nateeja.xlsx and uploads)
MEDIA_URL = '/media/'
MEDIA_ROOT = BASE_DIR / 'media'

# ✅ Default auto field
DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'
