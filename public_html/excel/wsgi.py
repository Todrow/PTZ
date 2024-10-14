"""
WSGI config for excel project.

It exposes the WSGI callable as a module-level variable named ``application``.

For more information on this file, see
https://docs.djangoproject.com/en/5.1/howto/deployment/wsgi/
"""

import os
import sys
from django.core.wsgi import get_wsgi_application

sys.path.append('/home/pinchukna/projects/PTZ/PTZ/Excel')
sys.path.append('/home/pinchikna/projects/PTZ/PTZ/Excel/excel')
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'excel.settings')

application = get_wsgi_application()
