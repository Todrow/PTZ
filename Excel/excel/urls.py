"""
URL configuration for excel project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.1/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path, include
from django.views.generic import RedirectView

urlpatterns = [
    path('admin/', admin.site.urls),
    path('merge_files/', include('merge_files.urls')),
    path('format/', include('format_file.urls')),
    path('doc/', include('doc.urls')),
    path('', RedirectView.as_view(url='/merge_files/', permanent=True)),
    path('merge_files/format/', RedirectView.as_view(url='/format/', permanent=True)),
    path('doc/format/', RedirectView.as_view(url='/format/', permanent=True)),
    path('merge_files/doc/', RedirectView.as_view(url='/doc/', permanent=True)),
    path('format/merge_files/', RedirectView.as_view(url='/merge_files/', permanent=True)),
    path('doc/merge_files/', RedirectView.as_view(url='/merge_files/', permanent=True)),
    path('format/doc/', RedirectView.as_view(url='/doc/', permanent=True)),

]
# Используйте static() чтобы добавить соотношения для статических файлов
# Только на период разработки
from django.conf import settings
from django.conf.urls.static import static

urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)