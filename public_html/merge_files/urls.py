
from django.urls import path
from . import views

"""URL маршрутизатор, указывает на какой адрес перенаправлять пользователя,
при попадании на определенный url адрес
"""

urlpatterns = [
    path('', views.index, name='index'),
    path('format', views.index, name='index_2'),
    path('download/<id>', views.download_file, name='download')
]