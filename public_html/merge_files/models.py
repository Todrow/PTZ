from django.db import models
from django.contrib.auth.models import User

"""Models, отвечает за хранение файлов и работу с БД
"""


def user_directory_path(user, filename):
    return 'uploads/user_{0}/{1}'.format(user, filename)


# Модель с информацией узлов ПЭ


class moduleSU(models.Model):
    title = models.TextField(primary_key=True)
    status = models.BooleanField()
    bureau = models.ManyToManyField(to="Bureau")

# Модель бюро


class Bureau(models.Model):
    title = models.CharField(max_length=100, primary_key=True)
    modules = models.ManyToManyField(to=moduleSU)
