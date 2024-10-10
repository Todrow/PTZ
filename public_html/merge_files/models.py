from django.db import models
"""Models, отвечает за хранение файлов и работу с БД
"""

# Модель бюро


class Bureau(models.Model):
    title = models.CharField(max_length=100, primary_key=True)
    modules = models.ManyToManyField(to="ModuleSU")


# Модель с информацией узлов ПЭ

class ModuleSU(models.Model):
    title = models.TextField(primary_key=True)
    status = models.BooleanField()
    bureaus = models.ManyToManyField(to="Bureau")
    op_hours = models.IntegerField(default=0)
