from django.db import models
from django.contrib.auth.models import User

def user_directory_path(user, filename):
    return 'uploads/user_{0}/{1}'.format(user, filename)

class XL_file_format(models.Model):
    SOURCE_CHOICE = [
        ('W', 'Web-system'),
        ('B', 'Bitrix24')
    ]
    user = models.ForeignKey(to=User, null=True, on_delete=models.SET_NULL)
    title = models.CharField(max_length=100)
    source = models.CharField(max_length=1, choices=SOURCE_CHOICE, default='W')
    file = models.FileField(upload_to=user_directory_path(user, title))

    def __str__(self):
        return user_directory_path