# -- encoding: UTF-8 --
from django.db import models


class TestModel(models.Model):
    value = models.IntegerField(unique=True)
