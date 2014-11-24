from django.db import models
from django.utils.translation import ugettext_lazy

class MockAuthor(models.Model):
    name = models.TextField(verbose_name=ugettext_lazy("Name"))

class MockModel(models.Model):
    title = models.TextField()
    author = models.ForeignKey(MockAuthor, null=True)
