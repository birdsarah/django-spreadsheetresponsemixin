from django.db import models

class MockAuthor(models.Model):
    name = models.TextField()

class MockModel(models.Model):
    title = models.TextField()
    author = models.ForeignKey(MockAuthor, null=True)
