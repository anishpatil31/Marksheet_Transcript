from django.db import models

# Create your models here.
class RangeInput(models.Model) :
    left = models.CharField(max_length=200, null=True)
    right = models.CharField(max_length=200, null=True)
    stamp = models.ImageField(upload_to='image/', blank=True, null=True)
    sign = models.ImageField(upload_to='image/', blank=True, null=True)