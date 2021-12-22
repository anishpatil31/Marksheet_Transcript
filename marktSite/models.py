from django.db import models

# Create your models here.


class Marktsheet_data(models.Model):
    positive = models.DecimalField(max_digits=8, decimal_places=2, blank=True, null=True)
    negative = models.DecimalField(max_digits=8, decimal_places=2, blank=True, null=True)
    master_roll = models.FileField(upload_to='file/', blank=True, null=True)
    response_csv = models.FileField(upload_to='file/', blank=True, null=True)
    # response = models.FileField(upload_to='static/file/', blank=True, null=True)
