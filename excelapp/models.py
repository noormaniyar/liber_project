from django.db import models

# Create your models here.

class TableConfig(models.Model):
    num_rows = models.IntegerField()
    num_cols = models.IntegerField()
    cell_width = models.FloatField()
    cell_height = models.FloatField()