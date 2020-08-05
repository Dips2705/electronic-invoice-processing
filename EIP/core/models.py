from django.db import models

# Create your models here.
class FileUpload(models.Model):
	file_name = models.CharField(max_length=250)
	input_file = models.FileField()

	def __str__(self):
		return self.file_name