from django import forms
from .models import *

class UploadFileForm(forms.ModelForm):
	class Meta:
		model = FileUpload
		fields = ['file_name', 'input_file']