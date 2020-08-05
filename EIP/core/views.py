from django.shortcuts import render, redirect
from .forms import *
from .models import *
from django.shortcuts import get_object_or_404
from .Pdf_invoice import convertPDF2xls
# Create your views here.

def index(request):
	return render(request, 'core/index.html')


def pdf(request):
	if request.method == 'POST':
		form = UploadFileForm(request.POST, request.FILES)
		if form.is_valid():
			uploaded=form.save(commit=False)
			uploaded.save()

			return redirect('convertpdf')
		else:
			return redirect('pdf')

	else:
		form = UploadFileForm()
	return render(request, 'core/pdffile.html', {'form': form})


def convertpdf(request):
	convertPDF2xls("core/Sample_Invoice_2.pdf")