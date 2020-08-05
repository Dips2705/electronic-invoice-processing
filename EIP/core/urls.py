from django.urls import path
from . import views

urlpatterns = [
	path('', views.index, name='index'),
	path('pdffile', views.pdf, name ='pdf'),
	path('convertpdf', views.convertpdf, name ='convertpdf'),
]